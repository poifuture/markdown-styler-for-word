import Unified, * as UnifiedModule from "unified"
import * as Unist from "unist"
import RemarkFrontmatter from "remark-frontmatter"
import RemarkParse from "remark-parse"
import Prettier from "prettier/standalone"
import PrettierMarkdown from "prettier/parser-markdown"
import { hex, sleep } from "./utils"

const getCleanText = str =>
  str
    .replace(/(?:\r\n|\r|\n)/g, "\n") // crlf
    .replace(/\xA0/g, " ") // &nbsp;

const getDisplayText = str =>
  str.replace(/[ ]{2,}/g, (match: String) => "\xA0".repeat(match.length)) // &nbsp;

const getPointParagraph = async (
  range: Word.Range,
  point: Unist.Point
): Promise<Word.Paragraph> => {
  console.debug("getPointParagraph", point)
  range.paragraphs.load()
  await range.paragraphs.context.sync()
  console.debug("target", range.paragraphs.items[point.line - 1])
  return range.paragraphs.items[point.line - 1]
}

const getPointCursur = async (
  range: Word.Range,
  point: Unist.Point,
  options: { isEnd: boolean }
): Promise<Word.Range> => {
  console.debug("getPointCursur", point)
  const pointParagraph = await getPointParagraph(range, point)
  pointParagraph.load()
  await pointParagraph.context.sync()
  const charRanges = pointParagraph.getTextRanges([""])
  charRanges.load()
  await charRanges.context.sync()
  console.debug("cursur", charRanges)
  const pointCursor = options.isEnd
    ? charRanges.items[point.column - 2]
    : charRanges.items[point.column - 1]
  return pointCursor
}

const getNodeRange = async (
  range: Word.Range,
  node: Unist.Node
): Promise<Word.Range> => {
  const startParagraph = await getPointParagraph(range, node.position.start)
  startParagraph.load()
  await startParagraph.context.sync()
  const startCharRanges = startParagraph.getTextRanges([""])
  startCharRanges.load()
  await startCharRanges.context.sync()
  const startCursor = startCharRanges.items[node.position.start.column - 1]
  const endCursor =
    node.position.start.line == node.position.end.line
      ? startCharRanges.items[node.position.end.column - 2]
      : await getPointCursur(range, node.position.end, { isEnd: true })
  const nodeRange = startCursor.expandTo(
    endCursor.getRange(Word.RangeLocation.after)
  )
  return nodeRange
}

const expandToParagraph = (range: Word.Range): Word.Range => {
  const startCursor = range.paragraphs
    .getFirst()
    .getRange(Word.RangeLocation.start)
  const endCursor = range.paragraphs.getLast().getRange(Word.RangeLocation.end)
  return range.expandTo(startCursor).expandTo(endCursor)
}

const excludeEOF = async (
  range: Word.Range,
  eof: Word.Range
): Promise<Word.Range> => {
  range.load()
  await range.context.sync()
  const hitEOF = range.intersectWithOrNullObject(eof)
  hitEOF.load()
  await hitEOF.context.sync()
  if (hitEOF.isNullObject) {
    console.warn("Miss EOF")
    return range
  }
  console.error("Hit EOF")
  const startCursor = range.getRange(Word.RangeLocation.start)
  const endCursor = range.paragraphs
    .getLast()
    .getRange(Word.RangeLocation.start)
  return startCursor.expandTo(endCursor)
}

interface UnistParentNode extends Unist.Node {
  children?: Unist.Node[]
}
type VisitorFunction = (node: Unist.Node) => Promise<void>
const UnistDFS = async (node: Unist.Node, visitor: VisitorFunction) => {
  await visitor(node)
  const extendedNode = node as UnistParentNode
  if (extendedNode.children) {
    for (let index = 0; index < extendedNode.children.length; index++) {
      await UnistDFS(extendedNode.children[index], visitor)
    }
  }
}

const RemarkWord: UnifiedModule.Attacher = function(options: {
  // Cant use arrow function here because the context of 'this' will be different
  range: Word.Range
}) {
  const range = options.range
  const RemarkWordTransformer: UnifiedModule.Transformer = async (
    tree,
    _
  ): Promise<Unist.Node> => {
    console.debug("Tree: ", tree)
    range.paragraphs.load()
    await range.context.sync()
    await UnistDFS(tree, async (node: Unist.Node) => {
      console.debug("Node: ", node)
      switch (node.type) {
        case "yaml": {
          const nodeRange = await getNodeRange(range, node)
          nodeRange.font.color = "darkblue"
          nodeRange.load()
          await nodeRange.context.sync()
          const titleParagraph = nodeRange
            .search("title:")
            .getFirst()
            .paragraphs.getFirst()
            .load()
          await titleParagraph.context.sync()
          titleParagraph.styleBuiltIn = Word.Style.title
          break
        }
        case "heading": {
          const nodeHeading: any = node
          const WordHeadingStyles = [
            Word.Style.title,
            Word.Style.heading1,
            Word.Style.heading2,
            Word.Style.heading3,
            Word.Style.heading4,
            Word.Style.heading5,
            Word.Style.heading6,
            Word.Style.heading7,
            Word.Style.heading8,
            Word.Style.heading9,
          ]
          const nodeParagraph = await getPointParagraph(
            range,
            node.position.start
          )
          nodeParagraph.styleBuiltIn = WordHeadingStyles[nodeHeading.depth]
          break
        }
        case "strong": {
          const nodeRange = await getNodeRange(range, node)
          nodeRange.styleBuiltIn = Word.Style.strong
          nodeRange.font.color = "darkblue"
          break
        }
      }
    })
    return tree
  }
  this.Compiler = tree => {
    console.debug("Compiler Tree:", tree)
    return ""
  }
  return RemarkWordTransformer
}

const RemarkRange = async (range: Word.Range) => {
  const context = range.context
  console.debug("Getting remark range...")

  // This workaround **may** reduce the chance to hit the inline style bug
  // https://github.com/OfficeDev/office-js/issues/586
  const remarkRange = await excludeEOF(
    range,
    context.document.body.getRange(Word.RangeLocation.end)
  )

  console.debug("Clearing original format...")
  remarkRange.styleBuiltIn = Word.Style.normal

  console.debug("Fetching document content...")
  remarkRange.load()
  await context.sync()
  const originalText = getCleanText(remarkRange.text)
  if (originalText == "") {
    console.error("No text is selected")
  }
  console.info("Original Text: ", originalText, await hex(originalText))

  console.debug("Prettifying markdown document...")
  const prettyText = Prettier.format(originalText, {
    parser: "markdown",
    plugins: [PrettierMarkdown],
    proseWrap: "never", // [always,never,preserve]
  })
  console.info("Pretty Text: ", prettyText, await hex(prettyText))

  console.debug("Replacing markdown document...")
  remarkRange.insertText(
    getDisplayText(prettyText),
    Word.InsertLocation.replace
  )
  remarkRange.load()
  await context.sync()

  console.debug("Parsing markdown document...")
  const remarkText = getCleanText(remarkRange.text)
  const remarkPromise = new Promise((resolve, reject) => {
    Unified()
      .use(RemarkParse)
      .use(RemarkFrontmatter, ["yaml", "toml"])
      .use(RemarkWord, { range: remarkRange })
      .process(remarkText, (error, _) => {
        if (error) {
          reject(error)
        }
        resolve("")
      })
  })
  await remarkPromise
  console.info("Remark done.")

  await context.sync()
  await sleep(5000)
}

export const RemarkSelection = async (context: Word.RequestContext) => {
  console.debug("Getting selection range...")
  const inputRange = expandToParagraph(context.document.getSelection())
  return RemarkRange(inputRange)
}
export const RemarkDocument = async (context: Word.RequestContext) => {
  console.debug("Getting document range...")
  const inputRange = context.document.body.getRange()
  return RemarkRange(inputRange)
}

export default {
  RemarkSelection: RemarkSelection,
  RemarkDocument: RemarkDocument,
}
