import Unified, * as UnifiedModule from "unified"
import * as Unist from "unist"
import RemarkParse from "remark-parse"
import RemarkStringify from "remark-stringify"
import Prettier from "prettier/standalone"
import PrettierMarkdown from "prettier/parser-markdown"
import { hex, sleep } from "./utils"

const getCleanText = str =>
  str
    .replace(/(?:\r\n|\r|\n)/g, "\n") // crlf
    .replace(/\xA0/g, " ") // &nbsp;

const getDisplayText = str =>
  str.replace(/[ ]{2,}/g, (match: String) => "\xA0".repeat(match.length)) // &nbsp;

const getNodeParagraph = async (
  range: Word.Range,
  node: Unist.Node
): Promise<Word.Paragraph> => {
  range.paragraphs.load()
  await range.paragraphs.context.sync()
  return range.paragraphs.items[node.position.start.line - 1]
}

const getNodeRange = async (
  range: Word.Range,
  node: Unist.Node
): Promise<Word.Range> => {
  const nodeParagraph = await getNodeParagraph(range, node)
  nodeParagraph.load()
  await nodeParagraph.context.sync()
  const charRanges = nodeParagraph.getTextRanges([""])
  charRanges.load()
  await charRanges.context.sync()
  const startCursor = charRanges.items[node.position.start.column - 1]
  const endCursor = charRanges.items[node.position.end.column - 2]
  const nodeRange = startCursor.expandTo(endCursor)
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

const RemarkWord: UnifiedModule.Attacher = (options: { range: Word.Range }) => {
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
          const nodeParagraph = await getNodeParagraph(range, node)
          nodeParagraph.styleBuiltIn = WordHeadingStyles[nodeHeading.depth]
          break
        }
        case "strong": {
          try {
            const nodeRange = await getNodeRange(range, node)
            nodeRange.font.bold = true
            nodeRange.font.color = "darkblue"
            await nodeRange.parentBody.context.sync()
          } catch (error) {
            console.error(error)
          }
          break
        }
      }
    })
    return tree
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
      .use(RemarkWord, { range: remarkRange })
      .use(RemarkStringify)
      .process(remarkText, (error, remarkAST) => {
        if (error) {
          reject(error)
        }
        resolve(remarkAST)
      })
  })
  const remarkAST = await remarkPromise
  console.info("AST: ", remarkAST)

  await context.sync()
  await sleep(1000)
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
