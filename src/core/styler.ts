import Unified, * as UnifiedModule from "unified"
import * as Unist from "unist"
import RemarkFrontmatter from "remark-frontmatter"
import RemarkParse from "remark-parse"
import Prettier from "prettier/standalone"
import PrettierMarkdown from "prettier/parser-markdown"
import { hex } from "./utils"

const getCleanText = (str: string): string =>
  str
    .replace(/(?:\r\n|\r|\n)/g, "\n") // crlf
    .replace(/\xA0/g, " ") // &nbsp;

const getDisplayText = (str: string): string =>
  str.replace(/[ ]{2,}/g, (match: String) => "\xA0".repeat(match.length)) // &nbsp;

const getPointParagraph = async (
  range: Word.Range,
  point: Unist.Point
): Promise<Word.Paragraph> => {
  console.debug("Geting point paragraph... ", point)
  range.paragraphs.load("items")
  await range.paragraphs.context.sync()
  console.debug("Got paragraph: ", range.paragraphs.items[point.line - 1])
  return range.paragraphs.items[point.line - 1]
}

const getPointCursur = async (
  range: Word.Range,
  point: Unist.Point,
  options: { isEnd: boolean }
): Promise<Word.Range> => {
  console.debug("Getting point cursor... ", point)
  const pointParagraph = await getPointParagraph(range, point)
  const charRanges = pointParagraph.split([""])
  charRanges.load("items")
  await charRanges.context.sync()
  console.debug("Got paragraph char cursors: ", charRanges)
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
  const startCharRanges = startParagraph.split([""])
  startCharRanges.load("items")
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
          await nodeRange.context.sync()
          const titleParagraph = nodeRange
            .search("title:")
            .getFirst()
            .paragraphs.getFirst()
          titleParagraph.styleBuiltIn = Word.Style.title
          await titleParagraph.context.sync()
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

const ProcessStyler = async (range: Word.Range) => {
  await Word.run(async () => {
    console.debug("Parsing markdown document...")
    range.load()
    await range.context.sync()
    const remarkText = getCleanText(range.text)
    const remarkPromise = new Promise((resolve, reject) => {
      Unified()
        .use(RemarkParse)
        .use(RemarkFrontmatter, ["yaml", "toml"])
        .use(RemarkWord, { range: range })
        .process(remarkText, (error, _) => {
          if (error) {
            reject(error)
          }
          resolve("")
        })
    })
    await remarkPromise
    console.info("Walking through done.")
    await range.context.sync()
  })
  console.info("Styler process done.")
}
const ProcessPrettier = async (range: Word.Range) => {
  await Word.run(async () => {
    console.debug("Fetching document content...")
    range.load()
    await range.context.sync()
    const originalText = getCleanText(range.text)
    if (originalText == "") {
      console.error("No text is selected")
      return
    }
    console.info("Original Text: ", originalText, await hex(originalText))

    console.debug("Prettifying markdown document...")
    const prettyText: string = Prettier.format(originalText, {
      parser: "markdown",
      plugins: [PrettierMarkdown],
      proseWrap: "never", // [always,never,preserve]
    })
    console.info("Pretty Text: ", prettyText, await hex(prettyText))

    let fixedText = prettyText
    if (originalText.startsWith("\n") && !fixedText.startsWith("\n")) {
      fixedText = "\n" + fixedText
    }
    if (!originalText.endsWith("\n") && fixedText.endsWith("\n")) {
      fixedText = fixedText.slice(0, -1)
    }
    console.debug("Fixed Text: ", fixedText)

    console.debug("Replacing prettier document...")
    range.insertText(getDisplayText(fixedText), Word.InsertLocation.replace)
    await range.context.sync()
  })
  console.info("Prettier process done.")
}

const ProcessRange = async (range: Word.Range) => {
  await Word.run(async () => {
    console.debug("Reseting style ...")
    range.styleBuiltIn = Word.Style.normal
    await range.context.sync()
  })
  await ProcessPrettier(range)
  await ProcessStyler(range)
  console.info("All processes done.")
}
export const ProcessSelection = async () => {
  try {
    let remarkRange: Word.Range = undefined
    await Word.run(async context => {
      console.debug("Getting selection range...")
      const inputRange = expandToParagraph(context.document.getSelection())
      remarkRange = inputRange
      context.trackedObjects.add(remarkRange)
      await context.sync()
    })
    await ProcessRange(remarkRange)
  } catch (error) {
    console.error(error)
  }
}
export const ProcessDocument = async () => {
  try {
    let remarkRange: Word.Range = undefined
    await Word.run(async context => {
      console.debug("Getting document range...")
      const inputRange = context.document.body.getRange()
      remarkRange = inputRange
      context.trackedObjects.add(remarkRange)
      await context.sync()
    })
    await ProcessRange(remarkRange)
  } catch (error) {
    console.error(error)
  }
}

export default {
  ProcessSelection: ProcessSelection,
  ProcessDocument: ProcessDocument,
}
