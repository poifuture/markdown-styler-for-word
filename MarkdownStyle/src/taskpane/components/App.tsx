import * as React from "react"
import Unified, * as UnifiedModule from "unified"
import * as Unist from "unist"
import RemarkParse from "remark-parse"
import RemarkStringify from "remark-stringify"
import { Button, ButtonType } from "office-ui-fabric-react"
import HeroList, { HeroListItem } from "./HeroList"
import Progress from "./Progress"
import Prettier from "prettier/standalone"
import PrettierMarkdown from "prettier/parser-markdown"
import { hex, sleep } from "../utils"

export interface AppProps {
  title: string
  isOfficeInitialized: boolean
}

export interface AppState {
  listItems: HeroListItem[]
}

const ReadmeMarkdown = `---
title: Markdown Style for MS Word
---

Make MS Word a markdown friendly collaborative editor.

This word add-in aims to apply MS Word styles to your document without changing your markdown content. You can easily view your document with a better style while collaborating with others on the document. We are using it for writing our meeting notes.

# Usage

1. Insert Readme and read the warning
1. (Optional) Setup the document theme
1. Click "Remark Document"
1. (Optional) Customize the builtin styles (Normal, Heading1, etc.) of the document theme in MS Word

# Warning

We aim to apply only styles to your document without changing your content. However, your work might be lost if there are bugs in the add-in. If it happens, please remember to use the document history feature of MS Word to retrieve your work.

# Why MS Word (Online)

* Chinese friends cannot access Google Doc easily
* Good integration with MS products family
* Free! (For personal use) (From developer: We paid Office 365)
* ~~Rich functionality~~ Buggy

# What it does

1. Clear all pre-existing styles
1. Format your document with [Prettifier](https://github.com/prettier/prettier)
   1. Prettifier will format your markdown
   1. Prettifier will format your front matter
   1. Prettifier will format your code block
1. [Pending] Parse your markdown styles with [Remark](https://github.com/remarkjs/remark)
1. [Pending] Apply syntax highlights to your code block with [Highlightjs](https://github.com/highlightjs/highlight.js/)

# What setup does

* [Pending] Change the theme font of your document
  - Face: Courier New (A monospace font)
  - Size: 10 (To make each line contains >=80 chars)

# Examples

## Long Paragraph

A long paragraph will be rewrapped at column 80 if the the prettier.proseWrap is configured as always.
If there is no empty line, the two lines will be merged in markdown.
So please always remember to insert an empty line between your paragraphs.

## Headings

## A **Strong** Title

There is a **strong** word and **some phrases** in a sentence.

## List

## Table

Column 1 | Column 2 has a long head | c3 | c4
--- | --- | --- | ---
c1 | c2 | Column 3 is long | c4

## Code

\`\`\`javascript
const a=1
\`\`\`

# Known Issues

## Whitespcaces

As every web UI developer knows, a normal space (0x20) is different from a display space (0xA0, also known as &nbsp;). As a workaround, this Add-in will replace all nbsp to space before processing, and put nbsp back in document. It works fine for most cases, however, in rare scenarios, you will get nbsp in your clipboard. So becareful.

# Inline Style

Sometimes the inline style suddenly apply to the entire paragraph, this is a [bug](https://github.com/OfficeDev/office-js/issues/586) in Word Online. The workaround is not to remark the end of file.



## MS Word doesn't have a vim plugin

`

const devMarkdown = `
It's a **strong** word
`

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

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context)
    this.state = {
      listItems: [],
    }
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    })
  }

  click = async () => {
    return Word.run(async context => {
      const paragraph = context.document.body.insertText(
        devMarkdown,
        Word.InsertLocation.replace
      )
      paragraph.font.color = "blue"
      await context.sync()
    })
  }

  insertReadme = async () => {
    console.debug("Inserting readme...")
    return Word.run(async context => {
      context.document.body.insertText(
        ReadmeMarkdown,
        Word.InsertLocation.start
      )
      await context.sync()
    })
  }

  setupTheme = async () => {
    console.debug("Setting up document theme...")
    console.error("Not Implemented")
  }

  remarkSelection = async () => {
    console.debug("Remarking selection...")
    return this.remarkRange(false)
  }

  remarkDocument = async () => {
    console.debug("Remarking entire document...")
    return this.remarkRange(true)
  }

  remarkRange = async (entire: boolean) => {
    try {
      await Word.run(async context => {
        try {
          await context.sync()
          console.debug("Getting remark range...")
          const userRange = entire
            ? context.document.body.getRange()
            : expandToParagraph(context.document.getSelection())

          const remarkRange = await excludeEOF(
            // It **may** reduce the chance to hit the inline style bug
            userRange,
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
        } catch (error) {
          console.error(error)
        }
      })
    } catch (error) {
      console.error(error)
    }
  }

  render() {
    const { title, isOfficeInitialized } = this.props

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo="assets/logo-filled.png"
          message="Please sideload your addin to see app body."
        />
      )
    }

    return (
      <div className="ms-welcome">
        <HeroList
          message="Discover what Office Add-ins can do for you today!"
          items={this.state.listItems}
        >
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.insertReadme}
          >
            Insert Readme.md
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.setupTheme}
          >
            Setup Document Theme
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.remarkSelection}
          >
            Remark Selection
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.remarkDocument}
          >
            Remark Document
          </Button>
        </HeroList>
      </div>
    )
  }
}
