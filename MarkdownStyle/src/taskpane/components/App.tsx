import * as React from "react"
import Unified, * as UnifiedModule from "unified"
import * as Unist from "unist"
import UnistVisit from "unist-util-visit"
import RemarkParse from "remark-parse"
import RemarkStringify from "remark-stringify"
import { Button, ButtonType } from "office-ui-fabric-react"
import HeroList, { HeroListItem } from "./HeroList"
import Progress from "./Progress"
import Prettier from "prettier/standalone"
import PrettierMarkdown from "prettier/parser-markdown"

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

Refresh

## MS Word doesn't have a vim plugin

`

const devMarkdown = `
- Item1
  - Item1.1
- Item2
`

const hex = str => str.split("").map(ch => ch.charCodeAt(0))

const getCleanText = str =>
  str
    .replace(/(?:\r\n|\r|\n)/g, "\n") // crlf
    .replace(/\xA0/g, " ") // &nbsp;

const getDisplayText = str =>
  str.replace(/[ ]{2,}/g, (match: String) => "\xA0".repeat(match.length)) // &nbsp;

const RemarkWord: UnifiedModule.Attacher = (options: { range: Word.Range }) => {
  const range = options.range
  const RemarkWordTransformer: UnifiedModule.Transformer = async (
    tree,
    _
  ): Promise<Unist.Node> => {
    console.debug("Tree: ", tree)
    range.paragraphs.load()
    await range.context.sync()
    UnistVisit(tree, null, (node: Unist.Node) => {
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
          range.paragraphs.items[node.position.start.line - 1].styleBuiltIn =
            WordHeadingStyles[nodeHeading.depth]
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
    return Word.run(async context => {
      try {
        console.debug("Getting remark range...")
        const remarkRange = entire
          ? context.document.body.getRange()
          : context.document.getSelection()

        console.debug("Clearing original format...")
        remarkRange.styleBuiltIn = Word.Style.normal

        console.debug("Fetching document content...")
        remarkRange.load()
        await context.sync()
        const originalText = getCleanText(remarkRange.text)
        if (originalText == "") {
          console.error("No text is selected")
        }
        console.info("Original Text: ", originalText, hex(originalText))

        console.debug("Prettifying markdown document...")
        const prettyText = Prettier.format(originalText, {
          parser: "markdown",
          plugins: [PrettierMarkdown],
          proseWrap: "never", // [always,never,preserve]
        })
        console.info("Pretty Text: ", prettyText, hex(prettyText))

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
      } catch (error) {
        console.error(error)
      }
    })
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
