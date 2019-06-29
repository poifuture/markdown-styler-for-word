import * as React from "react"
// import Unified from "unified"
// import RemarkParse from "remark-parse"
// import RemarkStringify from "remark-stringify"
// import RemarkRecommended from "remark-preset-lint-recommended"
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

# Warning

We aim to apply only styles to your document without changing your content. However, your work might be lost if there are bugs in the add-in. If it happens, please remember to use the document history feature of MS Word to retrieve your work.

# Why MS Word (Online)

* Chinese friends cannot access Google Doc easily
* Good integration with MS products family
* Free! (For personal use)(From developer: We paid Office 365)
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

# Known Issues

## Consequent Whitespcaces

## MS Word doesn't have a vim plugin

`

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
      const paragraph = context.document.body.insertParagraph(
        "Hello World",
        Word.InsertLocation.end
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
        remarkRange.styleBuiltIn = Word.Style.noSpacing

        console.debug("Fetching document content...")
        remarkRange.load()
        await context.sync()
        const originalText = remarkRange.text
        if (originalText == "") {
          console.error("No text is selected")
        }
        console.info("Original Text: ", originalText)

        console.debug("Prettifying markdown document...")
        const prettyText = Prettier.format(originalText, {
          parser: "markdown",
          plugins: [PrettierMarkdown],
          proseWrap: "always",
        })

        // const prettifyPromise = new Promise((resolve, reject) => {
        //   Unified()
        //     .use(RemarkParse)
        //     .use(RemarkStringify)
        //     .use(RemarkRecommended)
        //     .process(originalText, (error, prettyText) => {
        //       if (error) {
        //         reject(error)
        //       }
        //       resolve(String(prettyText))
        //     })
        // })
        // const prettyText = await prettifyPromise

        console.info("Pretty Text: ", prettyText)

        remarkRange.insertText(prettyText, Word.InsertLocation.replace)

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
