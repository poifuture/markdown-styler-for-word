import * as React from "react"
import { Button, ButtonType, DefaultButton } from "office-ui-fabric-react"
import HeroList, { HeroListItem } from "./HeroList"
import Progress from "./Progress"
import Styler from "../../core/styler"
import ReadmeMarkdown from "raw-loader!../../README.md"

export interface AppProps {
  title: string
  isOfficeInitialized: boolean
}

export interface AppState {
  listItems: HeroListItem[]
}

const devMarkdown = `
It's a **strong** word
`
const CongratsText = `<!-- Congratulations! Your team's life becomes much easier! Now, click on "Remark Document" to continue reading. -->`

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context)
    this.state = {
      listItems: [],
    }
  }

  componentDidMount() {
    this.setState({
      listItems: [],
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
    let FilteredReadme = ReadmeMarkdown

    // FilteredReadme = FilteredReadme.replace(
    //   /<!-- CONGRATULATIONS PLACEHOLDER -->/gm,
    //   CongratsParagraph
    // )
    FilteredReadme = FilteredReadme.replace(
      /INSTALL SECTION BEGIN(.|\n)*INSTALL SECTION END/gm,
      ""
    )
    FilteredReadme = FilteredReadme.replace(/<!--.*-->\n/g, "")
    try {
      return Word.run(async context => {
        console.debug("start")
        context.document.body.insertText(
          FilteredReadme,
          Word.InsertLocation.start
        )
        await context.sync()
        context.document.body.paragraphs.load()
        await context.sync()
        const CongratsParagraph = context.document.body.paragraphs.items[3].insertParagraph(
          CongratsText,
          Word.InsertLocation.after
        )
        CongratsParagraph.font.color = "red"
        CongratsParagraph.font.bold = true
        await context.sync()
      })
    } catch (error) {
      console.error(error)
    }
  }

  setupTheme = async () => {
    console.debug("Setting up document theme...")
    console.error("Not Implemented")
  }

  remarkSelection = async () => {
    console.debug("Remarking selection...")
    try {
      Word.run(async context => {
        await Styler.RemarkSelection(context)
        await context.sync()
      })
    } catch (error) {
      console.error(error)
    }
  }

  remarkDocument = async () => {
    console.debug("Remarking entire document...")
    try {
      Word.run(async context => {
        await Styler.RemarkDocument(context)
        await context.sync()
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
          message="Make Word a markdown friendly collaborative editor"
        />
      )
    }

    return (
      <div className="ms-welcome">
        <HeroList
          message="Make Word a markdown friendly collaborative editor"
          items={this.state.listItems}
        >
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </DefaultButton>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "MarkDownLanguage" }}
            onClick={this.insertReadme}
          >
            Insert Readme.md
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "color" }}
            onClick={this.setupTheme}
          >
            Setup Document Theme
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronDown" }}
            onClick={this.remarkSelection}
          >
            Remark Selection
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronDownEnd6" }}
            onClick={this.remarkDocument}
          >
            Remark Document
          </Button>
        </HeroList>
      </div>
    )
  }
}
