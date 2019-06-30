import * as React from "react"
import { Button, ButtonType } from "office-ui-fabric-react"
import HeroList, { HeroListItem } from "./HeroList"
import Progress from "./Progress"
import Styler from "../../core/styler"
import ReadmeMarkdown from "raw-loader!../../README.md.src"

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
const CongratsParagraph = `Congratulations! Your team's life become much easier!
Now, click on "Remark Document" to continue reading.`

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
    let FilteredReadme = ReadmeMarkdown

    FilteredReadme = FilteredReadme.replace(
      /<!-- CONGRATULATIONS -->/gms,
      CongratsParagraph
    )
    FilteredReadme = FilteredReadme.replace(/INSTALL BEGIN.*INSTALL END/gms, "")
    FilteredReadme = FilteredReadme.replace(/<!--.*-->/g, "")
    try {
      return Word.run(async context => {
        context.document.body.insertText(
          FilteredReadme,
          Word.InsertLocation.start
        )
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
