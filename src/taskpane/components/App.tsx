import { Record, List, Map } from "immutable"
import * as React from "react"
import * as Fabric from "office-ui-fabric-react"
import HeroList from "./HeroList"
import Progress from "./Progress"
import Styler from "../../core/styler"
import ReadmeMarkdown from "raw-loader!../../README.md"

// Types

export interface AppPropsType {
  title: string
  isOfficeInitialized: boolean
}
interface MessageType {
  id: string
  content: any
  type: Fabric.MessageBarType
  display: boolean
  isDismissable: boolean
}
const MessageRecord = Record<MessageType>({
  id: "",
  content: "",
  type: Fabric.MessageBarType.info,
  display: true,
  isDismissable: true,
})
interface ProfileType {
  showGetStarted: boolean
  showMessages: Map<string, boolean>
}
const ProfileRecord = Record<ProfileType>({
  showGetStarted: true,
  showMessages: Map<string, boolean>(),
})
interface SettingsType {
  inlineStyle: boolean
  prettier: boolean
  proseWrap: boolean
}
const SettingsRecord = Record<SettingsType>({
  inlineStyle: true,
  prettier: true,
  proseWrap: false,
})
export interface AppStateType {
  messages: List<Record<MessageType>>
  profile: Record<ProfileType>
  settings: Record<SettingsType>
}

// Styles

const ButtonStyle: React.CSSProperties = {
  width: "100%",
}
class ElementContainer extends React.Component {
  render() {
    return (
      <div
        className="ms-Grid-col ms-sm12 ms-md6 ms-lg3"
        style={{
          justifyContent: "center",
          display: "flex",
          padding: "5px 5px 0 0",
        }}
      >
        <div style={{ width: "200px" }}>{this.props.children}</div>
      </div>
    )
  }
}

// Data

const devMarkdown = `
It's a **strong** word
`
const CongratsText = `<!-- Congratulations! Your team's life becomes much easier! Now, click on "Remark Document" to continue reading. -->`

// Component

export default class App extends React.Component<AppPropsType, AppStateType> {
  GetStartedSpan: HTMLSpanElement

  constructor(props, context) {
    super(props, context)
    this.state = {
      messages: List<Record<MessageType>>(),
      settings: SettingsRecord(),
      profile: ProfileRecord(),
    }
    console.debug("jason:appconst", this.state)
  }

  componentDidMount() {
    this.setState(() => {
      return {
        messages: List<Record<MessageType>>([
          MessageRecord({
            id: "VersionHistory",
            content:
              'Use OneDrive "Version History" to get your work back in case anything is missing',
            type: Fabric.MessageBarType.error,
            display: true,
            isDismissable: true,
          }),
          MessageRecord({
            id: "InlineStyleOnline",
            content: [
              <span key="text">
                Known issue: Inline style may apply to the entire paragraph due
                to a bug in MS Word Online. See
              </span>,
              <Fabric.Link
                key="link"
                href="https://github.com/OfficeDev/office-js/issues/586"
              >
                office-js/issue#586
              </Fabric.Link>,
            ],
            type: Fabric.MessageBarType.warning,
            display: true,
            isDismissable: false,
          }),
          MessageRecord({
            id: "InlineStyleDesktop",
            content:
              "Known issue: Inline style may apply to wrong range in Word Desktop",
            type: Fabric.MessageBarType.warning,
            display: true,
            isDismissable: true,
          }),
          MessageRecord({
            id: "Whitespace",
            content:
              "Known issue: Due to MS Word implementation, the space (ascii:32) might be replaced by a nb-space (nbsp, ascii:160) by accident. Use caution when you copy your content into clipboard",
            type: Fabric.MessageBarType.warning,
            display: true,
            isDismissable: true,
          }),
          MessageRecord({
            id: "FindReadme",
            content: "Above messages can be found in Readme",
            type: Fabric.MessageBarType.info,
            display: true,
            isDismissable: true,
          }),
        ]),
      }
    })
  }

  clickDevButton = async () => {
    return Word.run(async context => {
      console.log("State: ", this.state)
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
      /INSTALL SECTION BEGIN(.|\n)*INSTALL SECTION END/gm,
      ""
    )
    FilteredReadme = FilteredReadme.replace(/<!--.*-->\n/g, "")
    try {
      return Word.run(async context => {
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

  dismissGetStarted = async (insert?: boolean) => {
    console.debug("Dismissing GetStarted ...")
    this.setState(state => ({
      profile: state.profile.set("showGetStarted", false),
    }))
    if (insert) {
      await this.insertReadme()
    }
  }

  dismissMessage = async messageId => {
    console.debug("Dismissing message: ", messageId)
    this.setState(state => {
      return {
        profile: state.profile.set(
          "showMessages",
          state.profile.get("showMessages").set(messageId, false)
        ),
      }
    })
  }

  mergeSettings = async (partialSettings: Partial<SettingsType>) => {
    console.info("Merging settings: ", partialSettings)
    this.setState({ settings: this.state.settings.merge(partialSettings) })
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
          items={[]}
        ></HeroList>
        <main className="ms-Grid" dir="ltr">
          <Fabric.Stack>
            {this.state.messages.map(
              message =>
                this.state.profile.get("showMessages").get(message.get("id")) !=
                  false && (
                  <Fabric.StackItem key={message.get("id")}>
                    <Fabric.MessageBar
                      messageBarType={message.get("type")}
                      onDismiss={
                        message.get("isDismissable")
                          ? () => this.dismissMessage(message.get("id"))
                          : undefined
                      }
                    >
                      {message.get("content")}
                    </Fabric.MessageBar>
                  </Fabric.StackItem>
                )
            )}
          </Fabric.Stack>
          <Fabric.Separator># Home</Fabric.Separator>
          <div className="ms-Grid-row">
            <ElementContainer>
              <Fabric.DefaultButton
                style={ButtonStyle}
                iconProps={{ iconName: "ChevronRight" }}
                onClick={this.clickDevButton}
              >
                Dev Button
              </Fabric.DefaultButton>
            </ElementContainer>
            <ElementContainer>
              <span ref={span => (this.GetStartedSpan = span)}>
                <Fabric.DefaultButton
                  style={ButtonStyle}
                  iconProps={{ iconName: "MarkDownLanguage" }}
                  onClick={this.insertReadme}
                >
                  Insert Readme.md
                </Fabric.DefaultButton>
              </span>
              {this.state.profile.get("showGetStarted") && (
                <Fabric.TeachingBubble
                  target={this.GetStartedSpan}
                  primaryButtonProps={{
                    children: "Go ahead",
                    onClick: () => this.dismissGetStarted(true),
                  }}
                  secondaryButtonProps={{
                    children: "Later",
                    onClick: () => this.dismissGetStarted(),
                  }}
                  headline="Get Started"
                  onDismiss={() => this.dismissGetStarted()}
                >
                  Insert Readme.md to the top of your document
                </Fabric.TeachingBubble>
              )}
            </ElementContainer>
            <ElementContainer>
              <Fabric.DefaultButton
                style={ButtonStyle}
                iconProps={{ iconName: "Color" }}
                onClick={this.setupTheme}
              >
                Setup Theme
              </Fabric.DefaultButton>
            </ElementContainer>
            <ElementContainer>
              <Fabric.DefaultButton
                style={ButtonStyle}
                iconProps={{ iconName: "ChevronDown" }}
                onClick={this.remarkSelection}
              >
                Remark Selection
              </Fabric.DefaultButton>
            </ElementContainer>
            <ElementContainer>
              <Fabric.DefaultButton
                style={ButtonStyle}
                iconProps={{ iconName: "ChevronDownEnd6" }}
                onClick={this.remarkDocument}
              >
                Remark Document
              </Fabric.DefaultButton>
            </ElementContainer>
          </div>
          <Fabric.Separator># Settings</Fabric.Separator>
          <div className="ms-Grid-row">
            <ElementContainer>
              <Fabric.Toggle
                defaultChecked={this.state.settings.get("inlineStyle")}
                label="Inline Style"
                inlineLabel={true}
                onChange={(_, checked) => {
                  this.mergeSettings({ inlineStyle: checked })
                }}
              />
            </ElementContainer>
            <ElementContainer>
              <Fabric.Toggle
                defaultChecked={this.state.settings.get("prettier")}
                label="Prettier"
                inlineLabel={true}
                onChange={(_, checked) => {
                  this.mergeSettings({ prettier: checked })
                }}
              />
            </ElementContainer>
            <ElementContainer>
              <Fabric.Toggle
                defaultChecked={this.state.settings.get("proseWrap")}
                label="Prose Wrap"
                inlineLabel={true}
                onChange={(_, checked) => {
                  this.mergeSettings({ proseWrap: checked })
                }}
              />
            </ElementContainer>
          </div>
        </main>
      </div>
    )
  }
}
