import { Record, List, Map } from "immutable"
import * as React from "react"
import * as Fabric from "office-ui-fabric-react"
import HeroList from "./HeroList"
import Progress from "./Progress"
import Styler from "../../core/styler"
import ReadmeMarkdown from "raw-loader!../../README.md"
import DEBUG from "../../core/debug"

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
  enableStyler: boolean
  inlineStyle: boolean
  enablePrettier: boolean
  proseWrap: boolean
}
const SettingsRecord = Record<SettingsType>({
  enableStyler: true,
  inlineStyle: true,
  enablePrettier: true,
  proseWrap: false,
})
export interface AppStateType {
  settings: Record<SettingsType>
  profile: Record<ProfileType>
  messages: List<Record<MessageType>>
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

const CongratsText = `<!-- Congratulations! Your team's life becomes much easier! Now, click on "Remark Document" to continue reading. -->`

// Component

export default class App extends React.Component<AppPropsType, AppStateType> {
  GetStartedSpan: HTMLSpanElement

  constructor(props, context) {
    super(props, context)
    this.state = {
      settings: SettingsRecord(),
      profile: ProfileRecord(),
      messages: List<Record<MessageType>>(),
    }
    console.info("Initial props: ", this.props, "Initial state: ", this.state)
  }

  componentDidMount() {
    Office.context.document.settings.refreshAsync(result => {
      try {
        console.debug("Loaded addin settings: ", result)
        const settings = Office.context.document.settings.get("settings")
        const profile = Office.context.document.settings.get("profile")
        console.info("Loaded Settings: ", settings, "Profile: ", profile)
        profile.showMessages = Map(profile.showMessages)
        this.setState(
          {
            settings: SettingsRecord(settings),
            profile: ProfileRecord(profile),
          },
          () => {
            console.debug("Loaded state: ", this.state)
          }
        )
      } catch (error) {
        console.error(error)
      }
    })

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
            id: "DesktopSupport",
            content:
              "Known issue: Word Desktop is not supported for now as it's using IE as internal javascript engine",
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
            id: "DocumentSettings",
            content: "Add-in settings are stored per documend",
            type: Fabric.MessageBarType.info,
            display: true,
            isDismissable: true,
          }),
          MessageRecord({
            id: "FindReadme",
            content: "Above messages can be found in Readme",
            type: Fabric.MessageBarType.success,
            display: true,
            isDismissable: true,
          }),
        ]),
      }
    })
  }

  clickDevButton = async () => {
    console.log("State: ", this.state)
    const devMarkdown = `
It's a **strong** word
`
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
    await Styler.ProcessSelection()
  }

  remarkDocument = async () => {
    console.debug("Remarking entire document...")
    await Styler.ProcessDocument()
  }

  saveSettings = () => {
    console.debug("Saving settings...")
    const settings = this.state.settings.toJS()
    const profile = this.state.profile.toJS()
    console.debug("Settings: ", settings, "Profile: ", profile)
    Office.context.document.settings.set("settings", settings)
    Office.context.document.settings.set("profile", profile)
    Office.context.document.settings.saveAsync(result => {
      console.info("Settings saved: ", result)
    })
  }

  resetSettings = () => {
    console.debug("Reseting settings...")
    this.setState(
      {
        settings: SettingsRecord(),
        profile: ProfileRecord(),
      },
      this.saveSettings
    )
  }

  dismissGetStarted = async (insert?: boolean) => {
    console.debug("Dismissing GetStarted ...")
    this.setState(
      state => ({
        profile: state.profile.set("showGetStarted", false),
      }),
      this.saveSettings
    )
    if (insert) {
      await this.insertReadme()
    }
  }

  dismissMessage = async messageId => {
    console.debug("Dismissing message: ", messageId)
    this.setState(
      state => ({
        profile: state.profile.set(
          "showMessages",
          state.profile.get("showMessages").set(messageId, false)
        ),
      }),
      this.saveSettings
    )
  }

  mergeSettings = async (partialSettings: Partial<SettingsType>) => {
    console.info("Merging settings: ", partialSettings)
    this.setState(
      { settings: this.state.settings.merge(partialSettings) },
      this.saveSettings
    )
  }

  render() {
    if (!this.props.isOfficeInitialized) {
      return (
        <Progress
          title={this.props.title}
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
            {DEBUG && (
              <ElementContainer>
                <Fabric.DefaultButton
                  style={ButtonStyle}
                  iconProps={{ iconName: "ChevronRight" }}
                  onClick={this.clickDevButton}
                >
                  Dev Button
                </Fabric.DefaultButton>
              </ElementContainer>
            )}
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
                checked={this.state.settings.get("enableStyler")}
                label="Enable Styler"
                inlineLabel={true}
                onChange={(_, checked) => {
                  this.mergeSettings({ enableStyler: checked })
                }}
              />
            </ElementContainer>
            <ElementContainer>
              <Fabric.Toggle
                checked={this.state.settings.get("inlineStyle")}
                label="Styler: Inline Style"
                inlineLabel={true}
                onChange={(_, checked) => {
                  this.mergeSettings({ inlineStyle: checked })
                }}
              />
            </ElementContainer>
            <ElementContainer>
              <Fabric.Toggle
                checked={this.state.settings.get("enablePrettier")}
                label="Enable Prettier"
                inlineLabel={true}
                onChange={(_, checked) => {
                  this.mergeSettings({ enablePrettier: checked })
                }}
              />
            </ElementContainer>
            <ElementContainer>
              <Fabric.Toggle
                checked={this.state.settings.get("proseWrap")}
                label="Prettier: Prose Wrap"
                inlineLabel={true}
                onChange={(_, checked) => {
                  this.mergeSettings({ proseWrap: checked })
                }}
              />
            </ElementContainer>
            <ElementContainer>
              <Fabric.DefaultButton
                style={ButtonStyle}
                iconProps={{ iconName: "ChevronDownEnd6" }}
                onClick={this.resetSettings}
              >
                Reset Settings
              </Fabric.DefaultButton>
            </ElementContainer>
          </div>
        </main>
      </div>
    )
  }
}
