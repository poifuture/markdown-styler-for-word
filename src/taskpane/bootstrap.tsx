import * as React from "react"
import "office-ui-fabric-react/dist/css/fabric.min.css"
import { initializeIcons } from "office-ui-fabric-react/lib/Icons"
import { Customizer } from "office-ui-fabric-react"
import { FluentCustomizations } from "@uifabric/fluent-theme"
import App, { AppProps } from "./components/App"

initializeIcons()

export interface BootstrapProps extends AppProps {
  app: typeof App
}

export class Bootstrap extends React.Component<BootstrapProps> {
  static defaultProps = {
    app: App,
  }
  render() {
    const Component = this.props.app
    return (
      <Customizer {...FluentCustomizations}>
        <Component {...this.props} />
      </Customizer>
    )
  }
}
