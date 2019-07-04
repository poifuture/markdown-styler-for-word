import * as React from "react"
import "office-ui-fabric-react/dist/css/fabric.min.css"
import { initializeIcons } from "office-ui-fabric-react/lib/Icons"
import { Customizer } from "office-ui-fabric-react"
import { FluentCustomizations } from "@uifabric/fluent-theme"
import App, { AppPropsType } from "./components/App"

initializeIcons()

export interface BootstrapPropsType extends AppPropsType {
  app: typeof App
}

export class Bootstrap extends React.Component<BootstrapPropsType> {
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
