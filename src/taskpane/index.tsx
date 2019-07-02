import "core-js/stable" // polyfill
import "regenerator-runtime/runtime" // polyfill
import { AppContainer } from "react-hot-loader"
import * as React from "react"
import * as ReactDOM from "react-dom"
// import { Bootstrap } from "./bootstrap"
const BootstrapModulePromise = import(
  /* webpackChunkName: "bootstrap" */ "./bootstrap"
)
let Bootstrap: typeof import("./bootstrap").Bootstrap = null

let isOfficeInitialized = false

const title = "Markdown Styler"

const render = (App?) => {
  ReactDOM.render(
    <AppContainer>
      {Bootstrap ? (
        <Bootstrap
          app={App}
          title={title}
          isOfficeInitialized={isOfficeInitialized}
        />
      ) : (
        <p>Splash Screen...</p>
      )}
    </AppContainer>,
    document.getElementById("container")
  )
}

/* Render application after modules initialize */
Office.initialize = () => {
  isOfficeInitialized = true
  console.info("Office initialized.")
  render()
}
BootstrapModulePromise.then(module => {
  console.info("App initialized")
  Bootstrap = module.Bootstrap
  render()
})

/* Initial render showing a progress bar */
render()

if ((module as any).hot) {
  ;(module as any).hot.accept("./components/App", () => {
    import("./components/App").then(module => {
      const NextApp = module.default
      render(NextApp)
    })
  })
}
