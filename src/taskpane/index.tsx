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
      {Bootstrap && isOfficeInitialized ? (
        <Bootstrap app={App} title={title} />
      ) : (
        <div>
          <img src="assets/logo-filled.png"></img>
          <h1>Markdown Styler</h1>
          <p>Make Word a markdown friendly collaborative editor</p>
          <p>
            Loading Markdown Styler App ...
            {Bootstrap && <span>done</span>}
          </p>
          <p>
            Loading Office API ...
            {isOfficeInitialized && <span>done</span>}
          </p>
        </div>
      )}
    </AppContainer>,
    document.getElementById("container")
  )
}

/* Render application after modules initialize */
Office.initialize = () => {
  isOfficeInitialized = true
  OfficeExtension.config.extendedErrorLogging = true
  console.info("Office initialized.")
  render()
}
BootstrapModulePromise.then(module => {
  console.info("App initialized")
  Bootstrap = module.Bootstrap
  render()
})

/* Initial render showing a splash screen */
render()

if ((module as any).hot) {
  ;(module as any).hot.accept("./components/App", () => {
    import("./components/App").then(module => {
      const NextApp = module.default
      render(NextApp)
    })
  })
}
