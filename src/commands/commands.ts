import "core-js/stable" // polyfill
import "regenerator-runtime/runtime" // polyfill
import Styler from "../core/styler"

Office.onReady(() => {})

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined
}

const g = getGlobal() as any

g.onClickRemarkSelection = (event: Office.AddinCommands.Event) => {
  console.debug("[Ribbon] Remarking selection...", event)
  Styler.ProcessSelection().finally(() => {
    event.completed()
  })
}

g.onClickRemarkDocument = (event: Office.AddinCommands.Event) => {
  console.debug("[Ribbon] Remarking document...", event)
  Styler.ProcessDocument().finally(() => {
    event.completed()
  })
}
