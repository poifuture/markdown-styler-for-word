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
  Word.run(async context => {
    await Styler.RemarkSelection(context)
    await context.sync()
  })
    .catch(error => {
      console.error(error)
    })
    .finally(() => {
      event.completed()
    })
}

g.onClickRemarkDocument = (event: Office.AddinCommands.Event) => {
  console.debug("[Ribbon] Remarking document...", event)
  Word.run(async context => {
    await Styler.RemarkDocument(context)
    await context.sync()
  })
    .catch(error => {
      console.error(error)
    })
    .finally(() => {
      event.completed()
    })
}
