import Styler from "../core/styler"
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
})

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  }

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "action",
    message
  )

  // Be sure to indicate when the add-in command function is complete
  event.completed()
}

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

// the add-in command functions need to be available in global scope
g.action = action

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
