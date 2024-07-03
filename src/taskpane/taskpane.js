/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    Office.context.mailbox.addHandlerAsync(
      Office.EventType.SelectedItemsChanged,
      SelectedItemsChangedHandler,
      asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.warn("Failed to add mutti item selection handler: " + asyncResult.error.message);
              return;
          }
          console.log("Multi Mail Event handler added.");
      }
    );
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}

export function SelectedItemsChangedHandler() {
  console.log("Outer SelectedItemsChangedHandler");
  Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
      console.log("Async SelectedItemsChangedHandler", asyncResult);
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.warn("Failed to handle multi select message: " + asyncResult.error.message);
          return;
      }
  });
}