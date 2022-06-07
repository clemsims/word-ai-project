/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // // Create an array of words.
    // var words = ["Hello", "World", "!"];
    // // Queue a command to load the selection and then create a proxy range object with the results.
    // var range = context.document.getSelection();

    // // Queue a command to load the range.
    // range.load("text");
    // // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.

    // // Queue a command to insert the specified text at the end of the range.
    // range.insertText(words[Math.floor(Math.random() * words.length)], "End");

    // // Queue a command to wrap the range in a rich text content control.
    // range.insertContentControl();

    // // Queue a command to load the id property of the content control.
    // range.contentControls.items[0].load("id");

    // // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    // // await context.sync();

    // // Queue a command to change the font color of the content control to red.
    // range.contentControls.items[0].font.color = "red";

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
    await context.sync();
  });
}
