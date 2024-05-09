/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { base64Image } from "../../base64Image";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    document.getElementById("apply-custom-style").onclick = () => tryCatch(applyCustomStyle);
    document.getElementById("change-font").onclick = () => tryCatch(changeFont);
    document.getElementById("insert-text-into-range").onclick = () => tryCatch(insertTextIntoRange);
    document.getElementById("insert-text-outside-range").onclick = () => tryCatch(insertTextBeforeRange);
    document.getElementById("replace-text").onclick = () => tryCatch(replaceText);
      document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
      document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
      document.getElementById("insert-table").onclick = () => tryCatch(insertTable);
      document.getElementById("create-content-control").onclick = () => tryCatch(createContentControl);
      document.getElementById("replace-content-in-control").onclick = () => tryCatch(replaceContentInControl);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function insertParagraph() {
    await Word.run(async (context) => {
        // TODO1: Queue commands to insert a paragraph into the document.
        const docBody = context.document.body;
        docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
            Word.InsertLocation.start);

        await context.sync();
    });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}

async function applyStyle() {
    await Word.run(async (context) => {
        const firstParagraph = context.document.body.paragraphs.getFirst();
        firstParagraph.styleBuiltIn = Word.Style.intenseReference;
        // TODO1: Queue commands to style text.
        await context.sync();
    });
}

async function applyCustomStyle() {
    await Word.run(async (context) => {
        // TODO1: Queue commands to apply the custom style.
        const lastParagraph = context.document.body.paragraphs.getLast();
        lastParagraph.style = "MyCustomStyle";
        await context.sync();
    });
}

async function changeFont() {
    await Word.run(async (context) => {
        // TODO1: Queue commands to apply a different font.
        const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
        secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
        await context.sync();
    });
}

async function insertTextIntoRange() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert text into a selected range.
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (M365)", Word.InsertLocation.end);
        // TODO2: Load the text of the range and sync so that the
        originalRange.load("text");
        await context.sync();
        //        current range text can be read.

        // TODO3: Queue commands to repeat the text of the original
        doc.body.insertParagraph("Original range: " + originalRange.text, Word.InsertLocation.end);
        //        range at the end of the document.

        await context.sync();
    });
}

async function insertTextBeforeRange() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert a new range before the
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText("Office 2019, ", Word.InsertLocation.before);
        //        selected range.

        // TODO2: Load the text of the original range and sync so that the
        originalRange.load("text");
        await context.sync();

// TODO3: Queue commands to insert the original range as a
//        paragraph at the end of the document.
        doc.body.insertParagraph("Current text of original range: " + originalRange.text, Word.InsertLocation.end);
// TODO4: Make a final call of context.sync here and ensure
//        that it runs after the insertParagraph has been queued.
        await context.sync();
        //        range text can be read and inserted.

    });
}

async function replaceText() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to replace the text.
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText("many", Word.InsertLocation.replace);
        await context.sync();
    });
}

async function insertImage() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert an image.
        context.document.body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);
        await context.sync();
    });
}

async function insertHTML() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert a string of HTML.
        const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
        blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', Word.InsertLocation.end);
        await context.sync();
    });
}

async function insertTable() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to get a reference to the paragraph
        //        that will precede the table.
        const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
        // TODO2: Queue commands to create a table and populate it with data.
        const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
        secondParagraph.insertTable(3, 3, Word.InsertLocation.after, tableData);
        await context.sync();
    });
}

async function createContentControl() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to create a content control.
        const serviceNameRange = context.document.getSelection();
        const serviceNameContentControl = serviceNameRange.insertContentControl();
        serviceNameContentControl.title = "Service Name";
        serviceNameContentControl.tag = "serviceName";
        serviceNameContentControl.appearance = "Tags";
        serviceNameContentControl.color = "blue";

        await context.sync();
    });
}

async function replaceContentInControl() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to replace the text in the Service Name
        //        content control.
        const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
        serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", Word.InsertLocation.replace);
        await context.sync();
    });
}