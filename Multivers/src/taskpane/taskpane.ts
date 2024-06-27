/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
export const encodeString = (str: string): string => {
  // Chuyển đổi chuỗi thành mảng byte
  const bytes = new Uint8Array(str.split("").map((char) => char.charCodeAt(0)));

  // Sử dụng TextDecoder để giải mã từ 'latin1' sang 'utf-8'
  const latin1Decoder = new TextDecoder("iso-8859-1");
  const utf8Bytes = latin1Decoder.decode(bytes);

  // Chuyển đổi từ UTF-8 bytes sang chuỗi gốc
  const utf8Decoder = new TextDecoder("utf-8");
  return utf8Decoder.decode(new TextEncoder().encode(utf8Bytes));
};

Office.onReady(async () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  document.getElementById("show_file_metadata_button").onclick = showFileMetadataPage;
  document.getElementById("show_web_view_button").onclick = showWebViewPage;

  showFileMetadataPage();
});

export class fileMetadata {
  creator: string;
  version: string;
  title: string;
  subject: string;
  description: string;
  identifier: number;
  lastModifiedBy: string;
  created: Date;
  modified: Date;
}

export function getFileMetadataTable(params: Object): string {
  return `
  <table>
  <tr>
  <th>Field</th>
  <th>Value</th>
  </tr>
  ${
    Object.keys(params).length > 0
      ? `${Object.keys(params)
          .map(
            (key) => `
  <tr>
  <td>${key}</td>
  <td>${params[key]}</td>
  </tr>
  `
          )
          .join(``)}`
      : ""
  }
  </table>
  `;
}

export async function showFileMetadataPage() {
  showNotification();

  await Excel.run(async (context) => {
    const metadata = context.workbook.properties;
    metadata.load(["title", "comments", "author", "subject", "lastAuthor", "creationDate"]);
    let other = {
      version: null,
      description: null,
      documentId: null,
    };
    await context.sync();

    try {
      other = JSON.parse(metadata.comments);
    } catch (err) {}

    document.getElementById("main_content").innerHTML = getFileMetadataTable({
      title: metadata.title,
      author: metadata.author,
      object: metadata.subject,
      lastModifiedBy: metadata.lastAuthor,
      createdAt: metadata.creationDate,
      version: other?.version || "",
      description: other?.description || "",
      documentId: other?.documentId || "",
    });
  });
}

export async function showWebViewPage() {
  document.getElementById("main_content").innerHTML = `
  <h2>
Web View
  </h2>
  `;
}

export function showNotification() {
  Office.context.ui.displayDialogAsync(
    "http://localhost:3000/dialog.html",
    { width: 100, height: 100, displayInIframe: true },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        let dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    }
  );
}

function processMessage(arg) {
  console.log(arg.message);
  // Handle the message from the dialog
}
