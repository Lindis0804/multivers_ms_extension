/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import * as bootstrap from "bootstrap";
import axios from "axios";

export function openInBrowser(event: Event) {
  event.preventDefault();

  const downloadLink = (event.target as HTMLAnchorElement).getAttribute("href");
  if (downloadLink) {
    window.open(downloadLink, "_blank"); // Open link in new tab
  }
}

interface Version {
  id: string;
  name: string;
  downloadLink: string;
}

const sleep = (ms) => {
  return new Promise((resolve) => setTimeout(resolve, ms));
};

const host = "http://localhost:2000";

const getNewestVersion = async (documentId: number) => {
  const response = await axios.get(`${host}/versions/document/${documentId}/newest`);
  return response.data?.result?.version;
};

const formatDownloadLink = (link: string) => {
  return `${host}${link}`;
};

const showToast = () => {
  const toast = document.getElementById("liveToast");
  const toastLive = bootstrap.Toast.getOrCreateInstance(toast);
  toastLive.show();
};

const notifyNewVersion = async ({ documentId, versionName }: { documentId: number; versionName: string }) => {
  while (1) {
    try {
      const version: Version = await getNewestVersion(documentId);

      if (versionName != version.name) {
        document.getElementById("toast-body").innerHTML = `New version ${version.name} is available. 
         <a href="${formatDownloadLink(version.downloadLink)}" id="download-link" 
        >Download here.</a>`;

        document.getElementById("download-link").onclick = openInBrowser;
      }

      showToast();
    } catch (err) {
      document.getElementById("toast-body").innerHTML = `Error in get version: ${err}`;

      showToast();
    }

    await sleep(120000);
  }
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("show_file_metadata_button").onclick = showFileMetadataPage;
    document.getElementById("show_web_view_button").onclick = showWebViewPage;

    showFileMetadataPage();
  }
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
  // showNotification();

  await Word.run(async (context) => {
    const metadata = context.document.properties;
    metadata.load(["title", "comments", "author", "subject", "lastAuthor", "creationDate", "keywords"]);
    let other: {
      version?: string;
      description?: string;
      documentId?: number;
      signature?: string;
    } = {
      version: null,
      description: null,
      documentId: null,
      signature: null,
    };
    await context.sync();

    try {
      other = JSON.parse(metadata.comments);
    } catch (err) {}

    document.getElementById("main_content").innerHTML = getFileMetadataTable({
      title: metadata.title || "",
      author: metadata.author || "",
      object: metadata.subject || "",
      lastModifiedBy: metadata.lastAuthor || "",
      createdAt: metadata.creationDate || "",
      version: other?.version || "",
      description: other?.description || "",
      documentId: other?.documentId || "",
      signature: metadata.keywords,
    });

    notifyNewVersion({ documentId: other.documentId, versionName: other.version });
  });
}

export async function showWebViewPage() {
  document.getElementById("main_content").innerHTML = `
  <h2>
Web View
  </h2>
  `;
}
