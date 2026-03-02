/* global document, Office */

let jsonAttachments: any[] = [];

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  const appBody = document.getElementById("app-body");
  if (appBody) {
    appBody.style.display = "block";
  }

  try {
    const params = new URL(globalThis.location.href).searchParams;
    if (params.get("openFirstAttachment") === "true") {
      setTimeout(() => autoOpenFirstJsonAttachment(), 250);
    }
  } catch {
    showError("Impossibile leggere i parametri della pagina");
  }
});

function autoOpenFirstJsonAttachment() {
  const item = Office.context.mailbox.item as any;
  const attachments = item?.attachments || [];

  if (!attachments.length) {
    showError("Nessun allegato presente");
    return;
  }

  jsonAttachments = attachments.filter((attachment: any) => {
    const name = (attachment.name || attachment.displayName || "").toLowerCase();
    const contentType = (attachment.contentType || "").toLowerCase();
    return name.endsWith(".json") || contentType.includes("json");
  });

  if (!jsonAttachments.length) {
    showError("Nessun allegato JSON trovato");
    renderAttachmentSelector([]);
    return;
  }

  renderAttachmentSelector(jsonAttachments);
  fetchAttachment(jsonAttachments[0].id);
}

function renderAttachmentSelector(attachments: any[]) {
  const picker = document.getElementById("attachment-picker");
  const select = document.getElementById("attachment-select") as HTMLSelectElement | null;
  if (!picker || !select) {
    return;
  }

  select.innerHTML = "";
  if (!attachments.length || attachments.length === 1) {
    picker.style.display = "none";
    return;
  }

  attachments.forEach((attachment, index) => {
    const option = document.createElement("option");
    option.value = attachment.id || "";
    option.text = attachment.name || attachment.displayName || `JSON ${index + 1}`;
    select.add(option);
  });

  select.onchange = () => {
    const selectedId = select.value;
    if (selectedId) {
      fetchAttachment(selectedId);
    }
  };

  picker.style.display = "flex";
}

function fetchAttachment(attachmentId: string) {
  clearError();
  const item = Office.context.mailbox.item as any;

  item.getAttachmentContentAsync(attachmentId, (result: any) => {
    if (result?.status !== "succeeded") {
      const message = result?.error?.message || "Errore nel recupero dell'allegato";
      showError(message);
      return;
    }

    const value = result.value || {};
    let text = "";

    if (value.format === "text") {
      text = value.content || "";
    } else if (value.format === "base64") {
      try {
        text = atob(value.content || "");
      } catch {
        showError("Impossibile decodificare allegato base64");
        return;
      }
    } else {
      text = value.content || "";
    }

    renderJson(text);
  });
}

function renderJson(text: string) {
  const viewer = document.getElementById("viewer");
  if (!viewer) {
    return;
  }

  if (!text) {
    viewer.textContent = "";
    showError("Allegato vuoto");
    return;
  }

  try {
    const parsed = JSON.parse(text);
    const formatted = JSON.stringify(parsed, null, 2);
    viewer.innerHTML = highlightJson(formatted);
    clearError();
  } catch (error) {
    viewer.textContent = `Errore parsing JSON: ${String(error)}\n\n${text}`;
    showError("JSON non valido");
  }
}

function highlightJson(json: string): string {
  let output = "";
  let index = 0;

  while (index < json.length) {
    const stringToken = tryReadStringToken(json, index);
    if (stringToken) {
      output += stringToken.html;
      index = stringToken.nextIndex;
      continue;
    }

    const numberToken = tryReadNumberToken(json, index);
    if (numberToken) {
      output += numberToken.html;
      index = numberToken.nextIndex;
      continue;
    }

    const keywordToken = tryReadKeywordToken(json, index);
    if (keywordToken) {
      output += keywordToken.html;
      index = keywordToken.nextIndex;
      continue;
    }

    output += escapeHtml(json[index]);
    index += 1;
  }

  return output;
}

function escapeHtml(value: string): string {
  return value
    .split("&")
    .join("&amp;")
    .split("<")
    .join("&lt;")
    .split(">")
    .join("&gt;");
}

function tryReadStringToken(text: string, startIndex: number): { html: string; nextIndex: number } | null {
  if (text[startIndex] !== '"') {
    return null;
  }

  let endIndex = startIndex + 1;
  while (endIndex < text.length) {
    if (text[endIndex] === "\\") {
      endIndex += 2;
      continue;
    }
    if (text[endIndex] === '"') {
      endIndex += 1;
      break;
    }
    endIndex += 1;
  }

  const token = text.slice(startIndex, endIndex);
  const cssClass = isJsonKey(text, endIndex) ? "json-key" : "json-string";
  return {
    html: `<span class="${cssClass}">${escapeHtml(token)}</span>`,
    nextIndex: endIndex,
  };
}

function tryReadNumberToken(text: string, startIndex: number): { html: string; nextIndex: number } | null {
  const current = text[startIndex];
  if (!isNumberStart(current, text, startIndex)) {
    return null;
  }

  let endIndex = startIndex + 1;
  while (endIndex < text.length && /[\deE+\-.]/.test(text[endIndex])) {
    endIndex += 1;
  }

  return {
    html: `<span class="json-number">${escapeHtml(text.slice(startIndex, endIndex))}</span>`,
    nextIndex: endIndex,
  };
}

function tryReadKeywordToken(text: string, startIndex: number): { html: string; nextIndex: number } | null {
  const keyword = readKeyword(text, startIndex);
  if (!keyword) {
    return null;
  }

  const cssClass = keyword === "null" ? "json-null" : "json-boolean";
  return {
    html: `<span class="${cssClass}">${keyword}</span>`,
    nextIndex: startIndex + keyword.length,
  };
}

function isJsonKey(text: string, afterStringIndex: number): boolean {
  let probe = afterStringIndex;
  while (probe < text.length && /\s/.test(text[probe])) {
    probe += 1;
  }
  return probe < text.length && text[probe] === ":";
}

function readKeyword(text: string, startIndex: number): string | null {
  if (text.startsWith("true", startIndex)) return "true";
  if (text.startsWith("false", startIndex)) return "false";
  if (text.startsWith("null", startIndex)) return "null";
  return null;
}

function isNumberStart(char: string, text: string, index: number): boolean {
  if (/\d/.test(char)) return true;
  if (char !== "-") return false;
  const next = text[index + 1];
  return !!next && /\d/.test(next);
}

function showError(message: string) {
  const errorElement = document.getElementById("error");
  if (errorElement) {
    errorElement.textContent = message;
  }
}

function clearError() {
  const errorElement = document.getElementById("error");
  if (errorElement) {
    errorElement.textContent = "";
  }
}
