/**
 * Outlook .msg adapter for GESMAIL.
 * Converts MSG files in memory to EML so existing RFQ duplicate detection
 * continues to work; data stays in the browser.
 */
const MSG_READER_URL = 'https://esm.sh/@npeersab/msgreader@5.1.0?bundle';
const MAX_MSG_BYTES = 35 * 1024 * 1024;
let readerPromise;

const isMsg = file => /\.msg$/i.test(file?.name || '');
const cleanHeader = value => String(value ?? '').replace(/[\r\n\0]+/g, ' ').trim();
const fieldFromHeaders = (headers, name) => {
  const match = String(headers ?? '').match(new RegExp('(?:^|\\r?\\n)' + name + ':\\s*([^\\r\\n]*)', 'i'));
  return match?.[1]?.trim() || '';
};
function rtfToPlain(rtf) {
  return String(rtf || '')
    .replace(/\\par\b\s?/g, '\n').replace(/\\line\b\s?/g, '\n').replace(/\\tab\b\s?/g, '\t')
    .replace(/\\u(-?\d+)\??/g, (_, n) => String.fromCharCode((Number(n) + 65536) % 65536))
    .replace(/\\'[0-9a-f]{2}/gi, ' ')
    .replace(/\\[a-z]+-?\d*\s?/gi, '')
    .replace(/\\[{}\\]/g, '')
    .replace(/[{}]/g, '')
    .replace(/[ \t]{2,}/g, ' ').trim();
}
function htmlToPlain(html) {
  if (!html) return '';
  if (typeof DOMParser === 'undefined') return String(html).replace(/<[^>]*>/g, ' ');
  const doc = new DOMParser().parseFromString(String(html), 'text/html');
  doc.querySelectorAll('script,style,noscript').forEach(el => el.remove());
  return (doc.body?.textContent || '').trim();
}
export function messageAsEml(info, reader) {
  const headers = String(info?.headers || '');
  const subject = cleanHeader(info?.subject || info?.normalizedSubject || fieldFromHeaders(headers, 'Subject'));
  const fromName = cleanHeader(info?.senderName || info?.sentRepresentingName);
  const fromEmail = cleanHeader(info?.senderSmtpAddress || info?.senderEmail || info?.sentRepresentingEmail);
  const sender = (fromEmail && /^[^\s<>@]+@[^\s<>@]+$/.test(fromEmail))
    ? (fromName ? fromName.replace(/[<>]/g, '') + ' ' : '') + '<' + fromEmail + '>'
    : (fromName || fromEmail || fieldFromHeaders(headers, 'From'));
  const msgId = cleanHeader(info?.internetMessageId || fieldFromHeaders(headers, 'Message-ID'));
  const time = info?.clientSubmitTime || info?.deliverTime || fieldFromHeaders(headers, 'Date');
  const date = time ? new Date(time) : null;
  let body = String(info?.body || '').trim();
  if (!body && typeof reader?.getHtmlBody === 'function') {
    try { body = htmlToPlain(reader.getHtmlBody()); } catch { /* try RTF */ }
  }
  if (!body && typeof reader?.getRtfBody === 'function') {
    try { body = rtfToPlain(reader.getRtfBody()); } catch { /* no RTF */ }
  }
  if (!subject && !sender) throw new Error('No se ha podido leer el asunto ni el remitente del archivo MSG.');
  const lines = [
    'From: ' + (sender || '(Sin remitente)'),
    'Subject: ' + (subject || '(Sin asunto)'),
    ...(msgId ? ['Message-ID: <' + msgId.replace(/^<|>$/g, '') + '>'] : []),
    ...(date && !Number.isNaN(date.getTime()) ? ['Date: ' + date.toUTCString()] : []),
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset=utf-8',
    'Content-Transfer-Encoding: 8bit',
    '', body
  ];
  return lines.join('\r\n');
}
async function parseMsg(file) {
  if (file.size > MAX_MSG_BYTES) throw new Error('El correo supera el límite de 35 MB.');
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (bytes.length < 8 || !signature.every((value, i) => bytes[i] === value)) {
    throw new Error('Este archivo no contiene un mensaje MSG de Outlook válido.');
  }
  if (!readerPromise) {
    readerPromise = import(MSG_READER_URL).catch(err => { readerPromise = null; throw err; });
  }
  let module;
  try { module = await readerPromise; }
  catch { throw new Error('No se pudo cargar el lector MSG. Comprueba la conexión a Internet e inténtalo otra vez.'); }
  const MsgReader = module.MsgReader || module.default;
  if (typeof MsgReader !== 'function') throw new Error('El lector de MSG no está disponible.');
  let eml;
  try {
    const reader = new MsgReader(buffer);
    eml = messageAsEml(reader.getFileData(), reader);
  } catch (err) {
    throw new Error('No se pudo interpretar el MSG: ' + (err?.message || 'formato incompatible'));
  }
  return new File([eml], file.name + '.eml', { type: 'message/rfc822' });
}
function showError(message) {
  const host = document.getElementById('toastHost');
  if (!host) { window.alert(message); return; }
  host.replaceChildren();
  const box = document.createElement('div');
  box.className = 'toast error';
  box.textContent = message;
  host.append(box);
  setTimeout(() => box.remove(), 11000);
}
const isEml = file => /\.eml$/i.test(file?.name || '');
const isOutlookMime = file => /^(application\/vnd\.ms-outlook|application\/x-ole-storage|application\/msoutlook)$/i.test(file?.type || '');
const MSG_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function hasExternalFileDrag(dataTransfer) {
  if (!dataTransfer) return false;
  // Outlook clásico ofrece ficheros virtuales; Chromium los expone como "Files".
  return Array.from(dataTransfer.types || []).includes('Files') ||
    Array.from(dataTransfer.items || []).some(item => item.kind === 'file');
}

function filesFromTransfer(dataTransfer) {
  // Durante dragover pueden no estar disponibles: leer sólo al soltar.
  const files = Array.from(dataTransfer?.files || []).filter(Boolean);
  if (files.length) return files;
  return Array.from(dataTransfer?.items || [])
    .filter(item => item.kind === 'file' && typeof item.getAsFile === 'function')
    .map(item => item.getAsFile()).filter(Boolean);
}

async function normalizeEmailFile(file) {
  if (isEml(file) || isMsg(file)) return file;
  // Algunos arrastres de Outlook no proporcionan la extensión .msg.
  if (isOutlookMime(file)) {
    return new File([file], (file.name || 'Correo Outlook') + '.msg',
      { type: 'application/vnd.ms-outlook' });
  }
  if (file.size >= 8 && file.size <= MAX_MSG_BYTES) {
    try {
      const bytes = new Uint8Array(await file.slice(0, 8).arrayBuffer());
      if (MSG_SIGNATURE.every((value, index) => bytes[index] === value)) {
        return new File([file], (file.name || 'Correo Outlook') + '.msg',
          { type: 'application/vnd.ms-outlook' });
      }
    } catch { /* No es un MSG legible; no aceptar otro tipo de fichero. */ }
  }
  return null;
}

async function convertedFiles(files) {
  const result = [];
  const failed = [];
  for (const file of files) {
    try {
      const email = await normalizeEmailFile(file);
      if (!email) {
        failed.push((file.name || 'Archivo sin nombre') + ': sólo se admiten correos .msg o .eml.');
        continue;
      }
      result.push(isMsg(email) ? await parseMsg(email) : email);
    } catch (err) {
      failed.push((file.name || 'Correo Outlook') + ': ' + (err?.message || 'error de lectura'));
    }
  }
  if (failed.length) showError(failed.join(' | '));
  return result;
}

function sendToOriginalImporter(files) {
  if (!files.length) return;
  const input = document.getElementById('mailInput');
  if (!input) { showError('No se encuentra el importador de correos de GESMAIL.'); return; }
  const dt = new DataTransfer();
  files.forEach(file => dt.items.add(file));
  input.files = dt.files;
  // Seguimos usando el importador original: RFQ, duplicados y similitud.
  input.dispatchEvent(new Event('change', { bubbles: true }));
}

/**
 * Mantiene los avisos y las acciones existentes, pero los sitúa
 * en el centro de la pantalla, también en móviles.
 */
function centerToastNotifications() {
  if (document.getElementById('gesmailCenteredToasts')) return;
  const style = document.createElement('style');
  style.id = 'gesmailCenteredToasts';
  style.textContent = `
    #toastHost {
      position: fixed !important;
      top: 50% !important;
      left: 50% !important;
      right: auto !important;
      bottom: auto !important;
      transform: translate(-50%, -50%) !important;
      width: min(92vw, 540px) !important;
      max-width: calc(100vw - 24px) !important;
      max-height: 80vh !important;
      padding: 0 !important;
      display: flex !important;
      flex-direction: column !important;
      justify-content: center !important;
      align-items: stretch !important;
      gap: 12px !important;
      overflow-y: auto !important;
      z-index: 2147483600 !important;
      pointer-events: none !important;
    }
    #toastHost .toast {
      box-sizing: border-box;
      width: 100% !important;
      max-width: 100% !important;
      margin: 0 !important;
      pointer-events: auto !important;
    }
  `;
  document.head.appendChild(style);
}

export function enableMsgSupport() {
  const input = document.getElementById('mailInput');
  const dropZone = document.getElementById('dropZone');
  if (!input || !dropZone) return false;
  if (document.getElementById('outlookDropOverlay')) return true;

  input.accept = '.eml,.msg,message/rfc822,application/vnd.ms-outlook';
  const hint = dropZone.querySelector('p');
  if (hint) hint.textContent = 'Arrastra correos directamente desde Outlook clásico (Edge o Chrome), o archivos .msg / .eml';

  const overlay = document.createElement('div');
  overlay.id = 'outlookDropOverlay';
  overlay.setAttribute('aria-hidden', 'true');
  overlay.textContent = 'Suelta aquí el correo de Outlook para crear una RFQ';
  Object.assign(overlay.style, {
    position: 'fixed',
    inset: '0',
    zIndex: '2147483646',
    boxSizing: 'border-box',
    display: 'none',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '32px',
    border: '6px dashed #2563eb',
    background: 'rgba(15, 23, 42, 0.86)',
    color: '#fff',
    font: 'bold 24px system-ui, sans-serif',
    textAlign: 'center',
    pointerEvents: 'none'
  });
  document.body.appendChild(overlay);

  let dragDepth = 0;
  const hideOverlay = () => { dragDepth = 0; overlay.style.display = 'none'; };
  const showOverlay = () => { overlay.style.display = 'flex'; };

  // Mantener el diálogo "Importar correos" y su comportamiento anterior.
  document.addEventListener('change', async event => {
    if (event.target !== input) return;
    const files = Array.from(input.files || []);
    if (!files.some(file => isMsg(file) || isOutlookMime(file))) return;
    event.stopImmediatePropagation();
    const converted = await convertedFiles(files);
    sendToOriginalImporter(converted);
  }, true);

  // Captura global sólo para ficheros externos; no interfiere con las
  // tarjetas RFQ que se arrastran dentro del tablero.
  document.addEventListener('dragenter', event => {
    if (!hasExternalFileDrag(event.dataTransfer)) return;
    dragDepth++;
    showOverlay();
  }, true);

  document.addEventListener('dragover', event => {
    if (!hasExternalFileDrag(event.dataTransfer)) return;
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
    showOverlay();
  }, true);

  document.addEventListener('dragleave', event => {
    if (!hasExternalFileDrag(event.dataTransfer)) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) hideOverlay();
  }, true);

  document.addEventListener('drop', async event => {
    const transfer = event.dataTransfer;
    if (!hasExternalFileDrag(transfer)) return;
    // Leer las entradas ANTES de la primera espera asíncrona: DataTransfer
    // deja de estar disponible tras finalizar el evento drop.
    const files = filesFromTransfer(transfer);
    event.preventDefault();
    event.stopImmediatePropagation();
    hideOverlay();

    if (!files.length) {
      showError('Outlook no ha entregado el correo al navegador. Prueba con Microsoft Edge o Chrome actualizado. Como alternativa, arrastra el correo al escritorio y después a GESMAIL.');
      return;
    }
    const converted = await convertedFiles(files);
    sendToOriginalImporter(converted);
  }, true);

  window.addEventListener('blur', hideOverlay);
  return true;
}

if (typeof document !== 'undefined') {
  centerToastNotifications();
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', enableMsgSupport, { once: true });
  else enableMsgSupport();
}
