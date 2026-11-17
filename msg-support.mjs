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
/**
 * Convierte los avisos del tablero en tarjetas visuales sin cambiar la función
 * toast() original ni su detección de duplicados o temporizador de cierre.
 */
function upgradeToast(toast) {
  if (!toast || toast.dataset.gesmailPopup === 'true') return;
  toast.dataset.gesmailPopup = 'true';
  toast.classList.add('gesmail-popup');
  toast.setAttribute('role', 'status');

  const messageText = toast.querySelector('span')?.textContent || toast.textContent || '';
  const rawMessage = toast.querySelector('span');
  const originalButton = toast.querySelector('button');
  const isDuplicate = /duplicad[oa]|message-id idéntico|message-id identico/i.test(messageText);
  const isSimilar = /similitud|parecid[oa]|similar/i.test(messageText);
  const kind = toast.classList.contains('error') ? 'error'
    : toast.classList.contains('success') ? 'success'
    : toast.classList.contains('warning') ? 'warning' : 'info';

  const title = isDuplicate ? 'Duplicado detectado'
    : isSimilar ? 'Correo similar detectado'
    : kind === 'warning' ? 'Atención'
    : kind === 'error' ? 'No se pudo completar'
    : kind === 'success' ? 'Operación completada' : 'Notificación';

  const icon = document.createElement('div');
  icon.className = 'gesmail-popup-icon';
  icon.setAttribute('aria-hidden', 'true');
  icon.textContent = kind === 'success' ? '✓' : kind === 'error' ? '×'
    : kind === 'warning' ? '!' : 'i';

  const content = document.createElement('div');
  content.className = 'gesmail-popup-content';

  const heading = document.createElement('div');
  heading.className = 'gesmail-popup-heading';
  heading.textContent = title;

  const message = rawMessage || document.createElement('span');
  if (!rawMessage) message.textContent = messageText;
  message.className = 'gesmail-popup-message';

  const actions = document.createElement('div');
  actions.className = 'gesmail-popup-actions';
  const close = originalButton || document.createElement('button');
  close.className = 'gesmail-popup-close';
  close.type = 'button';
  close.textContent = 'Entendido';
  close.setAttribute('aria-label', 'Cerrar aviso');
  // El botón original conserva su onclick. Los avisos del lector MSG
  // carecen de botón, por lo que añadimos un cierre en esos casos.
  if (!originalButton) close.addEventListener('click', () => toast.remove());

  toast.replaceChildren();
  actions.appendChild(close);
  content.append(heading, message, actions);
  toast.append(icon, content);
}

function centerToastNotifications() {
  if (document.getElementById('gesmailPopupStyle')) return;
  const style = document.createElement('style');
  style.id = 'gesmailPopupStyle';
  style.textContent = `
    html body #toastHost {
      position: static !important;
      top: auto !important;
      left: auto !important;
      right: auto !important;
      bottom: auto !important;
      transform: none !important;
      width: auto !important;
      max-width: none !important;
      max-height: none !important;
      overflow: visible !important;
      display: block !important;
      padding: 0 !important;
      pointer-events: none !important;
    }

    html body #toastHost:has(> .toast.gesmail-popup)::before {
      content: '';
      position: fixed;
      inset: 0;
      z-index: 2147483598;
      background: rgba(15, 23, 42, .14);
      pointer-events: none;
    }

    html body #toastHost > .toast.gesmail-popup {
      --popup-accent: #2563eb;
      --popup-light: #eaf2ff;
      position: fixed !important;
      top: 50% !important;
      left: 50% !important;
      right: auto !important;
      bottom: auto !important;
      transform: translate(-50%, -50%) !important;
      z-index: 2147483600 !important;
      box-sizing: border-box !important;
      display: flex !important;
      align-items: flex-start !important;
      gap: 16px !important;
      width: min(460px, calc(100vw - 32px)) !important;
      max-width: calc(100vw - 32px) !important;
      max-height: min(80vh, 640px) !important;
      margin: 0 !important;
      padding: 22px !important;
      overflow-y: auto !important;
      border-radius: 18px !important;
      border: 1px solid #e4eaf2 !important;
      background: #fff !important;
      color: #1f2937 !important;
      box-shadow: 0 22px 65px rgba(16, 32, 56, .24) !important;
      font: 400 13px/1.55 system-ui, -apple-system, "Segoe UI", sans-serif !important;
      pointer-events: auto !important;
      animation: gesmailPopupAppear .2s ease-out both;
    }

    html body #toastHost > .toast.gesmail-popup.warning {
      --popup-accent: #b7791f;
      --popup-light: #fff3d9;
    }

    html body #toastHost > .toast.gesmail-popup.error {
      --popup-accent: #dc3545;
      --popup-light: #fde9eb;
    }

    html body #toastHost > .toast.gesmail-popup.success {
      --popup-accent: #148657;
      --popup-light: #e4f7ee;
    }

    #toastHost .gesmail-popup-icon {
      flex: 0 0 46px;
      width: 46px;
      height: 46px;
      box-sizing: border-box;
      border-radius: 13px;
      display: grid;
      place-items: center;
      background: var(--popup-light);
      color: var(--popup-accent);
      font-size: 23px;
      font-weight: 800;
      line-height: 1;
    }

    #toastHost .gesmail-popup-content {
      flex: 1;
      min-width: 0;
    }

    #toastHost .gesmail-popup-heading {
      padding-top: 1px;
      margin-bottom: 8px;
      font-size: 16px;
      line-height: 1.3;
      font-weight: 750;
      color: #182c45;
    }

    #toastHost .gesmail-popup-message {
      display: block;
      font-size: 13px;
      line-height: 1.65;
      color: #526276;
      overflow-wrap: anywhere;
      white-space: pre-wrap;
    }

    #toastHost .gesmail-popup-actions {
      margin-top: 18px;
      display: flex;
      justify-content: flex-end;
    }

    html body #toastHost .gesmail-popup-close {
      margin: 0 !important;
      padding: 9px 18px !important;
      border: 0 !important;
      border-radius: 10px !important;
      background: #2468e8 !important;
      color: #fff !important;
      font-size: 12px !important;
      line-height: 1.4 !important;
      font-weight: 700 !important;
      cursor: pointer;
    }

    html body #toastHost .gesmail-popup-close:hover {
      background: #1d58cb !important;
    }

    html body #toastHost .gesmail-popup-close:focus-visible {
      outline: 3px solid #93baff !important;
      outline-offset: 3px;
    }

    @keyframes gesmailPopupAppear {
      from { opacity: 0; translate: 0 12px; scale: .97; }
      to { opacity: 1; translate: 0 0; scale: 1; }
    }

    @media (prefers-reduced-motion: reduce) {
      html body #toastHost > .toast.gesmail-popup { animation: none !important; }
    }
  `;
  document.head.appendChild(style);

  const host = document.getElementById('toastHost');
  if (!host) return;
  const process = () => {
    Array.from(host.children).forEach(element => {
      if (element.classList?.contains('toast')) upgradeToast(element);
    });
  };
  process();
  new MutationObserver(process).observe(host, { childList: true });
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
