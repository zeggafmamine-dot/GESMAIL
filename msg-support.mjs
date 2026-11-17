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
// Convierte texto y firmas HTML de Outlook a MIME sin subir archivos.
function bytesBase64(bytes) {
  let binary = '';
  for (let pos = 0; pos < bytes.length; pos += 8192) {
    binary += String.fromCharCode(...bytes.subarray(pos, pos + 8192));
  }
  return btoa(binary);
}
const utf8Base64 = text => bytesBase64(new TextEncoder().encode(String(text ?? '')));
function encodedMimeLines(text) {
  return String(text).match(/.{1,76}/g)?.join('\r\n') || '';
}

function normalizeImageRef(value) {
  let valueText=String(value||'').trim().replace(/^['"]|['"]$/g,'').replace(/^cid:/i,'').replace(/^<|>$/g,'');
  try {valueText=decodeURIComponent(valueText)} catch { /* conservar valor */ }
  return valueText.replace(/\\/g,'/').replace(/^\.\//,'').trim().toLowerCase();
}
function sniffImageMime(bytes) {
  if(bytes.length>=8&&bytes[0]===137&&bytes[1]===80&&bytes[2]===78&&bytes[3]===71)return 'image/png';
  if(bytes.length>=3&&bytes[0]===255&&bytes[1]===216&&bytes[2]===255)return 'image/jpeg';
  if(bytes.length>=6&&bytes[0]===71&&bytes[1]===73&&bytes[2]===70&&bytes[3]===56)return 'image/gif';
  if(bytes.length>=12&&bytes[0]===82&&bytes[1]===73&&bytes[2]===70&&bytes[3]===70
    &&bytes[8]===87&&bytes[9]===69&&bytes[10]===66&&bytes[11]===80)return 'image/webp';
  return '';
}
function inlineMsgImages(info,reader,html) {
  const list=Array.isArray(info?.attachments)?info.attachments:[];
  let resultHtml=String(html||'');
  const images=[];
  let remaining=1200*1024;
  const refs=[...resultHtml.matchAll(/(?:\b(?:src|background)\s*=\s*["']?([^"'\s>]+))|(?:url\(\s*["']?([^"')\s]+))/gi)]
    .map(match=>match[1]||match[2]).filter(Boolean);
  const seen=new Set();
  for(let index=0;index<list.length&&images.length<25;index++){
    const entry=list[index];
    const aliases=[entry.pidContentId,entry.contentId,entry.attachContentLocation,entry.fileName,entry.fileNameShort]
      .filter(Boolean).map(normalizeImageRef);
    if(!aliases.length)continue;
    const matched=refs.find(ref=>{
      const normalized=normalizeImageRef(ref);
      const basename=normalized.split('/').pop();
      return aliases.some(alias=>normalized===alias||basename===alias||
        ((normalized.startsWith(alias+'@')||basename.startsWith(alias+'@'))&&alias.includes('.')));
    });
    if(!matched||entry.contentLength>remaining)continue;
    let data;
    try {data=reader.getAttachment(entry)}catch {
      try {data=reader.getAttachment(index)}catch {continue}
    }
    const content=data?.content;
    const bytes=content instanceof Uint8Array?content:
      content instanceof ArrayBuffer?new Uint8Array(content):
      Array.isArray(content)?new Uint8Array(content):null;
    if(!bytes?.length||bytes.length>700*1024||bytes.length>remaining)continue;
    const mime=sniffImageMime(bytes);
    if(!mime)continue;
    const originalCid=normalizeImageRef(entry.pidContentId||entry.contentId);
    const matchedCid=/^cid:/i.test(matched)?normalizeImageRef(matched):'';
    let cid=originalCid||matchedCid||('gesmail-image-'+index+'@inline');
    if(!/^[^\s<>"]{1,200}$/.test(cid))cid='gesmail-image-'+index+'@inline';
    if(seen.has(cid))continue;
    seen.add(cid);
    if(normalizeImageRef(matched)!==cid){
      resultHtml=resultHtml.replace(/((?:src|background)\s*=\s*["']?)([^"'\s>]+)/gi,
        (full,prefix,value)=>normalizeImageRef(value)===normalizeImageRef(matched)?prefix+'cid:'+cid:full);
      resultHtml=resultHtml.replace(/(url\(\s*["']?)([^"')\s]+)/gi,
        (full,prefix,value)=>normalizeImageRef(value)===normalizeImageRef(matched)?prefix+'cid:'+cid:full);
    }else if(!/^cid:/i.test(matched)){
      resultHtml=resultHtml.replace(/((?:src|background)\s*=\s*["']?)([^"'\s>]+)/gi,
        (full,prefix,value)=>normalizeImageRef(value)===cid?prefix+'cid:'+cid:full);
    }
    remaining-=bytes.length;
    images.push({cid,mime,base64:bytesBase64(bytes),
      filename:cleanHeader(entry.fileName||entry.fileNameShort||('imagen-'+index+'.png')).replace(/[";\\]/g,'_')});
  }
  return {images,html:resultHtml};
}
export function messageAsEml(info, reader) {
  const headers = String(info?.headers || '');
  const subject = cleanHeader(info?.subject || info?.normalizedSubject || fieldFromHeaders(headers, 'Subject'));
  const fromName = cleanHeader(info?.senderName || info?.sentRepresentingName);
  const fromEmail = cleanHeader(info?.senderSmtpAddress || info?.senderEmail || info?.sentRepresentingEmail);
  const sender = (fromEmail && /^[^\s<>@]+@[^\s<>@]+$/.test(fromEmail))
    ? (fromName ? fromName.replace(/[<>]/g, '') + ' ' : '') + '<' + fromEmail + '>'
    : (fromName || fromEmail || fieldFromHeaders(headers, 'From'));
  const recipientList = Array.isArray(info?.recipients) ? info.recipients : [];
  const recipientText = type => recipientList
    .filter(r => String(r.recipType || r.recipientType || r.type || '').toLowerCase() === type)
    .map(r => cleanHeader(r.smtpAddress || r.email || r.name || '')).filter(Boolean).join(', ');
  const to = cleanHeader(fieldFromHeaders(headers, 'To') || info?.displayTo || recipientText('to'));
  const cc = cleanHeader(fieldFromHeaders(headers, 'Cc') || info?.displayCc || recipientText('cc'));
  const inReplyTo = cleanHeader(fieldFromHeaders(headers, 'In-Reply-To') || info?.inReplyTo || '');
  const references = cleanHeader(fieldFromHeaders(headers, 'References') || info?.references || '');
  const msgId = cleanHeader(info?.internetMessageId || fieldFromHeaders(headers, 'Message-ID'));
  const time = info?.clientSubmitTime || info?.deliverTime || fieldFromHeaders(headers, 'Date');
  const date = time ? new Date(time) : null;
  let html = '';
  try { if (typeof reader?.getHtmlBody === 'function') html = String(reader.getHtmlBody() || ''); }
  catch { /* algunos MSG contienen solamente texto o RTF */ }
  if (!html && typeof info?.bodyHtml === 'string') html = info.bodyHtml;
  if (!html && info?.bodyHtml instanceof Uint8Array) html = new TextDecoder('utf-8').decode(info.bodyHtml);
  let body = String(info?.body || '').trim();
  if (!body && html) body = htmlToPlain(html);
  if (!body && typeof reader?.getRtfBody === 'function') {
    try { body = rtfToPlain(reader.getRtfBody()); } catch { /* sin RTF */ }
  }
  if (!subject && !sender) throw new Error('No se ha podido leer el asunto ni el remitente del archivo MSG.');
  const lines = [
    'From: ' + (sender || '(Sin remitente)'),
    'Subject: ' + (subject || '(Sin asunto)'),
    ...(to ? ['To: ' + to] : []),
    ...(cc ? ['Cc: ' + cc] : []),
    ...(inReplyTo ? ['In-Reply-To: ' + inReplyTo] : []),
    ...(references ? ['References: ' + references] : []),
    ...(msgId ? ['Message-ID: <' + msgId.replace(/^<|>$/g, '') + '>'] : []),
    ...(date && !Number.isNaN(date.getTime()) ? ['Date: ' + date.toUTCString()] : []),
    'MIME-Version: 1.0'
  ];
  if (!html) {
    lines.push('Content-Type: text/plain; charset=utf-8','Content-Transfer-Encoding: base64','',encodedMimeLines(utf8Base64(body)));
    return lines.join('\r\n');
  }
  const token = Math.random().toString(36).slice(2);
  const outer = 'gesmail_related_' + token;
  const alternative = 'gesmail_alternative_' + token;
  const embedded=inlineMsgImages(info,reader,html);
  html=embedded.html;
  const images=embedded.images;
  lines.push('Content-Type: multipart/related; boundary="' + outer + '"','',
    '--' + outer,
    'Content-Type: multipart/alternative; boundary="' + alternative + '"','',
    '--' + alternative,
    'Content-Type: text/plain; charset=utf-8',
    'Content-Transfer-Encoding: base64','',
    encodedMimeLines(utf8Base64(body)),
    '--' + alternative,
    'Content-Type: text/html; charset=utf-8',
    'Content-Transfer-Encoding: base64','',
    encodedMimeLines(utf8Base64(html)),
    '--' + alternative + '--');
  for (const image of images) {
    lines.push('--' + outer,'Content-Type: ' + image.mime + '; name="' + image.filename + '"',
      'Content-Disposition: inline; filename="' + image.filename + '"',
      'Content-ID: <' + image.cid + '>',
      'Content-Transfer-Encoding: base64','',
      encodedMimeLines(image.base64));
  }
  lines.push('--' + outer + '--','');
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
  if (typeof window.gesmailShowToast === 'function') {
    window.gesmailShowToast(message, 'error');
    return;
  }
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

if (typeof window !== 'undefined') window.gesmailConvertEmailFiles = convertedFiles;

if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', enableMsgSupport, { once: true });
  else enableMsgSupport();
}
