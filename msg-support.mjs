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
async function convertedFiles(files) {
  const result = [];
  const failed = [];
  for (const file of files) {
    if (!isMsg(file)) { result.push(file); continue; }
    try { result.push(await parseMsg(file)); }
    catch (err) { failed.push(file.name + ': ' + err.message); }
  }
  if (failed.length) showError(failed.join(' | '));
  return result;
}
function sendToOriginalImporter(files) {
  if (!files.length) return;
  const input = document.getElementById('mailInput');
  if (!input) return;
  const dt = new DataTransfer();
  files.forEach(file => dt.items.add(file));
  input.files = dt.files;
  input.dispatchEvent(new Event('change', { bubbles: true }));
}
export function enableMsgSupport() {
  const input = document.getElementById('mailInput');
  const dropZone = document.getElementById('dropZone');
  if (!input || !dropZone) return false;
  input.accept = '.eml,.msg,message/rfc822,application/vnd.ms-outlook';
  const hint = dropZone.querySelector('p');
  if (hint) hint.textContent = 'Archivos .msg de Outlook y .eml · Detección de duplicados y similitud';
  document.addEventListener('change', async event => {
    if (event.target !== input) return;
    const files = Array.from(input.files || []);
    if (!files.some(isMsg)) return;
    event.stopImmediatePropagation();
    const converted = await convertedFiles(files);
    sendToOriginalImporter(converted);
  }, true);
  document.addEventListener('drop', async event => {
    if (!dropZone.contains(event.target)) return;
    const files = Array.from(event.dataTransfer?.files || []);
    if (!files.some(isMsg)) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    const converted = await convertedFiles(files);
    sendToOriginalImporter(converted);
  }, true);
  return true;
}
if (typeof document !== 'undefined') {
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', enableMsgSupport, { once: true });
  else enableMsgSupport();
}
