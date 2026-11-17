# GESMAIL — RFQ Equipo (demo GitHub Pages)

Demo web del tablero de solicitudes de oferta (RFQ) del departamento comercial.

## Funcionalidades
- Cuatro columnas: **Petición**, **Oferta enviada**, **Oferta perdida** y **Pedido ganado**.
- Importación de correos en formato **`.eml` y `.msg` de Outlook**, también mediante arrastre directo desde **Outlook clásico de Windows** a Chrome o Edge cuando el navegador proporciona el fichero virtual.
- Aviso y bloqueo de duplicados exactos (Message-ID, huella o fichero idéntico).
- **Detección de posibles hilos al arrastrar un correo**: cruza referencias `In-Reply-To`/`References`, asunto sin prefijos `RE:`/`FW:`, remitentes, destinatarios, CC y similitud del texto. Si hay indicios, abre una ventana para comparar el mensaje registrado (izquierda) y el nuevo (derecha), ambos en solo lectura. La puntuación es orientativa, no una probabilidad.
- En cada RFQ, **asunto**, **remitente**, **destinatarios** y **personas en copia (CC)** se obtienen automáticamente al importar un correo; **empresa** y **propietario/responsable** se asignan manualmente.
- **Lector de correo de solo lectura estilo Outlook**: desde la RFQ, pulsa **Leer correo completo** para abrir un visor grande con asunto, remitente, fecha, destinatarios, CC, cuerpo HTML y firmas originales. Permite pasar al correo anterior o siguiente dentro del hilo, ampliar la ventana, ajustar el zoom entre 60 % y 160 %, restablecer el 100 % o alternar entre HTML original y texto. No permite editar, responder ni enviar correos. Las imágenes externas se bloquean por defecto; las incrustadas se reproducen desde los datos del mensaje cuando están disponibles.
- **Hilo de correos** dentro de cada ficha: mensajes con fecha, remitente, To, CC y vista previa del cuerpo. Incluso si la referencia de respuesta coincide, GESMAIL solicita aprobación antes de vincular el nuevo correo. Puedes elegir **Sí, añadir al hilo**, **No, crear RFQ aparte** o **Descartar correo**. Si hay varias coincidencias, puedes seleccionar la RFQ con la que comparar.
- **Archivos adjuntos**: el lector muestra los adjuntos de cada mensaje, con nombre, tipo y tamaño. Permite descargar cualquier archivo guardado y obtener vista previa de imágenes (PNG/JPEG/GIF/WebP), PDF o texto/CSV. Los archivos binarios se almacenan localmente en IndexedDB, no en GitHub.
- **Usuarios del equipo** configurables desde el menú lateral, guardados **sólo en ese navegador** (no hay todavía registro, autenticación ni permisos reales).
- Creación manual, edición y cambio de estado de RFQ.
- **Borrar negocio:** abre una RFQ y selecciona **Borrar negocio** en la esquina inferior izquierda. Se solicita confirmación y sólo se elimina esa RFQ, junto con sus mensajes asociados, notas e historial. Es irreversible; exporta una copia JSON antes si deseas conservarla.
- Búsqueda, exportación y restauración de copias JSON, incluyendo los datos binarios de adjuntos locales que se hayan guardado correctamente (con límite aproximado de 80 MB por copia).
- Botón **Cargar ejemplos** con correos ficticios para pruebas.

## Publicar en GitHub Pages
El archivo `index.html` está en la raíz de la rama `main`. Para activar el sitio:
1. Abre **Settings → Pages** en este repositorio.
2. En **Build and deployment**, elige **Deploy from a branch**.
3. Selecciona la rama **main** y la carpeta **/(root)**.
4. Pulsa **Save**.

Dirección prevista: https://zeggafmamine-dot.github.io/GESMAIL/

## Avisos importantes
- **Esta versión NO utiliza Supabase ni una base de datos compartida.**
- Los datos se guardan exclusivamente en el **localStorage de cada navegador**. Los tres compañeros no comparten RFQ en esta demo.
- No hay inicio de sesión ni controles de acceso.
- Los correos y sus adjuntos **no se sincronizan entre ordenadores ni entre los tres miembros del equipo**: están en el navegador donde se importan. Para trasladarlos utiliza **Exportar datos** e **Importar copia**; al eliminar un negocio también se eliminan sus adjuntos locales cuando ya no están referenciados por otras RFQ.
- Los adjuntos con tamaño superior al límite o que no se puedan guardar pueden figurar en la lista como **No guardado**; para recuperarlos hay que usar el `.msg`/`.eml` original. Los correos importados antes de esta actualización deben reimportarse para detectar los adjuntos.
- Si la importación de `.eml` cae al modo de texto básico por falta de acceso al analizador MIME, no se podrán extraer adjuntos de esa importación. GESMAIL avisa de esta situación.
- El repositorio y la página GitHub Pages son **públicos**. **No subas correos ni datos comerciales reales al repositorio.**
- La aplicación procesa los correos localmente, sin enviarlos a un servidor, pero sigue siendo una demostración, no un entorno de producción.
- La importación de `.msg` convierte el mensaje **en memoria** a MIME sin subir archivos a GitHub. Conserva el HTML original del mensaje cuando existe y las imágenes de firma incrustadas mediante CID dentro de límites de tamaño. Si sólo existe texto o RTF, se muestra una versión de texto. Ahora también se extraen los archivos adjuntos normales y se conservan localmente cuando el navegador dispone de espacio. El conversor admite hasta 15 MB por adjunto y 25 MB de adjuntos por mensaje para evitar bloquear la aplicación.
- Los `.eml` se analizan con `postal-mime` desde `esm.sh` para conservar HTML, firmas, imágenes CID y algunas imágenes de firma referenciadas mediante nombres de archivo locales. Si no se puede cargar el lector MIME, se utiliza una lectura básica de texto (sin formato). El lector de MSG continúa dependiendo del módulo externo indicado abajo.
- La vista HTML se abre en un iframe aislado y sin scripts: no hay edición, respuesta ni envío de correos. Las imágenes remotas se bloquean por defecto; el usuario puede elegir mostrarlas, con el riesgo habitual de que el servidor remoto registre la carga.
- Los mensajes HTML están sujetos a límites de almacenamiento por navegador; para ficheros grandes, algunas imágenes incrustadas podrían omitirse. El visor respeta las tablas y anchos originales, con desplazamiento horizontal cuando hace falta, sin comprimir firmas anchas como las de SIV Electro. No es un clon perfecto del motor Word de Outlook, por lo que ciertos elementos especiales pueden verse distintos.
- El lector `.msg` (`@npeersab/msgreader` v5.1.0) se descarga en el navegador desde `esm.sh` la primera vez que se importa uno: requiere conexión a Internet y acceso a ese dominio. Si la descarga falla, el programa muestra un aviso.
- **Arrastre directo desde Outlook clásico:** abre GESMAIL en **Microsoft Edge o Google Chrome actualizados**, coloca Outlook y el navegador lado a lado y arrastra el correo de la lista de mensajes a **cualquier parte de GESMAIL**. Aparece una capa azul indicando dónde soltarlo. El correo se convierte a EML en memoria y entra por el importador existente, respetando los controles de duplicados y similitud.
- Se admiten varios correos en un mismo arrastre si Outlook y el navegador los entregan. También siguen disponibles el selector de archivos y los `.msg`/`.eml` guardados en el explorador.
- **Para comprobar el comparador:** importa un correo y después arrastra una respuesta `RE:`. Si no aparece la comparación automática, suelta el correo sobre la tarjeta existente. En ambos casos GESMAIL pregunta antes de crear el vínculo.
- Si Outlook o el navegador **no proporcionan el archivo virtual**, JavaScript no puede recuperar mágicamente el MSG. Se muestra un aviso; utiliza Edge/Chrome o guarda el correo temporalmente en el escritorio para importarlo. El arrastre desde Outlook clásico a Firefox no está soportado de forma fiable. El nuevo Outlook tiene limitaciones distintas.
- La importación de `.msg` utiliza los mismos controles de duplicados y similitud que `.eml`; se basa en Message-ID, datos normalizados y contenido extraído. Los porcentajes son orientativos.
- El porcentaje de parecido es una heurística aproximada, no una garantía de duplicación.

El sitio utiliza `index.html` directamente, con HTML, CSS y JavaScript legibles. Los avisos se presentan como tarjetas emergentes centradas desde la función `toast()` del propio tablero (no dependen de CSS añadido después). El lector de MSG continúa en `msg-support.mjs`. Requiere un navegador moderno compatible con módulos JavaScript y `DataTransfer`.
## Cómo probar la nueva ficha y los hilos
1. Abre la web y pulsa **Usuarios del equipo**. Añade los nombres de tus compañeros (uno por línea). Elige tu usuario activo, si quieres usar **Mis solicitudes**.
2. Pulsa **Cargar ejemplos** en un navegador limpio: verás dos RFQ, una de ellas con dos mensajes enlazados y personas en CC.
3. Abre una tarjeta, rellena **Empresa** y selecciona **Propietario de la RFQ**. El remitente, asunto y CC vienen del correo importado.
   También puedes usar **+ Añadir correo al hilo** dentro de la ficha y seleccionar archivos `.msg` o `.eml` aunque no tengan cabeceras de respuesta.
4. Arrastra a cualquier parte de GESMAIL otro correo de Outlook del mismo hilo o importa su `.msg`/`.eml`. **También puedes soltarlo directamente sobre la tarjeta de una RFQ**, para forzar la comparación con esa RFQ cuando el sistema no encuentra coincidencias. Si se detecta una relación, se abrirá una **comparación de dos columnas**: a la izquierda el mensaje anterior y a la derecha el nuevo. Confirma **Sí, añadir al hilo**, elige **No, crear RFQ aparte** o **Descartar correo**. También se puede cambiar entre varias RFQ candidatas. El sistema nunca vincula automáticamente una posible respuesta.
5. En el hilo, pulsa **Leer correo completo** para abrir el mensaje en un visor grande de solo lectura. Puedes moverte entre correos del hilo y ampliar o reducir el zoom sin cambiar ningún dato. Las firmas e imágenes incrustadas aparecerán si los datos están presentes en el `.msg` / `.eml`. Las imágenes alojadas en servidores externos requieren pulsar **Mostrar imágenes externas**.
6. Los correos que se importaron antes de añadir este visor pueden conservar sólo una vista previa de texto o imágenes incompletas. Vuelve a importar su archivo original para mejorar la representación HTML y recuperar imágenes CID o referenciadas por nombre de archivo sin duplicar la RFQ si los identificadores coinciden.
7. En el lector, revisa el bloque **Archivos adjuntos** para ver, descargar o previsualizar los documentos admitidos. Vuelve a importar un correo antiguo para recuperar sus archivos sin crear una RFQ duplicada.
8. Exporta una copia JSON con **Exportar datos** para incluir también los archivos adjuntos almacenados correctamente en este navegador.

**Limitaciones:** importar un solo correo no descarga mágicamente todos los mensajes de Outlook. El hilo muestra únicamente los correos que has incorporado a GESMAIL; si en su cuerpo hay mensajes citados también pueden verse dentro de ese cuerpo. Las RFQ antiguas mantienen sus campos e historial, pero no es posible reconstruir CC, encabezados de respuesta o mensajes anteriores que no fueron guardados. Los usuarios configurados son nombres locales, no cuentas registradas. En particular, la web pública no está preparada para usar datos comerciales confidenciales reales.
