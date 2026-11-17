# GESMAIL — RFQ Equipo (demo GitHub Pages)

Demo web del tablero de solicitudes de oferta (RFQ) del departamento comercial.

## Funcionalidades
- Cuatro columnas: **Petición**, **Oferta enviada**, **Oferta perdida** y **Pedido ganado**.
- Importación de correos en formato **`.eml` y `.msg` de Outlook**, también mediante arrastre directo desde **Outlook clásico de Windows** a Chrome o Edge cuando el navegador proporciona el fichero virtual.
- Aviso y bloqueo de duplicados exactos (Message-ID, huella o fichero idéntico).
- Avisos con porcentaje estimado de similitud para correos parecidos.
- En cada RFQ, **asunto**, **remitente**, **destinatarios** y **personas en copia (CC)** se obtienen automáticamente al importar un correo; **empresa** y **propietario/responsable** se asignan manualmente.
- **Hilo de correos** dentro de cada ficha: mensajes con fecha, remitente, To, CC y vista previa del cuerpo. Las respuestas que contienen `In-Reply-To` o `References` se enlazan automáticamente mediante `Message-ID` a un mensaje ya importado. En avisos de similitud también se ofrece **Añadir al hilo** manualmente.
- **Usuarios del equipo** configurables desde el menú lateral, guardados **sólo en ese navegador** (no hay todavía registro, autenticación ni permisos reales).
- Creación manual, edición y cambio de estado de RFQ.
- Búsqueda, exportación y restauración de copias JSON.
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
- El repositorio y la página GitHub Pages son **públicos**. **No subas correos ni datos comerciales reales al repositorio.**
- La aplicación procesa los correos localmente, sin enviarlos a un servidor, pero sigue siendo una demostración, no un entorno de producción.
- La importación de `.msg` convierte el mensaje **en memoria** al formato que entiende el tablero, sin subir los archivos a GitHub. Extrae asunto, remitente, Para, CC, encabezados de respuesta y cuerpo de texto (HTML o RTF como alternativa). No se guardan los archivos ni sus adjuntos.
- El lector `.msg` (`@npeersab/msgreader` v5.1.0) se descarga en el navegador desde `esm.sh` la primera vez que se importa uno: requiere conexión a Internet y acceso a ese dominio. Si la descarga falla, el programa muestra un aviso.
- **Arrastre directo desde Outlook clásico:** abre GESMAIL en **Microsoft Edge o Google Chrome actualizados**, coloca Outlook y el navegador lado a lado y arrastra el correo de la lista de mensajes a **cualquier parte de GESMAIL**. Aparece una capa azul indicando dónde soltarlo. El correo se convierte a EML en memoria y entra por el importador existente, respetando los controles de duplicados y similitud.
- Se admiten varios correos en un mismo arrastre si Outlook y el navegador los entregan. También siguen disponibles el selector de archivos y los `.msg`/`.eml` guardados en el explorador.
- Si Outlook o el navegador **no proporcionan el archivo virtual**, JavaScript no puede recuperar mágicamente el MSG. Se muestra un aviso; utiliza Edge/Chrome o guarda el correo temporalmente en el escritorio para importarlo. El arrastre desde Outlook clásico a Firefox no está soportado de forma fiable. El nuevo Outlook tiene limitaciones distintas.
- La importación de `.msg` utiliza los mismos controles de duplicados y similitud que `.eml`; se basa en Message-ID, datos normalizados y contenido extraído. Los porcentajes son orientativos.
- El porcentaje de parecido es una heurística aproximada, no una garantía de duplicación.

El sitio utiliza `index.html` directamente, con HTML, CSS y JavaScript legibles. Los avisos se presentan como tarjetas emergentes centradas desde la función `toast()` del propio tablero (no dependen de CSS añadido después). El lector de MSG continúa en `msg-support.mjs`. Requiere un navegador moderno compatible con módulos JavaScript y `DataTransfer`.
## Cómo probar la nueva ficha y los hilos
1. Abre la web y pulsa **Usuarios del equipo**. Añade los nombres de tus compañeros (uno por línea). Elige tu usuario activo, si quieres usar **Mis solicitudes**.
2. Pulsa **Cargar ejemplos** en un navegador limpio: verás dos RFQ, una de ellas con dos mensajes enlazados y personas en CC.
3. Abre una tarjeta, rellena **Empresa** y selecciona **Propietario de la RFQ**. El remitente, asunto y CC vienen del correo importado.
4. Importa después una respuesta relacionada. Si incluye las cabeceras de respuesta, se añadirá al hilo automáticamente. Si no las incluye pero se parece a una RFQ existente, puedes seleccionar **Añadir al hilo de esta RFQ** en la ventana de similitud.
5. Exporta una copia JSON con **Exportar datos** si quieres conservarla fuera de este navegador.

**Limitaciones:** importar un solo correo no descarga mágicamente todos los mensajes de Outlook. El hilo muestra únicamente los correos que has incorporado a GESMAIL; si en su cuerpo hay mensajes citados también pueden verse dentro de ese cuerpo. Las RFQ antiguas mantienen sus campos e historial, pero no es posible reconstruir CC, encabezados de respuesta o mensajes anteriores que no fueron guardados. Los usuarios configurados son nombres locales, no cuentas registradas. En particular, la web pública no está preparada para usar datos comerciales confidenciales reales.
