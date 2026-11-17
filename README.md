# GESMAIL — RFQ Equipo (demo GitHub Pages)

Demo web del tablero de solicitudes de oferta (RFQ) del departamento comercial.

## Funcionalidades
- Cuatro columnas: **Petición**, **Oferta enviada**, **Oferta perdida** y **Pedido ganado**.
- Importación de correos en formato **`.eml` y `.msg` de Outlook**, también mediante arrastre directo desde **Outlook clásico de Windows** a Chrome o Edge cuando el navegador proporciona el fichero virtual.
- Aviso y bloqueo de duplicados exactos (Message-ID, huella o fichero idéntico).
- Avisos con porcentaje estimado de similitud para correos parecidos.
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
- La importación de `.msg` convierte el mensaje **en memoria** al formato que entiende el tablero, sin subir los archivos a GitHub. Extrae asunto, remitente y cuerpo de texto, HTML o RTF como alternativa. No se guardan los archivos ni sus adjuntos.
- El lector `.msg` (`@npeersab/msgreader` v5.1.0) se descarga en el navegador desde `esm.sh` la primera vez que se importa uno: requiere conexión a Internet y acceso a ese dominio. Si la descarga falla, el programa muestra un aviso.
- **Arrastre directo desde Outlook clásico:** abre GESMAIL en **Microsoft Edge o Google Chrome actualizados**, coloca Outlook y el navegador lado a lado y arrastra el correo de la lista de mensajes a **cualquier parte de GESMAIL**. Aparece una capa azul indicando dónde soltarlo. El correo se convierte a EML en memoria y entra por el importador existente, respetando los controles de duplicados y similitud.
- Se admiten varios correos en un mismo arrastre si Outlook y el navegador los entregan. También siguen disponibles el selector de archivos y los `.msg`/`.eml` guardados en el explorador.
- Si Outlook o el navegador **no proporcionan el archivo virtual**, JavaScript no puede recuperar mágicamente el MSG. Se muestra un aviso; utiliza Edge/Chrome o guarda el correo temporalmente en el escritorio para importarlo. El arrastre desde Outlook clásico a Firefox no está soportado de forma fiable. El nuevo Outlook tiene limitaciones distintas.
- La importación de `.msg` utiliza los mismos controles de duplicados y similitud que `.eml`; se basa en Message-ID, datos normalizados y contenido extraído. Los porcentajes son orientativos.
- El porcentaje de parecido es una heurística aproximada, no una garantía de duplicación.

El sitio utiliza `index.html` directamente, con HTML, CSS y JavaScript legibles. Los avisos se presentan como tarjetas emergentes centradas desde la función `toast()` del propio tablero (no dependen de CSS añadido después). El lector de MSG continúa en `msg-support.mjs`. Requiere un navegador moderno compatible con módulos JavaScript y `DataTransfer`.