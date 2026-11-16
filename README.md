# GESMAIL — RFQ Equipo (demo GitHub Pages)

Demo web del tablero de solicitudes de oferta (RFQ) del departamento comercial.

## Funcionalidades
- Cuatro columnas: **Petición**, **Oferta enviada**, **Oferta perdida** y **Pedido ganado**.
- Importación de correos en formato **`.eml` y `.msg` de Outlook** (archivos guardados o arrastrados desde el explorador).
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
- **Arrastrar un archivo `.msg` desde el explorador** funciona cuando el navegador entrega un archivo. El arrastre directo desde Outlook al navegador depende de la versión de Outlook y no está garantizado.
- La importación de `.msg` utiliza los mismos controles de duplicados y similitud que `.eml`; se basa en Message-ID, datos normalizados y contenido extraído. Los porcentajes son orientativos.
- El porcentaje de parecido es una heurística aproximada, no una garantía de duplicación.

El sitio carga un paquete HTML comprimido dentro del propio `index.html` y el módulo `msg-support.mjs` para MSG. Requiere un navegador moderno compatible con `DecompressionStream`, módulos JavaScript y `DataTransfer`.