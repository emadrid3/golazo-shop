# Golazo Shop · Catálogo digital

Catálogo informativo de **Golazo Shop · Pasión en cada camiseta**, para publicar gratis en GitHub Pages.
No es tienda en línea: no tiene carrito, pagos ni checkout. Los pedidos se hacen por Instagram **@golazo.oficial21**.

---

## 1. Qué hay en esta carpeta

```
golazo-shop/
├── index.html                  ← todo el sitio (diseño, buscador, filtros, detalle…)
├── datos/
│   └── productos.xlsx          ← SU EXCEL (reemplácelo cuando tenga uno nuevo)
├── imagenes/                   ← SUS FOTOS (ver punto 2)  ★
├── assets/                     ← logo, ícono y la imagen que sale al compartir el enlace
├── data/catalogo.json          ← se genera solo a partir del Excel + fotos (no se edita)
├── herramientas/actualizar_catalogo.py   ← el programa que genera el catálogo
├── ACTUALIZAR_CATALOGO.bat     ← Windows: doble clic para actualizar
├── actualizar_catalogo.command ← Mac: doble clic para actualizar
├── VER_SITIO_LOCAL.bat         ← Windows: ver el sitio en su computador
└── .github/workflows/publicar.yml  ← publicación automática en GitHub Pages
```

---

## 2. Cómo agregar las fotos de los productos ★

### La regla: misma ruta del Excel, sin `salida/`

En el Excel, la columna `carpeta_fotos` de cada producto dice dónde están sus fotos, por ejemplo:

| nombre | carpeta_fotos |
|---|---|
| 1971 AT Mineiro Home Retro Soccer Jersey | `salida/imagenes/Brazil/AT Mineiro/2176432` |
| 1983 AT Mineiro Home Retro Soccer Jersey | `salida/imagenes/Brazil/AT Mineiro/2182322` |

En el sitio, esas fotos van en la carpeta **`imagenes`**, con la misma ruta pero **quitando `salida/` del comienzo**:

```
golazo-shop/
└── imagenes/
    └── Brazil/
        └── AT Mineiro/
            ├── 2176432/          ← fotos de "1971 AT Mineiro Home Retro…"
            │   ├── 1.jpg         ← foto principal (la primera en orden)
            │   ├── 2.jpg         ← galería
            │   └── 3.jpg
            └── 2182322/          ← fotos de "1983 AT Mineiro Home Retro…"
                ├── 1.jpg
                └── 2.jpg
```

- Cada producto tiene **su propia carpeta**; dentro van **todas** sus fotos.
- Los nombres de las fotos pueden ser cualquiera (`1.jpg`, `foto_2.png`, `abc.webp`…). Se ordenan de forma natural (1, 2, 3 … 10) y **la primera es la foto principal** de la tarjeta.
- Formatos aceptados: `.jpg .jpeg .png .webp .gif .avif`.
- Respete mayúsculas, espacios y tildes de las carpetas (`AT Mineiro` ≠ `at mineiro`). Si hay diferencias, el programa las detecta y avisa.

### Lo más rápido

Si ya tiene la carpeta `salida/imagenes` que generó con su Excel:
**abra `salida/imagenes`, seleccione todo lo que hay dentro, cópielo y péguelo dentro de `golazo-shop/imagenes`.** Listo, no tiene que renombrar nada.

> ¿Por qué no basta con pegar las fotos? Una página en GitHub Pages no puede "ver" qué archivos hay dentro de una carpeta. Por eso el programa `actualizar_catalogo.py` recorre `imagenes/`, anota el nombre exacto de cada foto en `data/catalogo.json` y el sitio las carga desde ahí. Cada vez que agregue o cambie fotos o el Excel, hay que volver a generar ese archivo (automático con la opción A, o con doble clic con la opción B).

---

## 3. Publicar en GitHub Pages

Necesita una cuenta gratuita de GitHub y, para subir miles de fotos, **GitHub Desktop** (https://desktop.github.com). El cargador de archivos de la página web de GitHub solo acepta 100 archivos por vez.

### Opción A · Automática (recomendada, no necesita Python)

1. En GitHub Desktop: **File → Add local repository** → elija la carpeta `golazo-shop` → *create a repository* → **Publish repository** (desmarque *Keep this code private*).
2. En github.com, en su repositorio: **Settings → Pages → Build and deployment → Source: `GitHub Actions`**.
3. Vaya a la pestaña **Actions**: verá "Publicar catálogo" trabajando. Cuando termine (la primera vez puede tardar bastante porque optimiza todas las fotos; después solo procesa las nuevas), el sitio queda en:
   `https://SU-USUARIO.github.io/golazo-shop/`
4. **Para actualizar**: cambie el Excel en `datos/` y/o agregue fotos en `imagenes/` → en GitHub Desktop escriba un mensaje → **Commit to main → Push origin**. Se publica solo.
5. El reporte de inconsistencias se descarga desde Actions → la ejecución → *Artifacts* → `reporte-validacion`.

En esta opción GitHub convierte las fotos a una versión web liviana (WebP, sin recortar ni alterar escudos, logos ni diseños) y publica **solo esas versiones**, para no pasar el límite de 1 GB de GitHub Pages.

### Opción B · Desde su computador

1. Instale Python (https://www.python.org/downloads/, marque **Add Python to PATH**).
2. Doble clic en **`ACTUALIZAR_CATALOGO.bat`** (Mac: `actualizar_catalogo.command`). Genera `data/catalogo.json`, las fotos web en `img/` y los reportes en `reportes/`.
3. Doble clic en **`VER_SITIO_LOCAL.bat`** para revisarlo en `http://localhost:8000`.
4. Para publicar sin la Opción A: borre la línea `img/` del archivo `.gitignore`, suba todo con GitHub Desktop y en **Settings → Pages** elija *Deploy from a branch → main / (root)*.

> Abrir `index.html` con doble clic no funciona (el navegador bloquea la lectura del catálogo); use siempre `VER_SITIO_LOCAL.bat` o el sitio publicado.

---

## 4. Límites gratuitos de GitHub

- Sitio publicado: máximo **1 GB** (por eso se publican las fotos optimizadas). El reporte muestra cuánto pesan.
- Repositorio: se recomienda menos de **5 GB**, archivos de menos de 100 MB y subir en tandas de menos de 2 GB (si son muchas fotos, haga varios *commits*, por ejemplo una liga a la vez).
- Si sus fotos originales son demasiadas para el repositorio, use la Opción B (las originales se quedan en su computador y solo sube `img/`).

---

## 5. Revisar el reporte de validación

`reportes/reporte_validacion.md` y `reportes/inconsistencias.csv` (abre en Excel, separador `;`) indican fila por fila:

- **EXCLUIDO**: filas sin nombre válido o sin precio (no se publican).
- **SIN FOTOS · carpeta no encontrada**: muestra la carpeta exacta que se esperaba dentro de `imagenes/`.
- **SIN FOTOS · carpeta vacía**: la carpeta existe pero no tiene fotos.
- **CORREGIDO · mayúsculas/minúsculas**: la carpeta se encontró con otra escritura.
- **AVISO**: nombres posiblemente cortados, carpetas compartidas por dos productos, etc.

Los productos sin fotos se muestran igual, con una imagen genérica de "Foto no disponible".

---

## 6. Personalizar

- **Logo**: `assets/logo.png`/`logo.webp` (principal), `assets/logo-texto.png` (encabezado), `assets/favicon.png`, `assets/og.jpg` (al compartir el enlace).
- **Colores**: al comienzo del `<style>` en `index.html` (*1. IDENTIDAD VISUAL*): `--gold`, `--grad-gold`, `--ink`, `--paper`…
- **Instagram y redes**: objeto `CONFIG` al comienzo del `<script>` (Facebook, TikTok o WhatsApp se muestran si los llena).
- **Alias del buscador**: diccionarios `EQUIPOS`, `PAISES`, `SECCIONES`, `TIPOS`, `PRENDAS` en `index.html`.

## 7. Funciones del catálogo

Buscador inteligente (español/inglés, abreviaturas, sin tildes, corrige errores: `barcleona`, `holanda 88`, `galo`, `arquero`), filtros por categoría, sección, subcategoría, equipo, prenda, tipo de camiseta, versión, público, año, temporada y precio; paginación de 24/48/96; orden por relevancia, novedades, precio, nombre y año; detalle con galería y zoom; enlace propio por producto (`?p=ID`); favoritos sin registro; y botón **Pedir por Instagram**, que copia nombre, precio y enlace del producto para pegarlo en el chat (Instagram no permite pre-escribir mensajes).
