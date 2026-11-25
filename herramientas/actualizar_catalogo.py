#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Golazo Shop · Actualizador del catálogo
=======================================

Convierte el Excel de productos en  data/catalogo.json  (el archivo que lee
index.html), valida las rutas de las fotografías y genera un reporte de
inconsistencias.

Estructura esperada (raíz del sitio, donde está index.html):

    datos/productos.xlsx                      ← el Excel (cualquier nombre .xlsx)
    imagenes/Brazil/AT Mineiro/2176432/1.jpg  ← fotos, misma ruta que el Excel
                                                SIN el prefijo "salida/"

En el Excel la columna carpeta_fotos dice  salida/imagenes/Brazil/AT Mineiro/2176432
y el script la busca en                    imagenes/Brazil/AT Mineiro/2176432
(si no existe, prueba también la ruta tal cual, con "salida/").

Uso:
    python herramientas/actualizar_catalogo.py --optimizar      (recomendado)
    python herramientas/actualizar_catalogo.py                  (fotos originales)
    python herramientas/actualizar_catalogo.py --sin-fotos      (solo valida el Excel)

Opciones:
    --excel RUTA      Excel con columnas: categoria, subcategoria, ruta_categoria,
                      nombre, precio venta, carpeta_fotos. Si se omite, se usa
                      el .xlsx más reciente de la carpeta datos/.
    --carpeta-fotos   Carpeta donde están las fotos (por defecto: imagenes).
    --raiz RUTA       Carpeta raíz del sitio (por defecto: la carpeta actual).
                      Las rutas de carpeta_fotos se resuelven desde aquí.
    --optimizar       Genera versiones web de cada foto (requiere Pillow):
                        img/m/<id>/<n>.webp  -> miniatura (tarjetas), 520 px
                        img/g/<id>/<n>.webp  -> galería / zoom, 1500 px
                      Solo cambia tamaño y compresión: NO recorta ni altera
                      escudos, logos ni diseños. Es incremental (no repite
                      fotos ya convertidas).
    --sin-fotos       No busca fotos en disco (útil para revisar solo el Excel).
    --max-fotos N     Publica como máximo N fotos por producto (0 = todas).
    --ancho-galeria   Ancho máximo de las fotos de galería/zoom (por defecto 1500).
    --ancho-mini      Ancho de las miniaturas de las tarjetas (por defecto 520).
                      Con 0 no se crean miniaturas: las tarjetas usan la misma
                      foto de galería (ocupa casi la mitad de espacio).
    --calidad         Calidad WebP de la galería, 1-100 (por defecto 85).
    --prefijo-fotos URL
                      Opcional. Si las fotos se publican en OTRO sitio (p. ej. un
                      segundo repositorio de GitHub Pages porque superan 1 GB),
                      indique su dirección:  https://usuario.github.io/golazo-fotos/
                      Por defecto se usan rutas relativas al propio sitio.

Dependencias:  pip install openpyxl   (y  pip install pillow  para --optimizar)
"""
import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
import unicodedata

EXT_IMG = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".avif"}
NOMBRES_INVALIDOS = {"", "none", "null", "nan", "service unavailable", "n/a", "#n/a"}
LIMITE_PAGES_MB = 1024  # GitHub Pages: sitios publicados de máximo 1 GB
MINI_W, FULL_W = 520, 1500


# ----------------------------------------------------------------- utilidades
def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "", s.lower())


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def parse_precio(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return int(round(v)) if v > 0 else None
    t = re.sub(r"[^\d,\.]", "", str(v))
    if not t:
        return None
    # formatos 125.000 / 125,000 / 125000.00
    if re.fullmatch(r"\d{1,3}([\.,]\d{3})+", t):
        t = re.sub(r"[\.,]", "", t)
    else:
        t = t.replace(",", ".")
    try:
        n = int(round(float(t)))
        return n if n > 0 else None
    except ValueError:
        return None


def id_seguro(texto):
    t = unicodedata.normalize("NFD", texto).encode("ascii", "ignore").decode()
    t = re.sub(r"[^A-Za-z0-9_-]+", "-", t).strip("-")
    return t or "p"


class Disco:
    """Resuelve rutas respetando MAYÚSCULAS/minúsculas como lo hará GitHub Pages
    (Linux), aunque el equipo local sea Windows/macOS."""

    def __init__(self, raiz):
        self.raiz = os.path.abspath(raiz)
        self.cache = {}

    def listar(self, ruta_abs):
        if ruta_abs not in self.cache:
            try:
                self.cache[ruta_abs] = os.listdir(ruta_abs)
            except OSError:
                self.cache[ruta_abs] = None
        return self.cache[ruta_abs]

    def resolver(self, rel):
        """Devuelve (ruta_real_relativa | None, hubo_diferencia_de_mayusculas).
        Prueba primero la coincidencia exacta y luego variantes de mayúsculas."""
        partes = [p for p in re.split(r"[\\/]+", rel.strip()) if p not in ("", ".")]

        def buscar(actual, i):
            if i == len(partes):
                return ([], False) if os.path.isdir(actual) else None
            hijos = self.listar(actual)
            if hijos is None:
                return None
            p = partes[i]
            cands = ([p] if p in hijos else []) + [h for h in hijos if h != p and
                                                    h.strip().lower() == p.strip().lower()]
            for h in cands:
                r = buscar(os.path.join(actual, h), i + 1)
                if r is not None:
                    return [h] + r[0], r[1] or h != p
            return None

        r = buscar(self.raiz, 0)
        return (None, False) if r is None else ("/".join(r[0]), r[1])


def _convertir(t):
    """Convierte una foto a WebP (en paralelo). Solo cambia tamaño/compresión."""
    src, dst, ancho, q, clave, firma = t
    try:
        from PIL import Image, ImageOps
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with Image.open(src) as im:
            im = ImageOps.exif_transpose(im)
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA" if "transparency" in im.info else "RGB")
            if im.width > ancho:
                im = im.resize((ancho, round(im.height * ancho / im.width)), Image.LANCZOS)
            im.save(dst, "WEBP", quality=q, method=4)
        return clave, firma, None
    except Exception as e:  # foto corrupta o ilegible
        return clave, firma, f"{os.path.basename(src)}: {e}"


# ----------------------------------------------------------------- principal
def main():
    ap = argparse.ArgumentParser(description="Actualiza data/catalogo.json desde el Excel")
    ap.add_argument("--excel")
    ap.add_argument("--carpeta-fotos", default="imagenes")
    ap.add_argument("--raiz", default=".")
    ap.add_argument("--optimizar", action="store_true")
    ap.add_argument("--sin-fotos", action="store_true")
    ap.add_argument("--prefijo-fotos", default="")
    ap.add_argument("--max-fotos", type=int, default=0)
    ap.add_argument("--ancho-galeria", type=int, default=FULL_W)
    ap.add_argument("--ancho-mini", type=int, default=MINI_W)
    ap.add_argument("--calidad", type=int, default=85)
    a = ap.parse_args()

    try:
        import openpyxl
    except ImportError:
        sys.exit("Falta openpyxl:  pip install openpyxl")

    raiz = os.path.abspath(a.raiz)
    if not a.excel:
        cands = [os.path.join(raiz, "datos", f) for f in os.listdir(os.path.join(raiz, "datos"))
                 if f.lower().endswith(".xlsx") and not f.startswith("~$")] if os.path.isdir(os.path.join(raiz, "datos")) else []
        if not cands:
            sys.exit("No encontré el Excel. Póngalo en la carpeta  datos/  o use  --excel RUTA")
        a.excel = max(cands, key=os.path.getmtime)
    print(f"Excel: {a.excel}")
    carpeta_fotos = a.carpeta_fotos.strip("/\\") or "imagenes"
    if not a.sin_fotos and not os.path.isdir(os.path.join(raiz, carpeta_fotos)):
        print(f"⚠ No existe la carpeta '{carpeta_fotos}/' en {raiz}. Los productos quedarán sin fotos.")

    def candidatas(rel):
        """'salida/imagenes/Brazil/X/1' -> ['imagenes/Brazil/X/1', 'salida/imagenes/Brazil/X/1']"""
        rel = rel.replace("\\", "/").strip("/")
        partes = rel.split("/")
        out = []
        bajas = [x.lower() for x in partes]
        if "imagenes" in bajas:
            i = bajas.index("imagenes")
            out.append("/".join([carpeta_fotos] + partes[i + 1:]))
        else:
            out.append("/".join([carpeta_fotos] + partes))
        out.append(rel)
        return list(dict.fromkeys(out))

    wb = openpyxl.load_workbook(a.excel, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    filas = ws.iter_rows(values_only=True)
    enc = [norm(h) for h in next(filas)]
    req = {"categoria": "categoria", "subcategoria": "subcategoria",
           "rutacategoria": "ruta_categoria", "nombre": "nombre",
           "precioventa": "precio venta", "carpetafotos": "carpeta_fotos"}
    col = {}
    for k, visible in req.items():
        if k not in enc:
            sys.exit(f"No encuentro la columna '{visible}' en el Excel. Encabezados: {enc}")
        col[k] = enc.index(k)

    disco = Disco(raiz)
    incid = []  # (fila, tipo, detalle, nombre, carpeta)

    def inc(fila, tipo, detalle, nombre="", carpeta=""):
        incid.append((fila, tipo, detalle, nombre or "", carpeta or ""))

    cats, subs, rutas = [], [], []
    idx_c, idx_s, idx_r = {}, {}, {}

    def ix(lista, mapa, v):
        v = "" if v is None else str(v).strip()
        if v not in mapa:
            mapa[v] = len(lista)
            lista.append(v)
        return mapa[v]

    productos, ids_usados, carpetas_usadas = [], {}, {}
    total_bytes, total_fotos = 0, 0
    pendientes_opt = []
    bases = set()

    for nfila, r in enumerate(filas, start=2):
        if r is None or all(v in (None, "") for v in r):
            continue
        g = lambda k: r[col[k]] if col[k] < len(r) else None
        nombre = str(g("nombre") or "").strip()
        carpeta = str(g("carpetafotos") or "").strip()
        if carpeta.lower() in ("none", "null", "nan"):
            carpeta = ""
        precio_raw = g("precioventa")
        precio = parse_precio(precio_raw)
        cat = str(g("categoria") or "").strip()
        sub = str(g("subcategoria") or "").strip()
        ruta = str(g("rutacategoria") or "").strip()

        if nombre.lower() in NOMBRES_INVALIDOS:
            inc(nfila, "EXCLUIDO · nombre vacío o inválido",
                f"nombre='{nombre or '(vacío)'}', precio='{precio_raw}'", nombre, carpeta)
            continue
        if precio is None:
            inc(nfila, "EXCLUIDO · precio vacío o inválido", f"precio='{precio_raw}'", nombre, carpeta)
            continue
        if not cat:
            inc(nfila, "AVISO · sin categoría", "se publica en 'Sin categoría'", nombre, carpeta)
            cat = "Sin categoría"
        if not sub:
            inc(nfila, "AVISO · sin subcategoría", "", nombre, carpeta)
        if nombre.count("(") != nombre.count(")"):
            inc(nfila, "AVISO · nombre posiblemente truncado",
                "paréntesis sin cerrar; se muestra tal cual", nombre, carpeta)

        # id estable a partir del nombre de la carpeta
        base_id = id_seguro(os.path.basename(carpeta.rstrip("/\\"))) if carpeta else f"f{nfila}"
        pid = base_id
        if pid in ids_usados:
            ids_usados[pid] += 1
            pid = f"{base_id}-{ids_usados[base_id]}"
        else:
            ids_usados[pid] = 1

        if carpeta:
            if carpeta in carpetas_usadas:
                inc(nfila, "AVISO · carpeta compartida",
                    f"la misma carpeta también la usa la fila {carpetas_usadas[carpeta]}", nombre, carpeta)
            else:
                carpetas_usadas[carpeta] = nfila
            bases.add(carpeta.replace("\\", "/").split("/")[0])

        fotos, carpeta_real = [], None
        if not carpeta:
            inc(nfila, "SIN FOTOS · carpeta_fotos vacía", "se muestra con imagen genérica", nombre, "")
        elif not a.sin_fotos:
            carpeta_real, difiere = None, False
            for cand in candidatas(carpeta):
                carpeta_real, difiere = disco.resolver(cand)
                if carpeta_real:
                    break
            if carpeta_real is None:
                inc(nfila, "SIN FOTOS · carpeta no encontrada",
                    f"se esperaba la carpeta '{candidatas(carpeta)[0]}'", nombre, carpeta)
            else:
                if difiere:
                    inc(nfila, "CORREGIDO · mayúsculas/minúsculas",
                        f"en disco es '{carpeta_real}' (GitHub Pages distingue mayúsculas)", nombre, carpeta)
                ab = os.path.join(raiz, carpeta_real)
                archivos = sorted((f for f in (disco.listar(ab) or [])
                                   if os.path.splitext(f)[1].lower() in EXT_IMG
                                   and not f.startswith(".")), key=natural_key)
                if not archivos:
                    inc(nfila, "SIN FOTOS · carpeta vacía", "no hay .jpg/.png/.webp", nombre, carpeta_real)
                for f in archivos:
                    try:
                        tam = os.path.getsize(os.path.join(ab, f))
                    except OSError:
                        tam = 0
                    if tam == 0:
                        inc(nfila, "AVISO · foto vacía (0 bytes)", f, nombre, carpeta_real)
                        continue
                    if re.search(r"[#?%]", f):
                        inc(nfila, "AVISO · nombre de foto con # ? o %",
                            f"'{f}' funciona, pero conviene renombrarla", nombre, carpeta_real)
                    total_bytes += tam
                    fotos.append(f)
                if a.max_fotos > 0:
                    fotos = fotos[:a.max_fotos]
                total_fotos += len(fotos)
                if a.optimizar and fotos:
                    pendientes_opt.append((pid, ab, fotos))

        productos.append([pid, ix(cats, idx_c, cat), ix(subs, idx_s, sub), ix(rutas, idx_r, ruta),
                          nombre, precio, carpeta_real or (carpeta.replace("\\", "/") or None),
                          len(fotos) if a.optimizar else fotos])

    # ------------------------------------------------------- optimización
    modo = "ninguno" if a.sin_fotos else ("optimizado" if a.optimizar else "original")
    if a.optimizar:
        try:
            from PIL import Image, ImageOps
        except ImportError:
            sys.exit("Para --optimizar instale Pillow:  pip install pillow")
        total = sum(len(f) for _, _, f in pendientes_opt)
        man_path = os.path.join(raiz, "img", "manifiesto.json")
        try:
            with open(man_path, encoding="utf-8") as fh:
                manifiesto = json.load(fh)
        except (OSError, ValueError):
            manifiesto = {}
        tareas, destinos = [], []
        for pid, ab, fotos in pendientes_opt:
            for i, f in enumerate(fotos, start=1):
                src = os.path.join(ab, f)
                for carpeta_dest, ancho, q in (("m", a.ancho_mini, 75), ("g", a.ancho_galeria, a.calidad)):
                    if ancho <= 0:
                        continue
                    dst = os.path.join(raiz, "img", carpeta_dest, pid, f"{i}.webp")
                    firma = f"{os.path.relpath(src, raiz)}|{os.path.getsize(src)}|{ancho}|{q}"
                    clave = f"{carpeta_dest}/{pid}/{i}"
                    destinos.append(dst)
                    if not (os.path.exists(dst) and manifiesto.get(clave) == firma):
                        tareas.append((src, dst, ancho, q, clave, firma))
        print(f"Optimizando: {total:,} fotos, {len(tareas):,} versiones web por generar…", flush=True)

        def guardar_manifiesto():
            os.makedirs(os.path.dirname(man_path), exist_ok=True)
            tmp = man_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(manifiesto, fh, separators=(",", ":"))
            os.replace(tmp, man_path)

        from concurrent.futures import ProcessPoolExecutor
        hechos = 0
        with ProcessPoolExecutor(max_workers=max(1, (os.cpu_count() or 2))) as ex:
            for clave, firma, err in ex.map(_convertir, tareas, chunksize=16):
                hechos += 1
                if err:
                    inc("-", "AVISO · foto no se pudo convertir", err, clave, "")
                else:
                    manifiesto[clave] = firma
                if hechos % 500 == 0:
                    guardar_manifiesto()
                    print(f"  {hechos:,}/{len(tareas):,}", flush=True)
        guardar_manifiesto()
        opt_bytes = sum(os.path.getsize(d) for d in destinos if os.path.exists(d))
        print(f"  {hechos:,} versiones web nuevas generadas", flush=True)
        total_bytes = opt_bytes

    # ------------------------------------------------------- salida JSON
    os.makedirs(os.path.join(raiz, "data"), exist_ok=True)
    pref = a.prefijo_fotos.strip()
    if pref and not pref.endswith("/"):
        pref += "/"
    out = {
        "version": 2,
        "generado": dt.datetime.now().isoformat(timespec="seconds"),
        "total": len(productos),
        "img": {"modo": modo, "base": pref, "mini": pref + ("img/m/" if a.ancho_mini > 0 else "img/g/"), "full": pref + "img/g/"},
        "cat": cats, "sub": subs, "ruta": rutas,
        "campos": ["id", "cat", "sub", "ruta", "nombre", "precio", "carpeta", "fotos"],
        "p": productos,
    }
    destino = os.path.join(raiz, "data", "catalogo.json")
    with open(destino, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(",", ":"))

    # ------------------------------------------------------- reportes
    rep_dir = os.path.join(raiz, "reportes")
    os.makedirs(rep_dir, exist_ok=True)
    with open(os.path.join(rep_dir, "inconsistencias.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["fila_excel", "tipo", "detalle", "nombre", "carpeta"])
        w.writerows(incid)

    from collections import Counter
    resumen = Counter(t for _, t, *_ in incid)
    sin_fotos = sum(1 for p in productos if not p[7])
    mb = total_bytes / 1048576
    lineas = [
        "# Reporte de validación · Golazo Shop", "",
        f"Generado: {out['generado']}  ·  Excel: `{os.path.basename(a.excel)}`  ·  Modo de fotos: **{modo}**", "",
        "| Indicador | Valor |", "|---|---|",
        f"| Productos publicados | {len(productos):,} |",
        f"| Productos sin fotos | {sin_fotos:,}{' (no se buscaron fotos: --sin-fotos)' if a.sin_fotos else ''} |",
        f"| Fotos encontradas | {total_fotos:,} |",
        f"| Peso de fotos a publicar | {mb:,.0f} MB |",
        f"| Categorías / subcategorías / rutas | {len(cats)} / {len(subs)} / {len(rutas)} |", "",
    ]
    no_enc = resumen.get("SIN FOTOS · carpeta no encontrada", 0)
    if not a.sin_fotos and no_enc > len(productos) * 0.5:
        lineas += [f"> ⚠️ **{no_enc:,} carpetas no se encontraron.** ¿Ejecutó el script desde la "
                   f"carpeta raíz del sitio? Raíz usada: `{raiz}`. Debe existir `<raíz>/salida/imagenes/`.", ""]
    if not a.sin_fotos and mb > LIMITE_PAGES_MB * 0.85:
        lineas += [f"> ⚠️ **Las fotos pesan {mb:,.0f} MB.** GitHub Pages admite sitios de hasta "
                   f"1 GB. Ejecute con `--optimizar` y publique solo la carpeta `img/` "
                   "(vea el README).", ""]
    lineas += ["## Incidencias por tipo", "", "| Tipo | Cantidad |", "|---|---|"]
    lineas += [f"| {t} | {n:,} |" for t, n in sorted(resumen.items())] or ["| (ninguna) | 0 |"]
    lineas += ["", "El detalle fila por fila está en `reportes/inconsistencias.csv` "
               "(se abre en Excel; separador `;`).", ""]
    with open(os.path.join(rep_dir, "reporte_validacion.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lineas))

    print(f"✔ {destino}  ({os.path.getsize(destino)/1024:,.0f} KB, {len(productos):,} productos)")
    print(f"✔ reportes/reporte_validacion.md  ·  {len(incid):,} incidencias")
    for t, n in sorted(resumen.items()):
        print(f"   - {t}: {n:,}")
    if not a.sin_fotos:
        print(f"   Fotos: {total_fotos:,}  ·  {mb:,.0f} MB  ·  productos sin fotos: {sin_fotos:,}")


if __name__ == "__main__":
    main()
