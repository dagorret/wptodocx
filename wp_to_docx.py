#!/usr/bin/env python3
"""
wp_to_docx.py — Convierte un artículo de WordPress (con fórmulas del plugin
KaTeX) a un .docx con ecuaciones nativas de Word (OMML) e imágenes incrustadas.

Funciona con tres tipos de entrada, indistintamente:
  1. Una URL publicada            (ej: https://dagorret.com.ar/mi-articulo/)
  2. El HTML completo descargado  (lo que "Guardar como" del navegador produce)
  3. El código fuente en bloques  (el HTML que pegás en el editor de Gutenberg)

Requiere: pandoc instalado en el sistema (`apt install pandoc` /
`brew install pandoc`), y las librerías Python `beautifulsoup4` y `requests`
(`pip install beautifulsoup4 requests`).

USO:
    python wp_to_docx.py entrada.html salida.docx
    python wp_to_docx.py "https://dagorret.com.ar/mi-articulo/" salida.docx

CÓMO FUNCIONA:
  El plugin de KaTeX que usás deja el LaTeX crudo dentro de:
    - <span class="katex-eq" data-katex-display="false">...</span>  (inline)
    - <div class="wp-block-katex-display-block katex-eq">
        <pre>...</pre>
      </div>                                                        (bloque)
  y lo renderiza recién en el navegador, vía JavaScript. Pandoc no ejecuta
  JavaScript, así que si le diéramos el HTML tal cual, esas fórmulas
  quedarían como texto plano en el .docx.

  Este script las reescribe ANTES de llamar a pandoc:
    - inline  -> $...$
    - bloque  -> $$...$$   (como párrafo propio)
  y le pide a pandoc que lea con la extensión `tex_math_dollars`, que
  convierte ese LaTeX en ecuaciones nativas de Word (editables con el editor
  de ecuaciones de Word, no imágenes).

  Las imágenes remotas (si las hay) se descargan a una carpeta `media/`
  junto al .docx y se re-referencian con ruta local, para que pandoc las
  incruste en vez de dejar un link roto.
"""

import argparse
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.parse import urljoin, urlparse

from bs4 import BeautifulSoup

try:
    import requests
except ImportError:
    requests = None


# Selectores candidatos para el contenedor del artículo, en orden de
# preferencia. Si tu tema de WordPress usa otro, agregalo acá.
CONTENT_SELECTORS = [
    ".wp-block-post-content",
    ".entry-content",
    "article",
    "main",
]


def load_html(source: str) -> tuple[str, str | None]:
    """Devuelve (html, base_url). base_url es None si la fuente es un archivo local."""
    parsed = urlparse(source)
    if parsed.scheme in ("http", "https"):
        if requests is None:
            sys.exit("Falta 'requests'. Instalá con: pip install requests")
        resp = requests.get(source, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        return resp.text, source
    path = Path(source)
    if not path.exists():
        sys.exit(f"No encuentro el archivo: {source}")
    return path.read_text(encoding="utf-8"), None


def extract_content(soup: BeautifulSoup) -> BeautifulSoup:
    """Aísla el contenedor del artículo. Si no encuentra ninguno de los
    selectores conocidos, asume que el HTML ya es solo el contenido
    (por ejemplo, si pegaste el código fuente en bloques de Gutenberg)."""
    for sel in CONTENT_SELECTORS:
        found = soup.select_one(sel)
        if found is not None:
            return found
    return soup


def convert_katex(content) -> int:
    """Reemplaza los spans/divs de KaTeX por delimitadores $ / $$ que pandoc
    entiende. Devuelve la cantidad de fórmulas convertidas."""
    count = 0

    # Bloques de ecuación (una ecuación centrada, en su propio párrafo)
    for block in content.select("div.katex-eq, div[data-katex-display='true']"):
        pre = block.find("pre")
        latex = (pre.get_text() if pre else block.get_text()).strip()
        soup_root = block
        while soup_root.parent is not None:
            soup_root = soup_root.parent
        new_p = soup_root.new_tag("p")
        new_p.string = f"$${latex}$$"
        block.replace_with(new_p)
        count += 1

    # Fórmulas inline (dentro de una oración)
    for span in content.select("span.katex-eq"):
        latex = span.get_text().strip()
        span.replace_with(f"${latex}$")
        count += 1

    return count


def download_images(content, base_url: str | None, media_dir: Path) -> int:
    """Descarga imágenes remotas a media_dir y reescribe el src a ruta local.
    No hace nada si no hay base_url (fuente local) o si requests no está
    disponible; en ese caso pandoc igual intentará resolver URLs absolutas
    por su cuenta al convertir."""
    if requests is None:
        return 0

    count = 0
    media_dir.mkdir(parents=True, exist_ok=True)
    for i, img in enumerate(content.select("img[src]")):
        src = img["src"]
        abs_url = urljoin(base_url, src) if base_url else src
        if not abs_url.startswith(("http://", "https://")):
            continue  # ya es local, no hay nada que descargar
        try:
            resp = requests.get(abs_url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            resp.raise_for_status()
        except Exception as e:
            print(f"  ! No pude descargar {abs_url}: {e}", file=sys.stderr)
            continue

        ext = Path(urlparse(abs_url).path).suffix or ".jpg"
        local_name = f"img_{i:03d}{ext}"
        (media_dir / local_name).write_bytes(resp.content)
        img["src"] = f"{media_dir.name}/{local_name}"
        count += 1
    return count


def convert(source: str, output: str) -> None:
    output_path = Path(output).resolve()
    media_dir = output_path.parent / f"{output_path.stem}_media"

    html, base_url = load_html(source)
    soup = BeautifulSoup(html, "html.parser")
    content = extract_content(soup)

    n_formulas = convert_katex(content)
    n_images = download_images(content, base_url, media_dir)

    print(f"  {n_formulas} fórmula(s) convertida(s) a notación $...$/$$...$$")
    print(f"  {n_images} imagen(es) descargada(s) a {media_dir.name}/")

    with tempfile.NamedTemporaryFile(
        "w", suffix=".html", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write(str(content))
        tmp_path = tmp.name

    cmd = [
        "pandoc",
        "-f", "html+tex_math_dollars",
        "-t", "docx",
        "--resource-path", f".:{media_dir}",
        "-o", str(output_path),
        tmp_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    Path(tmp_path).unlink(missing_ok=True)

    if result.returncode != 0:
        sys.exit(f"pandoc falló:\n{result.stderr}")

    print(f"  Listo -> {output_path}")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("entrada", help="URL, o ruta a un archivo HTML local")
    parser.add_argument("salida", help="Ruta del .docx a generar")
    args = parser.parse_args()
    convert(args.entrada, args.salida)


if __name__ == "__main__":
    main()
