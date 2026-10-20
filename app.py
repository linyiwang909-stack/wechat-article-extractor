from __future__ import annotations

import io
import os
import re
import tempfile
import urllib.parse
import urllib.request
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
from lxml import html
from PIL import Image


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)
ARTICLE_CACHE: dict[str, dict] = {}

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "Chrome/131.0 Safari/537.36"
)


def fetch_page(url: str) -> str:
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in {"http", "https"} or "mp.weixin.qq.com" not in parsed.netloc:
        raise ValueError("目前仅支持 mp.weixin.qq.com 公众号文章链接")
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=25) as response:
        return response.read().decode("utf-8", errors="ignore")


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def parse_article(source: str, url: str) -> dict:
    tree = html.fromstring(source)
    title = clean_text(tree.xpath("string(//*[@id='activity-name'])"))
    author = clean_text(tree.xpath("string(//*[@id='js_name'])"))
    meta_text = clean_text(tree.xpath("string(//*[@id='meta_content'])"))
    date_match = re.search(r"\d{4}年\d{1,2}月\d{1,2}日(?:\s+\d{1,2}:\d{2})?", meta_text)
    date = date_match.group(0) if date_match else ""
    content_nodes = tree.xpath("//*[@id='js_content']")
    if not content_nodes:
        raise ValueError("没有找到文章正文，可能链接已失效或页面需要登录")
    root = content_nodes[0]
    blocks = []
    paragraph_runs = []
    paragraph_align = "left"
    block_tags = {"p", "section", "h1", "h2", "h3", "blockquote", "li"}

    def css(node, key: str) -> str:
        match = re.search(rf"(?:^|;)\s*{re.escape(key)}\s*:\s*([^;]+)", node.get("style", ""), re.I)
        return match.group(1).strip() if match else ""

    def flush_paragraph() -> None:
        nonlocal paragraph_runs
        if paragraph_runs:
            blocks.append({"type": "paragraph", "runs": paragraph_runs, "align": paragraph_align})
            paragraph_runs = []

    def walk(node, inherited=None) -> None:
        nonlocal paragraph_align
        inherited = dict(inherited or {})
        tag = (node.tag or "").lower()
        if tag in block_tags:
            flush_paragraph()
            paragraph_align = css(node, "text-align") or "left"
        if tag in {"strong", "b"}:
            inherited["bold"] = True
        if tag in {"em", "i"}:
            inherited["italic"] = True
        color = css(node, "color")
        size = css(node, "font-size")
        if color:
            inherited["color"] = color
        if size:
            inherited["size"] = size
        if tag == "img":
            flush_paragraph()
            src = node.get("data-src") or node.get("data-original") or node.get("src")
            if src and not src.startswith("data:"):
                blocks.append({"type": "image", "src": urllib.parse.urljoin(url, src), "alt": node.get("alt", "")})
            return
        if node.text and clean_text(node.text):
            paragraph_runs.append({"text": clean_text(node.text), **inherited})
        for child in node:
            if child.tag == "br":
                flush_paragraph()
            else:
                walk(child, inherited)
            if child.tail and clean_text(child.tail):
                paragraph_runs.append({"text": clean_text(child.tail), **inherited})
        if tag in block_tags:
            flush_paragraph()

    walk(root)
    flush_paragraph()
    return {"title": title or "微信公众号文章", "author": author, "date": date, "url": url, "blocks": blocks}


def download_image(url: str, folder: Path) -> Path | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Referer": "https://mp.weixin.qq.com/"})
        with urllib.request.urlopen(req, timeout=20) as response:
            payload = response.read()
        suffix = ".jpg"
        content_type = response.headers.get_content_type()
        if content_type == "image/png":
            suffix = ".png"
        elif content_type == "image/gif":
            suffix = ".gif"
        elif content_type == "image/webp":
            # python-docx cannot insert WebP directly; normalize it to PNG.
            image = Image.open(io.BytesIO(payload)).convert("RGB")
            path = folder / f"image-{len(list(folder.iterdir())) + 1}.png"
            image.save(path, format="PNG")
            return path
        path = folder / f"image-{len(list(folder.iterdir())) + 1}{suffix}"
        path.write_bytes(payload)
        return path
    except Exception:
        return None


def build_docx(article: dict, output_path: Path) -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(9.5)
    normal.paragraph_format.space_after = Pt(7)
    normal.paragraph_format.line_spacing = 1.35
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    title_style = doc.styles["Title"]
    title_style.font.name = "Microsoft YaHei"
    title_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(article["title"])
    run.font.name = "Microsoft YaHei"
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.color.rgb = RGBColor(0, 0, 0)
    with tempfile.TemporaryDirectory(prefix="wechat-images-") as temp_dir:
        image_dir = Path(temp_dir)
        for block in article["blocks"]:
            if block["type"] in {"text", "paragraph"}:
                p = doc.add_paragraph()
                align = block.get("align", "left")
                p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
                p.paragraph_format.first_line_indent = Pt(21)
                runs = block.get("runs") or [{"text": block.get("text", "")}]
                for run_data in runs:
                    run = p.add_run(run_data.get("text", ""))
                    run.font.name = "Microsoft YaHei"
                    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
                    run.bold = bool(run_data.get("bold"))
                    run.italic = bool(run_data.get("italic"))
                    size = run_data.get("size", "")
                    if size:
                        match = re.search(r"[0-9.]+", size)
                        if match:
                            run.font.size = Pt(float(match.group(0)))
                    color = run_data.get("color", "")
                    if color.startswith("#") and len(color) in {4, 7}:
                        value = color[1:]
                        if len(value) == 3:
                            value = "".join(ch * 2 for ch in value)
                        run.font.color.rgb = RGBColor.from_string(value.upper())
            else:
                image_path = download_image(block["src"], image_dir)
                if image_path:
                    try:
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run().add_picture(str(image_path), width=Inches(5.8))
                    except Exception:
                        # A malformed or unsupported image should not abort the document.
                        if p._element.getparent() is not None:
                            p._element.getparent().remove(p._element)
                        if block.get("alt"):
                            alt_paragraph = doc.add_paragraph()
                            alt_run = alt_paragraph.add_run(f"[图片：{block['alt']}]")
                            alt_run.font.name = "Microsoft YaHei"
                            alt_run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
                elif block.get("alt"):
                    alt_paragraph = doc.add_paragraph()
                    alt_run = alt_paragraph.add_run(f"[图片：{block['alt']}]")
                    alt_run.font.name = "Microsoft YaHei"
                    alt_run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    doc.save(output_path)


def main() -> None:
    import json
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    class Handler(BaseHTTPRequestHandler):
        def _json(self, data, status=200):
            payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

        def _file(self, path: Path):
            payload = path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f'attachment; filename="{urllib.parse.quote(path.name)}"')
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

        def do_GET(self):
            if self.path == "/" or self.path == "/index.html":
                path = ROOT / "index.html"
                payload = path.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
            elif self.path.startswith("/downloads/"):
                name = Path(urllib.parse.unquote(self.path.removeprefix("/downloads/"))).name
                path = OUTPUT_DIR / name
                if path.exists():
                    self._file(path)
                else:
                    self._json({"error": "文件不存在"}, 404)
            else:
                self._json({"error": "Not found"}, 404)

        def do_POST(self):
            length = int(self.headers.get("Content-Length", "0"))
            try:
                body = json.loads(self.rfile.read(length) or b"{}")
                url = str(body.get("url", "")).strip()
                article = parse_article(fetch_page(url), url)
                if self.path == "/api/extract":
                    ARTICLE_CACHE[url] = article
                    preview = dict(article)
                    preview["blocks"] = []
                    for block in article["blocks"]:
                        if block["type"] == "text":
                            preview["blocks"].append(block)
                        elif block["type"] == "paragraph":
                            preview["blocks"].append({"type": "text", "text": "".join(r.get("text", "") for r in block.get("runs", []))})
                        if len(preview["blocks"]) >= 80:
                            break
                    preview["imageCount"] = sum(1 for b in article["blocks"] if b["type"] == "image")
                    self._json(preview)
                elif self.path == "/api/export":
                    # Reuse the article that was just extracted. A second request to
                    # the public WeChat page is often blocked or rate-limited.
                    article = ARTICLE_CACHE.get(url)
                    if article is None:
                        article = parse_article(fetch_page(url), url)
                    stamp = __import__("datetime").datetime.now().strftime("%Y%m%d-%H%M%S")
                    safe = re.sub(r"[^\w\u4e00-\u9fff-]+", "-", article["title"]).strip("-")[:45] or "wechat-article"
                    path = OUTPUT_DIR / f"{safe}-{stamp}.docx"
                    build_docx(article, path)
                    self._json({"download": f"/downloads/{urllib.parse.quote(path.name)}", "filename": path.name})
                else:
                    self._json({"error": "Not found"}, 404)
            except Exception as exc:
                print(f"Request failed: {type(exc).__name__}: {exc!r}")
                self._json({"error": f"{type(exc).__name__}: {exc}"}, 400)

        def log_message(self, fmt, *args):
            print(fmt % args)

    server = ThreadingHTTPServer(("127.0.0.1", int(os.environ.get("PORT", "8765"))), Handler)
    print(f"Wechat extractor listening on http://127.0.0.1:{server.server_port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
