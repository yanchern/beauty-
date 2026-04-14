from __future__ import annotations

import json
import os
import re
import shutil
import uuid
from datetime import datetime
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

from fr_self_tax_sales_report import COUNTRY_CONFIGS, generate_self_tax_report


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "tax_web_data"
HOME_TEMPLATE_PATH = BASE_DIR / "fr_self_tax_web.html"
COUNTRY_TEMPLATE_PATH = BASE_DIR / "fr_self_tax_country.html"
COUNTRY_ORDER = ("fr", "es", "uk", "de", "it", "nl")
ALLOWED_EMBLEMS = {Path(COUNTRY_CONFIGS[code].emblem_path).name for code in COUNTRY_ORDER}
COUNTRY_BY_SLUG = {COUNTRY_CONFIGS[code].slug: COUNTRY_CONFIGS[code] for code in COUNTRY_ORDER}


def load_template(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def format_amount(value) -> str:
    return f"{value:.2f}"


def serialize_metrics(metrics: list[tuple[str, object]]) -> list[dict[str, str]]:
    return [{"label": label, "amount": format_amount(amount)} for label, amount in metrics]


def safe_filename(filename: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._\-\u4e00-\u9fff]+", "_", filename).strip("._")
    return cleaned or f"upload_{uuid.uuid4().hex}"


def accepted_suffixes(accept_value: str) -> set[str]:
    return {token.strip().lower() for token in accept_value.split(",") if token.strip().startswith(".")}


def render_country_cards() -> str:
    cards: list[str] = []
    for code in COUNTRY_ORDER:
        country = COUNTRY_CONFIGS[code]
        cards.append(
            f"""
      <a class="card" href="/{country.slug}">
        <div class="crest-wrap">
          <img class="crest" src="{country.emblem_path}" alt="{country.name_zh}国徽">
        </div>
        <span class="eyebrow">{country.slug.upper()}</span>
        <h2>{country.name_zh}</h2>
        <p>{country.description}</p>
        <div class="meta">进入{country.name_zh}计算入口</div>
      </a>""".rstrip()
        )
    return "\n".join(cards)


def render_home() -> str:
    template = load_template(HOME_TEMPLATE_PATH)
    return template.replace("__COUNTRY_CARDS__", render_country_cards())


def render_country_page(slug: str) -> str:
    country = COUNTRY_BY_SLUG[slug]
    template = load_template(COUNTRY_TEMPLATE_PATH)
    replacements = {
        "__COUNTRY_CODE__": country.code,
        "__COUNTRY_SLUG__": country.slug,
        "__COUNTRY_TITLE__": country.title,
        "__COUNTRY_NAME__": country.name_zh,
        "__COUNTRY_DESCRIPTION__": country.description,
        "__SALES_REPORT_LABEL__": country.sales_report_label,
        "__LOGIC_DOC_LABEL__": country.logic_doc_label,
        "__LOGIC_ACCEPT__": country.logic_doc_accept,
        "__COUNTRY_EMBLEM_SRC__": country.emblem_path,
        "__COUNTRY_EMBLEM_ALT__": f"{country.name_zh}国徽",
    }
    for placeholder, value in replacements.items():
        template = template.replace(placeholder, value)
    return template


def parse_multipart_form(content_type: str, body: bytes) -> dict[str, dict[str, object]]:
    message = BytesParser(policy=policy.default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )
    fields: dict[str, dict[str, object]] = {}
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        field_name = part.get_param("name", header="content-disposition")
        if not field_name:
            continue
        fields[field_name] = {
            "filename": part.get_filename() or "",
            "content_type": part.get_content_type(),
            "content": part.get_payload(decode=True) or b"",
        }
    return fields


def ensure_download_path(*parts: str) -> Path:
    path = (DATA_DIR / Path(*parts)).resolve()
    if DATA_DIR.resolve() not in path.parents:
        raise ValueError("非法下载路径")
    return path


class TaxWebHandler(BaseHTTPRequestHandler):
    server_version = "QuarterlyTaxWeb/2.0"

    def send_html(self, html: str, status: int = HTTPStatus.OK) -> None:
        payload = html.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def send_json(self, payload: dict[str, object], status: int = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path: Path, content_type: str) -> None:
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        route = self.path.split("?", 1)[0]
        if route == "/":
            self.send_html(render_home())
            return

        if route == "/healthz":
            payload = b"ok"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return

        slug = route.lstrip("/")
        if slug in COUNTRY_BY_SLUG:
            self.send_html(render_country_page(slug))
            return

        if slug in ALLOWED_EMBLEMS:
            self.send_file(BASE_DIR / slug, "image/svg+xml; charset=utf-8")
            return

        if route.startswith("/downloads/"):
            relative = unquote(route.removeprefix("/downloads/"))
            try:
                file_path = ensure_download_path(*Path(relative).parts)
            except ValueError:
                self.send_error(HTTPStatus.BAD_REQUEST, "非法下载路径")
                return
            if not file_path.exists() or not file_path.is_file():
                self.send_error(HTTPStatus.NOT_FOUND, "未找到下载文件")
                return
            self.send_file(
                file_path,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            return

        self.send_error(HTTPStatus.NOT_FOUND, "未找到页面")

    def do_POST(self) -> None:
        route = self.path.split("?", 1)[0]
        if not route.startswith("/process/"):
            self.send_error(HTTPStatus.NOT_FOUND, "未找到接口")
            return

        slug = route.removeprefix("/process/")
        country = COUNTRY_BY_SLUG.get(slug)
        if country is None:
            self.send_json({"message": "不支持的国家入口。"}, HTTPStatus.NOT_FOUND)
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.send_json({"message": "请求体长度无效。"}, HTTPStatus.BAD_REQUEST)
            return

        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            self.send_json({"message": "请使用表单上传文件。"}, HTTPStatus.BAD_REQUEST)
            return

        body = self.rfile.read(content_length)
        try:
            fields = parse_multipart_form(content_type, body)
        except Exception as exc:
            self.send_json({"message": f"无法解析上传内容：{exc}"}, HTTPStatus.BAD_REQUEST)
            return

        sales_file = fields.get("sales_report")
        logic_file = fields.get("logic_pdf")
        if not sales_file or not logic_file:
            self.send_json({"message": "请同时上传销售报告和税金逻辑文件。"}, HTTPStatus.BAD_REQUEST)
            return

        sales_name = safe_filename(str(sales_file.get("filename") or "sales_report.csv"))
        logic_name = safe_filename(str(logic_file.get("filename") or "logic_doc"))
        if Path(sales_name).suffix.lower() != ".csv":
            self.send_json({"message": "销售报告必须是 CSV 文件。"}, HTTPStatus.BAD_REQUEST)
            return

        allowed_logic_suffixes = accepted_suffixes(country.logic_doc_accept)
        logic_suffix = Path(logic_name).suffix.lower()
        if allowed_logic_suffixes and logic_suffix not in allowed_logic_suffixes:
            self.send_json(
                {"message": f"{country.logic_doc_label} 文件类型不正确，请上传 {', '.join(sorted(allowed_logic_suffixes))}。"},
                HTTPStatus.BAD_REQUEST,
            )
            return

        DATA_DIR.mkdir(parents=True, exist_ok=True)
        session_dir = DATA_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
        session_dir.mkdir(parents=True, exist_ok=False)

        sales_path = session_dir / sales_name
        logic_path = session_dir / logic_name
        output_path = session_dir / f"{Path(sales_name).stem}_{country.name_zh}税金汇总.xlsx"

        try:
            sales_path.write_bytes(bytes(sales_file["content"]))
            logic_path.write_bytes(bytes(logic_file["content"]))
            result = generate_self_tax_report(
                csv_path=sales_path,
                country_code=country.code,
                logic_pdf_path=logic_path,
                output_path=output_path,
            )
        except Exception as exc:
            shutil.rmtree(session_dir, ignore_errors=True)
            self.send_json({"message": f"处理失败：{exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return

        download_url = f"/downloads/{session_dir.name}/{output_path.name}"
        notes = [
            f"国家：{result['country_name']}",
            f"销售报告：{sales_name}",
            f"逻辑文件：{logic_name}",
        ]
        currency_codes = result.get("currency_codes") or []
        currency_label = "/".join(currency_codes) if currency_codes else ""

        self.send_json(
            {
                "message": f"{result['country_name']}税金计算完成，已生成新的 Excel 文件。",
                "row_count": result["row_count"],
                "download_url": download_url,
                "card_metrics": serialize_metrics(result["card_metrics"]),
                "summary_metrics": serialize_metrics(result["summary_metrics"]),
                "currency_label": currency_label,
                "notes": notes,
            }
        )

    def log_message(self, format: str, *args) -> None:
        return


def make_server(host: str, port: int) -> ThreadingHTTPServer:
    return ThreadingHTTPServer((host, port), TaxWebHandler)


def main() -> None:
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8000"))
    server = make_server(host, port)
    print(f"Serving tax web on http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
