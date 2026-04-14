from __future__ import annotations

import json
import os
import re
import tempfile
import threading
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from fr_self_tax_sales_report import COUNTRY_CONFIGS, generate_self_tax_report, get_country_config


BASE_DIR = Path(__file__).resolve().parent
HOME_HTML_PATH = BASE_DIR / "fr_self_tax_web.html"
COUNTRY_HTML_PATH = BASE_DIR / "fr_self_tax_country.html"
WORK_DIR = BASE_DIR / "tax_web_data"
UPLOAD_DIR = WORK_DIR / "uploads"
RESULT_DIR = WORK_DIR / "results"

RESULTS_LOCK = threading.Lock()
RESULTS_INDEX: dict[str, Path] = {}
SLUG_TO_COUNTRY = {config.slug: config.code for config in COUNTRY_CONFIGS.values()}


def ensure_dirs() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)


def safe_name(name: str, fallback: str) -> str:
    clean = Path(name or fallback).name
    clean = re.sub(r"[^\w\-.()\u4e00-\u9fff]+", "_", clean)
    return clean or fallback


def parse_multipart(handler: BaseHTTPRequestHandler) -> dict[str, dict[str, object]]:
    content_type = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        raise ValueError("请求必须是 multipart/form-data。")

    content_length = int(handler.headers.get("Content-Length", "0"))
    body = handler.rfile.read(content_length)
    message = BytesParser(policy=default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )

    fields: dict[str, dict[str, object]] = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        fields[name] = {
            "filename": filename,
            "value": part.get_content().strip() if not filename else None,
            "content": payload,
            "content_type": part.get_content_type(),
        }
    return fields


def save_uploaded_file(data: dict[str, object], prefix: str) -> Path:
    filename = safe_name(str(data.get("filename") or prefix), prefix)
    temp_dir = Path(tempfile.mkdtemp(prefix=f"{prefix}_", dir=str(UPLOAD_DIR)))
    path = temp_dir / filename
    path.write_bytes(data["content"])  # type: ignore[index]
    return path


def render_home() -> bytes:
    return HOME_HTML_PATH.read_bytes()


def render_country_page(country_code: str) -> bytes:
    country = get_country_config(country_code)
    html = COUNTRY_HTML_PATH.read_text(encoding="utf-8")
    replacements = {
        "__COUNTRY_CODE__": country.code,
        "__COUNTRY_TITLE__": country.title,
        "__COUNTRY_DESCRIPTION__": country.description,
        "__SALES_REPORT_LABEL__": country.sales_report_label,
        "__LOGIC_PDF_LABEL__": country.logic_pdf_label,
    }
    for old, new in replacements.items():
        html = html.replace(old, new)
    return html.encode("utf-8")


def build_result_payload(result: dict[str, object]) -> dict[str, object]:
    group_totals = [
        {"label": str(group), "amount": f"{result['group_totals'][group]:.2f}"}  # type: ignore[index]
        for group in result["group_order"]  # type: ignore[index]
    ]
    group_totals.append({"label": "自行缴税合计", "amount": f"{result['total_sales']:.2f}"})

    return {
        "message": "已完成计算，结果文件可直接下载。",
        "row_count": result["row_count"],
        "group_totals": group_totals,
        "download_url": f"/download?token={result['download_token']}",
        "notes": [
            f"国家：{result['country_name']}",
            f"输出文件：{Path(result['output_path']).name}",
            "结果文件包含汇总、规则说明、命中明细三个工作表。",
        ],
    }


class TaxWebHandler(BaseHTTPRequestHandler):
    server_version = "CountryTaxWeb/2.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/") or "/"

        if path == "/healthz":
            self._send_json({"ok": True, "service": "tax-web"})
            return

        if path == "/":
            self._send_html(render_home())
            return

        if path == "/download":
            params = parse_qs(parsed.query)
            token = params.get("token", [""])[0]
            self._handle_download(token)
            return

        country_code = SLUG_TO_COUNTRY.get(path.lstrip("/"))
        if country_code:
            self._send_html(render_country_page(country_code))
            return

        self._send_json({"message": "未找到页面。"}, status=HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        parts = [segment for segment in parsed.path.split("/") if segment]
        if len(parts) != 2 or parts[0] != "process":
            self._send_json({"message": "未找到接口。"}, status=HTTPStatus.NOT_FOUND)
            return

        country_code = parts[1]
        try:
            country = get_country_config(country_code)
        except ValueError as exc:
            self._send_json({"message": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return

        try:
            fields = parse_multipart(self)
            sales_report = fields.get("sales_report")
            logic_pdf = fields.get("logic_pdf")

            if not sales_report or not sales_report.get("filename"):
                raise ValueError(f"请上传{country.sales_report_label}。")
            if not logic_pdf or not logic_pdf.get("filename"):
                raise ValueError(f"请上传{country.logic_pdf_label}。")

            sales_report_path = save_uploaded_file(sales_report, f"{country.code}_sales_report.csv")
            logic_pdf_path = save_uploaded_file(logic_pdf, f"{country.code}_logic.pdf")

            output_stem = Path(str(sales_report.get("filename"))).stem or f"{country.name_zh}销售报告"
            output_name = safe_name(f"{output_stem}_自行缴税销售额汇总.xlsx", "result.xlsx")
            output_path = RESULT_DIR / output_name
            if output_path.exists():
                output_path = RESULT_DIR / safe_name(
                    f"{output_stem}_自行缴税销售额汇总_{next(tempfile._get_candidate_names())}.xlsx",
                    "result.xlsx",
                )

            result = generate_self_tax_report(
                csv_path=sales_report_path,
                country_code=country.code,
                logic_pdf_path=logic_pdf_path,
                output_path=output_path,
            )

            token = next(tempfile._get_candidate_names())
            with RESULTS_LOCK:
                RESULTS_INDEX[token] = Path(result["output_path"])
            result["download_token"] = token

            self._send_json(build_result_payload(result))
        except Exception as exc:
            self._send_json({"message": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def _handle_download(self, token: str) -> None:
        with RESULTS_LOCK:
            path = RESULTS_INDEX.get(token)

        if not path or not path.exists():
            self._send_json({"message": "结果文件不存在或已失效。"}, status=HTTPStatus.NOT_FOUND)
            return

        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Length", str(len(data)))
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="result.xlsx"; filename*=UTF-8\'\'{quote(path.name)}',
        )
        self.end_headers()
        self.wfile.write(data)

    def _send_html(self, body: bytes, status: HTTPStatus = HTTPStatus.OK) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, payload: dict[str, object], status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args) -> None:
        return


def run_server(host: str | None = None, port: int | None = None) -> None:
    ensure_dirs()
    host = host or os.environ.get("HOST", "0.0.0.0")
    port = port or int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer((host, port), TaxWebHandler)
    print(f"税金计算网页已启动: http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    run_server()
