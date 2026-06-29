from __future__ import annotations

import json
import os
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlparse

BASE_DIR = Path(__file__).resolve().parent
SEARCH_DIRS = [BASE_DIR, Path.cwd()]
for local_libs in [candidate / name for candidate in SEARCH_DIRS for name in ("libs", ".python-libs2", ".python-libs")]:
    if local_libs.exists():
        sys.path.insert(0, str(local_libs))

try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    load_workbook = None

import schedule_model as model

STATIC_DIR = BASE_DIR / "web"

EMPTY_DEFAULTS = {
    "resources": [{"period": "上旬", "total": "", "foreign": "", "domestic": "", "hrb500e": "", "six_hundred": "", "four_hundred": ""}],
    "calendar": [{"period": "上旬", "line": "", "days": ""}],
    "foreign": [],
    "efficiency": [],
    "cashflow": [],
    "demands": [],
    "ratios": [],
    "forecast": [],
    "message": "已打开空模板，请直接在网页中输入数据。",
}


def clean_rows(rows):
    return [
        {key: value for key, value in row.items() if value not in ("", None)}
        for row in rows
        if any(value not in ("", None) for value in row.values())
    ]


def to_float(value, default=0.0):
    if value in ("", None):
        return default
    return float(value)


def to_int(value, default=0):
    if value in ("", None):
        return default
    return int(round(float(value)))


def read_defaults() -> dict:
    return json.loads(json.dumps(EMPTY_DEFAULTS, ensure_ascii=False))


def parse_multipart_file(body: bytes, content_type: str) -> bytes:
    marker = "boundary="
    if marker not in content_type:
        raise model.ModelInputError("导入失败：没有找到上传文件边界。")
    boundary = content_type.split(marker, 1)[1].split(";", 1)[0].strip().strip('"')
    delimiter = ("--" + boundary).encode("utf-8")
    for part in body.split(delimiter):
        if b"filename=" not in part:
            continue
        header_end = part.find(b"\r\n\r\n")
        if header_end < 0:
            continue
        content = part[header_end + 4 :]
        if content.endswith(b"\r\n"):
            content = content[:-2]
        if content.endswith(b"--"):
            content = content[:-2]
        if not content:
            raise model.ModelInputError("导入失败：上传文件为空。")
        return content
    raise model.ModelInputError("导入失败：没有读取到 Excel 文件。")


def row_value(row: dict, names: list[str]):
    for name in names:
        value = row.get(name)
        if value not in ("", None):
            return value
    return ""


def import_excel_rows(table: str, file_bytes: bytes) -> list[dict]:
    if load_workbook is None:
        raise model.ModelInputError("导入失败：缺少 openpyxl 依赖。")
    wb = load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    if table == "foreign":
        rows = model.sheet_rows(ws, ["旬度", "产线", "牌号", "规格", "吨位"])
        return [
            {
                "period": model.norm_text(row_value(row, ["旬度"])),
                "line": model.norm_text(row_value(row, ["产线", "生产条线"])),
                "grade": model.norm_text(row_value(row, ["牌号"])),
                "spec": row_value(row, ["规格"]),
                "tons": row_value(row, ["吨位", "外贸吨位"]),
            }
            for row in rows
        ]
    if table == "efficiency":
        rows = model.sheet_rows(ws, ["产线", "牌号", "规格", "日产量"])
        return [
            {
                "line": model.norm_text(row_value(row, ["产线", "生产条线"])),
                "grade": model.norm_text(row_value(row, ["牌号"])),
                "spec": row_value(row, ["规格"]),
                "daily": row_value(row, ["日产量", "日均产量"]),
            }
            for row in rows
        ]
    if table == "cashflow":
        rows = model.sheet_rows(ws, ["牌号", "规格", "现金流"])
        return [
            {
                "grade": model.norm_text(row_value(row, ["牌号"])),
                "spec": row_value(row, ["规格"]),
                "cashflow": row_value(row, ["现金流"]),
            }
            for row in rows
        ]
    raise model.ModelInputError("导入失败：不支持的表格类型。")


def solve_payload(payload: dict) -> dict:
    resources = []
    for row in clean_rows(payload.get("resources", [])):
        period = model.norm_text(row.get("period"))
        total = model.tons_from_resource(row.get("total"), f"{period} 总资源量")
        foreign = model.tons_from_resource(row.get("foreign"), f"{period} 外贸资源量")
        domestic = total - foreign
        hrb500e = model.tons_from_resource(row.get("hrb500e"), f"{period} 500兆帕保供量")
        six = model.tons_from_resource(row.get("six_hundred"), f"{period} 600兆帕保供量")
        four_tons = domestic - hrb500e - six
        resources.append(model.ResourceRow(period, total, foreign, domestic, hrb500e, six, four_tons))

    calendar = {}
    for row in clean_rows(payload.get("calendar", [])):
        calendar[(model.norm_text(row.get("period")), model.norm_line(row.get("line")))] = to_float(row.get("days"))

    foreign_rows = []
    for row in clean_rows(payload.get("foreign", [])):
        foreign_rows.append(
            model.ForeignRow(
                model.norm_text(row.get("period")),
                model.norm_line(row.get("line")),
                model.norm_grade(row.get("grade")),
                to_int(row.get("spec")),
                to_float(row.get("tons")),
            )
        )

    efficiency = {}
    for row in clean_rows(payload.get("efficiency", [])):
        efficiency[(model.norm_line(row.get("line")), model.norm_grade(row.get("grade")), to_int(row.get("spec")))] = to_float(row.get("daily"))

    cashflow = {}
    for row in clean_rows(payload.get("cashflow", [])):
        cashflow[(model.norm_grade(row.get("grade")), to_int(row.get("spec")))] = to_float(row.get("cashflow"))

    demands = []
    for row in clean_rows(payload.get("demands", [])):
        tons = to_float(row.get("tons"))
        if tons == 0:
            continue
        demands.append(model.DemandRow(model.norm_text(row.get("period")), model.norm_grade(row.get("grade")), to_int(row.get("spec")), model.tons_to_units(tons, "保供需求吨位") * model.UNIT_TONS))

    ratios = []
    for row in clean_rows(payload.get("ratios", [])):
        ratios.append(model.RatioRow(to_int(row.get("spec")), to_float(row.get("lower")), to_float(row.get("upper"))))

    forecasts = []
    for row in clean_rows(payload.get("forecast", [])):
        tons = to_float(row.get("tons"))
        if tons == 0:
            continue
        forecasts.append(model.ForecastRow(model.norm_text(row.get("period")), model.norm_grade(row.get("grade")), to_int(row.get("spec")), tons))

    if not resources:
        raise model.ModelInputError("至少需要填写一行旬度资源")

    foreign_days = model.compute_foreign_days(foreign_rows, efficiency)
    all_plan = []
    all_checks = []
    for resource in resources:
        plan, checks = model.solve_period(resource, calendar, foreign_days, efficiency, cashflow, demands, ratios)
        all_plan.extend(plan)
        all_checks.extend(checks)
    suggestions = model.build_suggestions(all_plan, forecasts)
    return {"plan": all_plan, "checks": all_checks, "suggestions": suggestions}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):  # noqa: A003
        return

    def send_json(self, status: int, body: dict) -> None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/defaults":
            self.send_json(200, read_defaults())
            return
        target = STATIC_DIR / ("index.html" if parsed.path in ("/", "") else parsed.path.lstrip("/"))
        if not target.resolve().is_relative_to(STATIC_DIR.resolve()) or not target.exists():
            self.send_error(404)
            return
        content_type = "text/html; charset=utf-8" if target.suffix == ".html" else "text/plain; charset=utf-8"
        if target.suffix == ".css":
            content_type = "text/css; charset=utf-8"
        elif target.suffix == ".js":
            content_type = "application/javascript; charset=utf-8"
        data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):  # noqa: N802
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        try:
            if parsed.path == "/api/solve":
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                self.send_json(200, solve_payload(payload))
                return
            if parsed.path == "/api/import-excel":
                table = parse_qs(parsed.query).get("table", [""])[0]
                body = self.rfile.read(length)
                file_bytes = parse_multipart_file(body, self.headers.get("Content-Type", ""))
                self.send_json(200, {"rows": import_excel_rows(table, file_bytes)})
                return
            self.send_error(404)
        except model.ModelInputError as exc:
            self.send_json(400, {"error": str(exc)})
        except Exception as exc:  # pragma: no cover
            self.send_json(500, {"error": f"程序错误：{exc}"})


def main() -> None:
    port = int(os.environ.get("PORT", "8765"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"排产网页应用已启动：http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
