from __future__ import annotations

import json
import os
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

BASE_DIR = Path(__file__).resolve().parent
for local_libs in (BASE_DIR / "libs", BASE_DIR / ".python-libs2", BASE_DIR / ".python-libs"):
    if local_libs.exists():
        sys.path.insert(0, str(local_libs))

try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    load_workbook = None

import schedule_model as model


DEFAULT_EXCEL = Path.home() / "Desktop" / "排产所需文件 - 副本.xlsx"
STATIC_DIR = BASE_DIR / "web"


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
    empty = {
        "resources": [],
        "calendar": [],
        "foreign": [],
        "efficiency": [],
        "cashflow": [],
        "demands": [],
        "ratios": [],
        "forecast": [],
        "message": "未找到默认 Excel，页面已打开空模板。",
    }
    if load_workbook is None:
        empty["message"] = "缺少 openpyxl，无法预填 Excel 数据。"
        return empty
    if not DEFAULT_EXCEL.exists():
        return empty

    wb = load_workbook(DEFAULT_EXCEL, data_only=True)
    resource_ws = model.find_sheet(wb, ["旬度资源说明"])
    resources = model.parse_resources(resource_ws)
    default_period = resources[0].period if resources else "上旬"
    calendar = model.parse_calendar(model.find_sheet(wb, ["生产日历"]), default_period)
    foreign = model.parse_foreign(model.find_sheet(wb, ["外贸排产明细"]))
    efficiency = model.parse_efficiency(model.find_sheet(wb, ["产品生产效率表"]))
    cashflow = model.parse_cashflow(model.find_sheet(wb, ["现金流量表"]))
    demands = model.parse_demands(model.find_sheet(wb, ["HRB500E及600兆帕需求表", "保供量"], required=False))
    ratios = model.parse_ratios(wb, resource_ws)
    forecast = model.parse_forecast(model.find_sheet(wb, ["400兆帕客户需求预报"], required=False), default_period)

    return {
        "resources": [row.__dict__ for row in resources],
        "calendar": [{"period": period, "line": line, "days": days} for (period, line), days in sorted(calendar.items())],
        "foreign": [row.__dict__ for row in foreign],
        "efficiency": [{"line": line, "grade": grade, "spec": spec, "daily": daily} for (line, grade, spec), daily in sorted(efficiency.items())],
        "cashflow": [{"grade": grade, "spec": spec, "cashflow": cash} for (grade, spec), cash in sorted(cashflow.items())],
        "demands": [row.__dict__ for row in demands],
        "ratios": [row.__dict__ for row in ratios],
        "forecast": [row.__dict__ for row in forecast],
        "message": f"已从 {DEFAULT_EXCEL} 预填数据。",
    }


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
            try:
                self.send_json(200, read_defaults())
            except model.ModelInputError as exc:
                self.send_json(400, {"error": str(exc)})
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
        if urlparse(self.path).path != "/api/solve":
            self.send_error(404)
            return
        length = int(self.headers.get("Content-Length", "0"))
        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            self.send_json(200, solve_payload(payload))
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



