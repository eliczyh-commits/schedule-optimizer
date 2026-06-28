from __future__ import annotations

import argparse
import math
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
for LOCAL_LIBS in (BASE_DIR / "libs", BASE_DIR / ".python-libs2", BASE_DIR / ".python-libs"):
    if LOCAL_LIBS.exists():
        sys.path.insert(0, str(LOCAL_LIBS))

try:
    from openpyxl import Workbook, load_workbook
except ImportError:  # pragma: no cover
    print("缺少依赖 openpyxl。请先运行：py -m pip install -r requirements.txt", file=sys.stderr)
    raise SystemExit(2)

try:
    import pulp
except ImportError:  # pragma: no cover
    print("缺少依赖 pulp。请先运行：py -m pip install -r requirements.txt", file=sys.stderr)
    raise SystemExit(2)


UNIT_TONS = 100
FOUR_HUNDRED_GRADES = {"HRB400", "HRB400E"}
PRIORITY_GRADES = {"HRB500E", "T63E/E/G", "T63/E/G"}
SIX_HUNDRED_GRADES = {"T63E/E/G", "T63/E/G"}
T63E_SPECS = {12, 14, 16, 18, 20, 22, 25, 28, 32}
T63_SPEC = {10}


class ModelInputError(ValueError):
    pass


@dataclass(frozen=True)
class ResourceRow:
    period: str
    total: int
    foreign: int
    domestic: int
    hrb500e: int
    six_hundred: int
    four_hundred: int


@dataclass(frozen=True)
class ForeignRow:
    period: str
    line: str
    grade: str
    spec: int
    tons: float


@dataclass(frozen=True)
class DemandRow:
    period: str
    grade: str
    spec: int
    tons: int


@dataclass(frozen=True)
class RatioRow:
    spec: int
    lower: float
    upper: float


@dataclass(frozen=True)
class ForecastRow:
    period: str
    grade: str
    spec: int
    tons: float


def norm_text(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def norm_line(value: Any) -> str:
    text = norm_text(value).replace(" ", "")
    if text.endswith("线"):
        text = text[:-1]
    return text


def norm_grade(value: Any) -> str:
    return norm_text(value).replace(" ", "").upper()


def as_float(value: Any, field: str) -> float:
    if value is None or norm_text(value) == "":
        raise ModelInputError(f"{field} 不能为空")
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ModelInputError(f"{field} 必须是数字，当前值：{value}") from exc


def as_spec(value: Any, field: str = "规格") -> int:
    number = as_float(value, field)
    if abs(number - round(number)) > 1e-9:
        raise ModelInputError(f"{field} 必须是整数规格，当前值：{value}")
    return int(round(number))


def tons_from_resource(value: Any, field: str) -> int:
    number = as_float(value, field)
    tons = number * 10000 if abs(number) < 1000 else number
    return int(round(tons))


def tons_to_units(tons: float, field: str) -> int:
    units = tons / UNIT_TONS
    if abs(units - round(units)) > 1e-9:
        raise ModelInputError(f"{field} 必须是 {UNIT_TONS} 吨的整数倍，当前吨位：{tons}")
    return int(round(units))


def sheet_rows(ws, required_headers: list[str] | None = None) -> list[dict[str, Any]]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header_idx = None
    for idx, row in enumerate(rows[:30]):
        names = [norm_text(v) for v in row]
        if required_headers:
            if all(header in names for header in required_headers):
                header_idx = idx
                break
        elif any(names):
            header_idx = idx
            break
    if header_idx is None:
        return []
    headers = [norm_text(v) for v in rows[header_idx]]
    output = []
    for row in rows[header_idx + 1 :]:
        item = {headers[i]: row[i] for i in range(min(len(headers), len(row))) if headers[i]}
        if any(v is not None and norm_text(v) != "" for v in item.values()):
            output.append(item)
    return output


def find_sheet(wb, names: list[str], required: bool = True):
    name_map = {norm_text(name): name for name in wb.sheetnames}
    for name in names:
        if name in name_map:
            return wb[name_map[name]]
    if required:
        raise ModelInputError(f"缺少工作表：{' 或 '.join(names)}")
    return None


def first_present(row: dict[str, Any], names: list[str], field: str) -> Any:
    for name in names:
        if name in row and row[name] is not None and norm_text(row[name]) != "":
            return row[name]
    raise ModelInputError(f"缺少字段或字段为空：{field}")


def parse_resources(ws) -> list[ResourceRow]:
    rows = sheet_rows(ws)
    if rows and "旬度" in rows[0]:
        parsed = []
        for idx, row in enumerate(rows, start=2):
            period = norm_text(first_present(row, ["旬度"], "旬度"))
            total = tons_from_resource(first_present(row, ["总资源量", "总资源"], "总资源量"), f"旬度资源说明第{idx}行总资源量")
            foreign = tons_from_resource(first_present(row, ["外贸资源量", "外贸"], "外贸资源量"), f"旬度资源说明第{idx}行外贸资源量")
            domestic = tons_from_resource(first_present(row, ["内贸资源量", "内贸"], "内贸资源量"), f"旬度资源说明第{idx}行内贸资源量")
            hrb500e = tons_from_resource(first_present(row, ["500兆帕保供量", "HRB500E保供量", "HRB500E"], "500兆帕保供量"), f"旬度资源说明第{idx}行500兆帕保供量")
            six = tons_from_resource(first_present(row, ["600兆帕保供量", "600兆帕", "600MPA"], "600兆帕保供量"), f"旬度资源说明第{idx}行600兆帕保供量")
            four = row.get("HRB400/HRB400E资源量")
            four_tons = tons_from_resource(four, f"旬度资源说明第{idx}行HRB400/HRB400E资源量") if four is not None and norm_text(four) != "" else domestic - hrb500e - six
            parsed.append(ResourceRow(period, total, foreign, domestic, hrb500e, six, four_tons))
        return parsed

    period = norm_text(ws["C1"].value) or "上旬"
    total = foreign = domestic = None
    hrb500e = six = four = 0
    for row in ws.iter_rows(values_only=True):
        cells = [norm_text(v) for v in row]
        label = "".join(cells[:2])
        values = [v for v in row if isinstance(v, (int, float))]
        if len(cells) > 1 and cells[1] == "总资源" and values:
            total = tons_from_resource(values[-1], "总资源")
        elif "外贸" in label and values:
            foreign = tons_from_resource(values[-1], "外贸资源")
        elif "内贸" in label and values:
            domestic = tons_from_resource(values[-1], "内贸资源")
        elif ("500兆帕" in label or "500MPA" in label.upper() or "HRB500E" in label) and values:
            hrb500e = tons_from_resource(values[-1], "500兆帕保供量")
        elif ("600兆帕" in label or "600MPA" in label.upper() or "T63" in label) and values:
            six += tons_from_resource(values[-1], "600兆帕保供量")
        elif ("HRB400/HRB400E" in label or "HRB400、HRB400E" in label) and len(cells) > 1 and cells[1] == "螺纹" and values:
            four = tons_from_resource(values[-1], "HRB400/HRB400E资源量")
    if total is None or foreign is None or domestic is None:
        raise ModelInputError("旬度资源说明无法识别总资源、外贸资源、内贸资源")
    if four == 0:
        four = domestic - hrb500e - six
    return [ResourceRow(period, total, foreign, domestic, hrb500e, six, four)]


def parse_calendar(ws, default_period: str) -> dict[tuple[str, str], float]:
    rows = sheet_rows(ws, ["产线", "生产天数"])
    result = {}
    for idx, row in enumerate(rows, start=2):
        raw_line = row.get("产线") or row.get("生产条线")
        if raw_line is None or norm_text(raw_line) == "":
            continue
        period = norm_text(row.get("旬度")) or default_period
        line = norm_line(raw_line)
        if row.get("生产天数") is None or norm_text(row.get("生产天数")) == "":
            continue
        days = as_float(first_present(row, ["生产天数", period], "生产天数"), f"生产日历第{idx}行生产天数")
        result[(period, line)] = days
    if result:
        return result

    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or not row[0]:
            continue
        result[(default_period, norm_line(row[0]))] = as_float(row[1], "生产天数")
    return result


def parse_line_capacity(ws) -> dict[str, float]:
    rows = sheet_rows(ws, ["产线", "日产能"])
    result = {}
    for idx, row in enumerate(rows, start=2):
        line = norm_line(first_present(row, ["产线", "生产条线"], "产线"))
        result[line] = as_float(first_present(row, ["日产能", "日均产能"], "日产能"), f"产线基础信息第{idx}行日产能")
    return result


def parse_foreign(ws) -> list[ForeignRow]:
    result = []
    for idx, row in enumerate(sheet_rows(ws, ["旬度", "产线", "牌号", "规格", "吨位"]), start=2):
        result.append(
            ForeignRow(
                norm_text(first_present(row, ["旬度"], "旬度")),
                norm_line(first_present(row, ["产线", "生产条线"], "产线")),
                norm_grade(first_present(row, ["牌号"], "牌号")),
                as_spec(first_present(row, ["规格"], "规格")),
                as_float(first_present(row, ["吨位", "外贸吨位"], "吨位"), f"外贸排产明细第{idx}行吨位"),
            )
        )
    return result


def parse_efficiency(ws) -> dict[tuple[str, str, int], float]:
    result = {}
    for idx, row in enumerate(sheet_rows(ws, ["产线", "牌号", "规格", "日产量"]), start=2):
        line = norm_line(first_present(row, ["产线", "生产条线"], "产线"))
        grade = norm_grade(first_present(row, ["牌号"], "牌号"))
        spec = as_spec(first_present(row, ["规格"], "规格"))
        daily = as_float(first_present(row, ["日产量", "日均产量"], "日产量"), f"产品生产效率表第{idx}行日产量")
        if daily <= 0:
            raise ModelInputError(f"产品生产效率表第{idx}行日产量必须大于0")
        result[(line, grade, spec)] = daily
    return result


def parse_cashflow(ws) -> dict[tuple[str, int], float]:
    result = {}
    for idx, row in enumerate(sheet_rows(ws, ["牌号", "规格", "现金流"]), start=2):
        grade = norm_grade(first_present(row, ["牌号"], "牌号"))
        spec = as_spec(first_present(row, ["规格"], "规格"))
        result[(grade, spec)] = as_float(first_present(row, ["现金流"], "现金流"), f"现金流量表第{idx}行现金流")
    return result


def parse_demands(ws) -> list[DemandRow]:
    if ws is None:
        return []
    result = []
    last_period = ""
    last_grade = ""
    for idx, row in enumerate(sheet_rows(ws, ["旬度", "牌号", "规格"]), start=2):
        period = norm_text(row.get("旬度")) or last_period
        grade = norm_grade(row.get("牌号")) or last_grade
        if period:
            last_period = period
        if grade:
            last_grade = grade
        raw_tons = row.get("需求吨位") if "需求吨位" in row else row.get("吨位")
        if raw_tons is None or norm_text(raw_tons) == "":
            continue
        if grade not in PRIORITY_GRADES:
            raise ModelInputError(f"保供量第{idx}行牌号不允许：{grade}")
        spec = as_spec(first_present(row, ["规格"], "规格"))
        if grade == "T63E/E/G" and spec not in T63E_SPECS:
            raise ModelInputError(f"T63E/E/G 只允许 12 至 32 规格，当前规格：{spec}")
        if grade == "T63/E/G" and spec not in T63_SPEC:
            raise ModelInputError(f"T63/E/G 只允许 10 规格，当前规格：{spec}")
        tons = as_float(raw_tons, f"保供量第{idx}行需求吨位")
        result.append(DemandRow(period, grade, spec, tons_to_units(tons, f"保供量第{idx}行需求吨位") * UNIT_TONS))
    return result


def parse_ratios(wb, resource_ws) -> list[RatioRow]:
    ws = find_sheet(wb, ["400兆帕规格比例约束", "规格比例约束"], required=False)
    if ws is not None:
        ratios = []
        for idx, row in enumerate(sheet_rows(ws, ["牌号", "规格", "比例下限", "比例上限"]), start=2):
            grade_group = norm_text(row.get("牌号"))
            if grade_group and "400" not in grade_group and "HRB400" not in grade_group:
                continue
            if row.get("规格") is None or row.get("比例下限") is None or row.get("比例上限") is None:
                continue
            ratios.append(
                RatioRow(
                    as_spec(first_present(row, ["规格"], "规格")),
                    as_float(first_present(row, ["比例下限", "下限"], "比例下限"), f"400兆帕规格比例约束第{idx}行比例下限"),
                    as_float(first_present(row, ["比例上限", "上限"], "比例上限"), f"400兆帕规格比例约束第{idx}行比例上限"),
                )
            )
        if ratios:
            return ratios

    ratios = []
    for row in resource_ws.iter_rows(values_only=True):
        cells = list(row)
        if len(cells) >= 4 and isinstance(cells[1], (int, float)) and isinstance(cells[2], (int, float)) and isinstance(cells[3], (int, float)):
            ratios.append(RatioRow(as_spec(cells[1]), float(cells[2]), float(cells[3])))
    if not ratios:
        raise ModelInputError("缺少 400兆帕规格比例约束，且无法从旬度资源说明读取规格比例")
    return ratios


def parse_forecast(ws, default_period: str) -> list[ForecastRow]:
    if ws is None:
        return []
    result = []
    for idx, row in enumerate(sheet_rows(ws, ["牌号", "规格"]), start=2):
        raw_tons = row.get("客户预报吨位")
        if raw_tons is None:
            raw_tons = row.get("预报吨位")
        if raw_tons is None:
            raw_tons = row.get("需求吨位")
        if raw_tons is None or norm_text(raw_tons) == "":
            continue
        grade = norm_grade(first_present(row, ["牌号"], "牌号"))
        if grade not in FOUR_HUNDRED_GRADES:
            raise ModelInputError(f"400兆帕客户需求预报第{idx}行牌号只能是 HRB400 或 HRB400E")
        result.append(
            ForecastRow(
                norm_text(row.get("旬度")) or default_period,
                grade,
                as_spec(first_present(row, ["规格"], "规格")),
                as_float(raw_tons, f"400兆帕客户需求预报第{idx}行客户预报吨位"),
            )
        )
    return result


def compute_foreign_days(foreign: list[ForeignRow], efficiency: dict[tuple[str, str, int], float]) -> dict[tuple[str, str], float]:
    days = defaultdict(float)
    missing = []
    for row in foreign:
        key = (row.line, row.grade, row.spec)
        if key not in efficiency:
            missing.append(f"产线={row.line}，牌号={row.grade}，规格={row.spec}")
            continue
        days[(row.period, row.line)] += row.tons / efficiency[key]
    if missing:
        unique_missing = sorted(set(missing))
        detail = "；".join(unique_missing[:20])
        if len(unique_missing) > 20:
            detail += f"；另有{len(unique_missing) - 20}项"
        raise ModelInputError(f"外贸排产缺少生产效率：{detail}")
    return dict(days)


def eligible_lines(grade: str, spec: int, efficiency: dict[tuple[str, str, int], float]) -> list[str]:
    return sorted({line for line, g, s in efficiency if g == grade and s == spec})


def solve_period(
    resource: ResourceRow,
    calendar: dict[tuple[str, str], float],
    foreign_days: dict[tuple[str, str], float],
    efficiency: dict[tuple[str, str, int], float],
    cashflow: dict[tuple[str, int], float],
    demands: list[DemandRow],
    ratios: list[RatioRow],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    period = resource.period
    period_lines = sorted({line for p, line in calendar if p == period})
    if not period_lines:
        raise ModelInputError(f"生产日历中没有旬度 {period} 的产线天数")

    demand_for_period = [d for d in demands if d.period == period]
    if (resource.hrb500e > 0 or resource.six_hundred > 0) and not demand_for_period:
        raise ModelInputError(f"{period} 缺少 HRB500E及600兆帕需求表 数据")

    hrb500e_sum = sum(d.tons for d in demand_for_period if d.grade == "HRB500E")
    six_sum = sum(d.tons for d in demand_for_period if d.grade in SIX_HUNDRED_GRADES)
    if hrb500e_sum != resource.hrb500e:
        raise ModelInputError(f"{period} HRB500E需求合计 {hrb500e_sum} 吨，与保供量 {resource.hrb500e} 吨不一致")
    if six_sum != resource.six_hundred:
        raise ModelInputError(f"{period} 600兆帕需求合计 {six_sum} 吨，与保供量 {resource.six_hundred} 吨不一致")

    total_domestic = resource.hrb500e + resource.six_hundred + resource.four_hundred
    if total_domestic != resource.domestic:
        raise ModelInputError(f"{period} 内贸分项合计 {total_domestic} 吨，与内贸资源量 {resource.domestic} 吨不一致")

    four_units = tons_to_units(resource.four_hundred, f"{period} 400兆帕资源量")
    model = pulp.LpProblem(f"domestic_schedule_{period}", pulp.LpMaximize)
    variables: dict[tuple[str, int, str], pulp.LpVariable] = {}

    def add_var(grade: str, spec: int, line: str) -> pulp.LpVariable:
        name = f"x_{period}_{grade}_{spec}_{line}".replace("/", "_")
        var = pulp.LpVariable(name, lowBound=0, cat=pulp.LpInteger)
        variables[(grade, spec, line)] = var
        return var

    for demand in demand_for_period:
        lines = eligible_lines(demand.grade, demand.spec, efficiency)
        if not lines:
            raise ModelInputError(f"{period} 需求无可排产线：牌号={demand.grade}，规格={demand.spec}")
        demand_vars = [add_var(demand.grade, demand.spec, line) for line in lines]
        model += pulp.lpSum(demand_vars) == tons_to_units(demand.tons, f"{period} {demand.grade} {demand.spec}需求吨位")

    ratio_specs = sorted({r.spec for r in ratios})
    for grade in sorted(FOUR_HUNDRED_GRADES):
        for spec in ratio_specs:
            if (grade, spec) not in cashflow:
                raise ModelInputError(f"现金流量表缺少：牌号={grade}，规格={spec}")
            lines = eligible_lines(grade, spec, efficiency)
            if not lines:
                continue
            for line in lines:
                add_var(grade, spec, line)

    four_vars = [var for (grade, _spec, _line), var in variables.items() if grade in FOUR_HUNDRED_GRADES]
    model += pulp.lpSum(four_vars) == four_units

    for ratio in ratios:
        if not 0 <= ratio.lower <= ratio.upper <= 1:
            raise ModelInputError(f"规格 {ratio.spec} 的比例上下限不合法")
        spec_vars = [var for (grade, spec, _line), var in variables.items() if grade in FOUR_HUNDRED_GRADES and spec == ratio.spec]
        min_units = math.ceil(ratio.lower * four_units - 1e-9)
        max_units = math.floor(ratio.upper * four_units + 1e-9)
        if not spec_vars and min_units > 0:
            raise ModelInputError(f"400兆帕规格 {ratio.spec} 有比例下限，但没有可排产线")
        model += pulp.lpSum(spec_vars) >= min_units
        model += pulp.lpSum(spec_vars) <= max_units

    e_vars = [var for (grade, _spec, _line), var in variables.items() if grade == "HRB400E"]
    model += pulp.lpSum(e_vars) >= math.ceil(0.65 * four_units - 1e-9)

    for line in period_lines:
        capacity_days = calendar[(period, line)]
        fixed_days = foreign_days.get((period, line), 0.0)
        if fixed_days - capacity_days > 1e-8:
            raise ModelInputError(f"{period} {line} 外贸已占用 {fixed_days:.2f} 天，超过生产天数 {capacity_days:.2f} 天")
        line_terms = []
        for (grade, spec, var_line), var in variables.items():
            if var_line != line:
                continue
            daily = efficiency[(line, grade, spec)]
            line_terms.append(var * UNIT_TONS / daily)
        model += pulp.lpSum(line_terms) <= capacity_days - fixed_days

    model += pulp.lpSum(
        var * UNIT_TONS * cashflow.get((grade, spec), 0.0)
        for (grade, spec, _line), var in variables.items()
        if grade in FOUR_HUNDRED_GRADES
    )

    status = model.solve(pulp.PULP_CBC_CMD(msg=False))
    if pulp.LpStatus[status] != "Optimal":
        raise ModelInputError(f"{period} 无可行排产方案：{pulp.LpStatus[status]}。请检查产线天数、需求量、规格比例或日产量。")

    plan = []
    for (grade, spec, line), var in sorted(variables.items(), key=lambda item: (item[0][0], item[0][1], item[0][2])):
        value = int(round(var.value() or 0))
        if value > 0:
            plan.append({"旬度": period, "牌号": grade, "规格": spec, "吨位": value * UNIT_TONS, "生产条线": line})

    checks = []
    used_days = defaultdict(float)
    for row in plan:
        used_days[row["生产条线"]] += row["吨位"] / efficiency[(row["生产条线"], row["牌号"], row["规格"])]
    for line in period_lines:
        foreign = foreign_days.get((period, line), 0.0)
        domestic = used_days.get(line, 0.0)
        total = foreign + domestic
        capacity = calendar[(period, line)]
        checks.append({"校验项": "产线生产天数", "旬度": period, "产线": line, "结果": "通过" if total <= capacity + 1e-7 else "超限", "数值": round(total, 4), "说明": f"可用{capacity:.4f}天，外贸{foreign:.4f}天，内贸{domestic:.4f}天"})

    by_grade = defaultdict(int)
    by_spec = defaultdict(int)
    total_cash = 0.0
    for row in plan:
        by_grade[row["牌号"]] += row["吨位"]
        if row["牌号"] in FOUR_HUNDRED_GRADES:
            by_spec[row["规格"]] += row["吨位"]
            total_cash += row["吨位"] * cashflow[(row["牌号"], row["规格"])]

    checks.extend(
        [
            {"校验项": "HRB500E需求", "旬度": period, "产线": "", "结果": "通过", "数值": by_grade["HRB500E"], "说明": f"保供量{resource.hrb500e}吨"},
            {"校验项": "600兆帕需求", "旬度": period, "产线": "", "结果": "通过", "数值": by_grade["T63E/E/G"] + by_grade["T63/E/G"], "说明": f"保供量{resource.six_hundred}吨"},
            {"校验项": "400兆帕资源", "旬度": period, "产线": "", "结果": "通过", "数值": by_grade["HRB400"] + by_grade["HRB400E"], "说明": f"资源量{resource.four_hundred}吨"},
            {"校验项": "HRB400E抗震占比", "旬度": period, "产线": "", "结果": "通过", "数值": round(by_grade["HRB400E"] / resource.four_hundred, 4) if resource.four_hundred else 0, "说明": "下限0.65"},
            {"校验项": "400兆帕总现金流", "旬度": period, "产线": "", "结果": "通过", "数值": round(total_cash, 4), "说明": ""},
        ]
    )
    for ratio in ratios:
        value = by_spec[ratio.spec] / resource.four_hundred if resource.four_hundred else 0
        checks.append({"校验项": f"400兆帕规格{ratio.spec}比例", "旬度": period, "产线": "", "结果": "通过", "数值": round(value, 4), "说明": f"范围{ratio.lower}-{ratio.upper}"})
    return plan, checks


def build_suggestions(plan: list[dict[str, Any]], forecasts: list[ForecastRow]) -> list[dict[str, Any]]:
    plan_sum = defaultdict(int)
    for row in plan:
        if row["牌号"] in FOUR_HUNDRED_GRADES:
            plan_sum[(row["旬度"], row["牌号"], row["规格"])] += row["吨位"]
    suggestions = []
    for item in forecasts:
        planned = plan_sum[(item.period, item.grade, item.spec)]
        discount = planned / item.tons if item.tons else None
        if item.tons <= 0:
            reason = "客户预报吨位为0，无法计算折扣"
        elif planned >= item.tons:
            reason = "方案可覆盖客户预报量"
        elif planned == 0:
            reason = "方案未安排该牌号规格，通常受现金流、比例、抗震占比或产线能力约束影响"
        else:
            reason = "方案低于客户预报，建议按折扣承接；原方案保持现金流最优和约束达标"
        suggestions.append(
            {
                "旬度": item.period,
                "牌号": item.grade,
                "规格": item.spec,
                "客户预报吨位": item.tons,
                "排产方案吨位": planned,
                "折扣": round(discount, 4) if discount is not None else "",
                "简要原因": reason,
            }
        )
    return suggestions


def write_output(path: Path, plan: list[dict[str, Any]], checks: list[dict[str, Any]], suggestions: list[dict[str, Any]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "内贸排产方案"

    def write_table(sheet, rows: list[dict[str, Any]], headers: list[str]) -> None:
        sheet.append(headers)
        for row in rows:
            sheet.append([row.get(header, "") for header in headers])
        for column_cells in sheet.columns:
            width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 10), 45)

    write_table(ws, plan, ["旬度", "牌号", "规格", "吨位", "生产条线"])
    write_table(wb.create_sheet("方案校验"), checks, ["校验项", "旬度", "产线", "结果", "数值", "说明"])
    write_table(wb.create_sheet("400兆帕需求建议"), suggestions, ["旬度", "牌号", "规格", "客户预报吨位", "排产方案吨位", "折扣", "简要原因"])
    wb.save(path)


def run(input_path: Path, output_path: Path) -> None:
    wb = load_workbook(input_path, data_only=True)
    resource_ws = find_sheet(wb, ["旬度资源说明"])
    resources = parse_resources(resource_ws)
    default_period = resources[0].period

    calendar = parse_calendar(find_sheet(wb, ["生产日历"]), default_period)
    _line_capacity = parse_line_capacity(find_sheet(wb, ["产线基础信息"]))
    foreign = parse_foreign(find_sheet(wb, ["外贸排产明细"]))
    efficiency = parse_efficiency(find_sheet(wb, ["产品生产效率表"]))
    cashflow = parse_cashflow(find_sheet(wb, ["现金流量表"]))
    demands = parse_demands(find_sheet(wb, ["HRB500E及600兆帕需求表", "保供量"], required=False))
    ratios = parse_ratios(wb, resource_ws)
    forecasts = parse_forecast(find_sheet(wb, ["400兆帕客户需求预报"], required=False), default_period)
    foreign_days = compute_foreign_days(foreign, efficiency)

    all_plan = []
    all_checks = []
    for resource in resources:
        plan, checks = solve_period(resource, calendar, foreign_days, efficiency, cashflow, demands, ratios)
        all_plan.extend(plan)
        all_checks.extend(checks)

    suggestions = build_suggestions(all_plan, forecasts)
    write_output(output_path, all_plan, all_checks, suggestions)


def main() -> None:
    parser = argparse.ArgumentParser(description="旬度内贸排产优化模型")
    parser.add_argument("--input", required=True, help="输入 Excel 文件路径")
    parser.add_argument("--output", required=True, help="输出 Excel 文件路径")
    args = parser.parse_args()

    try:
        run(Path(args.input), Path(args.output))
    except ModelInputError as exc:
        print(f"输入数据错误：{exc}", file=sys.stderr)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
