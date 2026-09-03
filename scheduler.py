"""排课数据模型、校验与约束求解。"""

from __future__ import annotations

import json
import time
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

from ortools.sat.python import cp_model


DAYS = ["星期一", "星期二", "星期三", "星期四", "星期五"]
SUBJECT_COLUMNS = 16


def _clean_subject(name):
    if name is None:
        return ""
    return str(name).replace("\n", "").strip()


def slot_key(day_index, period_index):
    return f"{day_index}-{period_index}"


def parse_slot(key):
    day, period = key.split("-")
    return int(day), int(period)


class ScheduleData:
    """从工作簿整理出来的排课基础数据。"""

    def __init__(self):
        self.classes = []
        self.subjects = []
        self.hours = {}
        self.teachers = {}
        self.periods = []
        self.section_spans = []
        self.days = DAYS
        self.warnings = []

    @property
    def n_slots(self):
        return len(self.periods)

    @property
    def teacher_list(self):
        return sorted({t for t in self.teachers.values() if t})

    def teacher_loads(self):
        loads = Counter()
        for (cls, subject), teacher in self.teachers.items():
            if teacher and self.hours.get((cls, subject)):
                loads[teacher] += self.hours[(cls, subject)]
        return dict(loads)

    def to_dict(self):
        return {
            "classes": self.classes,
            "subjects": self.subjects,
            "teachers": self.teacher_list,
            "teacherLoads": self.teacher_loads(),
            "days": self.days,
            "periods": self.periods,
            "sectionSpans": self.section_spans,
            "nSlots": self.n_slots,
            "availableSlots": len(self.days) * self.n_slots,
            "hours": {c: {s: self.hours[(c, s)] for s in self.subjects} for c in self.classes},
            "teacherMap": {c: {s: self.teachers.get((c, s), "") for s in self.subjects} for c in self.classes},
            "warnings": self.warnings,
        }


def load_schedule_data(path):
    """读取基本数据.xlsx 的三张表。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws_assign = wb["各班任课"]
    ws_hours = wb["各班课时"]
    ws_schedule = wb["作息表"]

    data = ScheduleData()
    subjects = []
    for col in range(2, 2 + SUBJECT_COLUMNS):
        name = _clean_subject(ws_assign.cell(row=1, column=col).value)
        subjects.append(name)
    data.subjects = subjects

    for row in range(2, ws_assign.max_row + 1):
        cls = ws_assign.cell(row=row, column=1).value
        if cls is None or str(cls).strip() == "":
            continue
        cls = str(cls).strip()
        if cls.startswith("~"):
            continue
        data.classes.append(cls)
        for idx, subject in enumerate(subjects):
            teacher = ws_assign.cell(row=row, column=2 + idx).value
            if teacher is not None and str(teacher).strip():
                data.teachers[(cls, subject)] = str(teacher).strip()

    for row in range(2, ws_hours.max_row + 1):
        cls = ws_hours.cell(row=row, column=1).value
        if cls is None or str(cls).strip() == "":
            continue
        cls = str(cls).strip()
        for idx, subject in enumerate(subjects):
            value = ws_hours.cell(row=row, column=2 + idx).value
            data.hours[(cls, subject)] = int(value) if isinstance(value, (int, float)) else 0

    data.days = []
    for col in range(4, ws_schedule.max_column + 1):
        day = ws_schedule.cell(row=1, column=col).value
        if day is not None and str(day).strip():
            data.days.append(str(day).strip())
    if not data.days:
        data.days = list(DAYS[:5])

    for row in range(2, ws_schedule.max_row + 1):
        period_label = ws_schedule.cell(row=row, column=3).value
        time_label = ws_schedule.cell(row=row, column=2).value
        section_label = ws_schedule.cell(row=row, column=1).value
        if period_label is None and time_label is None and section_label is None:
            continue
        data.periods.append(
            {
                "period": str(period_label).strip() if period_label is not None else f"第{len(data.periods) + 1}节",
                "time": str(time_label).strip() if time_label is not None else "",
                "section": str(section_label).strip() if section_label is not None else "",
            }
        )

    if not data.classes:
        raise ValueError("各班任课表没有读取到班级数据")
    if not data.periods:
        raise ValueError("作息表没有读取到节次数据")

    merged_rows = {}
    for rng in ws_schedule.merged_cells.ranges:
        if rng.min_col == 1 and rng.max_col == 1:
            for row in range(rng.min_row, rng.max_row + 1):
                merged_rows[row] = (rng.min_row, rng.max_row)
    for idx, row in enumerate(range(2, 2 + len(data.periods))):
        if row in merged_rows:
            start, end = merged_rows[row]
            data.section_spans.append(end - start + 1 if row == start else 0)
        else:
            data.section_spans.append(1 if data.periods[idx]["section"] else 0)

    for cls in data.classes:
        total = sum(data.hours[(cls, subject)] for subject in data.subjects)
        if total > len(data.days) * len(data.periods):
            data.warnings.append(f"{cls} 周课时 {total} 超过作息表可用空位")

    wb.close()
    return data


def _valid_slot(key, data):
    try:
        day, period = parse_slot(key)
    except (ValueError, TypeError):
        return False
    return 0 <= day < len(data.days) and 0 <= period < data.n_slots


def normalize_settings(raw, data=None):
    """把前端传来的设置规整成内部结构。"""
    raw = raw or {}
    rules = raw.get("rules") or {}

    def clamp_int(value, default, low, high):
        try:
            value = int(value)
        except (TypeError, ValueError):
            return default
        return max(low, min(high, value))

    settings = {
        "rules": {
            "dailyMax": clamp_int(rules.get("dailyMax"), 5, 1, 12),
            "consecutiveMax": clamp_int(rules.get("consecutiveMax"), 2, 1, 12),
        },
        "classBlocked": {},
        "fixed": {},
        "manual": {},
        "teacherUnavailable": {},
        "subjectConstraints": {},
        "titles": {"classSuffix": "", "teacherSuffix": ""},
        "preferredConsecutive": {},
    }

    def collect_slots(value):
        if value is None or isinstance(value, str):
            return set()
        result = set()
        try:
            items = list(value)
        except TypeError:
            return set()
        for item in items:
            if data is None or _valid_slot(item, data):
                result.add(str(item))
        return result

    for cls, slots in (raw.get("classBlocked") or {}).items():
        if data is not None and cls not in data.classes:
            continue
        settings["classBlocked"][cls] = collect_slots(slots)

    for cls, mapping in (raw.get("fixed") or {}).items():
        if data is not None and cls not in data.classes:
            continue
        clean = {}
        for key, subject in (mapping or {}).items():
            if data is not None and not _valid_slot(key, data):
                continue
            subject = _clean_subject(subject)
            if data is not None and subject not in data.subjects:
                continue
            clean[key] = subject
        settings["fixed"][cls] = clean

    for cls, mapping in (raw.get("manual") or {}).items():
        if data is not None and cls not in data.classes:
            continue
        clean = {}
        for key, subject in (mapping or {}).items():
            if data is not None and not _valid_slot(key, data):
                continue
            subject = _clean_subject(subject)
            if data is not None and subject not in data.subjects:
                continue
            clean[key] = subject
        settings["manual"][cls] = clean

    for teacher, slots in (raw.get("teacherUnavailable") or {}).items():
        if data is not None and teacher not in data.teacher_list:
            continue
        settings["teacherUnavailable"][teacher] = collect_slots(slots)

    for subject, slots in (raw.get("subjectConstraints") or {}).items():
        if data is not None and subject not in data.subjects:
            continue
        settings["subjectConstraints"][subject] = collect_slots(slots)

    titles = raw.get("titles") or {}
    settings["titles"] = {
        "classSuffix": str(titles.get("classSuffix") or "").strip()[:80],
        "teacherSuffix": str(titles.get("teacherSuffix") or "").strip()[:80],
    }

    preferred = raw.get("preferredConsecutive") or {}
    for subject, value in preferred.items():
        if data is not None and subject not in data.subjects:
            continue
        try:
            n = int(value)
        except (TypeError, ValueError):
            continue
        settings["preferredConsecutive"][subject] = max(1, min(12, n))

    return settings


def merge_locked(settings):
    """固定课和手动锁定课合并成自动排课必须保留的课程。"""
    locked = {}
    for cls in set(settings["fixed"]) | set(settings["manual"]):
        merged = dict(settings["fixed"].get(cls, {}))
        merged.update(settings["manual"].get(cls, {}))
        locked[cls] = merged
    return locked


def validate_settings(data, settings):
    """检查设置是否自相矛盾或无法满足。"""
    errors = []
    n_slots = data.n_slots * len(data.days)

    for cls in settings["classBlocked"]:
        if cls not in data.classes:
            errors.append(f"班级 {cls} 不存在")

    locked = merge_locked(settings)
    for cls in locked:
        if cls not in data.classes:
            errors.append(f"班级 {cls} 不存在")
            continue
        for slot, subject in locked[cls].items():
            if not _valid_slot(slot, data):
                errors.append(f"{cls} 存在无效时间格 {slot}")
                continue
            if slot in settings["classBlocked"].get(cls, set()):
                errors.append(f"{cls} 的 {slot} 同时被设为不排课和固定课")
            if not data.teachers.get((cls, subject)):
                errors.append(f"{cls} 的 {subject} 没有任课教师，不能设为固定课")
        if cls in settings["fixed"] and cls in settings["manual"]:
            overlap = set(settings["fixed"][cls]) & set(settings["manual"][cls])
            if overlap:
                errors.append(f"{cls} 的固定课与手动锁定重复：{', '.join(sorted(overlap))}")

    for teacher in settings["teacherUnavailable"]:
        if teacher not in data.teacher_list:
            errors.append(f"教师 {teacher} 不存在")

    for subject, slots in settings["subjectConstraints"].items():
        if subject not in data.subjects:
            errors.append(f"科目 {subject} 不存在")
        for slot in slots:
            if not _valid_slot(slot, data):
                errors.append(f"科目 {subject} 存在无效时间格 {slot}")

    fixed_conflicts = Counter()
    for cls, slots in locked.items():
        for slot, subject in slots.items():
            teacher = data.teachers.get((cls, subject))
            if teacher:
                fixed_conflicts[(teacher, slot)] += 1
    for (teacher, slot), count in fixed_conflicts.items():
        if count > 1:
            errors.append(f"教师 {teacher} 在 {slot} 同时有 {count} 个固定课")

    for cls, slots in locked.items():
        for slot, subject in slots.items():
            teacher = data.teachers.get((cls, subject))
            if teacher and slot in settings["teacherUnavailable"].get(teacher, set()):
                errors.append(f"教师 {teacher} 在 {slot} 被设为不排课，但 {cls} 固定了 {subject}")
            if slot in settings["subjectConstraints"].get(subject, set()):
                errors.append(f"{cls} 的 {subject} 固定课落在科目时段约束 {slot}")

    daily_max = settings["rules"]["dailyMax"]
    consecutive_max = settings["rules"]["consecutiveMax"]
    for subject, count in settings["preferredConsecutive"].items():
        if subject not in data.subjects:
            errors.append(f"优先连堂科目 {subject} 不存在")
        if count > consecutive_max:
            errors.append(
                f"{subject} 优先连堂 {count} 节，超过同科连堂上限 {consecutive_max} 节"
            )
    for cls in data.classes:
        fixed_counts = Counter()
        for slot, subject in locked.get(cls, {}).items():
            fixed_counts[subject] += 1
        for subject in data.subjects:
            required = data.hours[(cls, subject)]
            if fixed_counts[subject] > required:
                errors.append(f"{cls} 的 {subject} 固定课 {fixed_counts[subject]} 节，超过周课时 {required} 节")
        for day in range(len(data.days)):
            for subject in data.subjects:
                count = sum(
                    1
                    for slot, sub in locked.get(cls, {}).items()
                    if parse_slot(slot)[0] == day and sub == subject
                )
                if count > daily_max:
                    errors.append(
                        f"{cls} {data.days[day]} 的 {subject} 固定了 {count} 节，超过同科同天上限 {daily_max}"
                    )

        available = n_slots - len(settings["classBlocked"].get(cls, set())) - len(locked.get(cls, {}))
        remaining = sum(data.hours[(cls, s)] for s in data.subjects) - sum(fixed_counts.values())
        if remaining < 0:
            errors.append(f"{cls} 固定课总数超过周课时")
        if remaining > available:
            errors.append(
                f"{cls} 剩余需排 {remaining} 节，但只有 {available} 个可用空格，请增加不排课或减少固定课"
            )
        for subject in data.subjects:
            need = data.hours[(cls, subject)] - fixed_counts[subject]
            if need <= 0:
                continue
            teacher = data.teachers.get((cls, subject))
            if not teacher:
                continue
            possible = 0
            for day in range(len(data.days)):
                for period in range(data.n_slots):
                    key = slot_key(day, period)
                    if key in settings["classBlocked"].get(cls, set()):
                        continue
                    if key in locked.get(cls, {}):
                        continue
                    if key in settings["teacherUnavailable"].get(teacher, set()):
                        continue
                    if key in settings["subjectConstraints"].get(subject, set()):
                        continue
                    possible += 1
            if possible < need:
                errors.append(
                    f"{cls} 的 {subject} 还差 {need} 节，但教师 {teacher} 可用的时间格只有 {possible} 个"
                )

    if consecutive_max > data.n_slots:
        errors.append(f"连堂上限 {consecutive_max} 超过每天节数 {data.n_slots}")
    return sorted(set(errors))


def _infeasible_suggestions(data, settings):
    """根据设置生成无解时的关键建议。"""
    suggestions = []
    daily_max = settings["rules"]["dailyMax"]
    consecutive_max = settings["rules"]["consecutiveMax"]
    if daily_max <= 2:
        suggestions.append("同科同天最多限制较严，可尝试提高到 3-5 节。")
    if consecutive_max <= 1:
        suggestions.append("连堂上限为 1，可尝试提高到 2 节，通常更容易排通。")
    if any(settings["teacherUnavailable"].values()):
        total_off = sum(len(v) for v in settings["teacherUnavailable"].values())
        suggestions.append(f"教师不排课时段共 {total_off} 个，可尝试减少一些不排课时段。")
    if any(settings["subjectConstraints"].values()):
        suggestions.append("科目时段约束较多，可尝试减少某些科目不出现的时段。")
    if settings["preferredConsecutive"]:
        suggestions.append("优先连堂会增加求解难度，可暂时取消优先连堂后再排一次。")
    if any(settings["classBlocked"].values()):
        blocked_total = sum(len(v) for v in settings["classBlocked"].values())
        suggestions.append(f"当前共设置 {blocked_total} 个班级不排课时间格，可尝试减少。")
    suggestions.append("也可以点击“清空本次排课”后用新的随机方案重试。")
    return suggestions


def solve_schedule(data, settings, variant=None):
    """用 CP-SAT 求解课表。"""
    settings = normalize_settings(settings, data)
    errors = validate_settings(data, settings)
    if errors:
        return {
            "ok": False,
            "errorKind": "validation",
            "errors": errors,
            "report": errors,
        }

    locked = merge_locked(settings)
    remaining = {}
    for cls in data.classes:
        fixed_counts = Counter()
        for slot, subject in locked.get(cls, {}).items():
            fixed_counts[subject] += 1
        remaining[cls] = {
            subject: data.hours[(cls, subject)] - fixed_counts[subject]
            for subject in data.subjects
        }

    model = cp_model.CpModel()
    assign = {}
    slot_vars = {}

    for cls in data.classes:
        for day in range(len(data.days)):
            for period in range(data.n_slots):
                key = slot_key(day, period)
                if key in settings["classBlocked"].get(cls, set()):
                    continue
                if key in locked.get(cls, {}):
                    continue
                vars_for_slot = []
                for subject in data.subjects:
                    need = remaining[cls][subject]
                    if need <= 0:
                        continue
                    teacher = data.teachers.get((cls, subject))
                    if not teacher:
                        continue
                    if key in settings["teacherUnavailable"].get(teacher, set()):
                        continue
                    if key in settings["subjectConstraints"].get(subject, set()):
                        continue
                    var = model.NewBoolVar(f"a_{cls}_{day}_{period}_{subject}")
                    assign[(cls, day, period, subject)] = var
                    vars_for_slot.append(var)
                if vars_for_slot:
                    slot_vars[(cls, day, period)] = vars_for_slot
                    model.Add(sum(vars_for_slot) <= 1)

    for cls in data.classes:
        for subject in data.subjects:
            need = remaining[cls][subject]
            if need <= 0:
                continue
            vars_for_count = [
                assign[(cls, day, period, subject)]
                for day in range(len(data.days))
                for period in range(data.n_slots)
                if (cls, day, period, subject) in assign
            ]
            model.Add(sum(vars_for_count) == need)

    for teacher in data.teacher_list:
        for day in range(len(data.days)):
            for period in range(data.n_slots):
                key = slot_key(day, period)
                fixed_count = 0
                for cls, slots in locked.items():
                    subject = slots.get(key)
                    if subject and data.teachers.get((cls, subject)) == teacher:
                        fixed_count += 1
                vars_for_teacher = [
                    assign[(cls, day, period, subject)]
                    for cls in data.classes
                    for subject in data.subjects
                    if data.teachers.get((cls, subject)) == teacher
                    and (cls, day, period, subject) in assign
                ]
                model.Add(sum(vars_for_teacher) + fixed_count <= 1)

    daily_max = settings["rules"]["dailyMax"]
    for cls in data.classes:
        for day in range(len(data.days)):
            for subject in data.subjects:
                fixed_count = sum(
                    1
                    for slot, sub in locked.get(cls, {}).items()
                    if parse_slot(slot)[0] == day and sub == subject
                )
                vars_for_day = [
                    assign[(cls, day, period, subject)]
                    for period in range(data.n_slots)
                    if (cls, day, period, subject) in assign
                ]
                model.Add(sum(vars_for_day) + fixed_count <= daily_max)

    consecutive_max = settings["rules"]["consecutiveMax"]
    for cls in data.classes:
        for day in range(len(data.days)):
            for subject in data.subjects:
                for start in range(0, data.n_slots - consecutive_max):
                    window = range(start, start + consecutive_max + 1)
                    fixed_count = sum(
                        1
                        for slot, sub in locked.get(cls, {}).items()
                        if parse_slot(slot)[0] == day
                        and parse_slot(slot)[1] in window
                        and sub == subject
                    )
                    vars_for_window = [
                        assign[(cls, day, period, subject)]
                        for period in window
                        if (cls, day, period, subject) in assign
                    ]
                    model.Add(sum(vars_for_window) + fixed_count <= consecutive_max)

    pair_vars = []
    for cls in data.classes:
        for day in range(len(data.days)):
            for subject, preferred_count in settings["preferredConsecutive"].items():
                if preferred_count < 2:
                    continue
                for period in range(data.n_slots - 1):
                    left = assign.get((cls, day, period, subject))
                    right = assign.get((cls, day, period + 1, subject))
                    if left is None or right is None:
                        continue
                    pair = model.NewBoolVar(f"pair_{cls}_{day}_{period}_{subject}")
                    model.Add(pair <= left)
                    model.Add(pair <= right)
                    model.Add(left + right <= pair + 1)
                    pair_vars.append(pair)
    if pair_vars:
        model.Maximize(sum(pair_vars))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0
    solver.parameters.num_search_workers = 8
    seed = variant if variant is not None else int(time.time_ns() % (2**31))
    solver.parameters.random_seed = seed
    status = solver.Solve(model)

    if status == cp_model.INFEASIBLE:
        suggestions = _infeasible_suggestions(data, settings)
        return {
            "ok": False,
            "errorKind": "infeasible",
            "errors": ["当前约束下没有可行课表。"] + suggestions,
            "report": ["当前约束下没有可行课表。"] + suggestions,
        }
    if status == cp_model.UNKNOWN:
        return {
            "ok": False,
            "errorKind": "timeout",
            "errors": ["求解超时：60 秒内未找到可行课表，请稍后重试或检查约束设置。"],
            "report": ["求解超时：60 秒内未找到可行课表，请稍后重试或检查约束设置。"],
        }

    schedule = {}
    for cls in data.classes:
        cells = {}
        for day in range(len(data.days)):
            for period in range(data.n_slots):
                key = slot_key(day, period)
                if key in locked.get(cls, {}):
                    continue
                for subject in data.subjects:
                    var = assign.get((cls, day, period, subject))
                    if var is not None and solver.Value(var) == 1:
                        cells[key] = {
                            "subject": subject,
                            "teacher": data.teachers[(cls, subject)],
                            "source": "auto",
                        }
        for slot, subject in locked.get(cls, {}).items():
            source = "manual" if slot in settings["manual"].get(cls, {}) else "fixed"
            cells[slot] = {
                "subject": subject,
                "teacher": data.teachers.get((cls, subject)) or "",
                "source": source,
            }
        schedule[cls] = cells

    class_stats = []
    teacher_load = Counter()
    for cls in data.classes:
        counts = Counter()
        for slot, cell in schedule[cls].items():
            counts[cell["subject"]] += 1
            teacher_load[cell["teacher"]] += 1
        class_stats.append(
            {
                "class": cls,
                "scheduled": sum(counts.values()),
                "required": sum(data.hours[(cls, s)] for s in data.subjects),
                "empty": len(data.days) * data.n_slots - sum(counts.values()),
                "counts": {s: counts[s] for s in data.subjects},
            }
        )

    report = [
        f"排课成功：{len(data.classes)} 个班，每天 {data.n_slots} 节，每周共 {len(data.days) * data.n_slots} 个时间格。",
        f"规则：同科同天最多 {daily_max} 节，连堂最多 {consecutive_max} 节。",
    ]
    for item in class_stats:
        report.append(
            f"{item['class']}：已排 {item['scheduled']} 节 / 应排 {item['required']} 节，空 {item['empty']} 格"
        )

    return {
        "ok": True,
        "status": "OPTIMAL" if status == cp_model.OPTIMAL else "FEASIBLE",
        "schedule": schedule,
        "stats": {"classes": class_stats, "teacherLoads": dict(teacher_load)},
        "report": report,
        "errors": [],
    }


def validate_result(data, settings, result):
    """校验最终课表（包括手动拖动后的结果）。"""
    issues = []
    settings = normalize_settings(settings, data)
    locked = merge_locked(settings)

    for cls in data.classes:
        cells = result.get(cls, {})
        counts = Counter()
        for slot, cell in cells.items():
            if not _valid_slot(slot, data):
                issues.append(f"{cls} 出现无效时间格 {slot}")
                continue
            subject = cell.get("subject", "")
            teacher = cell.get("teacher", "")
            if not subject:
                issues.append(f"{cls} 时间格 {slot} 缺少科目")
                continue
            if slot in settings["classBlocked"].get(cls, set()):
                issues.append(f"{cls} 的 {slot} 标为不排课，但仍有课程")
            expected_teacher = data.teachers.get((cls, subject))
            if expected_teacher and teacher != expected_teacher:
                issues.append(f"{cls} 的 {slot} 教师 {teacher} 与任课表不一致")
            if slot in locked.get(cls, {}) and locked[cls][slot] != subject:
                issues.append(f"{cls} 的 {slot} 与锁定课程 {locked[cls][slot]} 不一致")
            counts[subject] += 1
            if teacher and slot in settings["teacherUnavailable"].get(teacher, set()):
                issues.append(f"教师 {teacher} 在 {slot} 不排课，但 {cls} 仍安排了课程")
            if slot in settings["subjectConstraints"].get(subject, set()):
                issues.append(f"{cls} 的 {subject} 出现在科目时段约束 {slot}")
        for subject in data.subjects:
            if counts[subject] != data.hours[(cls, subject)]:
                issues.append(
                    f"{cls} 的 {subject} 实际 {counts[subject]} 节，应排 {data.hours[(cls, subject)]} 节"
                )

    occupied = {}
    for cls in data.classes:
        for slot, cell in result.get(cls, {}).items():
            teacher = cell.get("teacher", "")
            if teacher:
                if (teacher, slot) in occupied:
                    issues.append(
                        f"教师 {teacher} 在 {slot} 同时出现在 {occupied[(teacher, slot)]} 和 {cls}"
                    )
                else:
                    occupied[(teacher, slot)] = cls

    daily_max = settings["rules"]["dailyMax"]
    consecutive_max = settings["rules"]["consecutiveMax"]
    for cls in data.classes:
        cells = result.get(cls, {})
        for day in range(len(data.days)):
            for subject in data.subjects:
                day_count = sum(
                    1
                    for slot, cell in cells.items()
                    if parse_slot(slot)[0] == day and cell.get("subject") == subject
                )
                if day_count > daily_max:
                    issues.append(
                        f"{cls} {data.days[day]} 的 {subject} 达到 {day_count} 节，超过上限 {daily_max}"
                    )
            sequence = [
                cells.get(slot_key(day, period), {}).get("subject", "")
                for period in range(data.n_slots)
            ]
            run = 1
            for idx in range(1, len(sequence)):
                if sequence[idx] and sequence[idx] == sequence[idx - 1]:
                    run += 1
                    if run > consecutive_max:
                        issues.append(
                            f"{cls} {data.days[day]} 的 {sequence[idx]} 连堂达到 {run} 节，超过上限 {consecutive_max}"
                        )
                        run = 1
                else:
                    run = 1
    return sorted(set(issues))


def settings_to_json(settings):
    return {
        "rules": settings["rules"],
        "classBlocked": {k: sorted(v) for k, v in settings["classBlocked"].items()},
        "fixed": settings["fixed"],
        "manual": settings["manual"],
        "teacherUnavailable": {k: sorted(v) for k, v in settings["teacherUnavailable"].items()},
        "subjectConstraints": {k: sorted(v) for k, v in settings["subjectConstraints"].items()},
        "titles": settings["titles"],
        "preferredConsecutive": settings["preferredConsecutive"],
    }


def load_settings(path, data=None):
    if Path(path).exists():
        try:
            raw = json.loads(Path(path).read_text(encoding="utf-8"))
            return normalize_settings(raw, data)
        except (ValueError, json.JSONDecodeError):
            pass
    return normalize_settings({}, data)


def save_settings(path, settings):
    Path(path).write_text(json.dumps(settings_to_json(settings), ensure_ascii=False, indent=2), encoding="utf-8")
