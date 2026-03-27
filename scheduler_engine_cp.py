from __future__ import annotations
from typing import Any, Callable, Optional
from ortools.sat.python import cp_model

DAY_ORDER = ["Mon","Tue","Wed","Thu","Fri"]
PE_SUBJECT = "保体"


def _normalize_subject_name(subject: str) -> str:
    return str(subject or "").replace("　", "").replace(" ", "").strip()


def _build_excluded_subject_keywords(excluded_subjects: set[str]) -> set[str]:
    keywords = {_normalize_subject_name(s) for s in excluded_subjects if _normalize_subject_name(s)}
    if PE_SUBJECT in keywords:
        keywords.update({"体育", "保健体育"})
    return keywords


def _is_excluded_subject(subject: str, excluded_keywords: set[str]) -> bool:
    normalized = _normalize_subject_name(subject)
    if not normalized:
        return False
    if normalized in excluded_keywords:
        return True
    return any(keyword in normalized for keyword in excluded_keywords)


class SchedulerError(RuntimeError):
    pass


def grade_of(class_name:str)->str:
    return class_name.split("-")[0]


def parse_slot(slot:str)->tuple[str,int]:
    d,p=slot.split("-")
    return d,int(p)


def build_slots(day_periods:dict[str,int])->list[str]:
    slots=[]
    for d in DAY_ORDER:
        for p in range(1,int(day_periods[d])+1):
            slots.append(f"{d}-{p}")
    return slots


def _normalize_fixed_assignment_rows(fixed_raw: list[dict]) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        t = str(a.get("slot", "")).strip()
        s = str(a.get("subject", "")).strip()
        if c and t and s:
            out.append({"class": c, "slot": t, "subject": s})
    return out


def _assign_fixed_lessons_to_non_tt_teachers(
    *,
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_unavailable: dict[str, set[str]],
    teacher_discouraged: Optional[dict[str, set[str]]] = None,
    time_limit_sec: float,
) -> tuple[dict[tuple[str, str], str], list[str]]:
    teacher_discouraged = teacher_discouraged or {}
    fixed_list = _normalize_fixed_assignment_rows(fixed_raw)
    if not fixed_list:
        return {}, []

    non_tt_caps = {
        (k, c, s): int(h)
        for (k, c, s), h in non_tt.items()
        if int(h) > 0
    }

    lessons: list[dict] = []
    issues: list[str] = []
    for i, a in enumerate(fixed_list):
        c = a["class"]
        t = a["slot"]
        s = a["subject"]
        raw_candidates = sorted(
            k for (k, cc, ss), h in non_tt_caps.items()
            if cc == c and ss == s and int(h) > 0
        )
        if not raw_candidates:
            continue

        candidates = [k for k in raw_candidates if t not in teacher_unavailable.get(k, set())]
        if not candidates:
            issues.append(
                f"- {c} {t} {s}: 主担当 { '・'.join(raw_candidates) } が全員不可コマ"
            )
            continue

        lessons.append(
            {
                "index": i,
                "class": c,
                "slot": t,
                "subject": s,
                "candidates": candidates,
            }
        )

    if issues:
        return {}, issues

    if not lessons:
        return {}, []

    model: Any = cp_model.CpModel()
    assign_vars: dict[tuple[int, str], Any] = {}

    for lesson in lessons:
        idx = int(lesson["index"])
        for teacher in lesson["candidates"]:
            assign_vars[(idx, teacher)] = model.NewBoolVar(f"fixed_teacher_{idx}_{teacher}")
        model.Add(sum(assign_vars[(idx, teacher)] for teacher in lesson["candidates"]) == 1)

    teacher_slot_map: dict[tuple[str, str], list[Any]] = {}
    teacher_class_subject_map: dict[tuple[str, str, str], list[Any]] = {}
    for lesson in lessons:
        idx = int(lesson["index"])
        c = str(lesson["class"])
        t = str(lesson["slot"])
        s = str(lesson["subject"])
        for teacher in lesson["candidates"]:
            var = assign_vars[(idx, teacher)]
            teacher_slot_map.setdefault((teacher, t), []).append(var)
            teacher_class_subject_map.setdefault((teacher, c, s), []).append(var)

    for vars_t in teacher_slot_map.values():
        model.Add(sum(vars_t) <= 1)

    for key, vars_kcs in teacher_class_subject_map.items():
        model.Add(sum(vars_kcs) <= non_tt_caps.get(key, 0))

    discouraged_terms: list[Any] = []
    for lesson in lessons:
        idx = int(lesson["index"])
        slot = str(lesson["slot"])
        for teacher in lesson["candidates"]:
            if slot in teacher_discouraged.get(teacher, set()):
                discouraged_terms.append(assign_vars[(idx, teacher)])

    if discouraged_terms:
        model.Minimize(sum(discouraged_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(3.0, min(15.0, float(time_limit_sec)))
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        forced_conflicts: list[str] = []
        forced_by_teacher_slot: dict[tuple[str, str], list[str]] = {}
        for lesson in lessons:
            candidates = list(lesson["candidates"])
            if len(candidates) != 1:
                continue
            teacher = candidates[0]
            slot = str(lesson["slot"])
            forced_by_teacher_slot.setdefault((teacher, slot), []).append(
                f"{lesson['class']} {slot} {lesson['subject']}"
            )
        for (teacher, slot), items in sorted(forced_by_teacher_slot.items()):
            if len(items) >= 2:
                forced_conflicts.append(
                    f"- {teacher} は {slot} に固定{len(items)}件（{' / '.join(items[:4])}）"
                )
            if len(forced_conflicts) >= 3:
                break

        issues = ["- 固定コマどうしの主担当割当が両立しません"]
        issues.extend(forced_conflicts)
        return {}, issues

    fixed_teacher_assignments: dict[tuple[str, str], str] = {}
    for lesson in lessons:
        idx = int(lesson["index"])
        c = str(lesson["class"])
        t = str(lesson["slot"])
        for teacher in lesson["candidates"]:
            if solver.Value(assign_vars[(idx, teacher)]) == 1:
                fixed_teacher_assignments[(c, t)] = teacher
                break

    return fixed_teacher_assignments, []


def _build_fixed_teacher_usage(
    fixed_teacher_assignments: dict[tuple[str, str], str],
) -> tuple[dict[str, set[str]], dict[tuple[str, str], int], dict[tuple[str, str], int]]:
    fixed_teacher_busy: dict[str, set[str]] = {}
    fixed_main_h: dict[tuple[str, str], int] = {}
    fixed_teacher_day_load: dict[tuple[str, str], int] = {}

    for (c, slot), teacher in fixed_teacher_assignments.items():
        fixed_teacher_busy.setdefault(teacher, set()).add(slot)
        fixed_main_h[(teacher, c)] = fixed_main_h.get((teacher, c), 0) + 1
        day = parse_slot(slot)[0]
        fixed_teacher_day_load[(teacher, day)] = fixed_teacher_day_load.get((teacher, day), 0) + 1

    return fixed_teacher_busy, fixed_main_h, fixed_teacher_day_load


def _build_teacher_day_variance_objective_terms(
    *,
    model: Any,
    classes: list[str],
    slots: list[str],
    teachers: list[str],
    y: dict[tuple[str, str, str], Any],
    z: dict[tuple[str, str, str], Any],
    fixed_teacher_day_load: Optional[dict[tuple[str, str], int]] = None,
) -> tuple[dict[tuple[str, str], Any], Any, int]:
    slots_by_day = {
        d: [t for t in slots if parse_slot(t)[0] == d]
        for d in DAY_ORDER
    }

    teacher_day_load: dict[tuple[str, str], Any] = {}
    variance_terms: list[Any] = []
    weekly_load_upper_bound = len(slots)
    fixed_teacher_day_load = fixed_teacher_day_load or {}

    for k in teachers:
        day_load_vars: list[Any] = []
        day_sq_vars: list[Any] = []

        for d in DAY_ORDER:
            day_slots = slots_by_day[d]
            day_upper_bound = len(day_slots)
            fixed_load = int(fixed_teacher_day_load.get((k, d), 0))
            load_var = model.NewIntVar(fixed_load, day_upper_bound, f"teacher_day_load_{k}_{d}")
            day_terms = []
            for t in day_slots:
                day_terms.extend(
                    y[(k, c, t)]
                    for c in classes
                    if (k, c, t) in y
                )
                day_terms.extend(
                    z[(k, c, t)]
                    for c in classes
                    if (k, c, t) in z
                )

            if day_terms:
                model.Add(load_var == fixed_load + sum(day_terms))
            else:
                model.Add(load_var == fixed_load)

            sq_var = model.NewIntVar(0, day_upper_bound * day_upper_bound, f"teacher_day_load_sq_{k}_{d}")
            model.AddMultiplicationEquality(sq_var, [load_var, load_var])

            teacher_day_load[(k, d)] = load_var
            day_load_vars.append(load_var)
            day_sq_vars.append(sq_var)

        weekly_load_var = model.NewIntVar(0, weekly_load_upper_bound, f"teacher_week_load_{k}")
        model.Add(weekly_load_var == sum(day_load_vars))

        weekly_sq_var = model.NewIntVar(0, weekly_load_upper_bound * weekly_load_upper_bound, f"teacher_week_load_sq_{k}")
        model.AddMultiplicationEquality(weekly_sq_var, [weekly_load_var, weekly_load_var])

        variance_terms.append(len(DAY_ORDER) * sum(day_sq_vars) - weekly_sq_var)

    variance_upper_bound = max(
        1,
        len(teachers) * len(DAY_ORDER) * sum(len(slots_by_day[d]) ** 2 for d in DAY_ORDER),
    )
    variance_expr = cp_model.LinearExpr.Sum(variance_terms)
    return teacher_day_load, variance_expr, variance_upper_bound


def _build_subject_requirements(non_tt: dict[tuple[str, str, str], int]) -> dict[tuple[str, str], int]:
    req: dict[tuple[str, str], int] = {}
    for (_k, c, s), h in non_tt.items():
        req[(c, s)] = req.get((c, s), 0) + int(h)
    return req


def _find_class_open_slot_mismatches(
    *,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed: dict[tuple[str, str], str],
) -> list[tuple[str, int, int]]:
    demand_by_class: dict[str, int] = {c: 0 for c in classes}
    for (_teacher, cls, _subject), h in non_tt.items():
        if cls in demand_by_class:
            demand_by_class[cls] += int(h)

    fixed_count_by_class: dict[str, int] = {c: 0 for c in classes}
    for (cls, _slot) in fixed.keys():
        if cls in fixed_count_by_class:
            fixed_count_by_class[cls] += 1

    total_slots = len(slots)
    mismatches: list[tuple[str, int, int]] = []
    for c in classes:
        demand = demand_by_class.get(c, 0)
        free_slots = total_slots - fixed_count_by_class.get(c, 0)
        if demand != free_slots:
            mismatches.append((c, demand, free_slots))
    return mismatches


def _is_scenario_feasible_with_fixed(
    *,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_subject: dict[str, str],
    exempt_subjects: set[str],
    teacher_unavailable: dict[str, set[str]],
    time_limit_sec: float,
) -> bool:
    req = _build_subject_requirements(non_tt)
    fixed: dict[tuple[str, str], str] = {}
    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        t = str(a.get("slot", "")).strip()
        s = str(a.get("subject", "")).strip()
        if c and t and s:
            fixed[(c, t)] = s

    fixed_teacher_assignments, fixed_teacher_issues = _assign_fixed_lessons_to_non_tt_teachers(
        non_tt=non_tt,
        fixed_raw=fixed_raw,
        teacher_unavailable=teacher_unavailable,
        time_limit_sec=max(3.0, min(10.0, float(time_limit_sec))),
    )
    if fixed_teacher_issues:
        return False
    fixed_teacher_busy, fixed_main_h, _fixed_teacher_day_load = _build_fixed_teacher_usage(fixed_teacher_assignments)

    auto_exempt = {a.get("subject", "") for a in fixed_raw if a.get("subject")}
    exempt = set(exempt_subjects) | auto_exempt

    class_subjects = {c: set() for c in classes}
    for (c, s), _ in req.items():
        class_subjects[c].add(s)
    for (c, _), s in fixed.items():
        class_subjects[c].add(s)

    model: Any = cp_model.CpModel()
    x: dict[tuple[str, str, str], Any] = {}

    for c in classes:
        for t in slots:
            for s in class_subjects[c]:
                x[(c, t, s)] = model.NewBoolVar(f"x_{c}_{t}_{s}")

    for c in classes:
        for t in slots:
            if (c, t) in fixed:
                fs = fixed[(c, t)]
                for s in class_subjects[c]:
                    model.Add(x[(c, t, s)] == (1 if s == fs else 0))
            else:
                model.Add(sum(x[(c, t, s)] for s in class_subjects[c]) == 1)

    for (c, s), h in req.items():
        fixed_count = sum(1 for (cc, _t), ss in fixed.items() if cc == c and ss == s)
        target = h - fixed_count
        if target < 0:
            return False
        non_fixed = [t for t in slots if (c, t) not in fixed]
        model.Add(sum(x[(c, t, s)] for t in non_fixed) == target)

    for c in classes:
        for s in class_subjects[c]:
            for d in DAY_ORDER:
                dslots = [t for t in slots if parse_slot(t)[0] == d and (c, t) not in fixed]
                if dslots:
                    model.Add(sum(x[(c, t, s)] for t in dslots) <= 1)

    by_grade: dict[str, list[str]] = {}
    for c in classes:
        by_grade.setdefault(grade_of(c), []).append(c)

    subjects = sorted({s for (_c, s) in req} | set(fixed.values()))
    for _g, g_classes in by_grade.items():
        for t in slots:
            for s in subjects:
                if s in exempt:
                    continue
                vars_auto = []
                for c in g_classes:
                    if (c, t) in fixed:
                        continue
                    if (c, t, s) in x:
                        vars_auto.append(x[(c, t, s)])
                if vars_auto:
                    model.Add(sum(vars_auto) <= 1)

    y: dict[tuple[str, str, str], Any] = {}
    teachers_used = sorted({k for (k, _c, _s) in non_tt})
    for (k, c, _s), _ in non_tt.items():
        for t in slots:
            if (c, t) in fixed:
                continue
            y[(k, c, t)] = model.NewBoolVar(f"y_{k}_{c}_{t}")

    for c in classes:
        for t in slots:
            if (c, t) in fixed:
                continue
            cand = [k for k in teachers_used if (k, c, t) in y]
            if cand:
                model.Add(sum(y[(k, c, t)] for k in cand) == 1)

    for c in classes:
        for t in slots:
            for s in class_subjects[c]:
                teachers_for_subject = [k for k in teachers_used if teacher_subject.get(k) == s and (k, c, t) in y]
                if teachers_for_subject:
                    model.Add(x[(c, t, s)] == sum(y[(k, c, t)] for k in teachers_for_subject))

    for k in teachers_used:
        for t in slots:
            vars_t = [y[(k, c, t)] for c in classes if (k, c, t) in y]
            if vars_t:
                model.Add(sum(vars_t) <= 1)
            if t in fixed_teacher_busy.get(k, set()) and vars_t:
                model.Add(sum(vars_t) == 0)
            if t in teacher_unavailable.get(k, set()) and vars_t:
                model.Add(sum(vars_t) == 0)

    main_h: dict[tuple[str, str], int] = {}
    for (k, c, _s), h in non_tt.items():
        main_h[(k, c)] = main_h.get((k, c), 0) + int(h)
    for (k, c), h in main_h.items():
        target = h - fixed_main_h.get((k, c), 0)
        if target < 0:
            return False
        vars_kc = [y[(k, c, t)] for t in slots if (k, c, t) in y]
        model.Add(sum(vars_kc) == target)

    # 可解診断ではTTは評価対象外（本体モデルでのみ最適化対象）。
    tt: dict[tuple[str, str, str], int] = {}

    # -----------------------------
    # TT VARIABLES / CONSTRAINTS
    # -----------------------------

    z: dict[tuple[str, str, str], Any] = {}

    for (k,c,s),h in tt.items():
        if h<=0:
            continue
        for t in slots:
            if (c,t,s) not in x:
                continue
            z[(k,c,t)] = model.NewBoolVar(f"z_{k}_{c}_{t}")
            # TT参加は当該クラスで自担当教科が実施されるコマに限定。
            model.Add(z[(k,c,t)] <= x[(c,t,s)])
            if t in teacher_unavailable.get(k,set()):
                model.Add(z[(k,c,t)] == 0)

    # 非TT(y) と TT(z) を同時に担当させない。
    all_teachers = sorted(set(teachers_used) | {k for (k,_c,_s) in tt})
    for k in all_teachers:
        for t in slots:
            vars_t=[]
            vars_t.extend(
                y[(k,c,t)]
                for c in classes
                if (k,c,t) in y
            )
            vars_t.extend(
                z[(k,c,t)]
                for c in classes
                if (k,c,t) in z
            )
            if vars_t:
                model.Add(sum(vars_t) <= 1)
            if t in fixed_teacher_busy.get(k, set()) and vars_t:
                model.Add(sum(vars_t) == 0)

    # TT時間は要求時間を上限とし、過剰に割り当てない。
    for (k,c,_s),h in tt.items():
        vars_kc=[
            z[(k,c,t)]
            for t in slots
            if (k,c,t) in z
        ]
        if vars_kc:
            model.Add(sum(vars_kc) <= int(h))

    # 可能な範囲でTT割当を最大化する（不可な場合は不足を許容）。
    if z:
        model.Maximize(sum(z.values()))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(3.0, float(time_limit_sec))
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)
    return status in (cp_model.OPTIMAL, cp_model.FEASIBLE)


def _suggest_relax_fixed_candidates(
    *,
    scenario_id: str,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_subject: dict[str, str],
    exempt_subjects: set[str],
    teacher_unavailable: dict[str, set[str]],
    time_limit_sec: float,
    max_candidates: int = 3,
    trial_limit: int = 30,
) -> list[str]:
    fixed_list = [
        {
            "class": str(a.get("class", "")).strip(),
            "slot": str(a.get("slot", "")).strip(),
            "subject": str(a.get("subject", "")).strip(),
        }
        for a in fixed_raw
        if str(a.get("class", "")).strip() and str(a.get("slot", "")).strip() and str(a.get("subject", "")).strip()
    ]
    if not fixed_list:
        return []

    # 固定入力順を尊重しつつ、診断時間の上限を設ける。
    candidates: list[str] = []
    for i, victim in enumerate(fixed_list[: max(1, int(trial_limit))]):
        test_fixed = [a for j, a in enumerate(fixed_list) if j != i]
        ok = _is_scenario_feasible_with_fixed(
            classes=classes,
            slots=slots,
            non_tt=non_tt,
            fixed_raw=test_fixed,
            teacher_subject=teacher_subject,
            exempt_subjects=exempt_subjects,
            teacher_unavailable=teacher_unavailable,
            time_limit_sec=max(5.0, min(20.0, float(time_limit_sec) / 2.0)),
        )
        if ok:
            candidates.append(f"{victim['class']} {victim['slot']} {victim['subject']} を一旦外す")
            if len(candidates) >= max(1, int(max_candidates)):
                break
    return candidates


def _diagnose_infeasible_fixed_core(
    *,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_subject: dict[str, str],
    exempt_subjects: set[str],
    teacher_unavailable: dict[str, set[str]],
    time_limit_sec: float,
    max_candidates: int = 3,
) -> list[str]:
    # assumption API がない環境では診断をスキップ。
    if not fixed_raw:
        return []
    if not hasattr(cp_model.CpModel, "AddAssumptions"):
        return []

    req = _build_subject_requirements(non_tt)
    fixed_list = [
        {
            "class": str(a.get("class", "")).strip(),
            "slot": str(a.get("slot", "")).strip(),
            "subject": str(a.get("subject", "")).strip(),
        }
        for a in fixed_raw
        if str(a.get("class", "")).strip() and str(a.get("slot", "")).strip() and str(a.get("subject", "")).strip()
    ]
    if not fixed_list:
        return []

    model: Any = cp_model.CpModel()
    class_subjects = {c: set() for c in classes}
    for (c, s), _ in req.items():
        class_subjects[c].add(s)
    for a in fixed_list:
        class_subjects[a["class"]].add(a["subject"])

    x: dict[tuple[str, str, str], Any] = {}
    for c in classes:
        for t in slots:
            for s in class_subjects[c]:
                x[(c, t, s)] = model.NewBoolVar(f"x_{c}_{t}_{s}")

    for c in classes:
        for t in slots:
            model.Add(sum(x[(c, t, s)] for s in class_subjects[c]) == 1)

    for (c, s), h in req.items():
        model.Add(sum(x[(c, t, s)] for t in slots) == h)

    for c in classes:
        for s in class_subjects[c]:
            for d in DAY_ORDER:
                dslots = [t for t in slots if parse_slot(t)[0] == d]
                if dslots:
                    model.Add(sum(x[(c, t, s)] for t in dslots) <= 1)

    by_grade: dict[str, list[str]] = {}
    for c in classes:
        by_grade.setdefault(grade_of(c), []).append(c)
    subjects = sorted({s for (_c, s) in req} | {a["subject"] for a in fixed_list})
    auto_exempt = {a["subject"] for a in fixed_list}
    exempt = set(exempt_subjects) | auto_exempt

    for _g, g_classes in by_grade.items():
        for t in slots:
            for s in subjects:
                if s in exempt:
                    continue
                vars_auto = [x[(c, t, s)] for c in g_classes if (c, t, s) in x]
                if vars_auto:
                    model.Add(sum(vars_auto) <= 1)

    y: dict[tuple[str, str, str], Any] = {}
    teachers_used = sorted({k for (k, _c, _s) in non_tt})
    for (k, c, _s), _ in non_tt.items():
        for t in slots:
            y[(k, c, t)] = model.NewBoolVar(f"y_{k}_{c}_{t}")

    for c in classes:
        for t in slots:
            cand = [k for k in teachers_used if (k, c, t) in y]
            if cand:
                model.Add(sum(y[(k, c, t)] for k in cand) == 1)

    for c in classes:
        for t in slots:
            for s in class_subjects[c]:
                teachers_for_subject = [k for k in teachers_used if teacher_subject.get(k) == s and (k, c, t) in y]
                if teachers_for_subject:
                    model.Add(x[(c, t, s)] == sum(y[(k, c, t)] for k in teachers_for_subject))

    for k in teachers_used:
        for t in slots:
            vars_t = [y[(k, c, t)] for c in classes if (k, c, t) in y]
            if vars_t:
                model.Add(sum(vars_t) <= 1)
            if t in teacher_unavailable.get(k, set()) and vars_t:
                model.Add(sum(vars_t) == 0)

    assumptions: list[Any] = []
    lit_to_fixed: dict[int, dict[str, str]] = {}
    for i, a in enumerate(fixed_list):
        c = a["class"]
        t = a["slot"]
        s = a["subject"]
        if (c, t, s) not in x:
            continue
        lit = model.NewBoolVar(f"assume_fix_{i}")
        model.Add(x[(c, t, s)] == 1).OnlyEnforceIf(lit)
        assumptions.append(lit)
        lit_to_fixed[lit.Index()] = a

    if not assumptions:
        return []

    solver0 = cp_model.CpSolver()
    solver0.parameters.max_time_in_seconds = max(5.0, min(15.0, float(time_limit_sec)))
    solver0.parameters.num_search_workers = 8
    base_status = solver0.Solve(model)
    if base_status == cp_model.INFEASIBLE:
        return ["固定配置を外しても不可解（教員時数・不可コマ・同日重複などを確認）"]

    model.AddAssumptions(assumptions)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(8.0, min(30.0, float(time_limit_sec)))
    solver.parameters.num_search_workers = 8

    status = solver.Solve(model)
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return []
    if status != cp_model.INFEASIBLE:
        return []

    core = list(solver.SufficientAssumptionsForInfeasibility())
    out: list[str] = []
    for lit_idx in core:
        fixed = lit_to_fixed.get(int(lit_idx))
        if not fixed:
            continue
        out.append(f"{fixed['class']} {fixed['slot']} {fixed['subject']} を見直す")
        if len(out) >= max(1, int(max_candidates)):
            break
    return out


def _diagnose_teacher_capacity_shortage(
    *,
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    teacher_unavailable: dict[str, set[str]],
    fixed_raw: list[dict],
    max_lines: int = 12,
) -> list[str]:
    teacher_required: dict[str, int] = {}
    teacher_class_required: dict[tuple[str, str], int] = {}
    for (teacher, cls, _subject), h in non_tt.items():
        teacher_required[teacher] = teacher_required.get(teacher, 0) + int(h)
        key = (teacher, cls)
        teacher_class_required[key] = teacher_class_required.get(key, 0) + int(h)

    # クラス固定コマはそのクラスの担当可能枠を減らす（y変数は fixed 上で作られないため）。
    fixed_by_class: dict[str, int] = {}
    for a in fixed_raw:
        cls = str(a.get("class", "")).strip()
        slot = str(a.get("slot", "")).strip()
        if cls and slot:
            fixed_by_class[cls] = fixed_by_class.get(cls, 0) + 1

    lines: list[str] = []
    total_slots = len(slots)
    for teacher, need in sorted(teacher_required.items()):
        unavail = len(teacher_unavailable.get(teacher, set()))
        avail = total_slots - unavail
        if avail < need:
            lines.append(
                f"教員不足: {teacher} 可用{avail} < 担当{need}（不可コマ{unavail}）"
            )

    for (teacher, cls), need in sorted(teacher_class_required.items()):
        unavail = len(teacher_unavailable.get(teacher, set()))
        fixed = fixed_by_class.get(cls, 0)
        # 教員不可とクラス固定は単純加算の上界評価（不足検知のため保守的に使用）。
        avail_for_class = total_slots - unavail - fixed
        if avail_for_class < need:
            lines.append(
                f"担当枠不足: {teacher}/{cls} 可用見積{avail_for_class} < 担当{need} "
                f"（不可{unavail}, 固定{fixed}）"
            )

    if not lines:
        return ["可用コマ不足は検出されませんでした（他制約の衝突の可能性）。"]
    return lines[: max(1, int(max_lines))]


def _diagnose_constraint_collision_candidates(
    *,
    classes: list[str],
    slots: list[str],
    req: dict[tuple[str, str], int],
    fixed_raw: list[dict],
    exempt_subjects: set[str],
    max_examples: int = 3,
) -> list[str]:
    fixed: dict[tuple[str, str], str] = {}
    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        t = str(a.get("slot", "")).strip()
        s = str(a.get("subject", "")).strip()
        if c and t and s:
            fixed[(c, t)] = s

    auto_exempt = {str(a.get("subject", "")).strip() for a in fixed_raw if str(a.get("subject", "")).strip()}
    exempt = set(exempt_subjects) | auto_exempt

    lines: list[str] = []

    # 同日重複: 1クラス同一教科は1日1回まで。
    daily_hits: list[tuple[str, str, int, int]] = []
    for c in classes:
        fixed_count_by_subj: dict[str, int] = {}
        fixed_days_by_subj: dict[str, set[str]] = {}
        for (cc, slot), subj in fixed.items():
            if cc != c:
                continue
            fixed_count_by_subj[subj] = fixed_count_by_subj.get(subj, 0) + 1
            day = parse_slot(slot)[0]
            fixed_days_by_subj.setdefault(subj, set()).add(day)

        for (cc, subj), h in req.items():
            if cc != c:
                continue
            target = int(h) - fixed_count_by_subj.get(subj, 0)
            if target <= 0:
                continue
            used_days = len(fixed_days_by_subj.get(subj, set()))
            available_days = max(0, len(DAY_ORDER) - used_days)
            if target > available_days:
                daily_hits.append((c, subj, target, available_days))

    if daily_hits:
        examples = ", ".join(
            [f"{c}/{s} 要{need}>日{cap}" for c, s, need, cap in daily_hits[: max(1, int(max_examples))]]
        )
        lines.append(f"同日重複候補: {len(daily_hits)}件（例: {examples}）")
    else:
        lines.append("同日重複候補: 0件")

    # 学年重複: 同学年同一教科は同一コマで重複不可（exempt除外）。
    grade_map: dict[str, list[str]] = {}
    for c in classes:
        grade_map.setdefault(grade_of(c), []).append(c)

    class_subjects: dict[str, set[str]] = {c: set() for c in classes}
    for (c, s), _h in req.items():
        class_subjects[c].add(s)
    for (c, _slot), s in fixed.items():
        class_subjects[c].add(s)

    grade_hits: list[tuple[str, str, int, int]] = []
    for g, g_classes in grade_map.items():
        all_subjects = sorted({s for c in g_classes for s in class_subjects.get(c, set())})
        for subj in all_subjects:
            if subj in exempt:
                continue

            need_total = 0
            for c in g_classes:
                h = int(req.get((c, subj), 0))
                if h <= 0:
                    continue
                fixed_count = sum(1 for (cc, _t), ss in fixed.items() if cc == c and ss == subj)
                need_total += max(0, h - fixed_count)

            if need_total <= 0:
                continue

            cap = 0
            for t in slots:
                exists = False
                for c in g_classes:
                    if (c, t) in fixed:
                        continue
                    if subj in class_subjects.get(c, set()):
                        exists = True
                        break
                if exists:
                    cap += 1

            if need_total > cap:
                grade_hits.append((g, subj, need_total, cap))

    if grade_hits:
        examples = ", ".join(
            [f"{g}年/{s} 要{need}>枠{cap}" for g, s, need, cap in grade_hits[: max(1, int(max_examples))]]
        )
        lines.append(f"学年重複候補: {len(grade_hits)}件（例: {examples}）")
    else:
        lines.append("学年重複候補: 0件")

    return lines


def _diagnose_teacher_slot_conflicts_top3(
    *,
    non_tt: dict[tuple[str, str, str], int],
    teacher_unavailable: dict[str, set[str]],
    fixed_raw: list[dict],
    max_items: int = 3,
) -> list[str]:
    fixed_slots = {
        (str(a.get("class", "")).strip(), str(a.get("slot", "")).strip())
        for a in fixed_raw
        if str(a.get("class", "")).strip() and str(a.get("slot", "")).strip()
    }

    teacher_classes: dict[str, set[str]] = {}
    for (teacher, cls, _subj), h in non_tt.items():
        if int(h) <= 0:
            continue
        teacher_classes.setdefault(teacher, set()).add(cls)

    hits: list[tuple[int, str, str, list[str]]] = []
    for teacher, slots_ng in teacher_unavailable.items():
        cls_set = teacher_classes.get(teacher, set())
        if not cls_set:
            continue
        for slot in sorted(slots_ng):
            candidates = [c for c in sorted(cls_set) if (c, slot) not in fixed_slots]
            if not candidates:
                continue
            hits.append((len(candidates), teacher, slot, candidates))

    if not hits:
        return ["不可コマx同時担当候補: 0件"]

    hits.sort(key=lambda x: (-x[0], x[1], x[2]))
    out: list[str] = []
    for cnt, teacher, slot, classes in hits[: max(1, int(max_items))]:
        out.append(
            f"不可コマx同時担当候補: {teacher} {slot} で{cnt}クラス候補（{', '.join(classes[:6])}）"
        )
    return out


def _diagnose_subject_shortage_from_req_fixed(
    *,
    req: dict[tuple[str, str], int],
    fixed_raw: list[dict],
    max_lines: int = 8,
) -> list[str]:
    fixed_count: dict[tuple[str, str], int] = {}
    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        s = str(a.get("subject", "")).strip()
        if c and s:
            key = (c, s)
            fixed_count[key] = fixed_count.get(key, 0) + 1

    by_subject: dict[str, tuple[int, int, int]] = {}
    for (c, s), need in req.items():
        req_h = int(need)
        fixed_h = fixed_count.get((c, s), 0)
        remain = max(0, req_h - fixed_h)
        old_req, old_fixed, old_remain = by_subject.get(s, (0, 0, 0))
        by_subject[s] = (old_req + req_h, old_fixed + fixed_h, old_remain + remain)

    shortage = [
        (s, req_h, fixed_h, remain_h)
        for s, (req_h, fixed_h, remain_h) in by_subject.items()
        if remain_h > 0
    ]
    shortage.sort(key=lambda x: (-x[3], x[0]))

    if not shortage:
        return ["教科別不足内訳(REQ-固定): 不足0"]

    out: list[str] = []
    for s, req_h, fixed_h, remain_h in shortage[: max(1, int(max_lines))]:
        out.append(f"教科別不足: {s} 残{remain_h} (REQ{req_h} / 固定{fixed_h})")
    return out


def _diagnose_structural_conflicts(
    *,
    classes: list[str],
    slots: list[str],
    non_tt: dict[tuple[str, str, str], int],
    fixed_raw: list[dict],
    teacher_unavailable: dict[str, set[str]],
    exempt_subjects: set[str],
    max_lines: int = 12,
) -> list[str]:
    lines: list[str] = []

    fixed_by_cell: dict[tuple[str, str], list[str]] = {}
    fixed_by_class_subject_day: dict[tuple[str, str, str], int] = {}
    fixed_grade_slot_subject: dict[tuple[str, str, str], int] = {}

    for a in fixed_raw:
        c = str(a.get("class", "")).strip()
        t = str(a.get("slot", "")).strip()
        s = str(a.get("subject", "")).strip()
        if not c or not t or not s:
            continue
        fixed_by_cell.setdefault((c, t), []).append(s)
        if "-" in t:
            d = parse_slot(t)[0]
            fixed_by_class_subject_day[(c, s, d)] = fixed_by_class_subject_day.get((c, s, d), 0) + 1
        if s not in exempt_subjects:
            g = grade_of(c)
            fixed_grade_slot_subject[(g, t, s)] = fixed_grade_slot_subject.get((g, t, s), 0) + 1

    cell_conflicts = [(c, t, sorted(set(subjs))) for (c, t), subjs in fixed_by_cell.items() if len(set(subjs)) > 1]
    for c, t, subjs in cell_conflicts[:3]:
        lines.append(f"固定衝突: {c} {t} に複数教科固定 ({'/'.join(subjs)})")

    daily_conflicts = [(c, s, d, cnt) for (c, s, d), cnt in fixed_by_class_subject_day.items() if cnt >= 2]
    for c, s, d, cnt in sorted(daily_conflicts)[:3]:
        lines.append(f"同日重複確定: {c} {d} {s} が{cnt}コマ固定")

    grade_conflicts = [(g, t, s, cnt) for (g, t, s), cnt in fixed_grade_slot_subject.items() if cnt >= 2]
    for g, t, s, cnt in sorted(grade_conflicts)[:3]:
        lines.append(f"学年同時重複確定: {g}年 {t} {s} が{cnt}クラス固定")

    fixed_cells = set(fixed_by_cell.keys())
    req_teacher_class: dict[tuple[str, str], int] = {}
    for (teacher, c, _s), h in non_tt.items():
        req_teacher_class[(teacher, c)] = req_teacher_class.get((teacher, c), 0) + int(h)

    cap_shortages: list[tuple[str, str, int, int]] = []
    for (teacher, c), need in sorted(req_teacher_class.items()):
        cap = 0
        for t in slots:
            if (c, t) in fixed_cells:
                continue
            if t in teacher_unavailable.get(teacher, set()):
                continue
            cap += 1
        if cap < need:
            cap_shortages.append((teacher, c, cap, need))

    for teacher, c, cap, need in cap_shortages[:3]:
        lines.append(f"教師担当枠不足: {teacher}/{c} 可用{cap} < 必要{need}")

    if not lines:
        return ["構造矛盾チェック: 直接矛盾は未検出"]
    return lines[: max(1, int(max_lines))]


def solve_all_scenarios_cp(
        config_data:dict,
        exempt_subjects:set[str],
        progress_callback:Optional[Callable[[int,str],None]]=None,
        time_limit_sec:float=60.0
)->dict:

    log_file = open("debug_log.txt", "w", encoding='utf-8')

    log_file.write("\n==============================\n")
    log_file.write("CP-SAT TIMETABLE DEBUG START\n")
    log_file.write("==============================\n")

    classes=list(config_data["classes"])
    teachers=list(config_data["teachers"])
    scenarios=list(config_data["scenarios"])
    day_periods=config_data["day_periods"]

    slots=build_slots(day_periods)
    excluded_subject_keywords = _build_excluded_subject_keywords(set(exempt_subjects))

    log_file.write("CLASSES:" + str(classes) + "\n")
    log_file.write("SLOTS:" + str(len(slots)) + "\n")
    log_file.write("TEACHERS:" + str(len(teachers)) + "\n")
    log_file.write("SCENARIOS:" + str(len(scenarios)) + "\n")

    def _emit_progress(percent:int, message:str) -> None:
        if not progress_callback:
            return
        progress_callback(max(0, min(100, int(percent))), str(message))

    total = max(1, len(scenarios))

    def _scenario_progress(index:int, sid_now:str, percent:int, message:str) -> None:
        if not progress_callback:
            return
        base = int((index * 100) / total)
        span = max(1, int(100 / total))
        global_p = min(100, base + int((span * max(0, min(100, int(percent)))) / 100))
        progress_callback(global_p, f"{sid_now}: {message}")

    _emit_progress(1, "CP-SATの入力データを確認中")

    # -----------------------------
    # teacher demand
    # -----------------------------

    teacher_subject={}
    teacher_unavailable={}
    teacher_discouraged={}
    non_tt={}
    tt={}

    for t in teachers:

        name=t["name"]
        subject=t["subjects"][0]
        excluded_subject = _is_excluded_subject(subject, excluded_subject_keywords)

        teacher_subject[name]=subject
        teacher_unavailable[name]={
            str(slot).strip()
            for slot in t.get("unavailable_slots",[])
            if str(slot).strip()
        }
        teacher_discouraged[name]={
            str(slot).strip()
            for slot in t.get("discouraged_slots",[])
            if str(slot).strip() and str(slot).strip() not in teacher_unavailable[name]
        }

        for ca in t["class_assignments"]:

            cls=ca["class"]
            hours=int(ca["hours"])

            if excluded_subject:
                continue

            if ca.get("tt",False):

                key=(name,cls,subject)

                tt[key]=tt.get(key,0)+hours

            else:

                key=(name,cls,subject)

                non_tt[key]=non_tt.get(key,0)+hours

    _emit_progress(4, "教員需要を集計中")

    # -----------------------------
    # class subject hours
    # -----------------------------

    req={}

    for (_k,c,s),h in non_tt.items():
        req[(c,s)]=req.get((c,s),0)+h

    log_file.write("\nCLASS SUBJECT HOURS\n")

    for k,v in sorted(req.items()):
        log_file.write(str(k) + " " + str(v) + "\n")

    solved={}

    for idx, sc in enumerate(scenarios):

        sid=sc["id"]

        _scenario_progress(idx, sid, 6, "固定配置を読み込み中")

        log_file.write("\n==============================\n")
        log_file.write("SCENARIO:" + str(sid) + "\n")
        log_file.write("==============================\n")

        fixed_raw=sc.get("fixed_assignments",[])
        fixed={}

        for a in fixed_raw:
            fixed[(a["class"],a["slot"])]=a["subject"]

        log_file.write("FIXED LESSONS:" + str(len(fixed)) + "\n")

        # 分散目的の構築より先に、固定入力後の空きコマ数を最初に検証する。
        _scenario_progress(idx, sid, 8, "空きコマ数を確認中")
        mismatches = _find_class_open_slot_mismatches(
            classes=classes,
            slots=slots,
            non_tt=non_tt,
            fixed=fixed,
        )
        if mismatches:
            log_file.write("\nCLASS FREE SLOT CHECK (NG)\n")
            for c, demand, free in mismatches:
                log_file.write(
                    "MISMATCH " + str(c) + " demand=" + str(demand) + " free=" + str(free) + "\n"
                )
            log_file.close()
            lines = [
                f"{sid}: クラス別コマ数の整合チェックで不一致があります。",
                "教員タブのTTを含まない担当コマ数と、固定入力後の空きコマ数を一致させてください。",
            ]
            lines.extend(
                [f"- {c}において、教員タブでの授業の指定数は{demand}ですが、空きコマが{free}あります。" for c, demand, free in mismatches]
            )
            raise SchedulerError("\n".join(lines))

        _scenario_progress(idx, sid, 12, "固定コマと教員不可コマの整合を確認中")
        fixed_teacher_assignments, fixed_teacher_issues = _assign_fixed_lessons_to_non_tt_teachers(
            non_tt=non_tt,
            fixed_raw=fixed_raw,
            teacher_unavailable=teacher_unavailable,
            teacher_discouraged=teacher_discouraged,
            time_limit_sec=time_limit_sec,
        )
        if fixed_teacher_issues:
            log_file.write("\nFIXED TEACHER VALIDATION (NG)\n")
            for line in fixed_teacher_issues:
                log_file.write(line + "\n")
            log_file.close()
            lines = [
                f"{sid}: 固定コマが教員の不可コマ条件と両立しません。",
                "不可コマは絶対条件です。固定配置を見直してください。",
            ]
            lines.extend(fixed_teacher_issues)
            raise SchedulerError("\n".join(lines))
        fixed_teacher_busy, fixed_main_h, fixed_teacher_day_load = _build_fixed_teacher_usage(fixed_teacher_assignments)
        log_file.write("FIXED MAIN TEACHER ASSIGNMENTS:" + str(len(fixed_teacher_assignments)) + "\n")
        for (c, t), teacher in sorted(fixed_teacher_assignments.items()):
            log_file.write(f"FIXED TEACHER: {c} {t} {teacher}\n")

        _scenario_progress(idx, sid, 16, "教科候補を整理中")

        # -----------------------------
        # auto exempt
        # -----------------------------

        auto_exempt={a["subject"] for a in fixed_raw}
        exempt=set(exempt_subjects)|auto_exempt

        log_file.write("EXEMPT SUBJECTS:" + str(sorted(exempt)) + "\n")

        # -----------------------------
        # subject set
        # -----------------------------

        subj_set={s for (_c,s) in req}
        subj_set.update(fixed.values())

        subjects=sorted(subj_set)

        # -----------------------------
        # class subjects
        # -----------------------------

        class_subjects={c:set() for c in classes}

        for (c,s),_ in req.items():
            class_subjects[c].add(s)

        for (c,_),s in fixed.items():
            class_subjects[c].add(s)

        # -----------------------------
        # fixed vs req debug
        # -----------------------------

        log_file.write("\nREQ vs FIXED CHECK\n")

        for (c,s),h in req.items():

            fixed_count=sum(
                1 for (cc,t),ss in fixed.items()
                if cc==c and ss==s
            )

            if fixed_count>h:

                log_file.write("ERROR FIXED>REQ " + str(c) + " " + str(s) + " " + str(fixed_count) + " " + str(h) + "\n")

        # -----------------------------
        # MODEL
        # -----------------------------

        _scenario_progress(idx, sid, 26, "CP-SATモデルを構築中（授業枠）")

        model: Any = cp_model.CpModel()

        # -----------------------------
        # VARIABLES
        # -----------------------------

        x: dict[tuple[str, str, str], Any] = {}

        for c in classes:
            for t in slots:
                for s in class_subjects[c]:
                    x[(c,t,s)]=model.NewBoolVar(f"x_{c}_{t}_{s}")

        # -----------------------------
        # SLOT FILL
        # -----------------------------

        for c in classes:
            for t in slots:

                if (c,t) in fixed:

                    fs=fixed[(c,t)]

                    for s in class_subjects[c]:

                        model.Add(
                            x[(c,t,s)]==(1 if s==fs else 0)
                        )

                else:

                    model.Add(
                        sum(
                            x[(c,t,s)]
                            for s in class_subjects[c]
                        )==1
                    )

        # -----------------------------
        # WEEKLY HOURS
        # -----------------------------

        for (c,s),h in req.items():

            fixed_count=sum(
                1 for (cc,t),ss in fixed.items()
                if cc==c and ss==s
            )

            target=h-fixed_count

            if target<0:

                raise SchedulerError(
                    f"FIXED>{c} {s}"
                )

            non_fixed=[
                t for t in slots
                if (c,t) not in fixed
            ]

            model.Add(
                sum(x[(c,t,s)] for t in non_fixed)==target
            )

        # -----------------------------
        # DAILY SUBJECT
        # -----------------------------

        for c in classes:
            for s in class_subjects[c]:

                for d in DAY_ORDER:

                    dslots=[
                        t for t in slots
                        if parse_slot(t)[0]==d
                           and (c,t) not in fixed
                    ]

                    if dslots:

                        model.Add(
                            sum(x[(c,t,s)] for t in dslots)<=1
                        )

        # -----------------------------
        # GRADE UNIQUE
        # -----------------------------

        by_grade={}

        for c in classes:
            by_grade.setdefault(grade_of(c),[]).append(c)

        for g,g_classes in by_grade.items():
            for t in slots:
                for s in subjects:

                    if s in exempt:
                        continue

                    vars_auto=[]

                    for c in g_classes:

                        if (c,t) in fixed:
                            continue

                        if (c,t,s) in x:
                            vars_auto.append(x[(c,t,s)])

                    if vars_auto:

                        model.Add(sum(vars_auto)<=1)

        # -----------------------------
        # TEACHER VARIABLES
        # -----------------------------

        _scenario_progress(idx, sid, 44, "CP-SATモデルを構築中（教員割当）")

        y: dict[tuple[str, str, str], Any] = {}
        teachers_used=sorted({k for (k,_c,_s) in non_tt})

        for (k,c,s),_ in non_tt.items():

            for t in slots:

                if (c,t) in fixed:
                    continue

                y[(k,c,t)]=model.NewBoolVar(
                    f"y_{k}_{c}_{t}"
                )

        # -----------------------------
        # TEACHER ASSIGN
        # -----------------------------

        for c in classes:
            for t in slots:

                if (c,t) in fixed:
                    continue

                cand=[
                    k for k in teachers_used
                    if (k,c,t) in y
                ]

                if cand:

                    model.Add(
                        sum(y[(k,c,t)] for k in cand)==1
                    )

        # -----------------------------
        # TEACHER LINK
        # -----------------------------

        for c in classes:
            for t in slots:

                if (c,t) in fixed:
                    continue

                for s in class_subjects[c]:

                    teachers_for_subject = [
                        k for k in teachers_used
                        if teacher_subject[k] == s and (k,c,t) in y
                    ]

                    if teachers_for_subject:

                        model.Add(
                            x[(c,t,s)] ==
                            sum(y[(k,c,t)] for k in teachers_for_subject)
                        )

        # -----------------------------
        # TEACHER OVERLAP
        # -----------------------------

        for k in teachers_used:
            for t in slots:

                vars_t=[
                    y[(k,c,t)]
                    for c in classes
                    if (k,c,t) in y
                ]

                if vars_t:

                    model.Add(sum(vars_t)<=1)

                if t in fixed_teacher_busy.get(k,set()) and vars_t:

                    model.Add(sum(vars_t)==0)

                if t in teacher_unavailable.get(k,set()) and vars_t:

                    model.Add(sum(vars_t)==0)

        # -----------------------------
        # TEACHER HOURS
        # -----------------------------

        main_h={}

        for (k,c,_s),h in non_tt.items():
            main_h[(k,c)]=main_h.get((k,c),0)+h

        log_file.write("\nTEACHER HOURS DEMAND\n")

        for k,v in main_h.items():
            log_file.write(str(k) + " " + str(v) + "\n")

        log_file.write("\nTT HOURS DEMAND\n")
        for k,v in sorted(tt.items()):
            log_file.write(str(k) + " " + str(v) + "\n")

        for (k,c),h in main_h.items():

            target = h - fixed_main_h.get((k,c),0)

            if target<0:

                raise SchedulerError(
                    f"固定配置が主担当時数を超過: {k} {c} fixed={fixed_main_h.get((k,c),0)} > demand={h}"
                )

            vars_kc=[
                y[(k,c,t)]
                for t in slots
                if (k,c,t) in y
            ]

            model.Add(sum(vars_kc)==target)

        # -----------------------------
        # TT VARIABLES / CONSTRAINTS
        # -----------------------------

        _scenario_progress(idx, sid, 62, "CP-SATモデルを構築中（TT制約）")

        z: dict[tuple[str, str, str], Any] = {}

        for (k,c,s),h in tt.items():
            if h<=0:
                continue
            for t in slots:
                if (c,t,s) not in x:
                    continue
                z[(k,c,t)] = model.NewBoolVar(f"z_{k}_{c}_{t}")
                # TT参加は当該クラスで自担当教科が実施されるコマに限定。
                model.Add(z[(k,c,t)] <= x[(c,t,s)])
                if t in teacher_unavailable.get(k,set()):
                    model.Add(z[(k,c,t)] == 0)

        # 非TT(y) と TT(z) を同時に担当させない。
        all_teachers = sorted(set(teachers_used) | {k for (k,_c,_s) in tt})
        for k in all_teachers:
            for t in slots:
                vars_t=[]
                vars_t.extend(
                    y[(k,c,t)]
                    for c in classes
                    if (k,c,t) in y
                )
                vars_t.extend(
                    z[(k,c,t)]
                    for c in classes
                    if (k,c,t) in z
                )
                if vars_t:
                    model.Add(sum(vars_t) <= 1)
                if t in fixed_teacher_busy.get(k,set()) and vars_t:
                    model.Add(sum(vars_t) == 0)

        # TT時間は要求時間を上限とし、過剰に割り当てない。
        for (k,c,_s),h in tt.items():
            vars_kc=[
                z[(k,c,t)]
                for t in slots
                if (k,c,t) in z
            ]
            if vars_kc:
                model.Add(sum(vars_kc) <= int(h))

        _scenario_progress(idx, sid, 72, "CP-SATモデルを構築中（曜日別担当分散）")
        teacher_day_load, variance_expr, variance_upper_bound = _build_teacher_day_variance_objective_terms(
            model=model,
            classes=classes,
            slots=slots,
            teachers=all_teachers,
            y=y,
            z=z,
            fixed_teacher_day_load=fixed_teacher_day_load,
        )

        discouraged_terms = [
            var
            for (k, _c, t), var in y.items()
            if t in teacher_discouraged.get(k, set())
        ]
        discouraged_terms.extend(
            var
            for (k, _c, t), var in z.items()
            if t in teacher_discouraged.get(k, set())
        )
        discouraged_expr = cp_model.LinearExpr.Sum(discouraged_terms) if discouraged_terms else 0
        discouraged_upper_bound = len(discouraged_terms)

        # 優先順位:
        # 1. TT割当最大化
        # 2. △コマ割当最小化
        # 3. 曜日別担当時数分散最小化
        soft_weight = variance_upper_bound + 1
        if z:
            tt_weight = soft_weight * (discouraged_upper_bound + 1)
            model.Maximize(tt_weight * sum(z.values()) - soft_weight * discouraged_expr - variance_expr)
        else:
            tt_weight = None
            if discouraged_terms:
                model.Minimize(soft_weight * discouraged_expr + variance_expr)
            else:
                model.Minimize(variance_expr)

        # -----------------------------
        # SOLVER
        # -----------------------------

        solver=cp_model.CpSolver()

        solver.parameters.max_time_in_seconds=time_limit_sec
        solver.parameters.num_search_workers=8
        solver.parameters.log_search_progress=True
        solver.parameters.cp_model_presolve=True

        log_file.write("\nMODEL STATS\n")
        log_file.write(model.ModelStats() + "\n")
        log_file.write("\nOBJECTIVE\n")
        if z:
            log_file.write(
                "lexicographic-like: maximize(TT * " + str(tt_weight) + " - discouraged * " + str(soft_weight) + " - scaled_variance)\n"
            )
        else:
            if discouraged_terms:
                log_file.write("minimize(discouraged * " + str(soft_weight) + " + scaled_variance)\n")
            else:
                log_file.write("minimize(scaled_variance)\n")
        log_file.write("discouraged_term_count=" + str(discouraged_upper_bound) + "\n")
        log_file.write("scaled_variance_upper_bound=" + str(variance_upper_bound) + "\n")

        log_file.write("\nSTART SOLVING...\n\n")

        _scenario_progress(idx, sid, 82, "解探索中（制約を探索しています）")

        status=solver.Solve(model)

        log_file.write("SOLVER STATUS:" + solver.StatusName(status) + "\n")

        if status not in (cp_model.OPTIMAL,cp_model.FEASIBLE):
            _scenario_progress(idx, sid, 90, "解なし診断を作成中")
            log_file.write("\nDEBUG: MODEL UNSAT\n")
            diagnostics = _diagnose_structural_conflicts(
                classes=classes,
                slots=slots,
                non_tt=non_tt,
                fixed_raw=fixed_raw,
                teacher_unavailable=teacher_unavailable,
                exempt_subjects=exempt_subjects,
                max_lines=12,
            )
            diagnostics.extend(
                _diagnose_teacher_capacity_shortage(
                    slots=slots,
                    non_tt=non_tt,
                    teacher_unavailable=teacher_unavailable,
                    fixed_raw=fixed_raw,
                    max_lines=12,
                )
            )
            diagnostics.extend(
                _diagnose_teacher_slot_conflicts_top3(
                    non_tt=non_tt,
                    teacher_unavailable=teacher_unavailable,
                    fixed_raw=fixed_raw,
                    max_items=3,
                )
            )
            diagnostics.extend(
                _diagnose_subject_shortage_from_req_fixed(
                    req=req,
                    fixed_raw=fixed_raw,
                    max_lines=8,
                )
            )
            diagnostics.extend(
                _diagnose_constraint_collision_candidates(
                    classes=classes,
                    slots=slots,
                    req=req,
                    fixed_raw=fixed_raw,
                    exempt_subjects=exempt_subjects,
                    max_examples=3,
                )
            )
            log_file.write("\nAUTO DIAGNOSTICS\n")
            for dline in diagnostics:
                log_file.write("- " + dline + "\n")
            base_ok = _is_scenario_feasible_with_fixed(
                classes=classes,
                slots=slots,
                non_tt=non_tt,
                fixed_raw=[],
                teacher_subject=teacher_subject,
                exempt_subjects=exempt_subjects,
                teacher_unavailable=teacher_unavailable,
                time_limit_sec=max(5.0, min(20.0, float(time_limit_sec) / 2.0)),
            )
            if not base_ok:
                log_file.close()
                lines = [
                    f"{sid}: 固定配置を全解除しても解が見つかりません。",
                    "教員時数・不可コマ・同日重複制約のいずれかが矛盾しています。",
                    "",
                    "自動診断:",
                ]
                lines.extend([f"- {d}" for d in diagnostics])
                raise SchedulerError("\n".join(lines))
            # 解なし時は、固定配置を1件ずつ緩和する最小診断を実施して候補を返す。
            suggestions = _suggest_relax_fixed_candidates(
                scenario_id=sid,
                classes=classes,
                slots=slots,
                non_tt=non_tt,
                fixed_raw=fixed_raw,
                teacher_subject=teacher_subject,
                exempt_subjects=exempt_subjects,
                teacher_unavailable=teacher_unavailable,
                time_limit_sec=time_limit_sec,
                max_candidates=3,
                trial_limit=30,
            )
            if not suggestions:
                suggestions = _diagnose_infeasible_fixed_core(
                    classes=classes,
                    slots=slots,
                    non_tt=non_tt,
                    fixed_raw=fixed_raw,
                    teacher_subject=teacher_subject,
                    exempt_subjects=exempt_subjects,
                    teacher_unavailable=teacher_unavailable,
                    time_limit_sec=time_limit_sec,
                    max_candidates=3,
                )
            log_file.close()
            if suggestions:
                lines = [
                    f"{sid}: 解が見つかりません",
                    "固定配置の異動候補（まずは1件のみ変更）:",
                ]
                lines.extend([f"- {s}" for s in suggestions])
                lines.extend(["", "自動診断:"])
                lines.extend([f"- {d}" for d in diagnostics])
                raise SchedulerError("\n".join(lines))
            lines = [
                f"{sid}: 解が見つかりません（固定配置1件の緩和では可解化できません）",
                "",
                "自動診断:",
            ]
            lines.extend([f"- {d}" for d in diagnostics])
            raise SchedulerError("\n".join(lines))

        log_file.write("SOLUTION FOUND\n")
        _scenario_progress(idx, sid, 93, "解を整理中（曜日別担当を集計）")
        total_tt_assigned = sum(solver.Value(var) for var in z.values()) if z else 0
        discouraged_value = sum(
            1
            for (c, slot), teacher in fixed_teacher_assignments.items()
            if slot in teacher_discouraged.get(teacher, set())
        )
        discouraged_value += sum(
            solver.Value(var)
            for (k, _c, t), var in y.items()
            if t in teacher_discouraged.get(k, set())
        )
        discouraged_value += sum(
            solver.Value(var)
            for (k, _c, t), var in z.items()
            if t in teacher_discouraged.get(k, set())
        )
        scaled_variance_value = 0
        log_file.write("\nTEACHER DAY LOADS\n")
        for k in all_teachers:
            day_loads = [solver.Value(teacher_day_load[(k, d)]) for d in DAY_ORDER]
            scaled_variance_value += len(DAY_ORDER) * sum(v * v for v in day_loads) - (sum(day_loads) ** 2)
            log_file.write(f"{k}: {dict(zip(DAY_ORDER, day_loads))}\n")
        log_file.write(f"TOTAL TT ASSIGNED: {total_tt_assigned}\n")
        log_file.write(f"DISCOURAGED SLOT ASSIGNMENTS: {discouraged_value}\n")
        log_file.write(f"SCALED VARIANCE: {scaled_variance_value}\n")

        # デバッグ: 変数値の確認
        log_file.write("\nDEBUG: VARIABLE VALUES\n")
        for (c, t, s), var in list(x.items())[:10]:  # 最初の10個を確認
            val = solver.Value(var)
            log_file.write(f"x[{c},{t},{s}] = {val} (type: {type(val)})\n")

        # 教科割り当ての抽出
        _scenario_progress(idx, sid, 96, "解を反映中（授業・教員割当）")
        assignments = {}
        for (c, t, s), var in x.items():
            val = solver.Value(var)
            if val == 1 or val >= 0.99:  # 浮動小数点対応
                assignments[(c, t)] = s
                log_file.write(f"ASSIGN: {c} {t} {s}\n")

        # 教師割り当ての抽出
        teacher_assignments = {}
        teacher_assignments.update(fixed_teacher_assignments)
        for (k, c, t), var in y.items():
            val = solver.Value(var)
            if val == 1 or val >= 0.99:
                teacher_assignments[(c, t)] = k
                log_file.write(f"TEACHER: {c} {t} {k}\n")

        # TT割り当ての抽出
        tt_assignments = {}
        for (k, c, t), var in z.items():
            val = solver.Value(var)
            if val == 1 or val >= 0.99:
                tt_assignments.setdefault((c, t), []).append(k)
                log_file.write(f"TT: {c} {t} {k}\n")

        log_file.write(f"TOTAL ASSIGNMENTS: {len(assignments)}\n")


        solved[sid] = {
            "assignments": assignments,
            "teacher_assignments": teacher_assignments,
            "tt_assignments": tt_assignments,
        }

        _scenario_progress(idx, sid, 100, "シナリオ完了")


    log_file.close()
    _emit_progress(100, "全シナリオ完了")
    return solved

