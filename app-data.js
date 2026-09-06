/*
 * app-data.js
 * Loads and parses the .xer / .xml / .xlsx schedule exports listed in data/manifest.json,
 * builds the same normalized model as the Power BI semantic model (see MODEL_SPEC.md in the
 * companion PBIP project), and computes the same KPI / DCMA-14 measures client-side.
 *
 * Mirrors, table for table and measure for measure, the Power Query M + DAX logic built for
 * the "ConstructionPortfolioKPI" Power BI dashboard.
 */

// ---------- generic helpers ----------

function toDateSafe(v) {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") {
    // Excel serial date, epoch 1899-12-30 (matches Power Query's Date.From on numbers)
    const epoch = Date.UTC(1899, 11, 30);
    return new Date(epoch + Math.round(v) * 86400000);
  }
  if (typeof v === "string") {
    let s = v.trim();
    if (!s) return null;
    if (s.includes("T")) s = s.split("T")[0];
    else if (s.includes(" ")) s = s.split(" ")[0];
    let d = new Date(s + "T00:00:00Z");
    if (!isNaN(d)) return d;
    const m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (m) {
      const months = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };
      const mi = months[m[2]];
      if (mi !== undefined) return new Date(Date.UTC(+m[3], mi, +m[1]));
    }
    return null;
  }
  return null;
}
function toNumSafe(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}
function isoDate(d) { return d ? d.toISOString().slice(0, 10) : null; }
function mapStatus(s) {
  const t = String(s || "").toLowerCase();
  if (t.includes("complete")) return "Complete";
  if (t.includes("progress") || t === "tk_active") return "In Progress";
  return "Not Started";
}

// ---------- .xer parsing ----------

function parseXerTables(text) {
  const lines = text.split(/\r\n|\r|\n/);
  let curTable = "", curFields = [];
  const tables = {};
  for (const line of lines) {
    if (!line) continue;
    const parts = line.split("\t");
    const tag = (parts[0] || "").trim();
    if (tag === "%T") { curTable = (parts[1] || "").trim(); curFields = []; if (!tables[curTable]) tables[curTable] = []; }
    else if (tag === "%F") { curFields = parts.slice(1); }
    else if (tag === "%R") {
      const vals = parts.slice(1);
      const rec = {};
      curFields.forEach((f, i) => { rec[f] = vals[i] !== undefined ? vals[i] : ""; });
      (tables[curTable] = tables[curTable] || []).push(rec);
    }
  }
  return tables;
}

function parseXerFile(text, pk, model) {
  const tables = parseXerTables(text);
  (tables.PROJECT || []).forEach(r => {
    model.projects.push({
      ProjectKey: pk, ProjectName: r.proj_short_name || pk, SourceFormat: "XER",
      PlanStart: toDateSafe(r.plan_start_date), PlanFinish: toDateSafe(r.plan_end_date)
    });
  });
  (tables.PROJWBS || []).forEach(r => {
    if (!r.wbs_id) return;
    model.wbs.push({
      WBSKey: pk + "|" + r.wbs_id, ProjectKey: pk,
      ParentWBSKey: r.parent_wbs_id ? pk + "|" + r.parent_wbs_id : null,
      WBSCode: r.wbs_short_name, WBSName: r.wbs_name
    });
  });
  (tables.TASK || []).forEach(r => {
    if (!r.task_id) return;
    model.activities.push({
      ActivityKey: pk + "|" + r.task_id, ProjectKey: pk, WBSKey: pk + "|" + r.wbs_id,
      ActivityCode: r.task_code, ActivityName: r.task_name,
      ActivityType: r.task_type === "TT_Mile" ? "Milestone" : "Task",
      OriginalDuration: toNumSafe(r.target_drtn_hr_cnt) / 8,
      RemainingDuration: toNumSafe(r.remain_drtn_hr_cnt) / 8,
      PercentComplete: toNumSafe(r.phys_complete_pct),
      Status: mapStatus(r.status_code),
      ActualStart: toDateSafe(r.act_start_date), ActualFinish: toDateSafe(r.act_end_date),
      EarlyStart: toDateSafe(r.early_start_date), EarlyFinish: toDateSafe(r.early_end_date),
      LateStart: toDateSafe(r.late_start_date), LateFinish: toDateSafe(r.late_end_date),
      TotalFloat: toNumSafe(r.total_float_hr_cnt) / 8,
      IsCritical: r.driving_path_flag === "Y" ? "Y" : "N"
    });
  });
  (tables.TASKPRED || []).forEach(r => {
    if (!r.task_pred_id) return;
    model.relationships.push({
      RelationshipKey: pk + "|" + r.task_pred_id, ProjectKey: pk,
      PredecessorActivityKey: pk + "|" + r.pred_task_id, SuccessorActivityKey: pk + "|" + r.task_id,
      RelationshipType: (r.pred_type || "").replace("PR_", ""), Lag: toNumSafe(r.lag_hr_cnt) / 8
    });
  });
}

// ---------- P6 XML parsing ----------

function mapXmlRelType(t) {
  t = t || "";
  if (t.includes("Finish to Start")) return "FS";
  if (t.includes("Start to Start")) return "SS";
  if (t.includes("Finish to Finish")) return "FF";
  if (t.includes("Start to Finish")) return "SF";
  return "FS";
}
function xmlChildText(el, tag) {
  const n = el.getElementsByTagName(tag)[0];
  return n ? n.textContent : null;
}
function parseXmlFile(text, pk, model) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const projEl = doc.getElementsByTagName("Project")[0];
  if (projEl) {
    model.projects.push({
      ProjectKey: pk, ProjectName: xmlChildText(projEl, "Name") || pk, SourceFormat: "XML",
      PlanStart: toDateSafe(xmlChildText(projEl, "StartDate")), PlanFinish: toDateSafe(xmlChildText(projEl, "FinishDate"))
    });
  }
  Array.from(doc.getElementsByTagName("WBS")).forEach(el => {
    const objId = xmlChildText(el, "ObjectId");
    if (!objId) return;
    model.wbs.push({
      WBSKey: pk + "|" + objId, ProjectKey: pk, ParentWBSKey: null,
      WBSCode: xmlChildText(el, "Code"), WBSName: xmlChildText(el, "Name")
    });
  });
  Array.from(doc.getElementsByTagName("Activity")).forEach(el => {
    const objId = xmlChildText(el, "ObjectId");
    if (!objId) return;
    const type = xmlChildText(el, "Type") || "";
    const isCrit = String(xmlChildText(el, "IsCritical") || "").toLowerCase() === "true";
    model.activities.push({
      ActivityKey: pk + "|" + objId, ProjectKey: pk, WBSKey: pk + "|" + xmlChildText(el, "WBSObjectId"),
      ActivityCode: xmlChildText(el, "Id"), ActivityName: xmlChildText(el, "Name"),
      ActivityType: type.includes("Milestone") ? "Milestone" : "Task",
      OriginalDuration: toNumSafe(xmlChildText(el, "PlannedDuration")),
      RemainingDuration: toNumSafe(xmlChildText(el, "RemainingDuration")),
      PercentComplete: toNumSafe(xmlChildText(el, "PercentComplete")),
      Status: mapStatus(xmlChildText(el, "Status")),
      ActualStart: toDateSafe(xmlChildText(el, "ActualStartDate")), ActualFinish: toDateSafe(xmlChildText(el, "ActualFinishDate")),
      EarlyStart: toDateSafe(xmlChildText(el, "StartDate")), EarlyFinish: toDateSafe(xmlChildText(el, "FinishDate")),
      LateStart: null, LateFinish: null,
      TotalFloat: toNumSafe(xmlChildText(el, "TotalFloat")),
      IsCritical: isCrit ? "Y" : "N"
    });
  });
  Array.from(doc.getElementsByTagName("Relationship")).forEach(el => {
    const objId = xmlChildText(el, "ObjectId");
    if (!objId) return;
    model.relationships.push({
      RelationshipKey: pk + "|" + objId, ProjectKey: pk,
      PredecessorActivityKey: pk + "|" + xmlChildText(el, "PredecessorActivityObjectId"),
      SuccessorActivityKey: pk + "|" + xmlChildText(el, "SuccessorActivityObjectId"),
      RelationshipType: mapXmlRelType(xmlChildText(el, "Type")), Lag: toNumSafe(xmlChildText(el, "Lag"))
    });
  });
}

// ---------- Excel (.xlsx) parsing, via SheetJS ----------

function sheetRows(wb, name) {
  if (!wb.SheetNames.includes(name)) return null;
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null, raw: true });
}
function parseXlsxFile(buf, pk, model) {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });

  const actRows = sheetRows(wb, "Activities") || [];
  actRows.forEach(r => {
    const id = r["Activity ID"];
    if (!id) return;
    model.activities.push({
      ActivityKey: pk + "|" + id, ProjectKey: pk, WBSKey: pk + "|" + (r["WBS ID"] ?? ""),
      ActivityCode: r["Activity Code"] ?? id, ActivityName: r["Activity Name"],
      ActivityType: r["Activity Type"],
      OriginalDuration: toNumSafe(r["Original Duration (days)"]),
      RemainingDuration: toNumSafe(r["Remaining Duration (days)"]),
      PercentComplete: toNumSafe(r["% Complete"]),
      Status: r["Status"],
      ActualStart: toDateSafe(r["Actual Start"]), ActualFinish: toDateSafe(r["Actual Finish"]),
      EarlyStart: toDateSafe(r["Early Start"]), EarlyFinish: toDateSafe(r["Early Finish"]),
      LateStart: toDateSafe(r["Late Start"]), LateFinish: toDateSafe(r["Late Finish"]),
      TotalFloat: toNumSafe(r["Total Float (days)"]),
      IsCritical: r["Critical"] === "Y" ? "Y" : "N"
    });
  });

  const wbsRows = sheetRows(wb, "WBS");
  if (wbsRows) {
    wbsRows.forEach(r => {
      const id = r["WBS ID"];
      if (!id) return;
      model.wbs.push({
        WBSKey: pk + "|" + id, ProjectKey: pk,
        ParentWBSKey: r["Parent WBS ID"] ? pk + "|" + r["Parent WBS ID"] : null,
        WBSCode: r["WBS Code"], WBSName: r["WBS Name"]
      });
    });
  } else {
    // No WBS sheet: synthesize one row per distinct WBS ID referenced by Activities
    const seen = new Set();
    actRows.forEach(r => {
      const id = r["WBS ID"];
      if (!id || seen.has(id)) return;
      seen.add(id);
      model.wbs.push({ WBSKey: pk + "|" + id, ProjectKey: pk, ParentWBSKey: null, WBSCode: id, WBSName: id });
    });
  }

  const relRows = sheetRows(wb, "Relationships");
  if (relRows) {
    relRows.forEach(r => {
      const p = r["Predecessor Activity ID"], s = r["Successor Activity ID"];
      if (!p) return;
      model.relationships.push({
        RelationshipKey: pk + "|" + p + "-" + s, ProjectKey: pk,
        PredecessorActivityKey: pk + "|" + p, SuccessorActivityKey: pk + "|" + s,
        RelationshipType: r["Relationship Type"], Lag: toNumSafe(r["Lag (days)"])
      });
    });
  }

  const resRows = sheetRows(wb, "Resources");
  if (resRows) {
    resRows.forEach(r => {
      const id = r["Resource ID"];
      if (!id) return;
      model.resources.push({
        ResourceKey: pk + "|" + id, ProjectKey: pk, ResourceName: r["Resource Name"],
        ResourceType: r["Resource Type"], UOM: r["Unit of Measure"], CostPerUnit: toNumSafe(r["Cost per Unit (EGP)"])
      });
    });
  }

  const asgRows = sheetRows(wb, "Assignments");
  if (asgRows) {
    asgRows.forEach(r => {
      const a = r["Activity ID"], res = r["Resource ID"];
      if (!a) return;
      model.assignments.push({
        AssignmentKey: pk + "|" + a + "-" + res, ActivityKey: pk + "|" + a, ResourceKey: pk + "|" + res,
        BudgetedUnits: toNumSafe(r["Budgeted Units"]), ActualUnits: toNumSafe(r["Actual Units"]),
        BudgetedCost: toNumSafe(r["Budgeted Cost (EGP)"]), ActualCost: toNumSafe(r["Actual Cost (EGP)"])
      });
    });
  }
}

// ---------- load + combine ----------

async function loadRawModel() {
  const manifest = await (await fetch("data/manifest.json")).json();
  const model = { projects: [], wbs: [], activities: [], relationships: [], resources: [], assignments: [] };

  for (const fname of manifest.files) {
    const pk = fname.split("_")[0];
    const snapMatch = fname.match(/_(\d{4}-\d{2}-\d{2})\./);
    const snap = snapMatch ? toDateSafe(snapMatch[1]) : null;
    const lower = fname.toLowerCase();
    const url = "data/" + fname;
    if (lower.endsWith(".xer")) {
      parseXerFile(await (await fetch(url)).text(), pk, model);
    } else if (lower.endsWith(".xml")) {
      parseXmlFile(await (await fetch(url)).text(), pk, model);
    } else if (lower.endsWith(".xlsx")) {
      parseXlsxFile(await (await fetch(url)).arrayBuffer(), pk, model);
    }
    const proj = model.projects.find(p => p.ProjectKey === pk);
    if (proj) proj.SnapshotDate = snap;
    else model.projects.push({ ProjectKey: pk, ProjectName: pk, SourceFormat: lower.split(".").pop().toUpperCase(), SnapshotDate: snap, PlanStart: null, PlanFinish: null });
  }
  // ensure every project row carries its snapshot date even if it was pushed before the date was known
  const snapByKey = {};
  manifest.files.forEach(fname => {
    const pk = fname.split("_")[0];
    const m = fname.match(/_(\d{4}-\d{2}-\d{2})\./);
    if (m) snapByKey[pk] = toDateSafe(m[1]);
  });
  model.projects.forEach(p => { if (!p.SnapshotDate) p.SnapshotDate = snapByKey[p.ProjectKey] || null; });

  return model;
}

// ---------- KPI / DCMA-14 computation (mirrors the DAX measures) ----------

function computeDashboardData(model) {
  const { projects, wbs, activities, relationships, resources, assignments } = model;

  const projByKey = Object.fromEntries(projects.map(p => [p.ProjectKey, p]));
  const activityByKey = Object.fromEntries(activities.map(a => [a.ActivityKey, a]));

  const activityCount = activities.length;
  const relCount = relationships.length;

  const weightedComplete = arr => {
    const durSum = arr.reduce((s, a) => s + a.OriginalDuration, 0);
    if (!durSum) return 0;
    return arr.reduce((s, a) => s + a.OriginalDuration * a.PercentComplete / 100, 0) / durSum;
  };

  // ----- per-project rollup (mirrors the PROJECTS table in the web app) -----
  const projEvPv = (acts, snap) => {
    const keys = new Set(acts.map(a => a.ActivityKey));
    const asg = assignments.filter(a => keys.has(a.ActivityKey));
    if (!asg.length) return { ev: null, cpi: null, spi: null };
    const budByAct = {};
    asg.forEach(a => { budByAct[a.ActivityKey] = (budByAct[a.ActivityKey] || 0) + a.BudgetedCost; });
    const ev = asg.reduce((s, a) => {
      const act = activityByKey[a.ActivityKey];
      return s + (act ? (act.PercentComplete / 100) * a.BudgetedCost : 0);
    }, 0);
    const actualCost = asg.reduce((s, a) => s + a.ActualCost, 0);
    let pv = null;
    if (snap) {
      pv = acts.reduce((s, a) => {
        const budCost = budByAct[a.ActivityKey];
        if (!budCost) return s;
        let frac;
        if (a.ActivityType === "Milestone") frac = (a.EarlyFinish && a.EarlyFinish <= snap) ? 1 : 0;
        else if (!a.EarlyStart) frac = 0;
        else {
          const dur = a.EarlyFinish ? (a.EarlyFinish - a.EarlyStart) / 86400000 : 0;
          const elapsed = (snap - a.EarlyStart) / 86400000;
          frac = dur > 0 ? Math.max(0, Math.min(1, elapsed / dur)) : (a.EarlyStart <= snap ? 1 : 0);
        }
        return s + frac * budCost;
      }, 0);
    }
    return { ev, cpi: actualCost ? ev / actualCost : null, spi: pv ? ev / pv : null };
  };
  const PROJECTS = projects.map(p => {
    const acts = activities.filter(a => a.ProjectKey === p.ProjectKey);
    const evPv = projEvPv(acts, p.SnapshotDate);
    return {
      name: p.ProjectName,
      complete: acts.filter(a => a.Status === "Complete").length,
      inProgress: acts.filter(a => a.Status === "In Progress").length,
      notStarted: acts.filter(a => a.Status === "Not Started").length,
      critical: acts.filter(a => a.IsCritical === "Y").length,
      total: acts.length,
      pctComplete: weightedComplete(acts),
      cpi: evPv.cpi, spi: evPv.spi
    };
  }).filter(p => p.total > 0);

  // ----- portfolio scalars -----
  const criticalCnt = activities.filter(a => a.IsCritical === "Y").length;
  const behindCount = activities.reduce((s, a) => {
    const snap = projByKey[a.ProjectKey] && projByKey[a.ProjectKey].SnapshotDate;
    if (a.Status !== "Complete" && snap && a.EarlyFinish && a.EarlyFinish < snap) return s + 1;
    return s;
  }, 0);
  const milestones = activities.filter(a => a.ActivityType === "Milestone");
  const negFloatCnt = activities.filter(a => a.TotalFloat < 0).length;
  const onTimeCnt = activities.filter(a => a.Status === "Complete" && a.TotalFloat >= 0).length;
  const totalBudgeted = assignments.reduce((s, a) => s + a.BudgetedCost, 0);
  const totalActual = assignments.reduce((s, a) => s + a.ActualCost, 0);
  const distinctResources = new Set(resources.map(r => r.ResourceKey)).size;

  // ----- earned value (EV/PV/CPI/SPI/EAC/ETC) -----
  // Mirrors the DAX added to the Power BI model: EV = % complete x that activity's budgeted
  // cost; PV = budgeted cost spread linearly across each activity's own early-date span, up to
  // the project's snapshot (data) date — an approximation, not a true baseline-driven Planned
  // Value (no baseline import exists yet; see Slippage page note).
  const activityBudgetedCost = {};
  assignments.forEach(a => { activityBudgetedCost[a.ActivityKey] = (activityBudgetedCost[a.ActivityKey] || 0) + a.BudgetedCost; });
  const ev = assignments.reduce((s, a) => {
    const act = activityByKey[a.ActivityKey];
    return s + (act ? (act.PercentComplete / 100) * a.BudgetedCost : 0);
  }, 0);
  const pv = activities.reduce((s, a) => {
    const budCost = activityBudgetedCost[a.ActivityKey];
    if (!budCost) return s;
    const snap = projByKey[a.ProjectKey] && projByKey[a.ProjectKey].SnapshotDate;
    if (!snap) return s;
    let frac;
    if (a.ActivityType === "Milestone") {
      frac = (a.EarlyFinish && a.EarlyFinish <= snap) ? 1 : 0;
    } else if (!a.EarlyStart) {
      frac = 0;
    } else {
      const dur = a.EarlyFinish ? (a.EarlyFinish - a.EarlyStart) / 86400000 : 0;
      const elapsed = (snap - a.EarlyStart) / 86400000;
      frac = dur > 0 ? Math.max(0, Math.min(1, elapsed / dur)) : (a.EarlyStart <= snap ? 1 : 0);
    }
    return s + frac * budCost;
  }, 0);
  const cpi = totalActual ? ev / totalActual : null;
  const spi = pv ? ev / pv : null;
  const eac = cpi ? totalActual + (totalBudgeted - ev) / cpi : null;
  const etc = eac != null ? eac - totalActual : null;

  const PORTFOLIO = {
    projects: PROJECTS.length,
    activities: activityCount,
    weightedComplete: weightedComplete(activities),
    criticalPct: activityCount ? criticalCnt / activityCount : 0,
    criticalCnt,
    behindCount,
    milestonesTotal: milestones.length,
    milestonesRemaining: milestones.filter(a => a.Status !== "Complete").length,
    avgFloat: activityCount ? activities.reduce((s, a) => s + a.TotalFloat, 0) / activityCount : 0,
    negFloatCnt,
    onTime: activityCount ? onTimeCnt / activityCount : 0,
    budgCost: totalBudgeted,
    actCost: totalActual,
    costVar: totalBudgeted ? (totalActual - totalBudgeted) / totalBudgeted : 0,
    distinctRes: distinctResources,
    ev, pv, cpi, spi, eac, etc,
    scheduleVarianceEGP: ev - pv,
    costVarianceEGP: ev - totalActual
  };

  // ----- DCMA 14-point check -----
  const linkedKeys = new Set();
  relationships.forEach(r => { linkedKeys.add(r.PredecessorActivityKey); linkedKeys.add(r.SuccessorActivityKey); });
  const unlinked = activities.filter(a => !linkedKeys.has(a.ActivityKey)).length;
  const leadsCnt = relationships.filter(r => r.Lag < 0).length;
  const lagsCnt = relationships.filter(r => r.Lag > 0).length;
  const nonFsCnt = relationships.filter(r => r.RelationshipType !== "FS").length;
  const highFloatCnt = activities.filter(a => a.TotalFloat > 44).length;
  const highDurCnt = activities.filter(a => a.ActivityType !== "Milestone" && a.OriginalDuration > 44).length;
  const invalidDatesCnt = activities.reduce((s, a) => {
    const snap = projByKey[a.ProjectKey] && projByKey[a.ProjectKey].SnapshotDate;
    if (a.Status !== "Complete" && snap && a.EarlyStart && a.EarlyStart < snap) return s + 1;
    return s;
  }, 0);
  const assignedActivityKeys = new Set(assignments.map(a => a.ActivityKey));
  const withResourcesCnt = activities.filter(a => assignedActivityKeys.has(a.ActivityKey)).length;
  const critPathBreaks = relationships.reduce((s, r) => {
    const pred = activityByKey[r.PredecessorActivityKey], succ = activityByKey[r.SuccessorActivityKey];
    if (pred && succ && pred.IsCritical === "Y" && succ.IsCritical !== "Y") return s + 1;
    return s;
  }, 0);

  const pct = (num, den) => (den ? num / den : 0);
  const DCMA = [
    { n: 1, name: "Logic %", val: pct(unlinked, activityCount), fmt: "pct", threshold: "< 5%", status: pct(unlinked, activityCount) < 0.05 ? "pass" : "watch" },
    { n: 2, name: "Leads %", val: pct(leadsCnt, relCount), fmt: "pct", threshold: "0%", status: leadsCnt === 0 ? "pass" : "watch" },
    { n: 3, name: "Lags %", val: pct(lagsCnt, relCount), fmt: "pct", threshold: "< 5%", status: pct(lagsCnt, relCount) < 0.05 ? "pass" : "watch" },
    { n: 4, name: "Non-FS Relationships %", val: pct(nonFsCnt, relCount), fmt: "pct", threshold: "< 10%", status: pct(nonFsCnt, relCount) < 0.10 ? "pass" : "watch" },
    { n: 5, name: "Hard Constraints", val: "No constraint data in source files", fmt: "txt", status: "na" },
    { n: 6, name: "High Float %", val: pct(highFloatCnt, activityCount), fmt: "pct", threshold: "< 5%", status: pct(highFloatCnt, activityCount) < 0.05 ? "pass" : "watch" },
    { n: 7, name: "Negative Float %", val: pct(negFloatCnt, activityCount), fmt: "pct", threshold: "0%", status: negFloatCnt === 0 ? "pass" : "watch" },
    { n: 8, name: "High Duration %", val: pct(highDurCnt, activityCount), fmt: "pct", threshold: "< 5%", status: pct(highDurCnt, activityCount) < 0.05 ? "pass" : "watch" },
    { n: 9, name: "Invalid Dates %", val: pct(invalidDatesCnt, activityCount), fmt: "pct", threshold: "0%", status: invalidDatesCnt === 0 ? "pass" : "watch" },
    { n: 10, name: "Missing Resources %", val: pct(activityCount - withResourcesCnt, activityCount), fmt: "pct", threshold: "< 5%", status: pct(activityCount - withResourcesCnt, activityCount) < 0.05 ? "pass" : "watch" },
    { n: 11, name: "Missed Tasks (Baseline)", val: "Requires a baseline snapshot — not yet available", fmt: "txt", status: "na" },
    { n: 12, name: "Critical Path Breaks", val: critPathBreaks, fmt: "int", threshold: "0", status: critPathBreaks === 0 ? "pass" : "watch" },
    { n: 13, name: "CPLI", val: "Requires baseline project finish date — not yet available", fmt: "txt", status: "na" },
    { n: 14, name: "BEI", val: "Requires baseline execution data — not yet available", fmt: "txt", status: "na" }
  ];

  // ----- cost -----
  const resourceByKey = Object.fromEntries(resources.map(r => [r.ResourceKey, r]));
  const costByTypeMap = {};
  assignments.forEach(a => {
    const res = resourceByKey[a.ResourceKey];
    const type = res ? res.ResourceType : "Unknown";
    if (!costByTypeMap[type]) costByTypeMap[type] = { type, budgeted: 0, actual: 0 };
    costByTypeMap[type].budgeted += a.BudgetedCost;
    costByTypeMap[type].actual += a.ActualCost;
  });
  const COST_BY_TYPE = Object.values(costByTypeMap);

  const costByProjectMap = {};
  PROJECTS.forEach(p => { costByProjectMap[p.name] = 0; });
  assignments.forEach(a => {
    const act = activityByKey[a.ActivityKey];
    if (!act) return;
    const proj = projByKey[act.ProjectKey];
    const name = proj ? proj.ProjectName : act.ProjectKey;
    costByProjectMap[name] = (costByProjectMap[name] || 0) + a.ActualCost;
  });
  const COST_BY_PROJECT = Object.entries(costByProjectMap).map(([name, actual]) => ({ name, actual }));

  // ----- milestone log -----
  const MILESTONES = milestones.map(a => {
    const proj = projByKey[a.ProjectKey];
    const date = a.Status === "Complete" ? a.ActualFinish : a.EarlyFinish;
    return { project: proj ? proj.ProjectName : a.ProjectKey, name: a.ActivityName, status: a.Status, date: isoDate(date) };
  });

  // ----- planned vs actual monthly cumulative S-curve -----
  const allDates = activities.flatMap(a => [a.EarlyFinish, a.ActualFinish]).filter(Boolean);
  let SCURVE = [];
  if (allDates.length) {
    const minD = new Date(Math.min(...allDates)), maxD = new Date(Math.max(...allDates));
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    let y = minD.getUTCFullYear(), m = minD.getUTCMonth();
    const endY = maxD.getUTCFullYear(), endM = maxD.getUTCMonth();
    while (y < endY || (y === endY && m <= endM)) {
      const monthEnd = new Date(Date.UTC(y, m + 1, 0));
      const planned = activities.filter(a => a.EarlyFinish && a.EarlyFinish <= monthEnd).length;
      const actual = activities.filter(a => a.Status === "Complete" && a.ActualFinish && a.ActualFinish <= monthEnd).length;
      SCURVE.push({ y, m: monthNames[m], planned, actual });
      m++; if (m > 11) { m = 0; y++; }
    }
  }

  // ----- flat activity list for the explorer -----
  const ACTIVITIES = activities.map(a => {
    const proj = projByKey[a.ProjectKey];
    const w = wbs.find(x => x.WBSKey === a.WBSKey);
    return {
      project: proj ? proj.ProjectName : a.ProjectKey,
      wbs: w ? w.WBSName : a.WBSKey,
      code: a.ActivityCode, name: a.ActivityName, type: a.ActivityType, status: a.Status,
      pct: a.PercentComplete, start: isoDate(a.ActualStart), finish: isoDate(a.ActualFinish),
      float: a.TotalFloat, critical: a.IsCritical
    };
  });

  // ----- per-resource utilization (actual vs. budgeted units/cost, across all its assignments) -----
  const resAgg = {};
  assignments.forEach(a => {
    if (!resAgg[a.ResourceKey]) resAgg[a.ResourceKey] = { budgetedUnits: 0, actualUnits: 0, budgetedCost: 0, actualCost: 0, any: false };
    const agg = resAgg[a.ResourceKey];
    agg.budgetedUnits += a.BudgetedUnits; agg.actualUnits += a.ActualUnits;
    agg.budgetedCost += a.BudgetedCost; agg.actualCost += a.ActualCost;
    agg.any = true;
  });
  const RESOURCES = resources.map(r => {
    const agg = resAgg[r.ResourceKey];
    return {
      name: r.ResourceName, type: r.ResourceType,
      budgetedUnits: agg ? agg.budgetedUnits : null, actualUnits: agg ? agg.actualUnits : null,
      budgetedCost: agg ? agg.budgetedCost : null, actualCost: agg ? agg.actualCost : null
    };
  });

  // ----- header/footer metadata -----
  const snapDates = projects.map(p => p.SnapshotDate).filter(Boolean);
  const dataDate = snapDates.length ? new Date(Math.max(...snapDates)) : null;
  const sourceFormats = [...new Set(projects.map(p => p.SourceFormat).filter(Boolean))].sort();

  return { PROJECTS, PORTFOLIO, DCMA, COST_BY_TYPE, COST_BY_PROJECT, MILESTONES, SCURVE, ACTIVITIES, RESOURCES, dataDate, sourceFormats };
}

async function loadDashboardData() {
  const model = await loadRawModel();
  const computed = computeDashboardData(model);
  computed.model = model; // exposed for rollup() — see below
  return computed;
}

// =====================================================================================
// ROLLUP ENGINE — aggregates the raw model into one node of measures for any scope:
// portfolio (all activities), project (ProjectKey), wbs (ProjectKey+WBSKey), or a single
// activity. Every KPI_REGISTRY formula reads from this node — no KPI re-derives its own
// slice of the model, so "SPI 0.88" and "SPI 0.88 on the MEP WBS of CairoHQ" are the same
// formula against two different nodes.
// =====================================================================================

function scopeActivities(level, scopeKey, model) {
  switch (level) {
    case "project": return model.activities.filter(a => a.ProjectKey === scopeKey.project);
    case "wbs": return model.activities.filter(a => a.ProjectKey === scopeKey.project && a.WBSKey === scopeKey.wbs);
    case "activity": return model.activities.filter(a => a.ActivityKey === scopeKey.activity);
    case "portfolio":
    default: return model.activities;
  }
}

function rollup(level, scopeKey, model) {
  const { projects, activities: allActivities, relationships, resources, assignments } = model;
  const acts = scopeActivities(level, scopeKey, model);
  const actKeySet = new Set(acts.map(a => a.ActivityKey));
  const projByKey = Object.fromEntries(projects.map(p => [p.ProjectKey, p]));
  const activityByKey = Object.fromEntries(allActivities.map(a => [a.ActivityKey, a]));

  // ----- earned value -----
  const asg = assignments.filter(a => actKeySet.has(a.ActivityKey));
  const budgetCost = asg.reduce((s, a) => s + a.BudgetedCost, 0);
  const actualCost = asg.reduce((s, a) => s + a.ActualCost, 0);
  const activityBudgetedCost = {};
  asg.forEach(a => { activityBudgetedCost[a.ActivityKey] = (activityBudgetedCost[a.ActivityKey] || 0) + a.BudgetedCost; });
  const ev = asg.reduce((s, a) => {
    const act = activityByKey[a.ActivityKey];
    return s + (act ? (act.PercentComplete / 100) * a.BudgetedCost : 0);
  }, 0);
  const pv = acts.reduce((s, a) => {
    const budCost = activityBudgetedCost[a.ActivityKey];
    if (!budCost) return s;
    const snap = projByKey[a.ProjectKey] && projByKey[a.ProjectKey].SnapshotDate;
    if (!snap) return s;
    let frac;
    if (a.ActivityType === "Milestone") frac = (a.EarlyFinish && a.EarlyFinish <= snap) ? 1 : 0;
    else if (!a.EarlyStart) frac = 0;
    else {
      const dur = a.EarlyFinish ? (a.EarlyFinish - a.EarlyStart) / 86400000 : 0;
      const elapsed = (snap - a.EarlyStart) / 86400000;
      frac = dur > 0 ? Math.max(0, Math.min(1, elapsed / dur)) : (a.EarlyStart <= snap ? 1 : 0);
    }
    return s + frac * budCost;
  }, 0);
  const cpi = actualCost ? ev / actualCost : null;
  const spi = pv ? ev / pv : null;
  const eac = cpi ? actualCost + (budgetCost - ev) / cpi : null;

  // ----- schedule / progress -----
  const count = acts.length;
  const durSum = acts.reduce((s, a) => s + a.OriginalDuration, 0);
  const weightedComplete = durSum ? acts.reduce((s, a) => s + a.OriginalDuration * a.PercentComplete / 100, 0) / durSum : null;
  const completeCount = acts.filter(a => a.Status === "Complete").length;
  const criticalCount = acts.filter(a => a.IsCritical === "Y").length;
  const negFloatActs = acts.filter(a => a.TotalFloat < 0);
  const onTimeCount = acts.filter(a => a.Status === "Complete" && a.TotalFloat >= 0).length;
  const floatSum = acts.reduce((s, a) => s + a.TotalFloat, 0);
  const milestones = acts.filter(a => a.ActivityType === "Milestone");
  const overdueMilestones = milestones.filter(a => {
    if (a.Status === "Complete") return false;
    const snap = projByKey[a.ProjectKey] && projByKey[a.ProjectKey].SnapshotDate;
    return snap && a.EarlyFinish && a.EarlyFinish < snap;
  });

  // ----- integrity (DCMA-style, scoped to relationships touching this activity set) -----
  const relInScope = relationships.filter(r => actKeySet.has(r.PredecessorActivityKey) || actKeySet.has(r.SuccessorActivityKey));
  const linkedKeys = new Set();
  relInScope.forEach(r => { linkedKeys.add(r.PredecessorActivityKey); linkedKeys.add(r.SuccessorActivityKey); });
  const unlinkedCount = acts.filter(a => !linkedKeys.has(a.ActivityKey)).length;
  const leadsCount = relInScope.filter(r => r.Lag < 0).length;
  const lagsCount = relInScope.filter(r => r.Lag > 0).length;
  const highFloatCount = acts.filter(a => a.TotalFloat > 44).length;

  // ----- resources -----
  const resKeysInScope = new Set(asg.map(a => a.ResourceKey));
  const resInScope = resources.filter(r => resKeysInScope.has(r.ResourceKey));
  const resAgg = {};
  asg.forEach(a => {
    if (!resAgg[a.ResourceKey]) resAgg[a.ResourceKey] = { budgetedUnits: 0, actualUnits: 0 };
    resAgg[a.ResourceKey].budgetedUnits += a.BudgetedUnits;
    resAgg[a.ResourceKey].actualUnits += a.ActualUnits;
  });
  const resUtilList = resInScope
    .map(r => { const agg = resAgg[r.ResourceKey]; return agg && agg.budgetedUnits ? agg.actualUnits / agg.budgetedUnits : null; })
    .filter(u => u != null);
  const avgUtilization = resUtilList.length ? resUtilList.reduce((s, u) => s + u, 0) / resUtilList.length : null;
  const overAllocatedCount = resUtilList.filter(u => u > 1).length;

  return {
    level, scopeKey, count, acts,
    ev, pv, budgetCost, actualCost, cpi, spi, eac,
    etc: eac != null ? eac - actualCost : null,
    weightedComplete, completeCount, criticalCount,
    negFloatCount: negFloatActs.length, negFloatActs,
    onTimeCount, avgFloat: count ? floatSum / count : null,
    milestonesTotal: milestones.length,
    milestonesRemaining: milestones.filter(a => a.Status !== "Complete").length,
    overdueMilestones,
    unlinkedCount, leadsCount, lagsCount, highFloatCount, relCount: relInScope.length,
    avgUtilization, overAllocatedCount, resCount: resInScope.length
  };
}

// =====================================================================================
// KPI REGISTRY — single source of truth. Each entry: {id, aspect, name, shortName,
// definition, formula(node)=>value|null, unit, decimals, target, direction, rag:{green,
// amber}|null, levels, isHeadline}. `rag: null` or `target: null` marks a KPI that's
// informational only — shown without a color/threshold rather than a fabricated one.
// =====================================================================================

const KPI_REGISTRY = [
  // ---- schedule: on time? ----
  { id: "spi", aspect: "schedule", name: "Schedule Performance Index", shortName: "SPI",
    definition: "Earned value ÷ planned value — below 1.0 means behind schedule in cost-weighted terms. Planned Value here is a linear spread of budgeted cost across each activity's early-date span, not a true baseline.",
    formula: n => n.spi, unit: "ratio", decimals: 2, target: 1.0, direction: "higher-better",
    rag: { green: 0.95, amber: 0.85 }, levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "on_time_pct", aspect: "schedule", name: "Completed On-Time %", shortName: "On-Time %",
    definition: "Share of completed activities finished with total float ≥ 0.",
    formula: n => n.completeCount ? n.onTimeCount / n.completeCount : null, unit: "pct", decimals: 0,
    target: 1.0, direction: "higher-better", rag: { green: 0.9, amber: 0.75 },
    levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "avg_float", aspect: "schedule", name: "Average Total Float", shortName: "Avg Float",
    definition: "Mean total float across scoped activities, in days.",
    formula: n => n.avgFloat, unit: "days", decimals: 1, target: 5, direction: "higher-better",
    rag: { green: 5, amber: 0 }, levels: ["portfolio", "project", "wbs"], isHeadline: false },
  { id: "critical_pct", aspect: "schedule", name: "% Critical Activities", shortName: "% Critical",
    definition: "Share of activities on the critical path. Not inherently bad — a tight network is normal — but very high values leave no flexibility to absorb slips.",
    formula: n => n.count ? n.criticalCount / n.count : null, unit: "pct", decimals: 0, target: 0.8,
    direction: "lower-better", rag: { green: 0.7, amber: 0.9 }, levels: ["portfolio", "project", "wbs"], isHeadline: false },

  // ---- cost: on budget? ----
  { id: "cpi", aspect: "cost", name: "Cost Performance Index", shortName: "CPI",
    definition: "Earned value ÷ actual cost — below 1.0 means the work performed cost more than it earned.",
    formula: n => n.cpi, unit: "ratio", decimals: 2, target: 1.0, direction: "higher-better",
    rag: { green: 0.95, amber: 0.85 }, levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "cv_pct", aspect: "cost", name: "Cost Variance %", shortName: "CV %",
    definition: "(Actual − Budgeted) ÷ Budgeted cost.",
    formula: n => n.budgetCost ? (n.actualCost - n.budgetCost) / n.budgetCost : null, unit: "pct", decimals: 1,
    target: 0, direction: "lower-better", rag: { green: 0, amber: 0.05 },
    levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "vac_pct", aspect: "cost", name: "Variance at Completion %", shortName: "VAC %",
    definition: "(Budgeted − Estimate at Completion) ÷ Budgeted — positive means forecast to finish under budget.",
    formula: n => (n.budgetCost && n.eac != null) ? (n.budgetCost - n.eac) / n.budgetCost : null, unit: "pct",
    decimals: 1, target: 0, direction: "higher-better", rag: { green: 0, amber: -0.05 },
    levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "eac", aspect: "cost", name: "Estimate at Completion", shortName: "EAC",
    definition: "Actual cost to date plus the remaining work forecast at the current cost efficiency.",
    formula: n => n.eac, unit: "egp", decimals: 0, target: null, direction: "lower-better", rag: null,
    levels: ["portfolio", "project", "wbs"], isHeadline: false },
  { id: "etc", aspect: "cost", name: "Estimate to Complete", shortName: "ETC",
    definition: "EAC minus actual cost to date — the forecast remaining spend.",
    formula: n => n.etc, unit: "egp", decimals: 0, target: null, direction: "lower-better", rag: null,
    levels: ["portfolio", "project", "wbs"], isHeadline: false },

  // ---- progress: done vs plan? ----
  { id: "weighted_pct_complete", aspect: "progress", name: "Weighted % Complete", shortName: "% Complete",
    definition: "Duration-weighted percent complete. Shown without a target or RAG color — there's no defensible pace target without a baseline to compare against.",
    formula: n => n.weightedComplete, unit: "pct", decimals: 0, target: null, direction: "higher-better",
    rag: null, levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "completion_rate", aspect: "progress", name: "Activities Complete %", shortName: "Complete %",
    definition: "Share of scoped activities marked Complete (unweighted by duration).",
    formula: n => n.count ? n.completeCount / n.count : null, unit: "pct", decimals: 0, target: null,
    direction: "higher-better", rag: null, levels: ["portfolio", "project", "wbs"], isHeadline: false },

  // ---- integrity: trust the plan? (DCMA 14-point, scoped) ----
  { id: "dcma_logic", aspect: "integrity", name: "DCMA 1 · Logic %", shortName: "Logic %",
    definition: "Share of activities missing a predecessor or successor.",
    formula: n => n.count ? n.unlinkedCount / n.count : null, unit: "pct", decimals: 1, target: 0.05,
    direction: "lower-better", rag: { green: 0.05, amber: 0.1 }, levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "dcma_leads", aspect: "integrity", name: "DCMA 2 · Leads %", shortName: "Leads %",
    definition: "Share of relationships carrying a negative lag.",
    formula: n => n.relCount ? n.leadsCount / n.relCount : null, unit: "pct", decimals: 1, target: 0,
    direction: "lower-better", rag: { green: 0, amber: 0.02 }, levels: ["portfolio", "project", "wbs"], isHeadline: false },
  { id: "dcma_lags", aspect: "integrity", name: "DCMA 3 · Lags %", shortName: "Lags %",
    definition: "Share of relationships carrying a positive lag.",
    formula: n => n.relCount ? n.lagsCount / n.relCount : null, unit: "pct", decimals: 1, target: 0.05,
    direction: "lower-better", rag: { green: 0.05, amber: 0.1 }, levels: ["portfolio", "project", "wbs"], isHeadline: false },
  { id: "dcma_high_float", aspect: "integrity", name: "DCMA 6 · High Float %", shortName: "High Float %",
    definition: "Share of activities with total float over 44 working days.",
    formula: n => n.count ? n.highFloatCount / n.count : null, unit: "pct", decimals: 1, target: 0.05,
    direction: "lower-better", rag: { green: 0.05, amber: 0.1 }, levels: ["portfolio", "project", "wbs"], isHeadline: false },
  { id: "dcma_neg_float", aspect: "integrity", name: "DCMA 7 · Negative Float %", shortName: "Neg. Float %",
    definition: "Share of activities with total float below zero.",
    formula: n => n.count ? n.negFloatCount / n.count : null, unit: "pct", decimals: 1, target: 0,
    direction: "lower-better", rag: { green: 0, amber: 0.02 }, levels: ["portfolio", "project", "wbs"], isHeadline: true },

  // ---- resources: adequate? ----
  { id: "resource_utilization", aspect: "resources", name: "Avg Resource Utilization", shortName: "Utilization",
    definition: "Actual vs. budgeted units, averaged across resources with assignment data in scope.",
    formula: n => n.avgUtilization, unit: "pct", decimals: 0, target: 0.85, direction: "higher-better",
    rag: { green: 0.7, amber: 0.5 }, levels: ["portfolio", "project"], isHeadline: true },
  { id: "over_allocated_count", aspect: "resources", name: "Over-Allocated Resources", shortName: "Over-Allocated",
    definition: "Resources where actual units consumed exceed budgeted units.",
    formula: n => n.overAllocatedCount, unit: "count", decimals: 0, target: 0, direction: "lower-better",
    rag: { green: 0, amber: 2 }, levels: ["portfolio", "project"], isHeadline: false },

  // ---- risk: what's at risk? ----
  { id: "neg_float_count", aspect: "risk", name: "Negative-Float Activities", shortName: "Neg. Float",
    definition: "Activities currently behind their logic network — the earliest signal of trouble.",
    formula: n => n.negFloatCount, unit: "count", decimals: 0, target: 0, direction: "lower-better",
    rag: { green: 0, amber: 2 }, levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "milestones_at_risk", aspect: "risk", name: "Overdue Milestones", shortName: "Overdue",
    definition: "Milestones past their planned finish that are still not complete.",
    formula: n => n.overdueMilestones.length, unit: "count", decimals: 0, target: 0, direction: "lower-better",
    rag: { green: 0, amber: 1 }, levels: ["portfolio", "project", "wbs"], isHeadline: true },
  { id: "open_issues_count", aspect: "risk", name: "Open Issues", shortName: "Open Issues",
    definition: "Items on the manually maintained Issues & Actions log that aren't resolved. Not derived from P6 data — the index.html page attaches this count onto the rollup node itself before formulas run.",
    formula: n => n.openIssuesCount != null ? n.openIssuesCount : null, unit: "count", decimals: 0, target: 0,
    direction: "lower-better", rag: { green: 0, amber: 3 }, levels: ["portfolio", "project"], isHeadline: false }
];
