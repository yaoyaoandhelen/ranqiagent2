import json
from pathlib import Path

import pandas as pd


OUTPUT_DIR = Path("outputs/chongqing_research_award_dashboard")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

TOTAL_WORKS_ASSUMPTION = 120

IMPORT_PATH = "/Users/luosiyao/Documents/work/work/项目文件/数字重庆-社科院/发展研究奖操作说明书/成果导入.xlsx"
AWARD_FILES = [
    (
        "/Users/luosiyao/Library/Containers/com.tencent.xinWeChat/Data/Library/Caches/com.tencent.xinWeChat/2.0b4.0.9/3a082e0dbd57afacfcda605ed43f569f/SaveTemp/386ad4fcd9478e78e64dc1a37d7499d5/研究成果_评审结果_一等奖.xlsx",
        "一等奖",
    ),
    (
        "/Users/luosiyao/Library/Containers/com.tencent.xinWeChat/Data/Library/Caches/com.tencent.xinWeChat/2.0b4.0.9/3a082e0dbd57afacfcda605ed43f569f/SaveTemp/6314ac9722cf640e1c5f87ab85a29f9b/研究成果_评审结果_二等奖.xlsx",
        "二等奖",
    ),
    (
        "/Users/luosiyao/Library/Containers/com.tencent.xinWeChat/Data/Library/Caches/com.tencent.xinWeChat/2.0b4.0.9/3a082e0dbd57afacfcda605ed43f569f/SaveTemp/15bdaa899b317ee3c511170b6da6341c/研究成果_评审结果_三等奖.xlsx",
        "三等奖",
    ),
]


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


import_df = pd.read_excel(IMPORT_PATH)
import_df = import_df.rename(columns={"成果编号": "编号"})
for col in ["编号", "成果名称", "第一完成单位", "成果形式", "所属领域", "申报主体类别", "成果分组结果"]:
    if col not in import_df.columns:
        import_df[col] = ""
    import_df[col] = import_df[col].map(clean_text)

standard_forms = {"研究报告", "政策建议", "调研报告", "著作", "论文"}
import_df["成果形式"] = import_df["成果形式"].map(
    lambda value: value if value in standard_forms else "待核验"
)

award_frames = []
for path, award in AWARD_FILES:
    df = pd.read_excel(path)
    df["奖项"] = award
    award_frames.append(df)

award_df = pd.concat(award_frames, ignore_index=True)
award_df["编号"] = award_df["编号"].map(clean_text)
award_df["成果名称"] = award_df["成果名称"].map(clean_text)
award_order = {"特等奖": 0, "一等奖": 1, "二等奖": 2, "三等奖": 3, "未获奖": 4}
award_df["奖项序"] = award_df["奖项"].map(award_order)

merged = import_df.merge(
    award_df[["编号", "奖项", "小组评分", "大会评分", "综合最终得分"]],
    on="编号",
    how="left",
)
merged["奖项"] = merged["奖项"].fillna("未获奖")
for col in ["小组评分", "大会评分", "综合最终得分"]:
    merged[col] = pd.to_numeric(merged[col], errors="coerce")
merged["大会-小组差"] = merged["大会评分"] - merged["小组评分"]

award_counts = {"特等奖": 0}
for award in ["一等奖", "二等奖", "三等奖"]:
    award_counts[award] = int((award_df["奖项"] == award).sum())
awarded_total = sum(award_counts.values())
award_counts["未获奖"] = max(TOTAL_WORKS_ASSUMPTION - awarded_total, 0)

award_summary = []
for award in ["特等奖", "一等奖", "二等奖", "三等奖", "未获奖"]:
    count = award_counts[award]
    award_summary.append(
        {
            "award": award,
            "count": count,
            "share": count / TOTAL_WORKS_ASSUMPTION if TOTAL_WORKS_ASSUMPTION else 0,
            "color": {
                "特等奖": "#A855F7",
                "一等奖": "#FF7A59",
                "二等奖": "#38BDF8",
                "三等奖": "#34D399",
                "未获奖": "#94A3B8",
            }[award],
        }
    )


def grouped_awards(df, group_col, top_n=12):
    rows = []
    grouped = df.groupby(group_col, dropna=False)
    for name, g in grouped:
        name = clean_text(name) or "未填写"
        counts = {award: int((g["奖项"] == award).sum()) for award in ["特等奖", "一等奖", "二等奖", "三等奖", "未获奖"]}
        submitted = int(len(g))
        awarded = counts["特等奖"] + counts["一等奖"] + counts["二等奖"] + counts["三等奖"]
        rows.append(
            {
                "name": name,
                "submitted": submitted,
                "awarded": awarded,
                "rate": awarded / submitted if submitted else 0,
                "counts": counts,
            }
        )
    rows.sort(key=lambda x: (x["awarded"], x["submitted"], x["rate"]), reverse=True)
    return rows[:top_n]


unit_stats = grouped_awards(merged, "第一完成单位", 14)
form_stats = grouped_awards(merged, "成果形式", 10)
field_stats = grouped_awards(merged, "所属领域", 10)

matrix_rows = []
name_group = merged.groupby("成果名称", dropna=False)
for title, g in name_group:
    title = clean_text(title) or "未命名成果"
    counts = {award: int((g["奖项"] == award).sum()) for award in ["特等奖", "一等奖", "二等奖", "三等奖", "未获奖"]}
    total = int(len(g))
    best_award = min(g["奖项"].tolist(), key=lambda a: award_order.get(a, 99))
    avg_score = g["综合最终得分"].dropna().mean()
    max_score = g["综合最终得分"].dropna().max()
    matrix_rows.append(
        {
            "title": title,
            "total": total,
            "bestAward": best_award,
            "avgScore": None if pd.isna(avg_score) else round(float(avg_score), 2),
            "maxScore": None if pd.isna(max_score) else round(float(max_score), 2),
            "counts": counts,
            "units": sorted([u for u in g["第一完成单位"].dropna().unique().tolist() if clean_text(u)])[:3],
        }
    )
matrix_rows.sort(
    key=lambda x: (
        award_order.get(x["bestAward"], 99),
        -(x["avgScore"] or 0),
        -x["total"],
        x["title"],
    )
)

top_awards = award_df.sort_values(
    ["奖项序", "综合最终得分", "大会评分"], ascending=[True, False, False]
).head(18)
top_works = top_awards[
    ["编号", "成果名称", "奖项", "小组评分", "大会评分", "综合最终得分"]
].to_dict(orient="records")

payload = {
    "meta": {
        "totalAssumption": TOTAL_WORKS_ASSUMPTION,
        "importedRows": int(len(import_df)),
        "awardedTotal": int(awarded_total),
        "matchedImportedAwards": int((merged["奖项"] != "未获奖").sum()),
        "unmatchedAwardIds": sorted(list(set(award_df["编号"]) - set(import_df["编号"]))),
    },
    "awardSummary": award_summary,
    "unitStats": unit_stats,
    "formStats": form_stats,
    "fieldStats": field_stats,
    "matrixRows": matrix_rows,
    "topWorks": top_works,
}

(OUTPUT_DIR / "award_dashboard_data.json").write_text(
    json.dumps(payload, ensure_ascii=False, indent=2),
    encoding="utf-8",
)

html_template = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>发展研究奖评审后台分析</title>
  <script id="dashboard-data" type="application/json">__DATA__</script>
  <style>
    :root {
      --bg: #ffffff;
      --panel: #ffffff;
      --muted: #64748b;
      --text: #0f172a;
      --border: #e2e8f0;
      --line: #f1f5f9;
      --brand: #1d4ed8;
      --brand-soft: #eff6ff;
      --orange: #FF7A59;
      --green: #34D399;
      --purple: #A855F7;
      --shadow: 0 10px 28px rgba(15, 23, 42, 0.08);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      color: var(--text);
      background: var(--bg);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      letter-spacing: 0;
    }
    .shell { display: grid; grid-template-columns: 232px 1fr; min-height: 100vh; }
    .sidebar {
      border-right: 1px solid var(--border);
      background: #fbfdff;
      padding: 22px 18px;
      position: sticky;
      top: 0;
      height: 100vh;
    }
    .brand { display: flex; gap: 12px; align-items: center; margin-bottom: 28px; }
    .brand-mark {
      width: 38px; height: 38px; border-radius: 8px;
      display: grid; place-items: center; color: #fff; background: var(--brand);
      font-weight: 800;
    }
    .brand strong { display: block; font-size: 15px; }
    .brand span { color: var(--muted); font-size: 12px; }
    .nav { display: grid; gap: 6px; }
    .nav a {
      text-decoration: none; color: #334155; padding: 10px 12px; border-radius: 8px;
      display: flex; align-items: center; gap: 10px; font-size: 14px;
    }
    .nav a.active, .nav a:hover { background: var(--brand-soft); color: var(--brand); }
    .dot { width: 8px; height: 8px; border-radius: 999px; background: currentColor; }
    main { padding: 24px 28px 40px; min-width: 0; }
    .topbar {
      display: flex; justify-content: space-between; align-items: flex-start; gap: 18px;
      border-bottom: 1px solid var(--border); padding-bottom: 18px; margin-bottom: 20px;
    }
    h1 { font-size: 24px; line-height: 1.28; margin: 0 0 8px; }
    .subtle { color: var(--muted); font-size: 13px; line-height: 1.6; }
    .award-tags { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
    .award-tag {
      display: inline-flex; align-items: center; gap: 6px;
      border: 1px solid var(--border); border-radius: 6px;
      padding: 4px 8px; color: #334155; background: #fff; font-size: 12px; font-weight: 700;
    }
    .award-tag i { width: 10px; height: 10px; border-radius: 2px; display: inline-block; }
    .toolbar { display: flex; gap: 10px; flex-wrap: wrap; justify-content: flex-end; }
    select, input {
      height: 38px; border: 1px solid var(--border); background: #fff; color: var(--text);
      border-radius: 8px; padding: 0 12px; min-width: 142px; font-size: 14px;
    }
    input { min-width: 220px; }
    .grid { display: grid; gap: 16px; }
    .kpis { grid-template-columns: repeat(5, minmax(140px, 1fr)); }
    .card {
      background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
      box-shadow: var(--shadow);
      min-width: 0;
    }
    .kpi { padding: 16px; }
    .kpi label { display: block; color: var(--muted); font-size: 12px; margin-bottom: 8px; }
    .kpi strong { font-size: 26px; line-height: 1; }
    .kpi small { display: block; color: var(--muted); margin-top: 8px; }
    .award-activity {
      grid-column: span 2;
      padding: 16px;
      border: 1px solid #bfdbfe;
      background: linear-gradient(135deg, #eff6ff 0%, #ffffff 58%, #f8fafc 100%);
    }
    .award-activity-head { display: flex; justify-content: space-between; gap: 10px; align-items: baseline; margin-bottom: 12px; }
    .award-activity-head label { color: #1e3a8a; font-size: 12px; font-weight: 800; }
    .award-activity-head span { color: var(--muted); font-size: 12px; }
    .award-mini-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 8px; }
    .award-mini {
      min-width: 0;
      border-radius: 8px;
      color: #fff;
      padding: 10px 10px 9px;
      min-height: 74px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }
    .award-mini b { font-size: 12px; line-height: 1.2; }
    .award-mini strong { font-size: 22px; line-height: 1; }
    .award-mini small { color: rgba(255,255,255,.86); font-size: 11px; }
    .section { margin-top: 18px; }
    .section-head {
      display: flex; justify-content: space-between; gap: 12px; align-items: center;
      padding: 16px 16px 0;
    }
    .section-head h2 { margin: 0; font-size: 16px; }
    .section-head span { color: var(--muted); font-size: 12px; }
    .two-col { grid-template-columns: minmax(0, 1.1fr) minmax(340px, 0.9fr); align-items: stretch; }
    .two-col > .card { display: flex; flex-direction: column; }
    .two-col > .card .chart-wrap { flex: 1; display: flex; align-items: center; }
    .award-summary-card { min-width: 0; }
    .award-combined {
      display: grid;
      grid-template-columns: minmax(0, 1.05fr) minmax(0, .95fr);
      gap: 0;
      padding: 8px 16px 16px;
    }
    .award-panel {
      min-width: 0;
      display: grid;
      grid-template-rows: auto 1fr;
      gap: 8px;
      align-items: center;
    }
    .award-panel + .award-panel { border-left: 1px solid var(--border); padding-left: 16px; margin-left: 16px; }
    .award-panel-title {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      color: #0f172a;
      font-size: 13px;
      font-weight: 800;
    }
    .award-panel-title span { color: var(--muted); font-size: 12px; font-weight: 600; }
    .award-panel .chart-wrap { padding: 0; display: flex; align-items: center; }
    .three-col { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    .chart-wrap { padding: 10px 16px 14px; min-width: 0; overflow: hidden; }
    .award-pie { display: block; width: 100%; min-width: 0; }
    .award-bars { display: grid; gap: 10px; }
    .award-row { display: grid; grid-template-columns: 72px 1fr 84px; gap: 12px; align-items: center; }
    .track { height: 26px; background: #f1f5f9; border-radius: 4px; overflow: hidden; }
    .fill { height: 100%; border-radius: 4px; display: flex; align-items: center; padding-left: 10px; color: #fff; font-size: 12px; white-space: nowrap; }
    .count { text-align: right; font-variant-numeric: tabular-nums; }
    .bars { display: grid; gap: 10px; }
    .bar-row { display: grid; grid-template-columns: minmax(92px, 1fr) minmax(96px, 2fr) 58px; gap: 10px; align-items: center; }
    .bar-label { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 13px; }
    .bar-track { height: 22px; background: #f1f5f9; border-radius: 3px; overflow: hidden; }
    .bar-fill { height: 100%; background: var(--brand); border-radius: 3px; }
    .stack { display: flex; height: 22px; border-radius: 2px; overflow: hidden; background: #f1f5f9; }
    .seg { height: 100%; min-width: 2px; }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 10px 12px; border-bottom: 1px solid var(--border); text-align: left; vertical-align: middle; font-size: 13px; }
    th { color: #475569; background: #f8fafc; font-weight: 700; position: sticky; top: 0; z-index: 1; }
    td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
    .table-scroll { overflow: auto; max-height: 520px; border-top: 1px solid var(--border); margin-top: 12px; }
    .pill { display: inline-flex; align-items: center; padding: 4px 8px; border-radius: 999px; font-size: 12px; font-weight: 700; background: #f1f5f9; color: #334155; white-space: nowrap; }
    .pill.super { background: #f3e8ff; color: var(--purple); }
    .pill.first { background: #ffedd5; color: var(--orange); }
    .pill.second { background: #dbeafe; color: var(--brand); }
    .pill.third { background: #dcfce7; color: var(--green); }
    .treemap { padding: 12px 16px 18px; min-width: 0; overflow: hidden; }
    .treemap svg { display: block; width: 100%; height: auto; background: #fff; }
    .treemap-rect { stroke: #fff; stroke-width: 2.2; vector-effect: non-scaling-stroke; }
    .treemap-label { pointer-events: none; fill: #fff; font-size: 16px; font-weight: 800; text-anchor: middle; dominant-baseline: middle; }
    .treemap-sub { pointer-events: none; fill: rgba(255,255,255,.9); font-size: 12px; text-anchor: middle; dominant-baseline: middle; }
    .legend { display: flex; gap: 10px; flex-wrap: wrap; padding: 0 16px 14px; }
    .legend span { color: var(--muted); font-size: 12px; display: inline-flex; align-items: center; gap: 6px; }
    .swatch { width: 10px; height: 10px; border-radius: 3px; display: inline-block; }
    @media (max-width: 1100px) {
      .shell { grid-template-columns: 1fr; }
      .sidebar { position: static; height: auto; border-right: 0; border-bottom: 1px solid var(--border); }
      .nav { grid-template-columns: repeat(3, minmax(0, 1fr)); }
      .kpis, .two-col, .three-col { grid-template-columns: 1fr; }
      .award-combined { grid-template-columns: 1fr; gap: 16px; }
      .award-panel + .award-panel { border-left: 0; border-top: 1px solid var(--border); padding-left: 0; padding-top: 16px; margin-left: 0; }
      .award-activity { grid-column: span 1; }
      .topbar { flex-direction: column; }
      .toolbar { justify-content: flex-start; }
    }
    @media (max-width: 640px) {
      main { padding: 18px 14px 28px; }
      .nav { grid-template-columns: 1fr 1fr; }
      .toolbar, select, input { width: 100%; }
      .award-mini-grid { grid-template-columns: 1fr 1fr; }
      .treemap { padding: 10px 12px 14px; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <aside class="sidebar">
      <div class="brand">
        <div class="brand-mark">奖</div>
        <div><strong>发展研究奖</strong><span>评审管理后台</span></div>
      </div>
      <nav class="nav">
        <a href="#overview" class="active"><span class="dot"></span>评审结果分析</a>
      </nav>
    </aside>
    <main>
      <section class="topbar" id="overview">
        <div>
          <h1>重庆社科科学院发展研究奖评审结果分析</h1>
          <div class="subtle" id="metaText"></div>
        </div>
      </section>

      <section class="grid kpis" id="kpis"></section>

      <section class="section">
        <div class="card award-summary-card">
          <div class="section-head"><span id="awardTotalNote">本次评审作品总量：--份</span></div>
          <div class="award-combined">
            <div class="award-panel">
              <div class="award-panel-title">奖项占比与数量<span>按 120 份统计</span></div>
              <div class="chart-wrap"><svg class="award-pie" id="awardPie" viewBox="0 0 520 200" width="100%" height="200" role="img" aria-label="奖项占比与数量饼图"></svg></div>
            </div>
            <div class="award-panel">
              <div class="award-panel-title">奖项结构<span>特等奖本次未产生</span></div>
              <div class="chart-wrap"><svg id="donut" viewBox="0 0 420 200" width="100%" height="200" role="img" aria-label="奖项结构环形图"></svg></div>
            </div>
          </div>
        </div>
      </section>

      <section class="grid three-col section" id="units">
        <div class="card">
          <div class="section-head"><h2>单位获奖数量排名</h2></div>
          <div class="chart-wrap"><div class="bars" id="unitBars"></div></div>
        </div>
        <div class="card" id="forms">
          <div class="section-head"><h2>成果形式获奖排名</h2></div>
          <div class="chart-wrap"><div class="bars" id="formBars"></div></div>
        </div>
        <div class="card">
          <div class="section-head"><h2>所属领域获奖排名</h2></div>
          <div class="chart-wrap"><div class="bars" id="fieldBars"></div></div>
        </div>
      </section>

    </main>
  </div>
  <script>
    const data = JSON.parse(document.getElementById("dashboard-data").textContent);
    const awards = ["特等奖", "一等奖", "二等奖", "三等奖", "未获奖"];
    const awardColors = { "特等奖": "#A855F7", "一等奖": "#FF7A59", "二等奖": "#38BDF8", "三等奖": "#34D399", "未获奖": "#94A3B8" };
    data.awardSummary.forEach(item => { item.color = awardColors[item.award] || item.color; });
    const pillClass = { "特等奖": "super", "一等奖": "first", "二等奖": "second", "三等奖": "third", "未获奖": "" };

    const fmtPct = v => `${(v * 100).toFixed(1)}%`;
    const fmtNum = v => Number(v || 0).toLocaleString("zh-CN");
    const esc = value => String(value ?? "").replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[ch]));

    function renderKpis() {
      const total = data.meta.totalAssumption;
      const imported = data.meta.importedRows;
      const awarded = data.meta.awardedTotal;
      const noAward = Math.max(total - awarded, 0);
      const cards = [
        ["评审作品总量", fmtNum(total), `导入表可追踪 ${fmtNum(imported)} 份`],
        ["获奖总量", fmtNum(awarded), `总体获奖率 ${fmtPct(awarded / total)}`],
        ["未获奖数量", fmtNum(noAward), `占比 ${fmtPct(noAward / total)}`],
      ];
      const awardActivity = ["特等奖", "一等奖", "二等奖", "三等奖"].map(award => {
        const item = data.awardSummary.find(x => x.award === award) || { count: 0, share: 0 };
        return `<div class="award-mini" style="background:${awardColors[award]}">
          <b>${award}</b>
          <strong>${fmtNum(item.count)}</strong>
          <small>${fmtPct(item.share)}</small>
        </div>`;
      }).join("");
      document.getElementById("kpis").innerHTML = cards.map(([label, value, note]) =>
        `<article class="card kpi"><label>${label}</label><strong>${value}</strong><small>${note}</small></article>`
      ).join("") + `<article class="card award-activity">
        <div class="award-activity-head"><label>奖项分布</label></div>
        <div class="award-mini-grid">${awardActivity}</div>
      </article>`;
      document.getElementById("metaText").innerHTML = `<div class="award-tags">${awards.map(award =>
        `<span class="award-tag"><i style="background:${awardColors[award]}"></i>${award}</span>`
      ).join("")}</div>`;
    }

    function renderAwardBars() {
      const reviewTotal = data.awardSummary.reduce((sum, x) => sum + x.count, 0);
      document.getElementById("awardTotalNote").textContent = `本次评审作品总量：${fmtNum(reviewTotal)}份`;
      const svg = document.getElementById("awardPie");
      const total = reviewTotal || 1;
      const cx = 128, cy = 100, r = 66, stroke = 30, circumference = 2 * Math.PI * r;
      let offset = 0;
      const circles = data.awardSummary.filter(item => item.count > 0).map(item => {
        const length = item.count / total * circumference;
        const circle = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${item.color}" stroke-width="${stroke}" stroke-dasharray="${length} ${circumference - length}" stroke-dashoffset="${-offset}" transform="rotate(-90 ${cx} ${cy})" />`;
        offset += length;
        return circle;
      }).join("");
      const list = data.awardSummary.map((item, index) => {
        const y = 45 + index * 28;
        return `<g>
          <rect x="292" y="${y - 11}" width="10" height="10" rx="3" fill="${item.color}" />
          <text x="310" y="${y - 2}" font-size="12" font-weight="700" fill="#334155">${item.award}</text>
          <text x="382" y="${y - 2}" font-size="12" fill="#64748b">${fmtPct(item.share)}</text>
          <text x="462" y="${y - 2}" text-anchor="end" font-size="12" font-weight="800" fill="#0f172a">${fmtNum(item.count)}份</text>
        </g>`;
      }).join("");
      svg.innerHTML = `
        <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#f1f5f9" stroke-width="${stroke}" />
        ${circles}
        <text x="${cx}" y="${cy - 8}" text-anchor="middle" font-size="26" font-weight="800" fill="#0f172a">${fmtNum(reviewTotal)}</text>
        <text x="${cx}" y="${cy + 18}" text-anchor="middle" font-size="12" fill="#64748b">评审总量</text>
        ${list}
      `;
    }

    function renderDonut() {
      const svg = document.getElementById("donut");
      const total = data.awardSummary.reduce((sum, x) => sum + x.count, 0) || 1;
      const cx = 120, cy = 110, r = 72, stroke = 28, circumference = 2 * Math.PI * r;
      let offset = 0;
      const circles = data.awardSummary.filter(x => x.count > 0).map(item => {
        const length = item.count / total * circumference;
        const circle = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${item.color}" stroke-width="${stroke}" stroke-dasharray="${length} ${circumference - length}" stroke-dashoffset="${-offset}" transform="rotate(-90 ${cx} ${cy})" />`;
        offset += length;
        return circle;
      }).join("");
      offset = 0;
      const labels = data.awardSummary.filter(x => x.count > 0).map(item => {
        const length = item.count / total * circumference;
        const mid = ((offset + length / 2) / circumference) * Math.PI * 2 - Math.PI / 2;
        const x1 = cx + Math.cos(mid) * (r + stroke / 2 + 2);
        const y1 = cy + Math.sin(mid) * (r + stroke / 2 + 2);
        const x2 = cx + Math.cos(mid) * (r + stroke / 2 + 18);
        const y2 = cy + Math.sin(mid) * (r + stroke / 2 + 18);
        const rightSide = Math.cos(mid) >= 0;
        let x3 = x2 + (rightSide ? 18 : -18);
        let anchor = rightSide ? "start" : "end";
        if (!rightSide && x3 < 54) {
          x3 = 12;
          anchor = "start";
        }
        offset += length;
        return `<polyline points="${x1.toFixed(1)},${y1.toFixed(1)} ${x2.toFixed(1)},${y2.toFixed(1)} ${x3.toFixed(1)},${y2.toFixed(1)}" fill="none" stroke="${item.color}" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" />
          <text x="${x3}" y="${(y2 - 3).toFixed(1)}" text-anchor="${anchor}" font-size="10.5" font-weight="700" fill="${item.color}">${item.award}</text>
          <text x="${x3}" y="${(y2 + 10).toFixed(1)}" text-anchor="${anchor}" font-size="10.5" fill="#334155">${fmtPct(item.share)}</text>`;
      }).join("");
      svg.innerHTML = `
        <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#f1f5f9" stroke-width="${stroke}" />
        ${circles}
        ${labels}
        <text x="${cx}" y="${cy - 8}" text-anchor="middle" font-size="26" font-weight="800" fill="#0f172a">${fmtNum(data.meta.awardedTotal)}</text>
        <text x="${cx}" y="${cy + 18}" text-anchor="middle" font-size="12" fill="#64748b">获奖成果</text>
        <text x="240" y="82" font-size="13" fill="#334155">获奖率</text>
        <text x="240" y="116" font-size="34" font-weight="800" fill="#1d4ed8">${fmtPct(data.meta.awardedTotal / data.meta.totalAssumption)}</text>
        <text x="240" y="144" font-size="12" fill="#64748b">以 ${data.meta.totalAssumption} 份为总基数</text>
      `;
    }

    function renderStackBars(targetId, rows) {
      const max = Math.max(...rows.map(x => x.submitted), 1);
      document.getElementById(targetId).innerHTML = rows.map(row => {
        const segments = awards.map(award => {
          const count = row.counts[award] || 0;
          const width = row.submitted ? count / row.submitted * 100 : 0;
          return `<div class="seg" title="${award} ${count}" style="width:${width}%; background:${awardColors[award]}"></div>`;
        }).join("");
        return `<div class="bar-row">
          <div class="bar-label" title="${esc(row.name)}">${esc(row.name)}</div>
          <div class="stack" style="max-width:${Math.max(row.submitted / max * 100, 8)}%">${segments}</div>
          <div class="count">${row.awarded}/${row.submitted}</div>
        </div>`;
      }).join("");
    }

    function filteredTreemapRows() {
      return data.matrixRows.filter(row => row.total > 1);
    }

    function splitTreemap(items, x, y, w, h) {
      if (!items.length) return [];
      if (items.length === 1) return [{ ...items[0], x, y, w, h }];
      const total = items.reduce((sum, item) => sum + item.value, 0);
      let acc = 0;
      let splitIndex = 1;
      let bestDiff = Infinity;
      for (let i = 1; i < items.length; i += 1) {
        acc += items[i - 1].value;
        const diff = Math.abs(total / 2 - acc);
        if (diff < bestDiff) {
          bestDiff = diff;
          splitIndex = i;
        }
      }
      const first = items.slice(0, splitIndex);
      const second = items.slice(splitIndex);
      const firstTotal = first.reduce((sum, item) => sum + item.value, 0);
      const ratio = firstTotal / total;
      if (w >= h) {
        const w1 = w * ratio;
        return splitTreemap(first, x, y, w1, h).concat(splitTreemap(second, x + w1, y, w - w1, h));
      }
      const h1 = h * ratio;
      return splitTreemap(first, x, y, w, h1).concat(splitTreemap(second, x, y + h1, w, h - h1));
    }

    function shortLabel(text, maxChars) {
      if (text.length <= maxChars) return text;
      return `${text.slice(0, Math.max(2, maxChars - 1))}…`;
    }

    function treemapReviewLines(item) {
      const counts = item.counts || {};
      return [
        `参与评审：${fmtNum(item.total)}份`,
        `特等奖${fmtNum(counts["特等奖"] || 0)}份 / 一等奖${fmtNum(counts["一等奖"] || 0)}份 / 二等奖${fmtNum(counts["二等奖"] || 0)}份`,
        `三等奖${fmtNum(counts["三等奖"] || 0)}份`,
      ];
    }

    function treemapColor(total, minTotal, maxTotal) {
      const span = Math.max(maxTotal - minTotal, 1);
      const ratio = (total - minTotal) / span;
      const hue = 198 + ratio * 142;
      const saturation = 48 + ratio * 8;
      const lightness = 60 + ratio * 6;
      return `hsl(${hue.toFixed(1)} ${saturation.toFixed(1)}% ${lightness.toFixed(1)}%)`;
    }

    function renderTreemap() {
      const rows = filteredTreemapRows();
      const container = document.getElementById("treemapGrid");
      const items = rows.map(row => ({
        ...row,
        value: Math.max(row.total, 1)
      })).sort((a, b) => b.value - a.value || (b.avgScore || 0) - (a.avgScore || 0));
      const totals = items.map(item => item.total);
      const minTotal = Math.min(...totals);
      const maxTotal = Math.max(...totals, 1);
      const width = 1000;
      const height = Math.max(430, Math.min(660, 360 + items.length * 7));
      const rects = splitTreemap(items, 0, 0, width, height);
      const body = rects.map((item, index) => {
        const reviewLines = treemapReviewLines(item);
        const maxScoreLine = item.maxScore ? `最高分：${item.maxScore}` : "最高分：暂无";
        const tooltip = `${item.title}\n${reviewLines.join("\n")}\n${maxScoreLine}`;
        const fill = treemapColor(item.total, minTotal, maxTotal);
        const cx = item.x + item.w / 2;
        const cy = item.y + item.h / 2;
        const labelChars = Math.max(2, Math.floor(item.w / 18));
        const showTitle = item.w > 70 && item.h > 42;
        const showMeta = item.w > 150 && item.h > 96;
        return `<g>
          <title>${esc(tooltip)}</title>
          <rect class="treemap-rect" x="${item.x.toFixed(1)}" y="${item.y.toFixed(1)}" width="${Math.max(0, item.w).toFixed(1)}" height="${Math.max(0, item.h).toFixed(1)}" fill="${fill}" />
          ${showTitle ? `<text class="treemap-label" x="${cx.toFixed(1)}" y="${(showMeta ? cy - 28 : cy).toFixed(1)}">${esc(shortLabel(item.title, labelChars))}</text>` : ""}
          ${showMeta ? `<text class="treemap-sub" x="${cx.toFixed(1)}" y="${(cy - 1).toFixed(1)}">${esc(reviewLines[0])}</text>
          <text class="treemap-sub" x="${cx.toFixed(1)}" y="${(cy + 17).toFixed(1)}">${esc(shortLabel(reviewLines[1], labelChars + 6))}</text>
          <text class="treemap-sub" x="${cx.toFixed(1)}" y="${(cy + 35).toFixed(1)}">${esc(reviewLines[2])}</text>` : ""}
        </g>`;
      }).join("");
      container.innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="成果名称获奖矩形树图">${body}</svg>`;
    }

    function init() {
      renderKpis();
      renderAwardBars();
      renderDonut();
      renderStackBars("unitBars", data.unitStats);
      renderStackBars("formBars", data.formStats);
      renderStackBars("fieldBars", data.fieldStats);
    }
    init();
  </script>
</body>
</html>
"""

(OUTPUT_DIR / "社科院发展研究奖_后台分析看板.html").write_text(
    html_template.replace("__DATA__", json.dumps(payload, ensure_ascii=False)),
    encoding="utf-8",
)
