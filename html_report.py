"""Interactive, touch-friendly HTML report (Power BI style). One self-contained file, works offline."""
from __future__ import annotations

import json
from datetime import date

import pandas as pd

def pct_yes(s: pd.Series) -> float:
    answered = s.isin(["Yes", "No"]).sum()
    return (s == "Yes").sum() / answered * 100 if answered else 0.0


def kpi_items(df, summary):
    items = [(f"{len(df):,}", "Responses", False), (str(len(summary)), "Yes / No questions", False)]
    if "Highlight" in summary:
        for _, r in summary[summary["Highlight"]].head(3).iterrows():
            items.append((f"{r['% Yes']:.0f}%", _clean(r["Question"]), True))
    return items


def _clean(q: str) -> str:
    t = " ".join(str(q).split())
    if len(t) > 3 and t[0].isdigit() and "." in t[:4]:
        t = t.split(".", 1)[1].strip()
    return t


ALL_SECTIONS = ["summary", "notes", "profile", "yesno", "compare"]


def build_html(df, report_title, summary, cat_cols, filters_text, group=None, chosen=None,
               notes=None, sections=None) -> bytes:
    secs = list(sections or ALL_SECTIONS)
    label_of = dict(zip(summary["Original"], summary["Question"])) if "Original" in summary else {}
    data = {
        "sections": secs,
        "notes": {"findings": list((notes or {}).get("findings") or []),
                  "recommendations": list((notes or {}).get("recommendations") or [])},
        "title": report_title,
        "filters": filters_text,
        "n": int(len(df)),
        "generated": f"{date.today():%d %B %Y}",
        "kpis": [{"value": v, "label": l, "red": bool(r)} for v, l, r in kpi_items(df, summary)],
        "profile": [
            {"name": c.rstrip(": "),
             "items": [{"label": str(k), "count": int(v)} for k, v in df[c].value_counts().head(8).items()]}
            for c in cat_cols[:6]
        ],
        "questions": [
            {"q": _clean(r["Question"]), "yes": int(r["Yes"]), "no": int(r["No"]), "pct": float(r["% Yes"]),
             "red": bool(r.get("Highlight", False))}
            for _, r in summary.iterrows()
        ],
        "compare": None,
    }
    if group and chosen:
        sizes = df.groupby(group).size()
        table = df.groupby(group)[chosen].agg(pct_yes)
        data["compare"] = {
            "group": group.rstrip(": "),
            "groups": [{"label": str(g), "n": int(sizes[g])} for g in table.index],
            "rows": [{"q": _clean(label_of.get(q, q)), "vals": [round(float(table.loc[g, q]), 1) for g in table.index]}
                     for q in chosen],
        }
    payload = json.dumps(data).replace("</", "<\\/")
    return TEMPLATE.replace("__DATA__", payload).replace("__TITLE__", _html_escape(report_title.replace("\n", " – "))).encode("utf-8")


def _html_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


TEMPLATE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
:root{--blue:#118DFF;--navy:#12239E;--red:#D64550;--no:#E6E6E6;--canvas:#F3F2F1;--ink:#252423;--muted:#605E5C;--grid:#EDEBE9}
*{box-sizing:border-box}
body{margin:0;background:var(--canvas);color:var(--ink);font-family:"Segoe UI",-apple-system,Helvetica,Arial,sans-serif;-webkit-tap-highlight-color:transparent}
header{background:var(--blue);color:#fff;padding:22px 20px 18px;border-left:8px solid var(--navy)}
header{position:relative;overflow:hidden;padding:28px 24px 22px}
header:after{content:"";position:absolute;right:-90px;top:-120px;width:340px;height:340px;border-radius:50%;background:rgba(255,255,255,.08)}
header .lbl{font-size:11px;letter-spacing:.35em;font-weight:700;color:#CFE6FF;margin-bottom:10px}
header h1{margin:0;font-size:clamp(24px,4.4vw,36px);font-weight:700;letter-spacing:-.01em}
header .t2{font-size:clamp(17px,2.6vw,22px);font-weight:400;color:#E3F1FF;margin-top:4px}
header .by{margin-top:14px;padding-top:10px;border-top:1px solid rgba(255,255,255,.45);display:inline-block;min-width:240px}
header .by small{display:block;font-size:10.5px;letter-spacing:.3em;font-weight:700;color:#CFE6FF}
header .by{margin-top:26px}
header .by b{display:block;font-size:clamp(17px,2.6vw,22px);margin-top:4px}
header .by span{display:block;color:#E3F1FF;font-size:14px;margin-top:2px}
header p{margin:6px 0 0;color:#E3F1FF;font-size:14px}
nav{position:sticky;top:0;z-index:5;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.12);display:flex;gap:4px;overflow-x:auto;padding:0 12px}
nav a{padding:12px 14px;color:var(--muted);text-decoration:none;font-size:14px;font-weight:600;white-space:nowrap;border-bottom:3px solid transparent}
nav a:hover{color:var(--ink);border-bottom-color:var(--blue)}
main{max-width:1200px;margin:0 auto;padding:16px}
section{scroll-margin-top:56px;margin-bottom:22px}
h2{font-size:18px;margin:6px 0 10px}
.card{background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.12);padding:16px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
.kpi{background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.12);border-top:4px solid var(--blue);padding:14px 16px}
.kpi.red{border-top-color:var(--red)}
.kpi .v{font-size:32px;font-weight:700}.kpi.red .v{color:var(--red)}
.kpi .l{font-size:12.5px;color:var(--muted);margin-top:4px;line-height:1.3}
.donuts{display:grid;grid-template-columns:repeat(auto-fit,minmax(290px,1fr));gap:12px}
.donut-card h3{margin:0 0 6px;font-size:15px}
.donut-wrap{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.donut-wrap svg{width:150px;height:150px;flex:none}
.donut-wrap path{cursor:pointer;transition:opacity .15s,transform .15s;transform-origin:75px 75px}
.donut-wrap path.dim{opacity:.3}
.donut-wrap path.on{transform:scale(1.06)}
.legend{list-style:none;margin:0;padding:0;font-size:13px;flex:1;min-width:120px}
.legend li{display:flex;align-items:center;gap:8px;padding:3px 4px;border-radius:4px;cursor:pointer}
.legend li.on{background:#EAF4FF}
.legend i{width:11px;height:11px;border-radius:2px;flex:none}
.legend b{margin-left:auto;font-weight:600}
.center-n{font-size:22px;font-weight:700;fill:var(--ink)}.center-l{font-size:10px;fill:var(--muted)}
.tools{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}
.tools input{flex:1;min-width:180px;padding:9px 12px;border:1px solid #C8C6C4;border-radius:6px;font-size:14px}
.seg{display:flex;border:1px solid #C8C6C4;border-radius:6px;overflow:hidden}
.seg button{border:0;background:#fff;padding:9px 12px;font-size:13px;cursor:pointer;color:var(--muted)}
.seg button.on{background:var(--blue);color:#fff}
.key{display:flex;gap:14px;font-size:12.5px;color:var(--muted);margin:2px 0 10px;flex-wrap:wrap}
.key span{display:flex;align-items:center;gap:6px}.key i{width:11px;height:11px;border-radius:2px}
.row{display:grid;grid-template-columns:minmax(140px,38%) 1fr;gap:12px;align-items:center;padding:6px 4px;border-radius:6px;cursor:pointer}
.row:hover,.row.on{background:#F5F9FF}
.row .q{font-size:13px;line-height:1.3}
.bar{display:flex;height:26px;border-radius:3px;overflow:hidden;background:var(--no)}
.bar .y{background:var(--blue);color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;padding-left:6px;white-space:nowrap;transition:width .5s ease}
.bar .y.red{background:var(--red)}
.bar .n{flex:1;color:var(--muted);font-size:12px;display:flex;align-items:center;justify-content:flex-end;padding-right:6px}
.bar .y.out{color:var(--ink);background:transparent;padding-left:4px}
.detail{grid-column:1/-1;font-size:12.5px;color:var(--muted);display:none;padding:2px 0 4px}
.row.on .detail{display:block}
@media (max-width:640px){.row{grid-template-columns:1fr}.row .q{font-size:13.5px}}
.cmp-row{padding:8px 4px;border-bottom:1px solid var(--grid)}
.cmp-row .q{font-size:13px;margin-bottom:6px}
.cbar{display:flex;align-items:center;gap:8px;margin:3px 0;cursor:pointer}
.cbar .track{flex:1;height:18px;background:#F3F2F1;border-radius:3px;overflow:hidden}
.cbar .fill{height:100%;transition:width .5s ease}
.cbar .val{width:44px;text-align:right;font-size:12.5px;font-weight:600}
.cbar.hide{display:none}
.glegend{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:8px}
.glegend button{display:flex;align-items:center;gap:6px;border:1px solid #C8C6C4;background:#fff;border-radius:16px;padding:6px 12px;font-size:13px;cursor:pointer}
.glegend button.off{opacity:.4}
.glegend i{width:11px;height:11px;border-radius:2px}
.notes{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:12px}
.notes .card{border-top:4px solid var(--blue)}.notes .card.rec{border-top-color:var(--navy)}
.notes h3{margin:0 0 10px;font-size:16px}.notes ol,.notes ul{margin:0;padding-left:20px}
.notes li{margin:0 0 8px;line-height:1.45;font-size:14.5px}
#tip{position:fixed;z-index:20;pointer-events:none;background:#252423;color:#fff;font-size:13px;line-height:1.4;padding:9px 11px;border-radius:6px;max-width:280px;box-shadow:0 4px 14px rgba(0,0,0,.25);opacity:0;transition:opacity .12s}
#tip.show{opacity:1}
footer{color:var(--muted);font-size:12px;text-align:center;padding:10px 0 24px}
.hint{font-size:12.5px;color:var(--muted);margin:-4px 0 10px}
</style></head>
<body>
<header><h1 id="t"></h1><p id="sub"></p></header>
<nav><a href="#summary">Summary</a><a href="#notes">Findings</a><a href="#profile">Respondent profile</a><a href="#yesno">Yes / No results</a><a href="#compare" id="navcmp">Comparison</a></nav>
<main>
<section id="summary"><h2>Summary</h2><div class="kpis" id="kpis"></div></section>
<section id="notes"><h2>Key findings &amp; recommendations</h2><div class="notes" id="notesBox"></div></section>
<section id="profile"><h2>Respondent profile</h2><p class="hint">Tap a slice or a legend item to see the numbers.</p><div class="donuts" id="donuts"></div></section>
<section id="yesno"><h2>Yes / No results</h2>
 <div class="card">
  <div class="tools"><input id="search" type="search" placeholder="Search questions…">
   <div class="seg" id="sort"><button data-s="orig" class="on">Original</button><button data-s="desc">Highest</button><button data-s="asc">Lowest</button></div></div>
  <div class="key"><span><i style="background:var(--blue)"></i>Yes</span><span><i style="background:var(--red)"></i>Yes (key indicator)</span><span><i style="background:var(--no)"></i>No</span><span id="count"></span></div>
  <p class="hint">Tap any question to see the exact counts.</p>
  <div id="rows"></div>
 </div></section>
<section id="compare"><h2 id="cmpTitle"></h2><div class="card"><p class="hint">Tap a group to show or hide it. Tap a bar for details.</p><div class="glegend" id="glegend"></div><div id="cmp"></div></div></section>
<footer id="foot"></footer>
</main>
<div id="tip"></div>
<script>
const D = __DATA__;
const PAL = ["#118DFF","#12239E","#E66C37","#6B007B","#E044A7","#744EC2","#D9B300","#D64550"];
const $ = s => document.querySelector(s);
const esc = s => String(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const fmt = v => Math.round(v) + "%";

const TL = D.title.split("\n").map(s => s.trim());
$("#t").textContent = TL[0] || "Survey Report";
$("#t").insertAdjacentHTML("beforebegin", '<div class="lbl">REPORT</div>');
{ const role = (TL[1] || "").trim(), name = (TL[2] || "").trim();
  if (role || name) $("#sub").insertAdjacentHTML("afterend",
    `<div class="by"><small>PREPARED BY</small>${name ? `<b>${esc(name)}</b>` : ""}${role ? `<span>${esc(role)}</span>` : ""}</div>`); }
$("#sub").textContent = D.n.toLocaleString() + " responses" + (D.filters ? " · " + D.filters : "");
$("#foot").textContent = TL.slice(0, 2).filter(Boolean).join(" – ") + " · Aggregate results only, no individual data";

/* ---------- tooltip (mouse hover + touch tap) ---------- */
const tip = $("#tip");
function showTip(html, x, y){
  tip.innerHTML = html; tip.classList.add("show");
  const r = tip.getBoundingClientRect();
  let left = Math.min(x + 14, window.innerWidth - r.width - 8), top = y + 14;
  if (top + r.height > window.innerHeight - 8) top = y - r.height - 14;
  tip.style.left = Math.max(8, left) + "px"; tip.style.top = Math.max(8, top) + "px";
}
function hideTip(){ tip.classList.remove("show"); }
function bindTip(el, htmlFn){
  el.addEventListener("pointermove", e => { if (e.pointerType === "mouse") showTip(htmlFn(), e.clientX, e.clientY); });
  el.addEventListener("pointerleave", e => { if (e.pointerType === "mouse") hideTip(); });
  el.addEventListener("click", e => { e.stopPropagation(); showTip(htmlFn(), e.clientX, e.clientY); });
}
document.addEventListener("click", hideTip);
window.addEventListener("scroll", hideTip, {passive:true});

/* ---------- Sections chosen in the app ---------- */
const NAV = {summary:"summary", notes:"notes", profile:"profile", yesno:"yesno", compare:"compare"};
Object.keys(NAV).forEach(k => { if (!D.sections.includes(k)) {
  $("#" + NAV[k]).style.display = "none";
  const a = document.querySelector(`nav a[href="#${NAV[k]}"]`); if (a) a.style.display = "none"; } });
{ const F = D.notes.findings, R = D.notes.recommendations;
  if (!F.length && !R.length) { $("#notes").style.display = "none"; document.querySelector('nav a[href="#notes"]').style.display = "none"; }
  $("#notesBox").innerHTML =
    (F.length ? `<div class="card"><h3>Key findings</h3><ul>${F.map(x => `<li>${esc(x)}</li>`).join("")}</ul></div>` : "") +
    (R.length ? `<div class="card rec"><h3>Recommendations</h3><ol>${R.map(x => `<li>${esc(x)}</li>`).join("")}</ol></div>` : ""); }

/* ---------- KPI cards ---------- */
$("#kpis").innerHTML = D.kpis.map(k =>
  `<div class="kpi ${k.red ? "red" : ""}"><div class="v">${esc(k.value)}</div><div class="l">${esc(k.label)}</div></div>`).join("");

/* ---------- Donuts ---------- */
function arc(cx, cy, r0, r1, a0, a1){
  const p = (r, a) => [cx + r * Math.sin(a), cy - r * Math.cos(a)];
  const large = a1 - a0 > Math.PI ? 1 : 0;
  const [x0,y0] = p(r1,a0), [x1,y1] = p(r1,a1), [x2,y2] = p(r0,a1), [x3,y3] = p(r0,a0);
  return `M${x0},${y0} A${r1},${r1} 0 ${large} 1 ${x1},${y1} L${x2},${y2} A${r0},${r0} 0 ${large} 0 ${x3},${y3} Z`;
}
D.profile.forEach((p, pi) => {
  const total = p.items.reduce((s, i) => s + i.count, 0);
  let a = 0, paths = "";
  p.items.forEach((it, i) => {
    let span = it.count / total * Math.PI * 2;
    if (span >= Math.PI * 2) span = Math.PI * 2 - 1e-4;
    paths += `<path data-i="${i}" d="${arc(75,75,44,72,a,a+span)}" fill="${PAL[i % PAL.length]}" stroke="#fff" stroke-width="1.5"></path>`;
    a += it.count / total * Math.PI * 2;
  });
  const card = document.createElement("div");
  card.className = "card donut-card";
  card.innerHTML = `<h3>${esc(p.name)}</h3><div class="donut-wrap">
    <svg viewBox="0 0 150 150">${paths}<text class="center-n" x="75" y="78" text-anchor="middle">${total}</text>
    <text class="center-l" x="75" y="94" text-anchor="middle">responses</text></svg>
    <ul class="legend">${p.items.map((it, i) =>
      `<li data-i="${i}"><i style="background:${PAL[i % PAL.length]}"></i>${esc(it.label)}<b>${(it.count/total*100).toFixed(0)}%</b></li>`).join("")}</ul></div>`;
  $("#donuts").appendChild(card);
  const pathsEl = card.querySelectorAll("path"), lis = card.querySelectorAll("li");
  const tipHtml = i => `<b>${esc(p.name)}: ${esc(p.items[i].label)}</b><br>${p.items[i].count} of ${total} (${(p.items[i].count/total*100).toFixed(1)}%)`;
  const focus = i => {
    pathsEl.forEach((el, j) => { el.classList.toggle("dim", i !== null && j !== i); el.classList.toggle("on", j === i); });
    lis.forEach((el, j) => el.classList.toggle("on", j === i));
  };
  [...pathsEl, ...lis].forEach(el => {
    const i = +el.dataset.i;
    bindTip(el, () => tipHtml(i));
    el.addEventListener("pointerenter", () => focus(i));
    el.addEventListener("click", () => focus(i));
  });
  card.addEventListener("pointerleave", () => focus(null));
});

/* ---------- Yes / No rows ---------- */
let sortMode = "orig";
function renderRows(){
  const term = $("#search").value.trim().toLowerCase();
  let qs = D.questions.map((q, i) => ({...q, i})).filter(q => q.q.toLowerCase().includes(term));
  if (sortMode === "desc") qs.sort((a, b) => b.pct - a.pct);
  if (sortMode === "asc") qs.sort((a, b) => a.pct - b.pct);
  $("#count").textContent = `${qs.length} of ${D.questions.length} questions`;
  $("#rows").innerHTML = qs.map(q => {
    const inside = q.pct >= 9;
    return `<div class="row" data-i="${q.i}"><div class="q">${esc(q.q)}</div>
      <div class="bar"><div class="y ${q.red ? "red" : ""} ${inside ? "" : "out"}" style="width:0" data-w="${q.pct}">${inside ? fmt(q.pct) : ""}</div>
      <div class="n">${!inside ? `<span style="margin-right:auto;padding-left:4px;color:var(--ink);font-weight:700">${fmt(q.pct)}</span>` : ""}${100 - q.pct >= 8 ? fmt(100 - q.pct) : ""}</div></div>
      <div class="detail">Yes: <b>${q.yes}</b> · No: <b>${q.no}</b> · Answered: <b>${q.yes + q.no}</b> · ${q.pct.toFixed(1)}% Yes</div></div>`;
  }).join("");
  requestAnimationFrame(() => document.querySelectorAll(".bar .y").forEach(el => el.style.width = el.dataset.w + "%"));
  document.querySelectorAll(".row").forEach(row => {
    const q = D.questions[+row.dataset.i];
    row.addEventListener("click", () => row.classList.toggle("on"));
    bindTip(row.querySelector(".bar"), () =>
      `<b>${esc(q.q)}</b><br>Yes: ${q.yes} (${q.pct.toFixed(1)}%)<br>No: ${q.no} (${(100 - q.pct).toFixed(1)}%)<br>Answered: ${q.yes + q.no}`);
  });
}
$("#search").addEventListener("input", renderRows);
document.querySelectorAll("#sort button").forEach(b => b.addEventListener("click", () => {
  sortMode = b.dataset.s;
  document.querySelectorAll("#sort button").forEach(x => x.classList.toggle("on", x === b));
  renderRows();
}));
renderRows();

/* ---------- Comparison ---------- */
if (!D.compare){
  $("#compare").style.display = "none"; $("#navcmp").style.display = "none";
} else {
  const C = D.compare, hidden = new Set();
  $("#cmpTitle").textContent = "% answering Yes by " + C.group;
  $("#glegend").innerHTML = C.groups.map((g, i) =>
    `<button data-g="${i}"><i style="background:${PAL[i % PAL.length]}"></i>${esc(g.label)} (n=${g.n})</button>`).join("");
  $("#cmp").innerHTML = C.rows.map(r => `<div class="cmp-row"><div class="q">${esc(r.q)}</div>${
    r.vals.map((v, i) => `<div class="cbar" data-g="${i}"><div class="track"><div class="fill" style="width:0;background:${PAL[i % PAL.length]}" data-w="${v}"></div></div><div class="val">${fmt(v)}</div></div>`).join("")}</div>`).join("");
  requestAnimationFrame(() => document.querySelectorAll(".cbar .fill").forEach(el => el.style.width = el.dataset.w + "%"));
  document.querySelectorAll(".cmp-row").forEach((rowEl, ri) => rowEl.querySelectorAll(".cbar").forEach(b => {
    const gi = +b.dataset.g, g = C.groups[gi], r = C.rows[ri];
    bindTip(b, () => `<b>${esc(r.q)}</b><br>${esc(g.label)} (n=${g.n}): <b>${r.vals[gi].toFixed(1)}%</b> Yes`);
  }));
  document.querySelectorAll("#glegend button").forEach(btn => btn.addEventListener("click", e => {
    e.stopPropagation();
    const gi = btn.dataset.g;
    hidden.has(gi) ? hidden.delete(gi) : hidden.add(gi);
    btn.classList.toggle("off", hidden.has(gi));
    document.querySelectorAll(`.cbar[data-g="${gi}"]`).forEach(b => b.classList.toggle("hide", hidden.has(gi)));
  }));
}
</script>
</body></html>
"""
