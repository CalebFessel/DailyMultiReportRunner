# -*- coding: utf-8 -*-
"""Generate the facility turnaround dashboard as a self-contained HTML page."""
import csv, json, os, sys

OFFLINE = "--offline" in sys.argv

SRC = r"C:\Reports\facility_turnaround_2026-08.csv"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "turnaround-dashboard-offline.html" if OFFLINE else "turnaround-dashboard.html")

rows = list(csv.DictReader(open(SRC, encoding="utf-8")))
def f(r, k):
    v = r.get(k, "")
    return float(v) if v not in ("", None) else None

data = []
for r in rows:
    data.append({
        "n": r["facility"],
        "v": int(r["visits"]),
        "m": round(f(r, "median_min"), 1),
        "p": round(f(r, "p90_min"), 1),
        "h": round(f(r, "hours"), 1),
        "d": round(f(r, "delta_min"), 1) if f(r, "delta_min") is not None else None,
        "pm": round(f(r, "prev_median_min"), 1) if f(r, "prev_median_min") is not None else None,
    })

fleet = round(f(rows[0], "fleet_median_min"), 1)
tot_h = round(sum(d["h"] for d in data))
tot_v = sum(d["v"] for d in data)
above = [d for d in data if d["m"] > fleet]
above_h = round(sum(d["h"] for d in above))

meta = {"fleet": fleet, "totalHours": tot_h, "totalVisits": tot_v,
        "facilities": len(data), "aboveCount": len(above), "aboveHours": above_h,
        "abovePct": round(100.0 * above_h / tot_h)}

HTML = """<title>Facility Turnaround</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Public+Sans:wght@400;500;600;700&family=Roboto+Mono:wght@400;500&display=swap">
<style>
:root{
  --ground:#f2f4f5; --surface:#fcfcfb; --surface-2:#e9edef;
  --ink:#14191d; --muted:#5c666e; --faint:#8b959c; --border:#d5dbdf; --rule:#c2cace;
  --series-1:#2a78d6; --series-2:#eb6834;
  --good:#0ca30c; --critical:#d03b3b;
  --grid:#e3e8ea;
}
@media (prefers-color-scheme: dark){ :root:not([data-theme="light"]){
  --ground:#111414; --surface:#1a1a19; --surface-2:#232624;
  --ink:#f2f3f1; --muted:#a4ada8; --faint:#78817d; --border:#2e3230; --rule:#3a3f3c;
  --series-1:#3987e5; --series-2:#d95926;
  --good:#0ca30c; --critical:#d03b3b;
  --grid:#2a2e2c;
}}
:root[data-theme="dark"]{
  --ground:#111414; --surface:#1a1a19; --surface-2:#232624;
  --ink:#f2f3f1; --muted:#a4ada8; --faint:#78817d; --border:#2e3230; --rule:#3a3f3c;
  --series-1:#3987e5; --series-2:#d95926;
  --good:#0ca30c; --critical:#d03b3b;
  --grid:#2a2e2c;
}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);
  font-family:"Public Sans",system-ui,sans-serif;font-size:15px;line-height:1.55;-webkit-font-smoothing:antialiased}
.wrap{max-width:1080px;margin:0 auto;padding:40px 22px 80px;display:flex;flex-direction:column;gap:26px}
.eyebrow{font-family:"Roboto Mono",monospace;font-size:11px;letter-spacing:.15em;text-transform:uppercase;color:var(--series-1);margin:0}
h1{font-size:clamp(28px,4.4vw,38px);font-weight:700;letter-spacing:-.02em;line-height:1.1;margin:8px 0 0;text-wrap:balance}
.sub{color:var(--muted);margin:10px 0 0;max-width:66ch}
header{border-bottom:2px solid var(--ink);padding-bottom:22px}
h2{font-size:12px;font-weight:600;letter-spacing:.11em;text-transform:uppercase;color:var(--muted);
  margin:0 0 4px;padding-bottom:8px;border-bottom:1px solid var(--rule)}
.note{color:var(--muted);font-size:13.5px;margin:0 0 14px;max-width:74ch}
section{display:flex;flex-direction:column}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(178px,1fr));gap:1px;background:var(--border);border:1px solid var(--border)}
.tile{background:var(--surface);padding:18px 18px 20px;display:flex;flex-direction:column;gap:5px}
.tile .k{font-family:"Roboto Mono",monospace;font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted)}
.tile .n{font-size:32px;font-weight:700;line-height:1;letter-spacing:-.02em;font-variant-numeric:tabular-nums}
.tile .s{font-size:13px;color:var(--muted);line-height:1.4}
.card{background:var(--surface);border:1px solid var(--border);padding:18px 18px 10px}
.legend{display:flex;gap:18px;flex-wrap:wrap;font-size:13px;color:var(--muted);margin:0 0 8px;align-items:center}
.legend i{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px;vertical-align:-1px}
.scroll{overflow-x:auto}
svg{display:block;max-width:100%;height:auto}
.tick{font-family:"Roboto Mono",monospace;font-size:10px;fill:var(--faint)}
.axlab{font-size:11.5px;fill:var(--muted);font-weight:500}
.rowlab{font-size:12px;fill:var(--ink)}
.val{font-family:"Roboto Mono",monospace;font-size:11px;fill:var(--muted)}
.tip{position:fixed;pointer-events:none;background:var(--surface);border:1px solid var(--rule);
  padding:9px 11px;font-size:12.5px;box-shadow:0 4px 14px rgba(0,0,0,.14);opacity:0;transition:opacity .1s;z-index:9;max-width:250px}
.tip b{display:block;font-size:13px;margin-bottom:3px}
.tip span{font-family:"Roboto Mono",monospace;color:var(--muted)}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th,td{padding:7px 10px;text-align:right;border-bottom:1px solid var(--border);font-variant-numeric:tabular-nums;white-space:nowrap}
th:first-child,td:first-child{text-align:left;white-space:normal}
thead th{font-family:"Roboto Mono",monospace;font-size:10px;letter-spacing:.07em;text-transform:uppercase;
  color:var(--muted);font-weight:500;border-bottom:1px solid var(--rule)}
tbody tr:last-child td{border-bottom:none}
details{margin-top:6px}
summary{cursor:pointer;font-size:13px;color:var(--series-1);padding:6px 0}
footer{border-top:1px solid var(--rule);padding-top:16px;font-size:13px;color:var(--muted)}
code{font-family:"Roboto Mono",monospace;font-size:.87em;background:var(--surface-2);padding:1px 5px}
</style>

<div class="wrap">
<header>
  <p class="eyebrow">Fleet operations &middot; August 2026</p>
  <h1>Facility Turnaround</h1>
  <p class="sub">How long crews spend at each facility per transport, measured from Samsara geofence
  entry to exit. Posting time and base visits are excluded, so what remains is time spent
  on a patient handoff.</p>
</header>

<section>
  <div class="tiles" id="tiles"></div>
</section>

<section>
  <h2>Where the hours actually go</h2>
  <p class="note">Each dot is a facility. Right means more transports; up means slower per transport.
  Dot size is total crew hours. <b>The top-right quadrant is where time is concentrated</b> &mdash; a slow
  facility you rarely visit costs less than a moderate one you visit constantly.</p>
  <div class="card">
    <div class="legend"><span><i style="background:var(--series-1)"></i>Facility</span>
      <span>Dashed line = fleet median</span><span>Larger dot = more crew hours</span></div>
    <div class="scroll"><svg id="scatter" viewBox="0 0 980 430" role="img"
      aria-label="Scatter of visits versus median turnaround, sized by total crew hours"></svg></div>
  </div>
</section>

<section>
  <h2>Slowest facilities, and how unpredictable they are</h2>
  <p class="note">Blue is the typical visit; orange is the 90th percentile &mdash; the bad-but-routine case
  that happens about one visit in ten. <b>A long bar means unpredictable</b>, which disrupts scheduling
  more than a slow-but-consistent facility does.</p>
  <div class="card">
    <div class="legend"><span><i style="background:var(--series-1)"></i>Median</span>
      <span><i style="background:var(--series-2)"></i>90th percentile</span></div>
    <div class="scroll"><svg id="dumb" viewBox="0 0 980 560" role="img"
      aria-label="Median and 90th percentile turnaround for the fifteen slowest facilities"></svg></div>
  </div>
</section>

<section>
  <h2>Biggest changes since July</h2>
  <p class="note">Facilities whose typical turnaround moved most month over month. Direction is labelled,
  not carried by colour alone.</p>
  <div class="card">
    <div class="scroll"><svg id="movers" viewBox="0 0 980 400" role="img"
      aria-label="Month over month change in median turnaround"></svg></div>
  </div>
</section>

<section>
  <h2>The numbers</h2>
  <div class="card scroll">
    <table id="tbl"><thead><tr><th>Facility</th><th>Visits</th><th>Median</th><th>P90</th>
      <th>Hours</th><th>vs July</th><th>vs fleet</th></tr></thead><tbody></tbody></table>
    <details><summary id="more"></summary><div class="scroll"><table id="tbl2"><tbody></tbody></table></div></details>
  </div>
</section>

<footer id="foot"></footer>
</div>
<div class="tip" id="tip"></div>

<script>
const DATA = __DATA__, META = __META__;
const $ = s => document.querySelector(s);
const tip = $("#tip");
function showTip(e, html){ tip.innerHTML = html; tip.style.opacity = 1;
  const p = 14; let x = e.clientX + p, y = e.clientY + p;
  const r = tip.getBoundingClientRect();
  if (x + r.width > innerWidth - 8) x = e.clientX - r.width - p;
  if (y + r.height > innerHeight - 8) y = e.clientY - r.height - p;
  tip.style.left = x + "px"; tip.style.top = y + "px"; }
function hideTip(){ tip.style.opacity = 0; }
const NS = "http://www.w3.org/2000/svg";
const el = (n, a) => { const e = document.createElementNS(NS, n);
  for (const k in a) e.setAttribute(k, a[k]); return e; };

/* ---- tiles ---- */
$("#tiles").innerHTML = [
  ["Fleet median", META.fleet + " min", "Typical turnaround across " + META.facilities + " facilities"],
  ["Transports", META.totalVisits.toLocaleString(), "Patient handoffs measured in August"],
  ["Crew hours", META.totalHours.toLocaleString(), "Total time spent at facilities"],
  ["Above median", META.aboveCount + " sites", META.abovePct + "% of all facility hours"]
].map(t => `<div class="tile"><span class="k">${t[0]}</span><span class="n">${t[1]}</span><span class="s">${t[2]}</span></div>`).join("");

/* ---- scatter ---- */
(function(){
  const s = $("#scatter"), W = 980, H = 430, L = 58, R = 20, T = 18, B = 46;
  const xs = DATA.map(d => d.v), ys = DATA.map(d => d.m), hs = DATA.map(d => d.h);
  const xMax = Math.max(...xs) * 1.05, yMax = Math.max(...ys) * 1.08, hMax = Math.max(...hs);
  const X = v => L + (v / xMax) * (W - L - R);
  const Y = v => H - B - (v / yMax) * (H - T - B);
  const Rr = h => 3.5 + Math.sqrt(h / hMax) * 15;
  for (let i = 0; i <= 4; i++){ const v = yMax * i / 4, y = Y(v);
    s.appendChild(el("line", {x1:L, x2:W-R, y1:y, y2:y, stroke:"var(--grid)", "stroke-width":1}));
    s.appendChild(el("text", {x:L-9, y:y+3.5, "text-anchor":"end", class:"tick"})).textContent = Math.round(v); }
  for (let i = 0; i <= 4; i++){ const v = xMax * i / 4;
    s.appendChild(el("text", {x:X(v), y:H-B+18, "text-anchor":"middle", class:"tick"})).textContent = Math.round(v); }
  const fy = Y(META.fleet);
  s.appendChild(el("line", {x1:L, x2:W-R, y1:fy, y2:fy, stroke:"var(--muted)", "stroke-width":1.5, "stroke-dasharray":"5 4"}));
  s.appendChild(el("text", {x:W-R-4, y:fy-7, "text-anchor":"end", class:"axlab"})).textContent = "fleet median " + META.fleet + " min";
  s.appendChild(el("text", {x:L, y:H-6, class:"axlab"})).textContent = "transports in August \\u2192";
  const yl = el("text", {x:0, y:0, class:"axlab", transform:"translate(14," + (T+120) + ") rotate(-90)"});
  yl.textContent = "median minutes per transport \\u2192"; s.appendChild(yl);
  DATA.slice().sort((a,b)=>b.h-a.h).forEach(d => {
    const c = el("circle", {cx:X(d.v), cy:Y(d.m), r:Rr(d.h),
      fill:"var(--series-1)", "fill-opacity":.5, stroke:"var(--surface)", "stroke-width":2, tabindex:"0"});
    const html = `<b>${d.n}</b><span>${d.v} transports &middot; median ${d.m} min<br>p90 ${d.p} min &middot; ${d.h} crew hours</span>`;
    c.addEventListener("mousemove", e => showTip(e, html));
    c.addEventListener("mouseleave", hideTip);
    c.addEventListener("focus", e => showTip({clientX:c.getBoundingClientRect().x, clientY:c.getBoundingClientRect().y}, html));
    c.addEventListener("blur", hideTip);
    s.appendChild(c); });
  [...DATA].sort((a,b)=>b.h-a.h).slice(0,5).forEach(d => {
    const t = el("text", {x:X(d.v), y:Y(d.m)-Rr(d.h)-6, "text-anchor":"middle", class:"val"});
    t.textContent = d.n.replace(/^\\*/,"").split(" ").slice(0,2).join(" "); s.appendChild(t); });
})();

/* ---- dumbbell ---- */
(function(){
  const s = $("#dumb"), W = 980, T = 26, B = 40, L = 250, R = 60;
  const top = [...DATA].sort((a,b)=>b.m-a.m).slice(0,15);
  const rowH = (560 - T - B) / top.length;
  const xMax = Math.max(...top.map(d=>d.p)) * 1.06;
  const X = v => L + (v / xMax) * (W - L - R);
  for (let i = 0; i <= 4; i++){ const v = xMax*i/4;
    s.appendChild(el("line", {x1:X(v), x2:X(v), y1:T-6, y2:560-B, stroke:"var(--grid)", "stroke-width":1}));
    s.appendChild(el("text", {x:X(v), y:560-B+18, "text-anchor":"middle", class:"tick"})).textContent = Math.round(v); }
  const fx = X(META.fleet);
  s.appendChild(el("line", {x1:fx, x2:fx, y1:T-6, y2:560-B, stroke:"var(--muted)", "stroke-width":1.5, "stroke-dasharray":"5 4"}));
  s.appendChild(el("text", {x:W/2, y:560-8, "text-anchor":"middle", class:"axlab"})).textContent = "minutes at facility";
  top.forEach((d,i) => {
    const y = T + i*rowH + rowH/2;
    s.appendChild(el("text", {x:L-14, y:y+4, "text-anchor":"end", class:"rowlab"})).textContent =
      (d.n.replace(/^\\*/,"").length > 32 ? d.n.replace(/^\\*/,"").slice(0,32)+"\\u2026" : d.n.replace(/^\\*/,""));
    s.appendChild(el("line", {x1:X(d.m), x2:X(d.p), y1:y, y2:y, stroke:"var(--rule)", "stroke-width":2.5, "stroke-linecap":"round"}));
    [[d.p,"var(--series-2)"],[d.m,"var(--series-1)"]].forEach(([v,col]) => {
      const c = el("circle", {cx:X(v), cy:y, r:5.5, fill:col, stroke:"var(--surface)", "stroke-width":2, tabindex:"0"});
      const html = `<b>${d.n}</b><span>median ${d.m} min &middot; p90 ${d.p} min<br>${d.v} transports &middot; ${d.h} crew hours</span>`;
      c.addEventListener("mousemove", e => showTip(e, html));
      c.addEventListener("mouseleave", hideTip);
      c.addEventListener("focus", () => { const r = c.getBoundingClientRect(); showTip({clientX:r.x, clientY:r.y}, html); });
      c.addEventListener("blur", hideTip);
      s.appendChild(c); });
    s.appendChild(el("text", {x:X(d.p)+11, y:y+4, class:"val"})).textContent = Math.round(d.m)+" / "+Math.round(d.p);
  });
})();

/* ---- movers ---- */
(function(){
  const s = $("#movers"), W = 980, T = 22, B = 34, L = 250, R = 90;
  const mv = DATA.filter(d => d.d !== null).sort((a,b)=>Math.abs(b.d)-Math.abs(a.d)).slice(0,12);
  const rowH = (400 - T - B) / mv.length;
  const mx = Math.max(...mv.map(d=>Math.abs(d.d))) * 1.15;
  const mid = L + (W - L - R)/2;
  const X = v => mid + (v/mx) * ((W-L-R)/2);
  s.appendChild(el("line", {x1:mid, x2:mid, y1:T-4, y2:400-B, stroke:"var(--rule)", "stroke-width":1.5}));
  mv.forEach((d,i) => {
    const y = T + i*rowH, h = Math.max(9, rowH-9);
    const worse = d.d > 0;
    const x0 = worse ? mid : X(d.d), w = Math.abs(X(d.d) - mid);
    const r = el("rect", {x:x0, y:y, width:Math.max(w,2), height:h, rx:3,
      fill: worse ? "var(--critical)" : "var(--good)", "fill-opacity":.85, tabindex:"0"});
    const html = `<b>${d.n}</b><span>${d.pm} \\u2192 ${d.m} min (${worse?"+":""}${d.d})<br>${d.v} transports in August</span>`;
    r.addEventListener("mousemove", e => showTip(e, html));
    r.addEventListener("mouseleave", hideTip);
    r.addEventListener("focus", () => { const b = r.getBoundingClientRect(); showTip({clientX:b.x, clientY:b.y}, html); });
    r.addEventListener("blur", hideTip);
    s.appendChild(r);
    s.appendChild(el("text", {x:L-14, y:y+h/2+4, "text-anchor":"end", class:"rowlab"})).textContent =
      (d.n.replace(/^\\*/,"").length > 32 ? d.n.replace(/^\\*/,"").slice(0,32)+"\\u2026" : d.n.replace(/^\\*/,""));
    const t = el("text", {x: worse ? X(d.d)+8 : X(d.d)-8, y:y+h/2+4,
      "text-anchor": worse ? "start" : "end", class:"val"});
    t.textContent = (worse ? "+" : "") + d.d + " min " + (worse ? "slower" : "faster");
    s.appendChild(t); });
  s.appendChild(el("text", {x:mid, y:400-10, "text-anchor":"middle", class:"axlab"})).textContent =
    "\\u2190 faster than July      slower than July \\u2192";
})();

/* ---- table ---- */
(function(){
  const fmt = d => `<tr><td>${d.n}</td><td>${d.v}</td><td>${d.m}</td><td>${d.p}</td><td>${d.h}</td>
    <td>${d.d===null?"&mdash;":(d.d>0?"+":"")+d.d}</td><td>${(d.m-META.fleet>0?"+":"")+ (Math.round((d.m-META.fleet)*10)/10)}</td></tr>`;
  const sorted = [...DATA].sort((a,b)=>b.m-a.m);
  $("#tbl tbody").innerHTML = sorted.slice(0,20).map(fmt).join("");
  $("#tbl2 tbody").innerHTML = sorted.slice(20).map(fmt).join("");
  $("#more").textContent = "Show the remaining " + (sorted.length-20) + " facilities";
  $("#foot").innerHTML = "Source: Samsara geofence entry/exit, archived locally. Visit = a stay of 2 minutes or "
    + "more inside a facility geofence; shorter passes are excluded as drive-throughs. Posting time at "
    + "crew bases is excluded via detected post periods. Facilities with fewer than 25 transports in "
    + "the month are not shown. Fleet median " + META.fleet + " min across " + META.facilities + " facilities.";
})();
</script>
"""

if OFFLINE:
    import re as _re
    HTML = _re.sub(r'<link rel="preconnect"[^>]*>\s*', "", HTML)
    HTML = _re.sub(r'<link rel="stylesheet" href="https://fonts\.googleapis[^>]*>\s*', "", HTML)
    HTML = HTML.replace('"Public Sans",system-ui,sans-serif',
                        'system-ui,"Segoe UI",Roboto,Helvetica,Arial,sans-serif')
    HTML = HTML.replace('"Roboto Mono",monospace', 'Consolas,"Cascadia Mono",monospace')

HTML = HTML.replace("__DATA__", json.dumps(data, separators=(",", ":")))
HTML = HTML.replace("__META__", json.dumps(meta, separators=(",", ":")))
open(OUT, "w", encoding="utf-8").write(HTML)
print("wrote", OUT, "%.1f KB" % (os.path.getsize(OUT)/1024))
