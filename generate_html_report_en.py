# -*- coding: utf-8 -*-
# generate_html_report_en.py
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_charts_en as gc

print("Generating charts...")
charts = {}
for name, fn in gc.CHARTS.items():
    charts[name] = gc.to_b64(fn())
    print(f"  Chart '{name}' done.")

print("Building HTML...")

html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>WASH Survey Report &ndash; C&ocirc;te d&rsquo;Ivoire 2026</title>
<style>
/* ===== CSS VARIABLES ===== */
:root {
  --blue: #0054A6;
  --teal: #00B0B9;
  --green: #70AD47;
  --orange: #ED7D31;
  --red: #C00000;
  --yellow: #FFC000;
  --darkblue: #1F3864;
  --gray: #595959;
  --lightgray: #F2F2F2;
  --white: #FFFFFF;
  --shadow: 0 2px 8px rgba(0,0,0,0.12);
  --radius: 8px;
}

/* ===== RESET & BASE ===== */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
  font-size: 14px;
  color: #222;
  background: #f5f6fa;
  line-height: 1.6;
}

h1, h2, h3, h4 {
  font-family: Georgia, serif;
  color: var(--darkblue);
}

a { color: var(--blue); text-decoration: none; }
a:hover { text-decoration: underline; }

/* ===== LAYOUT ===== */
.page-wrapper {
  display: flex;
  min-height: 100vh;
}

/* ===== SIDEBAR ===== */
.sidebar {
  width: 260px;
  min-width: 260px;
  background: var(--darkblue);
  color: #cdd5e0;
  position: sticky;
  top: 0;
  height: 100vh;
  overflow-y: auto;
  padding: 24px 0;
  z-index: 100;
}

.sidebar-brand {
  padding: 0 20px 20px;
  border-bottom: 1px solid rgba(255,255,255,0.1);
  margin-bottom: 16px;
}

.sidebar-brand h2 {
  font-family: 'Segoe UI', sans-serif;
  font-size: 13px;
  font-weight: 700;
  color: var(--teal);
  text-transform: uppercase;
  letter-spacing: 1px;
}

.sidebar-brand p {
  font-size: 11px;
  color: #8fa0b8;
  margin-top: 4px;
}

.sidebar nav ul {
  list-style: none;
}

.sidebar nav ul li a {
  display: block;
  padding: 8px 20px;
  color: #a8b8cc;
  font-size: 12.5px;
  transition: all 0.2s;
  border-left: 3px solid transparent;
}

.sidebar nav ul li a:hover,
.sidebar nav ul li a.active {
  background: rgba(255,255,255,0.07);
  color: #fff;
  border-left-color: var(--teal);
  text-decoration: none;
}

.sidebar nav ul li.section-header {
  padding: 12px 20px 4px;
  font-size: 10px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 1.5px;
  color: #5d7a9a;
  cursor: default;
}

/* ===== MAIN CONTENT ===== */
.main-content {
  flex: 1;
  overflow-x: hidden;
  padding: 0;
}

/* ===== COVER PAGE ===== */
.cover-page {
  background: linear-gradient(135deg, var(--darkblue) 0%, #0054A6 50%, #00B0B9 100%);
  color: white;
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  text-align: center;
  padding: 60px 40px;
  position: relative;
  overflow: hidden;
}

.cover-page::before {
  content: '';
  position: absolute;
  top: -50%;
  left: -50%;
  width: 200%;
  height: 200%;
  background: radial-gradient(circle at 30% 70%, rgba(0,176,185,0.15) 0%, transparent 50%),
              radial-gradient(circle at 70% 30%, rgba(0,84,166,0.2) 0%, transparent 50%);
  pointer-events: none;
}

.cover-logo {
  width: 80px;
  height: 80px;
  background: rgba(255,255,255,0.15);
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  margin: 0 auto 32px;
  font-size: 36px;
  border: 3px solid rgba(255,255,255,0.3);
}

.cover-page .org-badge {
  background: rgba(255,255,255,0.15);
  border: 1px solid rgba(255,255,255,0.3);
  border-radius: 20px;
  padding: 6px 20px;
  font-size: 12px;
  letter-spacing: 2px;
  text-transform: uppercase;
  margin-bottom: 32px;
  display: inline-block;
}

.cover-page h1 {
  font-family: Georgia, serif;
  font-size: 48px;
  font-weight: 700;
  color: white;
  line-height: 1.2;
  margin-bottom: 12px;
  text-shadow: 0 2px 4px rgba(0,0,0,0.2);
}

.cover-page .subtitle {
  font-size: 22px;
  color: rgba(255,255,255,0.85);
  margin-bottom: 40px;
  font-style: italic;
}

.cover-divider {
  width: 80px;
  height: 3px;
  background: var(--teal);
  margin: 0 auto 40px;
  border-radius: 2px;
}

.cover-meta {
  display: flex;
  gap: 40px;
  justify-content: center;
  flex-wrap: wrap;
  margin-bottom: 40px;
}

.cover-meta-item {
  text-align: center;
}

.cover-meta-item .label {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 1px;
  color: rgba(255,255,255,0.6);
  display: block;
}

.cover-meta-item .value {
  font-size: 16px;
  font-weight: 600;
  color: white;
  display: block;
  margin-top: 4px;
}

.cover-stats {
  display: flex;
  gap: 24px;
  justify-content: center;
  flex-wrap: wrap;
  background: rgba(255,255,255,0.1);
  border-radius: 12px;
  padding: 24px 32px;
  border: 1px solid rgba(255,255,255,0.2);
}

.cover-stat {
  text-align: center;
}

.cover-stat .num {
  font-size: 32px;
  font-weight: 700;
  color: var(--teal);
  font-family: Georgia, serif;
  display: block;
}

.cover-stat .desc {
  font-size: 12px;
  color: rgba(255,255,255,0.75);
  display: block;
  margin-top: 4px;
}

/* ===== SECTIONS ===== */
.section {
  background: white;
  margin: 24px 24px;
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  overflow: hidden;
}

.section-header {
  background: linear-gradient(90deg, var(--blue) 0%, var(--teal) 100%);
  padding: 20px 28px;
  color: white;
}

.section-header h2 {
  font-family: Georgia, serif;
  font-size: 22px;
  color: white;
  font-weight: 700;
}

.section-header .section-num {
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 2px;
  opacity: 0.8;
  margin-bottom: 6px;
}

.section-body {
  padding: 28px;
}

/* ===== KPI CARDS ===== */
.kpi-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 16px;
  margin-bottom: 24px;
}

.kpi-card {
  background: white;
  border-radius: var(--radius);
  padding: 20px;
  box-shadow: var(--shadow);
  border-top: 4px solid var(--blue);
  text-align: center;
  transition: transform 0.2s;
}

.kpi-card:hover { transform: translateY(-2px); }
.kpi-card.teal  { border-top-color: var(--teal); }
.kpi-card.green { border-top-color: var(--green); }
.kpi-card.orange{ border-top-color: var(--orange); }
.kpi-card.red   { border-top-color: var(--red); }
.kpi-card.yellow{ border-top-color: var(--yellow); }

.kpi-card .kpi-num {
  font-size: 36px;
  font-weight: 700;
  color: var(--blue);
  font-family: Georgia, serif;
  line-height: 1;
}

.kpi-card.teal  .kpi-num { color: var(--teal); }
.kpi-card.green .kpi-num { color: var(--green); }
.kpi-card.orange .kpi-num { color: var(--orange); }
.kpi-card.red   .kpi-num { color: var(--red); }
.kpi-card.yellow .kpi-num { color: #c89000; }

.kpi-card .kpi-label {
  font-size: 12px;
  color: var(--gray);
  margin-top: 8px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

.kpi-card .kpi-sub {
  font-size: 11px;
  color: #999;
  margin-top: 4px;
}

/* ===== ALERT BOXES ===== */
.alert-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 16px;
  margin-bottom: 20px;
}

.alert-box {
  border-radius: var(--radius);
  padding: 16px 20px;
  display: flex;
  gap: 14px;
  align-items: flex-start;
}

.alert-box.danger  { background: #fff5f5; border-left: 4px solid var(--red); }
.alert-box.warning { background: #fffbf0; border-left: 4px solid var(--yellow); }
.alert-box.success { background: #f0fff4; border-left: 4px solid var(--green); }
.alert-box.info    { background: #f0f8ff; border-left: 4px solid var(--blue); }

.alert-box .alert-icon {
  font-size: 20px;
  flex-shrink: 0;
  margin-top: 2px;
}

.alert-box .alert-title {
  font-weight: 700;
  font-size: 13px;
  margin-bottom: 4px;
  color: var(--darkblue);
}

.alert-box .alert-text {
  font-size: 12.5px;
  color: #555;
  line-height: 1.5;
}

/* ===== CHART CONTAINERS ===== */
.chart-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
  gap: 24px;
  margin: 20px 0;
}

.chart-card {
  background: #fafbfc;
  border-radius: var(--radius);
  padding: 16px;
  border: 1px solid #e8ecf0;
  text-align: center;
}

.chart-card img {
  max-width: 100%;
  height: auto;
  display: block;
  margin: 0 auto;
}

.chart-card .chart-caption {
  font-size: 11.5px;
  color: var(--gray);
  font-style: italic;
  margin-top: 8px;
}

.chart-full {
  background: #fafbfc;
  border-radius: var(--radius);
  padding: 16px;
  border: 1px solid #e8ecf0;
  text-align: center;
  margin: 20px 0;
}

.chart-full img {
  max-width: 100%;
  height: auto;
  display: block;
  margin: 0 auto;
}

/* ===== TABLES ===== */
.table-wrap { overflow-x: auto; margin: 20px 0; }

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
}

table caption {
  font-weight: 700;
  color: var(--darkblue);
  text-align: left;
  padding: 8px 0;
  font-size: 13.5px;
}

thead tr {
  background: var(--blue);
  color: white;
}

thead th {
  padding: 10px 14px;
  text-align: left;
  font-weight: 600;
  font-size: 12.5px;
  letter-spacing: 0.3px;
}

tbody tr:nth-child(even) { background: #f5f8ff; }
tbody tr:hover { background: #e8f0ff; }

tbody td {
  padding: 9px 14px;
  border-bottom: 1px solid #e8ecf0;
  color: #333;
}

.badge {
  display: inline-block;
  padding: 2px 10px;
  border-radius: 12px;
  font-size: 11px;
  font-weight: 600;
}

.badge-red    { background: #ffe0e0; color: var(--red); }
.badge-green  { background: #e0f5e8; color: #2d7a3f; }
.badge-yellow { background: #fff8d6; color: #9a7000; }
.badge-blue   { background: #ddeeff; color: var(--blue); }
.badge-teal   { background: #d6f5f7; color: #007880; }
.badge-orange { background: #fff0e5; color: #b34700; }

/* ===== PROGRESS BARS ===== */
.progress-item {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 10px;
}

.progress-label {
  min-width: 160px;
  font-size: 13px;
  color: #444;
}

.progress-bar-wrap {
  flex: 1;
  background: #e8ecf0;
  border-radius: 4px;
  height: 16px;
  overflow: hidden;
}

.progress-bar {
  height: 100%;
  border-radius: 4px;
  transition: width 0.5s;
}

.progress-val {
  min-width: 60px;
  font-size: 12px;
  color: var(--gray);
  text-align: right;
}

/* ===== SECTION-SPECIFIC ===== */
.subsection {
  margin-top: 28px;
  padding-top: 20px;
  border-top: 2px solid var(--lightgray);
}

.subsection h3 {
  font-family: Georgia, serif;
  font-size: 16px;
  color: var(--blue);
  margin-bottom: 14px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.subsection h3::before {
  content: '';
  display: inline-block;
  width: 4px;
  height: 18px;
  background: var(--teal);
  border-radius: 2px;
}

p { margin-bottom: 12px; color: #444; line-height: 1.7; }

/* ===== RECOMMENDATION CARDS ===== */
.reco-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
  gap: 20px;
  margin: 20px 0;
}

.reco-card {
  background: white;
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  overflow: hidden;
}

.reco-card-header {
  padding: 14px 18px;
  color: white;
  display: flex;
  align-items: center;
  gap: 10px;
}

.reco-card-header.priority-1 { background: var(--red); }
.reco-card-header.priority-2 { background: var(--orange); }
.reco-card-header.priority-3 { background: var(--blue); }
.reco-card-header.priority-4 { background: var(--teal); }
.reco-card-header.priority-5 { background: var(--green); }
.reco-card-header.priority-6 { background: var(--gray); }

.reco-card-header .priority-num {
  background: rgba(255,255,255,0.25);
  width: 28px;
  height: 28px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 700;
  font-size: 13px;
  flex-shrink: 0;
}

.reco-card-header .reco-title {
  font-weight: 700;
  font-size: 14px;
}

.reco-card-body {
  padding: 16px 18px;
}

.reco-card-body ul {
  list-style: none;
  padding: 0;
}

.reco-card-body ul li {
  padding: 5px 0;
  font-size: 12.5px;
  color: #444;
  border-bottom: 1px solid #f0f0f0;
  display: flex;
  align-items: flex-start;
  gap: 8px;
}

.reco-card-body ul li:last-child { border-bottom: none; }
.reco-card-body ul li::before { content: '&#10003;'; color: var(--green); font-weight: bold; flex-shrink: 0; }

/* ===== COMPARISON TABLE ===== */
.comparison-table thead tr { background: var(--darkblue); }
.comparison-highlight { background: #fff9e8 !important; }

/* ===== FOOTER ===== */
.report-footer {
  background: var(--darkblue);
  color: rgba(255,255,255,0.7);
  text-align: center;
  padding: 20px;
  font-size: 12px;
  margin: 24px 24px 0;
  border-radius: var(--radius) var(--radius) 0 0;
}

.report-footer strong { color: var(--teal); }

/* ===== PRINT STYLES ===== */
@media print {
  body { background: white; font-size: 12px; }
  .sidebar { display: none !important; }
  .page-wrapper { display: block; }
  .main-content { padding: 0; }
  .section { margin: 0; box-shadow: none; border-radius: 0; border: none; page-break-inside: avoid; }
  .section + .section { page-break-before: always; }
  .cover-page { page-break-after: always; min-height: auto; padding: 40px 30px; }
  .kpi-grid { grid-template-columns: repeat(4, 1fr); }
  .chart-grid { grid-template-columns: repeat(2, 1fr); }
  .alert-grid { grid-template-columns: repeat(2, 1fr); }
  .reco-grid { grid-template-columns: repeat(2, 1fr); }
  .report-footer { border-radius: 0; margin: 0; }
  a { color: inherit; }
  thead { display: table-header-group; }
  tr { page-break-inside: avoid; }
}

@media (max-width: 768px) {
  .sidebar { display: none; }
  .cover-page h1 { font-size: 32px; }
  .kpi-grid { grid-template-columns: repeat(2, 1fr); }
  .chart-grid { grid-template-columns: 1fr; }
}
</style>
</head>
<body>
<div class="page-wrapper">

<!-- SIDEBAR -->
<aside class="sidebar">
  <div class="sidebar-brand">
    <h2>WASH Report</h2>
    <p>C&ocirc;te d&rsquo;Ivoire &bull; 2026</p>
  </div>
  <nav>
    <ul>
      <li class="section-header">Navigation</li>
      <li><a href="#cover">Cover Page</a></li>
      <li><a href="#resume">Executive Summary</a></li>
      <li><a href="#contexte">Context &amp; Methodology</a></li>
      <li class="section-header">WASH Sections</li>
      <li><a href="#eau">Water &amp; Sanitation</a></li>
      <li><a href="#assainissement">Sanitation</a></li>
      <li><a href="#dechets">Waste Management</a></li>
      <li><a href="#pollution">Pollution &amp; Air</a></li>
      <li><a href="#ressources">Natural Resources</a></li>
      <li class="section-header">Analysis</li>
      <li><a href="#comparaison">Comparative Analysis</a></li>
      <li><a href="#recommandations">Recommendations</a></li>
      <li><a href="#conclusion">Conclusion</a></li>
    </ul>
  </nav>
</aside>

<!-- MAIN CONTENT -->
<main class="main-content">

<!-- ===== COVER PAGE ===== -->
<section id="cover" class="cover-page">
  <div class="org-badge">Technical Report &bull; 2026</div>
  <div class="cover-logo">&#9749;</div>
  <h1>WASH Survey Report</h1>
  <p class="subtitle">Water, Sanitation &amp; Hygiene &mdash; C&ocirc;te d&rsquo;Ivoire</p>
  <div class="cover-divider"></div>
  <div class="cover-meta">
    <div class="cover-meta-item">
      <span class="label">Study Area</span>
      <span class="value">Goh &amp; Loh Djiboua Regions</span>
    </div>
    <div class="cover-meta-item">
      <span class="label">Period</span>
      <span class="value">January &ndash; February 2026</span>
    </div>
    <div class="cover-meta-item">
      <span class="label">Sub-prefectures</span>
      <span class="value">8 localities</span>
    </div>
  </div>
  <div class="cover-stats">
    <div class="cover-stat">
      <span class="num">1 000</span>
      <span class="desc">Households surveyed</span>
    </div>
    <div class="cover-stat">
      <span class="num">166</span>
      <span class="desc">Variables collected</span>
    </div>
    <div class="cover-stat">
      <span class="num">2</span>
      <span class="desc">Departments</span>
    </div>
    <div class="cover-stat">
      <span class="num">40%</span>
      <span class="desc">Poor water quality</span>
    </div>
  </div>
</section>

<!-- ===== EXECUTIVE SUMMARY ===== -->
<section id="resume" class="section">
  <div class="section-header">
    <div class="section-num">Section 1</div>
    <h2>Executive Summary</h2>
  </div>
  <div class="section-body">
    <p>This WASH survey was conducted among <strong>1,000 households</strong> distributed across the departments of <strong>Goh (570)</strong> and <strong>Loh Djiboua (430)</strong> in C&ocirc;te d&rsquo;Ivoire, covering 8 sub-prefectures between January and February 2026. The data collected covers 166 indicators on water, sanitation, waste management, atmospheric pollution and natural resources availability.</p>

    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-num">1 000</div>
        <div class="kpi-label">Households surveyed</div>
        <div class="kpi-sub">Jan (358) + Feb (642)</div>
      </div>
      <div class="kpi-card red">
        <div class="kpi-num">40%</div>
        <div class="kpi-label">Poor water quality</div>
        <div class="kpi-sub">400 households affected</div>
      </div>
      <div class="kpi-card orange">
        <div class="kpi-num">70.9%</div>
        <div class="kpi-label">Untreated water</div>
        <div class="kpi-sub">709 households</div>
      </div>
      <div class="kpi-card yellow">
        <div class="kpi-num">33.6%</div>
        <div class="kpi-label">No latrine</div>
        <div class="kpi-sub">336 households</div>
      </div>
      <div class="kpi-card teal">
        <div class="kpi-num">49.4%</div>
        <div class="kpi-label">Atmospheric pollution</div>
        <div class="kpi-sub">494 reports</div>
      </div>
      <div class="kpi-card green">
        <div class="kpi-num">58%</div>
        <div class="kpi-label">Sufficient water</div>
        <div class="kpi-sub">580 satisfied households</div>
      </div>
    </div>

    <div class="alert-grid">
      <div class="alert-box danger">
        <div class="alert-icon">&#9888;</div>
        <div>
          <div class="alert-title">Critical Alert &ndash; Water Quality</div>
          <div class="alert-text">40% of households have poor quality water and 70.9% do not treat their water before consumption. High risk of waterborne diseases.</div>
        </div>
      </div>
      <div class="alert-box danger">
        <div class="alert-icon">&#9888;</div>
        <div>
          <div class="alert-title">Critical Alert &ndash; Sanitation</div>
          <div class="alert-text">33.6% of households have no latrines. Coverage of improved latrines remains very low (9%).</div>
        </div>
      </div>
      <div class="alert-box warning">
        <div class="alert-icon">&#9679;</div>
        <div>
          <div class="alert-title">Concern &ndash; Waterborne Diseases</div>
          <div class="alert-text">27.1% of households report water-related diseases. Diarrhea dominates (~180 cases) followed by typhoid (~120 cases).</div>
        </div>
      </div>
      <div class="alert-box warning">
        <div class="alert-icon">&#9679;</div>
        <div>
          <div class="alert-title">Concern &ndash; Natural Resources</div>
          <div class="alert-text">57.7% of households report resource degradation. 38.4% consider them scarce, and 16% consider them near depletion.</div>
        </div>
      </div>
      <div class="alert-box success">
        <div class="alert-icon">&#10003;</div>
        <div>
          <div class="alert-title">Positive Point &ndash; Water Availability</div>
          <div class="alert-text">58% of households have a sufficient quantity of water for their daily needs.</div>
        </div>
      </div>
      <div class="alert-box info">
        <div class="alert-icon">&#9432;</div>
        <div>
          <div class="alert-title">Information &ndash; Waste Management</div>
          <div class="alert-text">Collection/recycling is the most common practice (515 households), but open burning remains very widespread (499) with associated health risks.</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===== CONTEXT & METHODOLOGY ===== -->
<section id="contexte" class="section">
  <div class="section-header">
    <div class="section-num">Section 2</div>
    <h2>Context &amp; Methodology</h2>
  </div>
  <div class="section-body">
    <p>The 2026 WASH survey aims to assess the conditions of access to drinking water, the state of sanitation and hygiene practices in the Goh and Loh Djiboua regions. These two departments face significant challenges in terms of public health and environmental management.</p>

    <div class="chart-grid">
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['departements'] + """" alt="Sample distribution by department">
        <div class="chart-caption">Fig. 1 &ndash; Sample distribution by department</div>
      </div>
      <div>
        <div class="subsection" style="border-top:none; margin-top:0; padding-top:0;">
          <h3>Methodological Framework</h3>
          <div class="progress-item">
            <div class="progress-label">Sample size</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:100%;background:var(--blue)"></div></div>
            <div class="progress-val">N = 1 000</div>
          </div>
          <div class="progress-item">
            <div class="progress-label">Goh Department</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:57%;background:var(--blue)"></div></div>
            <div class="progress-val">570 (57%)</div>
          </div>
          <div class="progress-item">
            <div class="progress-label">Loh Djiboua Department</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:43%;background:var(--teal)"></div></div>
            <div class="progress-val">430 (43%)</div>
          </div>
          <div class="progress-item">
            <div class="progress-label">January surveys</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:35.8%;background:var(--orange)"></div></div>
            <div class="progress-val">358 (35.8%)</div>
          </div>
          <div class="progress-item">
            <div class="progress-label">February surveys</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:64.2%;background:var(--green)"></div></div>
            <div class="progress-val">642 (64.2%)</div>
          </div>
        </div>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Table 1 &ndash; Sample distribution by sub-prefecture</caption>
        <thead>
          <tr>
            <th>Sub-prefecture</th>
            <th>Department</th>
            <th>Surveys</th>
            <th>% of sample</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          <tr><td>Gagnoa</td><td>Goh</td><td>~115</td><td>~11.5%</td><td><span class="badge badge-blue">Capital</span></td></tr>
          <tr><td>Bayota</td><td>Goh</td><td>~95</td><td>~9.5%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr><td>Ouragahio</td><td>Goh</td><td>~90</td><td>~9.0%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr><td>Guitry</td><td>Goh</td><td>~88</td><td>~8.8%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr><td>Guiberoua</td><td>Goh</td><td>~85</td><td>~8.5%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr><td>Divo</td><td>Loh Djiboua</td><td>~155</td><td>~15.5%</td><td><span class="badge badge-blue">Capital</span></td></tr>
          <tr><td>Lakota</td><td>Loh Djiboua</td><td>~140</td><td>~14.0%</td><td><span class="badge badge-blue">Capital</span></td></tr>
          <tr><td>Hire</td><td>Loh Djiboua</td><td>~135</td><td>~13.5%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr style="font-weight:700;background:#f0f4ff"><td>TOTAL</td><td>2 departments</td><td>1 000</td><td>100%</td><td><span class="badge badge-green">Complete</span></td></tr>
        </tbody>
      </table>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Table 2 &ndash; Data collection tools and variables</caption>
        <thead>
          <tr><th>Theme</th><th>Key variables</th><th>Indicator type</th></tr>
        </thead>
        <tbody>
          <tr><td>Drinking water</td><td>Source, quality, access time, treatment, availability</td><td><span class="badge badge-blue">Quantitative</span></td></tr>
          <tr><td>Water health</td><td>Water-related diseases, frequency, type</td><td><span class="badge badge-orange">Mixed</span></td></tr>
          <tr><td>Sanitation</td><td>Latrine type, hygiene practices</td><td><span class="badge badge-blue">Quantitative</span></td></tr>
          <tr><td>Waste</td><td>Disposal method, satisfaction, exposure</td><td><span class="badge badge-blue">Quantitative</span></td></tr>
          <tr><td>Environment</td><td>Pollution, air quality, noise nuisances</td><td><span class="badge badge-teal">Qualitative</span></td></tr>
          <tr><td>Resources</td><td>Availability, degradation, stewardship</td><td><span class="badge badge-teal">Qualitative</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== WATER SECTION ===== -->
<section id="eau" class="section">
  <div class="section-header">
    <div class="section-num">Section 3</div>
    <h2>Water &amp; Sanitation</h2>
  </div>
  <div class="section-body">
    <p>Access to drinking water is one of the major challenges identified in this survey. While 58% of households have a sufficient quantity of water, quality remains concerning with <strong>40% poor quality water</strong> and only <strong>29.1% of households treating their water</strong>.</p>

    <div class="subsection">
      <h3>3.1 Water supply sources</h3>
      <div class="chart-grid">
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['sources'] + """" alt="Water supply sources">
          <div class="chart-caption">Fig. 2 &ndash; Water supply sources (N=1000)</div>
        </div>
        <div>
          <div class="table-wrap">
            <table>
              <caption>Table 3 &ndash; Water source details</caption>
              <thead><tr><th>Source</th><th>No. households</th><th>%</th><th>Assessment</th></tr></thead>
              <tbody>
                <tr><td>Traditional well</td><td>275</td><td>27.5%</td><td><span class="badge badge-red">High risk</span></td></tr>
                <tr><td>Improved well</td><td>200</td><td>20.0%</td><td><span class="badge badge-yellow">Moderate risk</span></td></tr>
                <tr><td>Borehole</td><td>196</td><td>19.6%</td><td><span class="badge badge-green">Safe</span></td></tr>
                <tr><td>River / Spring</td><td>167</td><td>16.7%</td><td><span class="badge badge-red">Very high risk</span></td></tr>
                <tr><td>Other</td><td>82</td><td>8.2%</td><td><span class="badge badge-orange">Undetermined</span></td></tr>
                <tr><td>SODECI (piped)</td><td>80</td><td>8.0%</td><td><span class="badge badge-green">Safe</span></td></tr>
                <tr style="font-weight:700"><td>Total</td><td>1 000</td><td>100%</td><td></td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <div class="subsection">
      <h3>3.2 Water quality by department</h3>
      <div class="chart-grid">
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['qualite_dept'] + """" alt="Water quality by department">
          <div class="chart-caption">Fig. 3 &ndash; Perceived water quality by department</div>
        </div>
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['traitement'] + """" alt="Water treatment">
          <div class="chart-caption">Fig. 4 &ndash; Household water treatment</div>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <caption>Table 4 &ndash; Water quality by department</caption>
          <thead><tr><th>Department</th><th>Good</th><th>Acceptable</th><th>Poor</th><th>Total</th><th>% poor</th></tr></thead>
          <tbody>
            <tr><td><strong>Goh</strong></td><td>118 (20.7%)</td><td>221 (38.8%)</td><td>231 (40.5%)</td><td>570</td><td><span class="badge badge-red">40.5%</span></td></tr>
            <tr><td><strong>Loh Djiboua</strong></td><td>90 (20.9%)</td><td>171 (39.8%)</td><td>169 (39.3%)</td><td>430</td><td><span class="badge badge-red">39.3%</span></td></tr>
            <tr style="font-weight:700"><td>Total</td><td>208 (20.8%)</td><td>392 (39.2%)</td><td>400 (40.0%)</td><td>1 000</td><td><span class="badge badge-red">40.0%</span></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="subsection">
      <h3>3.3 Water access and waterborne diseases</h3>
      <div class="chart-grid">
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['acces'] + """" alt="Water access time">
          <div class="chart-caption">Fig. 5 &ndash; Time to reach water source</div>
        </div>
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['maladies'] + """" alt="Waterborne diseases">
          <div class="chart-caption">Fig. 6 &ndash; Reported waterborne diseases by department</div>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <caption>Table 5 &ndash; Water access time and implications</caption>
          <thead><tr><th>Time bracket</th><th>Households</th><th>%</th><th>WHO standard</th><th>Status</th></tr></thead>
          <tbody>
            <tr><td>Under 30 minutes</td><td>386</td><td>38.6%</td><td>Compliant</td><td><span class="badge badge-green">Acceptable</span></td></tr>
            <tr><td>30 to 60 minutes</td><td>341</td><td>34.1%</td><td>Borderline</td><td><span class="badge badge-yellow">Caution</span></td></tr>
            <tr><td>Over 60 minutes</td><td>273</td><td>27.3%</td><td>Non-compliant</td><td><span class="badge badge-red">Critical</span></td></tr>
          </tbody>
        </table>
      </div>
      <div class="table-wrap">
        <table>
          <caption>Table 6 &ndash; Reported waterborne diseases</caption>
          <thead><tr><th>Disease</th><th>Goh</th><th>Loh Djiboua</th><th>Estimated total</th><th>Severity</th></tr></thead>
          <tbody>
            <tr><td>Diarrhea</td><td>~115</td><td>~87</td><td>~202</td><td><span class="badge badge-orange">Moderate</span></td></tr>
            <tr><td>Typhoid</td><td>~72</td><td>~55</td><td>~127</td><td><span class="badge badge-red">High</span></td></tr>
            <tr><td>Hepatitis A</td><td>~60</td><td>~43</td><td>~103</td><td><span class="badge badge-red">High</span></td></tr>
            <tr><td>Cholera</td><td>~50</td><td>~34</td><td>~84</td><td><span class="badge badge-red">Very high</span></td></tr>
            <tr style="font-weight:700"><td>Total reported</td><td>~297</td><td>~219</td><td>~516 cases</td><td></td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</section>

<!-- ===== SANITATION SECTION ===== -->
<section id="assainissement" class="section">
  <div class="section-header">
    <div class="section-num">Section 4</div>
    <h2>Sanitation</h2>
  </div>
  <div class="section-body">
    <p>Sanitation represents a major challenge in both departments. Nearly one third of households (33.6%) have no access to latrines, and only <strong>14.8% have improved latrines or modern toilets</strong>, far from SDG 6 targets.</p>

    <div class="chart-full">
      <img src="data:image/png;base64,""" + charts['latrines'] + """" alt="Sanitation infrastructure types">
      <div class="chart-caption">Fig. 7 &ndash; Sanitation infrastructure types (N=1000)</div>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Table 7 &ndash; Latrine types / sanitation infrastructure</caption>
        <thead><tr><th>Infrastructure type</th><th>No. households</th><th>%</th><th>SDG 6 compliance</th></tr></thead>
        <tbody>
          <tr><td>Simple latrine</td><td>424</td><td>42.4%</td><td><span class="badge badge-yellow">Partial</span></td></tr>
          <tr><td>No latrine (open defecation)</td><td>336</td><td>33.6%</td><td><span class="badge badge-red">Non-compliant</span></td></tr>
          <tr><td>Other type</td><td>92</td><td>9.2%</td><td><span class="badge badge-orange">Undetermined</span></td></tr>
          <tr><td>Improved latrine</td><td>90</td><td>9.0%</td><td><span class="badge badge-green">Compliant</span></td></tr>
          <tr><td>Modern toilet</td><td>58</td><td>5.8%</td><td><span class="badge badge-green">Compliant</span></td></tr>
          <tr style="font-weight:700"><td>Total</td><td>1 000</td><td>100%</td><td></td></tr>
        </tbody>
      </table>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Table 8 &ndash; Key sanitation indicators</caption>
        <thead><tr><th>Indicator</th><th>Value</th><th>SDG 6 target</th><th>Gap</th></tr></thead>
        <tbody>
          <tr><td>Access to latrines (all types)</td><td>66.4%</td><td>100%</td><td><span class="badge badge-red">-33.6 pts</span></td></tr>
          <tr><td>Improved latrines + modern toilets</td><td>14.8%</td><td>100%</td><td><span class="badge badge-red">-85.2 pts</span></td></tr>
          <tr><td>Open defecation (OD)</td><td>33.6%</td><td>0%</td><td><span class="badge badge-red">+33.6 pts</span></td></tr>
          <tr><td>Household water treatment</td><td>29.1%</td><td>100%</td><td><span class="badge badge-red">-70.9 pts</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== WASTE MANAGEMENT SECTION ===== -->
<section id="dechets" class="section">
  <div class="section-header">
    <div class="section-num">Section 5</div>
    <h2>Waste Management</h2>
  </div>
  <div class="section-body">
    <p>Waste management is a complex issue in both departments. While 52.3% of households are exposed to open dumps, collection and recycling practices show potential for improvement. Satisfaction remains mostly negative with <strong>62% of households somewhat or completely unsatisfied</strong>.</p>

    <div class="chart-grid">
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['gestion_modes'] + """" alt="Waste disposal methods">
        <div class="chart-caption">Fig. 8 &ndash; Waste disposal methods</div>
      </div>
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['dechets_sat'] + """" alt="Waste management satisfaction">
        <div class="chart-caption">Fig. 9 &ndash; Waste management satisfaction</div>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Table 9 &ndash; Waste disposal methods</caption>
        <thead><tr><th>Disposal method</th><th>Households</th><th>%</th><th>Environmental impact</th></tr></thead>
        <tbody>
          <tr><td>Collection / Recycling</td><td>515</td><td>51.5%</td><td><span class="badge badge-green">Positive</span></td></tr>
          <tr><td>Open dumping</td><td>501</td><td>50.1%</td><td><span class="badge badge-red">Very negative</span></td></tr>
          <tr><td>Burning</td><td>499</td><td>49.9%</td><td><span class="badge badge-red">Negative</span></td></tr>
          <tr><td>Reuse</td><td>287</td><td>28.7%</td><td><span class="badge badge-green">Positive</span></td></tr>
        </tbody>
      </table>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Table 10 &ndash; Waste management satisfaction</caption>
        <thead><tr><th>Satisfaction level</th><th>Households</th><th>%</th><th>Trend</th></tr></thead>
        <tbody>
          <tr><td>Somewhat satisfied</td><td>378</td><td>37.8%</td><td><span class="badge badge-orange">Negative</span></td></tr>
          <tr><td>Satisfied</td><td>271</td><td>27.1%</td><td><span class="badge badge-green">Positive</span></td></tr>
          <tr><td>Not satisfied</td><td>242</td><td>24.2%</td><td><span class="badge badge-red">Negative</span></td></tr>
          <tr><td>Very satisfied</td><td>109</td><td>10.9%</td><td><span class="badge badge-green">Positive</span></td></tr>
          <tr style="font-weight:700"><td>Total</td><td>1 000</td><td>100%</td><td></td></tr>
        </tbody>
      </table>
    </div>

    <div class="alert-box warning" style="margin-top:16px">
      <div class="alert-icon">&#9432;</div>
      <div>
        <div class="alert-title">Point of attention &ndash; Waste exposure</div>
        <div class="alert-text">52.3% of households (523) report exposure to open dumps or uncontrolled waste sites. This exposure constitutes a major health and environmental risk requiring urgent interventions.</div>
      </div>
    </div>
  </div>
</section>

<!-- ===== POLLUTION SECTION ===== -->
<section id="pollution" class="section">
  <div class="section-header">
    <div class="section-num">Section 6</div>
    <h2>Pollution &amp; Air Quality</h2>
  </div>
  <div class="section-body">
    <p>Atmospheric pollution is reported by nearly one household in two (49.4%). Air quality is judged poor or very poor by <strong>51.6% of respondents</strong>. Noise nuisances also affect a majority of households (57.3%), reflecting growing urbanization and poorly regulated industrial activities.</p>

    <div class="chart-grid">
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['pollution'] + """" alt="Atmospheric pollution">
        <div class="chart-caption">Fig. 10 &ndash; Reported atmospheric pollution by department</div>
      </div>
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['qualite_air'] + """" alt="Perceived air quality">
        <div class="chart-caption">Fig. 11 &ndash; Perceived air quality (N=1000)</div>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Table 11 &ndash; Pollution and air quality indicators</caption>
        <thead><tr><th>Indicator</th><th>Households</th><th>%</th><th>Risk level</th></tr></thead>
        <tbody>
          <tr><td>Atmospheric pollution reported</td><td>494</td><td>49.4%</td><td><span class="badge badge-red">Concerning</span></td></tr>
          <tr><td>Air quality: Poor</td><td>364</td><td>36.4%</td><td><span class="badge badge-red">Critical</span></td></tr>
          <tr><td>Air quality: Average</td><td>361</td><td>36.1%</td><td><span class="badge badge-yellow">Moderate</span></td></tr>
          <tr><td>Air quality: Very poor</td><td>152</td><td>15.2%</td><td><span class="badge badge-red">Very critical</span></td></tr>
          <tr><td>Air quality: Good</td><td>123</td><td>12.3%</td><td><span class="badge badge-green">Acceptable</span></td></tr>
          <tr><td>Noise nuisances</td><td>573</td><td>57.3%</td><td><span class="badge badge-orange">Concerning</span></td></tr>
        </tbody>
      </table>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Table 12 &ndash; Pollution by department</caption>
        <thead><tr><th>Department</th><th>Pollution reported</th><th>No pollution</th><th>Rate</th></tr></thead>
        <tbody>
          <tr><td><strong>Goh</strong></td><td>285</td><td>285</td><td><span class="badge badge-yellow">50.0%</span></td></tr>
          <tr><td><strong>Loh Djiboua</strong></td><td>209</td><td>221</td><td><span class="badge badge-yellow">48.6%</span></td></tr>
          <tr style="font-weight:700"><td>Total</td><td>494</td><td>506</td><td><span class="badge badge-orange">49.4%</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== NATURAL RESOURCES SECTION ===== -->
<section id="ressources" class="section">
  <div class="section-header">
    <div class="section-num">Section 7</div>
    <h2>Natural Resources</h2>
  </div>
  <div class="section-body">
    <p>The natural resources situation is alarming: <strong>57.7% of households report resource degradation</strong>, and only 12.1% consider them abundant. This situation, combined with 54.1% of households rating their environmental stewardship as strong or very strong, indicates awareness but practices that remain insufficient.</p>

    <div class="chart-full">
      <img src="data:image/png;base64,""" + charts['ressources'] + """" alt="Natural resources availability">
      <div class="chart-caption">Fig. 12 &ndash; Perceived natural resources availability (N=1000)</div>
    </div>

    <div class="chart-grid">
      <div>
        <div class="table-wrap">
          <table>
            <caption>Table 13 &ndash; Natural resources availability</caption>
            <thead><tr><th>Level</th><th>Households</th><th>%</th><th>Status</th></tr></thead>
            <tbody>
              <tr><td>Scarce</td><td>384</td><td>38.4%</td><td><span class="badge badge-red">Critical</span></td></tr>
              <tr><td>Moderately available</td><td>335</td><td>33.5%</td><td><span class="badge badge-yellow">Precarious</span></td></tr>
              <tr><td>Near depletion</td><td>160</td><td>16.0%</td><td><span class="badge badge-red">Emergency</span></td></tr>
              <tr><td>Abundant</td><td>121</td><td>12.1%</td><td><span class="badge badge-green">Favorable</span></td></tr>
            </tbody>
          </table>
        </div>
      </div>
      <div>
        <div class="table-wrap">
          <table>
            <caption>Table 14 &ndash; Environmental stewardship</caption>
            <thead><tr><th>Level</th><th>Households</th><th>%</th></tr></thead>
            <tbody>
              <tr><td>Strong</td><td>358</td><td>35.8%</td></tr>
              <tr><td>Moderate</td><td>308</td><td>30.8%</td></tr>
              <tr><td>Very strong</td><td>183</td><td>18.3%</td></tr>
              <tr><td>Weak</td><td>151</td><td>15.1%</td></tr>
            </tbody>
          </table>
        </div>
        <div class="alert-box danger" style="margin-top:12px">
          <div class="alert-icon">&#9888;</div>
          <div>
            <div class="alert-title">Confirmed degradation</div>
            <div class="alert-text">577 households (57.7%) report visible degradation of resources. Urgent actions required.</div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===== COMPARATIVE ANALYSIS ===== -->
<section id="comparaison" class="section">
  <div class="section-header">
    <div class="section-num">Section 8</div>
    <h2>Comparative Analysis</h2>
  </div>
  <div class="section-body">
    <p>The comparison of WASH indicators between the two departments reveals relatively similar profiles, with small gaps across most dimensions. The Goh department shows slightly more disease cases and a higher proportion of treated water, while Loh Djiboua shows better access within 30 minutes.</p>

    <div class="chart-full">
      <img src="data:image/png;base64,""" + charts['comparaison'] + """" alt="WASH indicators comparison">
      <div class="chart-caption">Fig. 13 &ndash; Comparison of key WASH indicators by department</div>
    </div>

    <div class="table-wrap">
      <table class="comparison-table">
        <caption>Table 15 &ndash; Full comparative table of WASH indicators</caption>
        <thead>
          <tr>
            <th>Indicator</th>
            <th>Overall (N=1000)</th>
            <th>Goh (N=570)</th>
            <th>Loh Djiboua (N=430)</th>
            <th>Gap</th>
          </tr>
        </thead>
        <tbody>
          <tr><td><strong>Good water quality</strong></td><td>20.8%</td><td>20.7%</td><td>20.9%</td><td><span class="badge badge-green">+0.2 pts</span></td></tr>
          <tr><td><strong>Poor water quality</strong></td><td>40.0%</td><td>40.5%</td><td>39.3%</td><td><span class="badge badge-yellow">-1.2 pts</span></td></tr>
          <tr class="comparison-highlight"><td><strong>Access &lt;30 min</strong></td><td>38.6%</td><td>37.5%</td><td>40.0%</td><td><span class="badge badge-teal">+2.5 pts</span></td></tr>
          <tr><td><strong>Access &gt;60 min</strong></td><td>27.3%</td><td>28.1%</td><td>26.3%</td><td><span class="badge badge-green">-1.8 pts</span></td></tr>
          <tr class="comparison-highlight"><td><strong>Treated water</strong></td><td>29.1%</td><td>30.0%</td><td>28.1%</td><td><span class="badge badge-orange">-1.9 pts</span></td></tr>
          <tr><td><strong>Waterborne diseases</strong></td><td>27.1%</td><td>26.7%</td><td>27.7%</td><td><span class="badge badge-yellow">+1.0 pts</span></td></tr>
          <tr class="comparison-highlight"><td><strong>No latrine (OD)</strong></td><td>33.6%</td><td>33.3%</td><td>33.3%</td><td><span class="badge badge-green">0.0 pts</span></td></tr>
          <tr><td><strong>Atmospheric pollution</strong></td><td>49.4%</td><td>50.0%</td><td>48.6%</td><td><span class="badge badge-yellow">-1.4 pts</span></td></tr>
          <tr class="comparison-highlight"><td><strong>Scarce resources</strong></td><td>38.4%</td><td>~39.0%</td><td>~37.5%</td><td><span class="badge badge-yellow">-1.5 pts</span></td></tr>
          <tr><td><strong>Resource degradation</strong></td><td>57.7%</td><td>~58.0%</td><td>~57.2%</td><td><span class="badge badge-yellow">-0.8 pts</span></td></tr>
          <tr><td><strong>Sufficient water</strong></td><td>58.0%</td><td>~58.5%</td><td>~57.2%</td><td><span class="badge badge-yellow">-1.3 pts</span></td></tr>
          <tr><td><strong>Exposed to waste</strong></td><td>52.3%</td><td>~53.0%</td><td>~51.4%</td><td><span class="badge badge-yellow">-1.6 pts</span></td></tr>
          <tr><td><strong>Noise nuisances</strong></td><td>57.3%</td><td>~57.9%</td><td>~56.5%</td><td><span class="badge badge-yellow">-1.4 pts</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== RECOMMENDATIONS ===== -->
<section id="recommandations" class="section">
  <div class="section-header">
    <div class="section-num">Section 9</div>
    <h2>Recommendations</h2>
  </div>
  <div class="section-body">
    <p>Based on the findings of this survey, the following recommendations are formulated, ranked in order of priority according to urgency and potential impact on public health and the environment.</p>

    <div class="reco-grid">
      <div class="reco-card">
        <div class="reco-card-header priority-1">
          <div class="priority-num">1</div>
          <div class="reco-title">Water quality and treatment</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Mass campaigns on household water treatment (chlorination, filtration, boiling)</li>
            <li>Construction of additional boreholes in high-risk areas</li>
            <li>Regular testing of water source quality</li>
            <li>Rehabilitation of existing traditional wells</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-2">
          <div class="priority-num">2</div>
          <div class="reco-title">Sanitation &amp; Latrines</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Open defecation free (ODF) elimination programme</li>
            <li>Subsidies for construction of improved latrines</li>
            <li>Awareness of waterborne diseases linked to lack of sanitation</li>
            <li>Training local masons in latrine construction</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-3">
          <div class="priority-num">3</div>
          <div class="reco-title">Waste management</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Establish structured collection systems in underserved areas</li>
            <li>Promote composting and organic waste recycling</li>
            <li>Ban and penalise open-air burning</li>
            <li>Regulated and secured transit depots</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-4">
          <div class="priority-num">4</div>
          <div class="reco-title">Atmospheric pollution</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Mapping and monitoring of pollution sources</li>
            <li>Regulation of industrial and artisanal emissions</li>
            <li>Promote improved cookstoves to reduce domestic smoke</li>
            <li>Buffer zones around urban agglomerations</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-5">
          <div class="priority-num">5</div>
          <div class="reco-title">Natural resources</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Local sustainable natural resource management plans</li>
            <li>Reforestation and degraded land restoration</li>
            <li>Community awareness on environmental preservation</li>
            <li>Participatory resource monitoring systems</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-6">
          <div class="priority-num">6</div>
          <div class="reco-title">Monitoring &amp; Governance</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>WASH indicator monitoring and evaluation systems</li>
            <li>Capacity building for local authorities</li>
            <li>Community involvement in infrastructure management</li>
            <li>Repeat the survey every 2 years to measure progress</li>
          </ul>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===== CONCLUSION ===== -->
<section id="conclusion" class="section">
  <div class="section-header">
    <div class="section-num">Section 10</div>
    <h2>Conclusion</h2>
  </div>
  <div class="section-body">
    <p>This WASH survey conducted among <strong>1,000 households</strong> in the Goh and Loh Djiboua departments of C&ocirc;te d&rsquo;Ivoire highlights considerable challenges in drinking water, sanitation and environmental management.</p>

    <p>The most concerning findings relate to <strong>water quality</strong> (40% poor quality), the <strong>low rate of household treatment</strong> (29.1%), and the <strong>prevalence of open defecation</strong> (33.6%). These combined factors largely explain the 27.1% of households reporting water-related diseases.</p>

    <p>Relative to the <strong>Sustainable Development Goals (SDG 6)</strong> &mdash; ensuring access to safe water and sanitation for all &mdash; both departments remain far from targets. Targeted investments and behaviour change programmes are essential to reduce access inequalities and protect population health.</p>

    <p>As the profiles of the two departments are very similar, interventions can be planned in a coordinated manner, while accounting for the local specificities of each sub-prefecture. A <strong>repeat survey in two years</strong> will allow progress to be measured and intervention strategies to be adapted.</p>

    <div class="kpi-grid" style="margin-top:24px">
      <div class="kpi-card red">
        <div class="kpi-num">3</div>
        <div class="kpi-label">Critical alerts</div>
        <div class="kpi-sub">Water, sanitation, health</div>
      </div>
      <div class="kpi-card orange">
        <div class="kpi-num">6</div>
        <div class="kpi-label">Recommendations</div>
        <div class="kpi-sub">Ranked by priority</div>
      </div>
      <div class="kpi-card teal">
        <div class="kpi-num">8</div>
        <div class="kpi-label">Sub-prefectures</div>
        <div class="kpi-sub">2 departments covered</div>
      </div>
      <div class="kpi-card green">
        <div class="kpi-num">166</div>
        <div class="kpi-label">Variables analysed</div>
        <div class="kpi-sub">Comprehensive coverage</div>
      </div>
    </div>
  </div>
</section>

<!-- FOOTER -->
<footer class="report-footer">
  <strong>WASH Survey Report &ndash; C&ocirc;te d&rsquo;Ivoire 2026</strong> &bull;
  WASH Survey &bull; Goh &amp; Loh Djiboua Regions &bull;
  N = 1,000 households &bull; January&ndash;February 2026 &bull;
  Generated 28 March 2026
</footer>

</main>
</div>

<script>
// Highlight active section in sidebar
(function() {
  var sections = document.querySelectorAll('section[id]');
  var navLinks = document.querySelectorAll('.sidebar nav a');
  window.addEventListener('scroll', function() {
    var scrollY = window.scrollY + 80;
    sections.forEach(function(sec) {
      var top = sec.offsetTop;
      var bottom = top + sec.offsetHeight;
      if (scrollY >= top && scrollY < bottom) {
        navLinks.forEach(function(a) { a.classList.remove('active'); });
        var match = document.querySelector('.sidebar nav a[href="#' + sec.id + '"]');
        if (match) match.classList.add('active');
      }
    });
  });
})();
</script>
</body>
</html>
"""

out_path = r'C:\Users\ITrebi\Music\PortFolio Upwork\Excel Dashboard\WASH_Report_EN.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print("HTML report saved to: " + out_path)
