# -*- coding: utf-8 -*-
# generate_html_report.py
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_charts as gc

print("Generating charts...")
charts = {}
for name, fn in gc.CHARTS.items():
    charts[name] = gc.to_b64(fn())
    print(f"  Chart '{name}' done.")

print("Building HTML...")

html = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Rapport WASH &ndash; C&ocirc;te d&rsquo;Ivoire 2026</title>
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
    <h2>Rapport WASH</h2>
    <p>C&ocirc;te d&rsquo;Ivoire &bull; 2026</p>
  </div>
  <nav>
    <ul>
      <li class="section-header">Navigation</li>
      <li><a href="#cover">Page de couverture</a></li>
      <li><a href="#resume">R&eacute;sum&eacute; ex&eacute;cutif</a></li>
      <li><a href="#contexte">Contexte &amp; M&eacute;thodologie</a></li>
      <li class="section-header">Sections WASH</li>
      <li><a href="#eau">Eau potable</a></li>
      <li><a href="#assainissement">Assainissement</a></li>
      <li><a href="#dechets">Gestion des d&eacute;chets</a></li>
      <li><a href="#pollution">Pollution &amp; Air</a></li>
      <li><a href="#ressources">Ressources naturelles</a></li>
      <li class="section-header">Analyse</li>
      <li><a href="#comparaison">Analyse comparative</a></li>
      <li><a href="#recommandations">Recommandations</a></li>
      <li><a href="#conclusion">Conclusion</a></li>
    </ul>
  </nav>
</aside>

<!-- MAIN CONTENT -->
<main class="main-content">

<!-- ===== COVER PAGE ===== -->
<section id="cover" class="cover-page">
  <div class="org-badge">Rapport Technique &bull; 2026</div>
  <div class="cover-logo">&#9749;</div>
  <h1>Rapport WASH</h1>
  <p class="subtitle">Eau, Assainissement &amp; Hygi&egrave;ne &mdash; C&ocirc;te d&rsquo;Ivoire</p>
  <div class="cover-divider"></div>
  <div class="cover-meta">
    <div class="cover-meta-item">
      <span class="label">Zone d&rsquo;&eacute;tude</span>
      <span class="value">R&eacute;gions Goh &amp; Loh Djiboua</span>
    </div>
    <div class="cover-meta-item">
      <span class="label">P&eacute;riode</span>
      <span class="value">Janvier &ndash; F&eacute;vrier 2026</span>
    </div>
    <div class="cover-meta-item">
      <span class="label">Sous-pr&eacute;fectures</span>
      <span class="value">8 localit&eacute;s</span>
    </div>
  </div>
  <div class="cover-stats">
    <div class="cover-stat">
      <span class="num">1 000</span>
      <span class="desc">M&eacute;nages enqu&ecirc;t&eacute;s</span>
    </div>
    <div class="cover-stat">
      <span class="num">166</span>
      <span class="desc">Variables collect&eacute;es</span>
    </div>
    <div class="cover-stat">
      <span class="num">2</span>
      <span class="desc">D&eacute;partements</span>
    </div>
    <div class="cover-stat">
      <span class="num">40%</span>
      <span class="desc">Eau de mauvaise qualit&eacute;</span>
    </div>
  </div>
</section>

<!-- ===== RESUME EXECUTIF ===== -->
<section id="resume" class="section">
  <div class="section-header">
    <div class="section-num">Section 1</div>
    <h2>R&eacute;sum&eacute; Ex&eacute;cutif</h2>
  </div>
  <div class="section-body">
    <p>Cette enqu&ecirc;te WASH a &eacute;t&eacute; conduite aupr&egrave;s de <strong>1 000 m&eacute;nages</strong> r&eacute;partis dans les d&eacute;partements du <strong>Goh (570)</strong> et de <strong>Loh Djiboua (430)</strong> en C&ocirc;te d&rsquo;Ivoire, couvrant 8 sous-pr&eacute;fectures entre janvier et f&eacute;vrier 2026. Les donn&eacute;es collectent 166 indicateurs sur l&rsquo;eau, l&rsquo;assainissement, la gestion des d&eacute;chets, la pollution atmosphrique et la disponibilit&eacute; des ressources naturelles.</p>

    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-num">1 000</div>
        <div class="kpi-label">M&eacute;nages enqu&ecirc;t&eacute;s</div>
        <div class="kpi-sub">Jan (358) + F&eacute;v (642)</div>
      </div>
      <div class="kpi-card red">
        <div class="kpi-num">40%</div>
        <div class="kpi-label">Eau mauvaise qualit&eacute;</div>
        <div class="kpi-sub">400 m&eacute;nages affect&eacute;s</div>
      </div>
      <div class="kpi-card orange">
        <div class="kpi-num">70.9%</div>
        <div class="kpi-label">Eau non trait&eacute;e</div>
        <div class="kpi-sub">709 m&eacute;nages</div>
      </div>
      <div class="kpi-card yellow">
        <div class="kpi-num">33.6%</div>
        <div class="kpi-label">Sans latrine</div>
        <div class="kpi-sub">336 m&eacute;nages</div>
      </div>
      <div class="kpi-card teal">
        <div class="kpi-num">49.4%</div>
        <div class="kpi-label">Pollution atmosph&eacute;rique</div>
        <div class="kpi-sub">494 signalements</div>
      </div>
      <div class="kpi-card green">
        <div class="kpi-num">58%</div>
        <div class="kpi-label">Eau suffisante</div>
        <div class="kpi-sub">580 m&eacute;nages satisfaits</div>
      </div>
    </div>

    <div class="alert-grid">
      <div class="alert-box danger">
        <div class="alert-icon">&#9888;</div>
        <div>
          <div class="alert-title">Alerte critique &ndash; Qualit&eacute; de l&rsquo;eau</div>
          <div class="alert-text">40% des m&eacute;nages ont une eau de mauvaise qualit&eacute; et 70,9% ne traitent pas leur eau avant consommation. Risque &eacute;lev&eacute; de maladies hydriques.</div>
        </div>
      </div>
      <div class="alert-box danger">
        <div class="alert-icon">&#9888;</div>
        <div>
          <div class="alert-title">Alerte critique &ndash; Assainissement</div>
          <div class="alert-text">33,6% des m&eacute;nages n&rsquo;ont pas de latrines. La couverture en latrines am&eacute;lior&eacute;es reste tr&egrave;s faible (9%).</div>
        </div>
      </div>
      <div class="alert-box warning">
        <div class="alert-icon">&#9679;</div>
        <div>
          <div class="alert-title">Pr&eacute;occupation &ndash; Maladies hydriques</div>
          <div class="alert-text">27,1% des m&eacute;nages d&eacute;clarent des maladies li&eacute;es &agrave; l&rsquo;eau. La diarrh&eacute;e domine (~180 cas) suivie de la typho&iuml;de (~120 cas).</div>
        </div>
      </div>
      <div class="alert-box warning">
        <div class="alert-icon">&#9679;</div>
        <div>
          <div class="alert-title">Pr&eacute;occupation &ndash; Ressources naturelles</div>
          <div class="alert-text">57,7% des m&eacute;nages signalent une d&eacute;gradation des ressources. 38,4% les jugent rares, et 16% les estiment en voie de disparition.</div>
        </div>
      </div>
      <div class="alert-box success">
        <div class="alert-icon">&#10003;</div>
        <div>
          <div class="alert-title">Point positif &ndash; Disponibilit&eacute; en eau</div>
          <div class="alert-text">58% des m&eacute;nages disposent d&rsquo;une quantit&eacute; d&rsquo;eau suffisante pour leurs besoins quotidiens.</div>
        </div>
      </div>
      <div class="alert-box info">
        <div class="alert-icon">&#9432;</div>
        <div>
          <div class="alert-title">Information &ndash; Gestion des d&eacute;chets</div>
          <div class="alert-text">La collecte/valorisation est la pratique majoritaire (515 m&eacute;nages), mais le br&ucirc;lage reste tr&egrave;s r&eacute;pandu (499) avec des risques sanitaires associ&eacute;s.</div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===== CONTEXTE & METHODOLOGIE ===== -->
<section id="contexte" class="section">
  <div class="section-header">
    <div class="section-num">Section 2</div>
    <h2>Contexte &amp; M&eacute;thodologie</h2>
  </div>
  <div class="section-body">
    <p>L&rsquo;enqu&ecirc;te WASH 2026 vise &agrave; &eacute;valuer les conditions d&rsquo;acc&egrave;s &agrave; l&rsquo;eau potable, l&rsquo;&eacute;tat de l&rsquo;assainissement et les pratiques d&rsquo;hygi&egrave;ne dans les r&eacute;gions du Goh et de Loh Djiboua. Ces deux d&eacute;partements pr&eacute;sentent des enjeux importants en mati&egrave;re de sant&eacute; publique et de gestion environnementale.</p>

    <div class="chart-grid">
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['departements'] + """" alt="Repartition par departement">
        <div class="chart-caption">Fig. 1 &ndash; R&eacute;partition de l&rsquo;&eacute;chantillon par d&eacute;partement</div>
      </div>
      <div>
        <div class="subsection" style="border-top:none; margin-top:0; padding-top:0;">
          <h3>Cadre m&eacute;thodologique</h3>
          <div class="progress-item">
            <div class="progress-label">Taille de l&rsquo;&eacute;chantillon</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:100%;background:var(--blue)"></div></div>
            <div class="progress-val">N = 1 000</div>
          </div>
          <div class="progress-item">
            <div class="progress-label">D&eacute;partement Goh</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:57%;background:var(--blue)"></div></div>
            <div class="progress-val">570 (57%)</div>
          </div>
          <div class="progress-item">
            <div class="progress-label">D&eacute;partement Loh Djiboua</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:43%;background:var(--teal)"></div></div>
            <div class="progress-val">430 (43%)</div>
          </div>
          <div class="progress-item">
            <div class="progress-label">Enqu&ecirc;tes janvier</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:35.8%;background:var(--orange)"></div></div>
            <div class="progress-val">358 (35.8%)</div>
          </div>
          <div class="progress-item">
            <div class="progress-label">Enqu&ecirc;tes f&eacute;vrier</div>
            <div class="progress-bar-wrap"><div class="progress-bar" style="width:64.2%;background:var(--green)"></div></div>
            <div class="progress-val">642 (64.2%)</div>
          </div>
        </div>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Tableau 1 &ndash; R&eacute;partition de l&rsquo;&eacute;chantillon par sous-pr&eacute;fecture</caption>
        <thead>
          <tr>
            <th>Sous-pr&eacute;fecture</th>
            <th>D&eacute;partement</th>
            <th>Enqu&ecirc;tes</th>
            <th>% de l&rsquo;&eacute;chantillon</th>
            <th>Statut</th>
          </tr>
        </thead>
        <tbody>
          <tr><td>Gagnoa</td><td>Goh</td><td>~115</td><td>~11.5%</td><td><span class="badge badge-blue">Chef-lieu</span></td></tr>
          <tr><td>Bayota</td><td>Goh</td><td>~95</td><td>~9.5%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr><td>Ouragahio</td><td>Goh</td><td>~90</td><td>~9.0%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr><td>Guitry</td><td>Goh</td><td>~88</td><td>~8.8%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr><td>Guiberoua</td><td>Goh</td><td>~85</td><td>~8.5%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr><td>Divo</td><td>Loh Djiboua</td><td>~155</td><td>~15.5%</td><td><span class="badge badge-blue">Chef-lieu</span></td></tr>
          <tr><td>Lakota</td><td>Loh Djiboua</td><td>~140</td><td>~14.0%</td><td><span class="badge badge-blue">Chef-lieu</span></td></tr>
          <tr><td>Hire</td><td>Loh Djiboua</td><td>~135</td><td>~13.5%</td><td><span class="badge badge-teal">Rural</span></td></tr>
          <tr style="font-weight:700;background:#f0f4ff"><td>TOTAL</td><td>2 d&eacute;partements</td><td>1 000</td><td>100%</td><td><span class="badge badge-green">Complet</span></td></tr>
        </tbody>
      </table>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Tableau 2 &ndash; Outils et variables de collecte</caption>
        <thead>
          <tr><th>Th&eacute;matique</th><th>Variables cl&eacute;s</th><th>Type d&rsquo;indicateur</th></tr>
        </thead>
        <tbody>
          <tr><td>Eau potable</td><td>Source, qualit&eacute;, temps d&rsquo;acc&egrave;s, traitement, disponibilit&eacute;</td><td><span class="badge badge-blue">Quantitatif</span></td></tr>
          <tr><td>Sant&eacute; hydrique</td><td>Maladies li&eacute;es &agrave; l&rsquo;eau, fr&eacute;quence, type</td><td><span class="badge badge-orange">Mixte</span></td></tr>
          <tr><td>Assainissement</td><td>Type de latrine, pratiques d&rsquo;hygi&egrave;ne</td><td><span class="badge badge-blue">Quantitatif</span></td></tr>
          <tr><td>D&eacute;chets</td><td>Mode de gestion, satisfaction, exposition</td><td><span class="badge badge-blue">Quantitatif</span></td></tr>
          <tr><td>Environnement</td><td>Pollution, qualit&eacute; air, nuisances sonores</td><td><span class="badge badge-teal">Qualitatif</span></td></tr>
          <tr><td>Ressources</td><td>Disponibilit&eacute;, d&eacute;gradation, respect</td><td><span class="badge badge-teal">Qualitatif</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== SECTION EAU ===== -->
<section id="eau" class="section">
  <div class="section-header">
    <div class="section-num">Section 3</div>
    <h2>Eau Potable</h2>
  </div>
  <div class="section-body">
    <p>L&rsquo;acc&egrave;s &agrave; l&rsquo;eau potable est l&rsquo;un des d&eacute;fis majeurs identifi&eacute;s dans cette enqu&ecirc;te. Si 58% des m&eacute;nages disposent d&rsquo;une eau en quantit&eacute; suffisante, la qualit&eacute; reste pr&eacute;occupante avec <strong>40% d&rsquo;eau de mauvaise qualit&eacute;</strong> et seulement <strong>29,1% des m&eacute;nages qui traitent leur eau</strong>.</p>

    <div class="subsection">
      <h3>3.1 Sources d&rsquo;approvisionnement</h3>
      <div class="chart-grid">
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['sources'] + """" alt="Sources d'approvisionnement">
          <div class="chart-caption">Fig. 2 &ndash; Sources d&rsquo;approvisionnement en eau (N=1000)</div>
        </div>
        <div>
          <div class="table-wrap">
            <table>
              <caption>Tableau 3 &ndash; D&eacute;tail des sources d&rsquo;eau</caption>
              <thead><tr><th>Source</th><th>Nb m&eacute;nages</th><th>%</th><th>&Eacute;valuation</th></tr></thead>
              <tbody>
                <tr><td>Puits traditionnel</td><td>275</td><td>27.5%</td><td><span class="badge badge-red">Risque &eacute;lev&eacute;</span></td></tr>
                <tr><td>Puits am&eacute;lior&eacute;</td><td>200</td><td>20.0%</td><td><span class="badge badge-yellow">Risque mod&eacute;r&eacute;</span></td></tr>
                <tr><td>Forage</td><td>196</td><td>19.6%</td><td><span class="badge badge-green">S&ucirc;r</span></td></tr>
                <tr><td>Rivi&egrave;re / Source</td><td>167</td><td>16.7%</td><td><span class="badge badge-red">Risque tr&egrave;s &eacute;lev&eacute;</span></td></tr>
                <tr><td>Autres</td><td>82</td><td>8.2%</td><td><span class="badge badge-orange">Ind&eacute;termin&eacute;</span></td></tr>
                <tr><td>SODECI (r&eacute;seau)</td><td>80</td><td>8.0%</td><td><span class="badge badge-green">S&ucirc;r</span></td></tr>
                <tr style="font-weight:700"><td>Total</td><td>1 000</td><td>100%</td><td></td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <div class="subsection">
      <h3>3.2 Qualit&eacute; de l&rsquo;eau par d&eacute;partement</h3>
      <div class="chart-grid">
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['qualite_dept'] + """" alt="Qualite de l'eau par departement">
          <div class="chart-caption">Fig. 3 &ndash; Qualit&eacute; de l&rsquo;eau per&ccedil;ue par d&eacute;partement</div>
        </div>
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['traitement'] + """" alt="Traitement de l'eau">
          <div class="chart-caption">Fig. 4 &ndash; Traitement de l&rsquo;eau au niveau domestique</div>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <caption>Tableau 4 &ndash; Qualit&eacute; de l&rsquo;eau par d&eacute;partement</caption>
          <thead><tr><th>D&eacute;partement</th><th>Bonne</th><th>Acceptable</th><th>Mauvaise</th><th>Total</th><th>% mauvaise</th></tr></thead>
          <tbody>
            <tr><td><strong>Goh</strong></td><td>118 (20.7%)</td><td>221 (38.8%)</td><td>231 (40.5%)</td><td>570</td><td><span class="badge badge-red">40.5%</span></td></tr>
            <tr><td><strong>Loh Djiboua</strong></td><td>90 (20.9%)</td><td>171 (39.8%)</td><td>169 (39.3%)</td><td>430</td><td><span class="badge badge-red">39.3%</span></td></tr>
            <tr style="font-weight:700"><td>Total</td><td>208 (20.8%)</td><td>392 (39.2%)</td><td>400 (40.0%)</td><td>1 000</td><td><span class="badge badge-red">40.0%</span></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="subsection">
      <h3>3.3 Acc&egrave;s &agrave; l&rsquo;eau et maladies hydriques</h3>
      <div class="chart-grid">
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['acces'] + """" alt="Temps d'acces a l'eau">
          <div class="chart-caption">Fig. 5 &ndash; Temps d&rsquo;acc&egrave;s &agrave; la source d&rsquo;eau</div>
        </div>
        <div class="chart-card">
          <img src="data:image/png;base64,""" + charts['maladies'] + """" alt="Maladies hydriques">
          <div class="chart-caption">Fig. 6 &ndash; Maladies hydriques d&eacute;clar&eacute;es par d&eacute;partement</div>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <caption>Tableau 5 &ndash; Temps d&rsquo;acc&egrave;s &agrave; l&rsquo;eau et implications</caption>
          <thead><tr><th>Tranche de temps</th><th>M&eacute;nages</th><th>%</th><th>Norme OMS</th><th>Statut</th></tr></thead>
          <tbody>
            <tr><td>Moins de 30 minutes</td><td>386</td><td>38.6%</td><td>Conforme</td><td><span class="badge badge-green">Acceptable</span></td></tr>
            <tr><td>30 &agrave; 60 minutes</td><td>341</td><td>34.1%</td><td>Limite</td><td><span class="badge badge-yellow">Attention</span></td></tr>
            <tr><td>Plus de 60 minutes</td><td>273</td><td>27.3%</td><td>Non conforme</td><td><span class="badge badge-red">Critique</span></td></tr>
          </tbody>
        </table>
      </div>
      <div class="table-wrap">
        <table>
          <caption>Tableau 6 &ndash; Maladies hydriques d&eacute;clar&eacute;es</caption>
          <thead><tr><th>Maladie</th><th>Goh</th><th>Loh Djiboua</th><th>Total estim&eacute;</th><th>Gravit&eacute;</th></tr></thead>
          <tbody>
            <tr><td>Diarrh&eacute;e</td><td>~115</td><td>~87</td><td>~202</td><td><span class="badge badge-orange">Mod&eacute;r&eacute;e</span></td></tr>
            <tr><td>Typho&iuml;de</td><td>~72</td><td>~55</td><td>~127</td><td><span class="badge badge-red">Haute</span></td></tr>
            <tr><td>H&eacute;patite A</td><td>~60</td><td>~43</td><td>~103</td><td><span class="badge badge-red">Haute</span></td></tr>
            <tr><td>Chol&eacute;ra</td><td>~50</td><td>~34</td><td>~84</td><td><span class="badge badge-red">Tr&egrave;s haute</span></td></tr>
            <tr style="font-weight:700"><td>Total d&eacute;clar&eacute;s</td><td>~297</td><td>~219</td><td>~516 cas</td><td></td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</section>

<!-- ===== SECTION ASSAINISSEMENT ===== -->
<section id="assainissement" class="section">
  <div class="section-header">
    <div class="section-num">Section 4</div>
    <h2>Assainissement</h2>
  </div>
  <div class="section-body">
    <p>L&rsquo;assainissement repr&eacute;sente un d&eacute;fi majeur dans les deux d&eacute;partements. Pr&egrave;s d&rsquo;un tiers des m&eacute;nages (33,6%) n&rsquo;ont pas acc&egrave;s &agrave; des latrines, et seulement <strong>14,8% disposent de latrines am&eacute;lior&eacute;es ou de WC modernes</strong>, loin des objectifs de l&rsquo;ODD 6.</p>

    <div class="chart-full">
      <img src="data:image/png;base64,""" + charts['latrines'] + """" alt="Types d'infrastructures sanitaires">
      <div class="chart-caption">Fig. 7 &ndash; Types d&rsquo;infrastructures sanitaires (N=1000)</div>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Tableau 7 &ndash; Types de latrines / infrastructures sanitaires</caption>
        <thead><tr><th>Type d&rsquo;infrastructure</th><th>Nb m&eacute;nages</th><th>%</th><th>Conformit&eacute; ODD 6</th></tr></thead>
        <tbody>
          <tr><td>Latrine simple</td><td>424</td><td>42.4%</td><td><span class="badge badge-yellow">Partielle</span></td></tr>
          <tr><td>Aucune latrine (d&eacute;f&eacute;cation &agrave; l&rsquo;air libre)</td><td>336</td><td>33.6%</td><td><span class="badge badge-red">Non conforme</span></td></tr>
          <tr><td>Autre type</td><td>92</td><td>9.2%</td><td><span class="badge badge-orange">Ind&eacute;termin&eacute;</span></td></tr>
          <tr><td>Latrines am&eacute;lior&eacute;es</td><td>90</td><td>9.0%</td><td><span class="badge badge-green">Conforme</span></td></tr>
          <tr><td>WC modernes</td><td>58</td><td>5.8%</td><td><span class="badge badge-green">Conforme</span></td></tr>
          <tr style="font-weight:700"><td>Total</td><td>1 000</td><td>100%</td><td></td></tr>
        </tbody>
      </table>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Tableau 8 &ndash; Indicateurs d&rsquo;assainissement cl&eacute;s</caption>
        <thead><tr><th>Indicateur</th><th>Valeur</th><th>Cible ODD 6</th><th>Ecart</th></tr></thead>
        <tbody>
          <tr><td>Acc&egrave;s &agrave; des latrines (tout type)</td><td>66.4%</td><td>100%</td><td><span class="badge badge-red">-33.6 pts</span></td></tr>
          <tr><td>Latrines am&eacute;lior&eacute;es + WC</td><td>14.8%</td><td>100%</td><td><span class="badge badge-red">-85.2 pts</span></td></tr>
          <tr><td>D&eacute;f&eacute;cation &agrave; l&rsquo;air libre (DAL)</td><td>33.6%</td><td>0%</td><td><span class="badge badge-red">+33.6 pts</span></td></tr>
          <tr><td>Traitement de l&rsquo;eau domestique</td><td>29.1%</td><td>100%</td><td><span class="badge badge-red">-70.9 pts</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== SECTION DECHETS ===== -->
<section id="dechets" class="section">
  <div class="section-header">
    <div class="section-num">Section 5</div>
    <h2>Gestion des D&eacute;chets</h2>
  </div>
  <div class="section-body">
    <p>La gestion des d&eacute;chets est une probl&eacute;matique complexe dans les deux d&eacute;partements. Si 52,3% des m&eacute;nages sont expos&eacute;s &agrave; des d&eacute;charges sauvages, les pratiques de collecte et valorisation montrent un potentiel d&rsquo;am&eacute;lioration. La satisfaction reste majoritairement n&eacute;gative avec <strong>62% de m&eacute;nages peu ou pas satisfaits</strong>.</p>

    <div class="chart-grid">
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['gestion_modes'] + """" alt="Modes de gestion des dechets">
        <div class="chart-caption">Fig. 8 &ndash; Modes de gestion des d&eacute;chets</div>
      </div>
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['dechets_sat'] + """" alt="Satisfaction gestion des dechets">
        <div class="chart-caption">Fig. 9 &ndash; Satisfaction quant &agrave; la gestion des d&eacute;chets</div>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Tableau 9 &ndash; Modes de gestion des d&eacute;chets</caption>
        <thead><tr><th>Mode de gestion</th><th>M&eacute;nages</th><th>%</th><th>Impact environnemental</th></tr></thead>
        <tbody>
          <tr><td>Collecte / Valorisation</td><td>515</td><td>51.5%</td><td><span class="badge badge-green">Positif</span></td></tr>
          <tr><td>Abandon / D&eacute;charge sauvage</td><td>501</td><td>50.1%</td><td><span class="badge badge-red">Tr&egrave;s n&eacute;gatif</span></td></tr>
          <tr><td>Br&ucirc;lage</td><td>499</td><td>49.9%</td><td><span class="badge badge-red">N&eacute;gatif</span></td></tr>
          <tr><td>R&eacute;utilisation</td><td>287</td><td>28.7%</td><td><span class="badge badge-green">Positif</span></td></tr>
        </tbody>
      </table>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Tableau 10 &ndash; Satisfaction de la gestion des d&eacute;chets</caption>
        <thead><tr><th>Niveau de satisfaction</th><th>M&eacute;nages</th><th>%</th><th>Tendance</th></tr></thead>
        <tbody>
          <tr><td>Peu satisfait</td><td>378</td><td>37.8%</td><td><span class="badge badge-orange">N&eacute;gatif</span></td></tr>
          <tr><td>Satisfait</td><td>271</td><td>27.1%</td><td><span class="badge badge-green">Positif</span></td></tr>
          <tr><td>Pas satisfait</td><td>242</td><td>24.2%</td><td><span class="badge badge-red">N&eacute;gatif</span></td></tr>
          <tr><td>Tr&egrave;s satisfait</td><td>109</td><td>10.9%</td><td><span class="badge badge-green">Positif</span></td></tr>
          <tr style="font-weight:700"><td>Total</td><td>1 000</td><td>100%</td><td></td></tr>
        </tbody>
      </table>
    </div>

    <div class="alert-box warning" style="margin-top:16px">
      <div class="alert-icon">&#9432;</div>
      <div>
        <div class="alert-title">Point d&rsquo;attention &ndash; Exposition aux d&eacute;chets</div>
        <div class="alert-text">52,3% des m&eacute;nages (523) d&eacute;clarent &ecirc;tre expos&eacute;s &agrave; des d&eacute;charges sauvages ou d&eacute;ch&eacute;teries non contr&ocirc;l&eacute;es. Cette exposition constitue un risque sanitaire et environnemental majeur n&eacute;cessitant des interventions urgentes.</div>
      </div>
    </div>
  </div>
</section>

<!-- ===== SECTION POLLUTION ===== -->
<section id="pollution" class="section">
  <div class="section-header">
    <div class="section-num">Section 6</div>
    <h2>Pollution &amp; Qualit&eacute; de l&rsquo;Air</h2>
  </div>
  <div class="section-body">
    <p>La pollution atmosph&eacute;rique est signal&eacute;e par pr&egrave;s d&rsquo;un m&eacute;nage sur deux (49,4%). La qualit&eacute; de l&rsquo;air est jug&eacute;e mauvaise ou tr&egrave;s mauvaise par <strong>51,6% des r&eacute;pondants</strong>. Les nuisances sonores affectent &eacute;galement une majorit&eacute; des m&eacute;nages (57,3%), refl&eacute;tant une urbanisation croissante et des activit&eacute;s industrielles peu r&eacute;gul&eacute;es.</p>

    <div class="chart-grid">
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['pollution'] + """" alt="Pollution atmospherique">
        <div class="chart-caption">Fig. 10 &ndash; Pollution atmosph&eacute;rique signal&eacute;e par d&eacute;partement</div>
      </div>
      <div class="chart-card">
        <img src="data:image/png;base64,""" + charts['qualite_air'] + """" alt="Qualite de l'air">
        <div class="chart-caption">Fig. 11 &ndash; Qualit&eacute; de l&rsquo;air per&ccedil;ue (N=1000)</div>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Tableau 11 &ndash; Indicateurs de pollution et qualit&eacute; de l&rsquo;air</caption>
        <thead><tr><th>Indicateur</th><th>M&eacute;nages</th><th>%</th><th>Niveau de risque</th></tr></thead>
        <tbody>
          <tr><td>Pollution atmosph&eacute;rique signal&eacute;e</td><td>494</td><td>49.4%</td><td><span class="badge badge-red">Pr&eacute;occupant</span></td></tr>
          <tr><td>Qualit&eacute; air : Mauvaise</td><td>364</td><td>36.4%</td><td><span class="badge badge-red">Critique</span></td></tr>
          <tr><td>Qualit&eacute; air : Moyenne</td><td>361</td><td>36.1%</td><td><span class="badge badge-yellow">Mod&eacute;r&eacute;</span></td></tr>
          <tr><td>Qualit&eacute; air : Tr&egrave;s mauvaise</td><td>152</td><td>15.2%</td><td><span class="badge badge-red">Tr&egrave;s critique</span></td></tr>
          <tr><td>Qualit&eacute; air : Bonne</td><td>123</td><td>12.3%</td><td><span class="badge badge-green">Acceptable</span></td></tr>
          <tr><td>Nuisances sonores</td><td>573</td><td>57.3%</td><td><span class="badge badge-orange">Pr&eacute;occupant</span></td></tr>
        </tbody>
      </table>
    </div>

    <div class="table-wrap">
      <table>
        <caption>Tableau 12 &ndash; Pollution par d&eacute;partement</caption>
        <thead><tr><th>D&eacute;partement</th><th>Pollution signal&eacute;e</th><th>Pas de pollution</th><th>Taux</th></tr></thead>
        <tbody>
          <tr><td><strong>Goh</strong></td><td>285</td><td>285</td><td><span class="badge badge-yellow">50.0%</span></td></tr>
          <tr><td><strong>Loh Djiboua</strong></td><td>209</td><td>221</td><td><span class="badge badge-yellow">48.6%</span></td></tr>
          <tr style="font-weight:700"><td>Total</td><td>494</td><td>506</td><td><span class="badge badge-orange">49.4%</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== SECTION RESSOURCES NATURELLES ===== -->
<section id="ressources" class="section">
  <div class="section-header">
    <div class="section-num">Section 7</div>
    <h2>Ressources Naturelles</h2>
  </div>
  <div class="section-body">
    <p>La situation des ressources naturelles est alarmante&nbsp;: <strong>57,7% des m&eacute;nages signalent une d&eacute;gradation des ressources</strong>, et seulement 12,1% les consid&egrave;rent abondantes. Cette situation, combin&eacute;e &agrave; un respect de l&rsquo;environnement jug&eacute; fort ou tr&egrave;s fort par 54,1% des m&eacute;nages, indique une prise de conscience mais des pratiques qui restent insuffisantes.</p>

    <div class="chart-full">
      <img src="data:image/png;base64,""" + charts['ressources'] + """" alt="Disponibilite des ressources naturelles">
      <div class="chart-caption">Fig. 12 &ndash; Disponibilit&eacute; per&ccedil;ue des ressources naturelles (N=1000)</div>
    </div>

    <div class="chart-grid">
      <div>
        <div class="table-wrap">
          <table>
            <caption>Tableau 13 &ndash; Disponibilit&eacute; des ressources naturelles</caption>
            <thead><tr><th>Niveau</th><th>M&eacute;nages</th><th>%</th><th>Statut</th></tr></thead>
            <tbody>
              <tr><td>Rares</td><td>384</td><td>38.4%</td><td><span class="badge badge-red">Critique</span></td></tr>
              <tr><td>Moyennement disponibles</td><td>335</td><td>33.5%</td><td><span class="badge badge-yellow">Pr&eacute;caire</span></td></tr>
              <tr><td>En voie de disparition</td><td>160</td><td>16.0%</td><td><span class="badge badge-red">Urgence</span></td></tr>
              <tr><td>Abondantes</td><td>121</td><td>12.1%</td><td><span class="badge badge-green">Favorable</span></td></tr>
            </tbody>
          </table>
        </div>
      </div>
      <div>
        <div class="table-wrap">
          <table>
            <caption>Tableau 14 &ndash; Respect de l&rsquo;environnement</caption>
            <thead><tr><th>Niveau</th><th>M&eacute;nages</th><th>%</th></tr></thead>
            <tbody>
              <tr><td>Fort</td><td>358</td><td>35.8%</td></tr>
              <tr><td>Moyen</td><td>308</td><td>30.8%</td></tr>
              <tr><td>Tr&egrave;s fort</td><td>183</td><td>18.3%</td></tr>
              <tr><td>Faible</td><td>151</td><td>15.1%</td></tr>
            </tbody>
          </table>
        </div>
        <div class="alert-box danger" style="margin-top:12px">
          <div class="alert-icon">&#9888;</div>
          <div>
            <div class="alert-title">D&eacute;gradation confirm&eacute;e</div>
            <div class="alert-text">577 m&eacute;nages (57.7%) signalent une d&eacute;gradation visible des ressources. Actions urgentes requises.</div>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ===== ANALYSE COMPARATIVE ===== -->
<section id="comparaison" class="section">
  <div class="section-header">
    <div class="section-num">Section 8</div>
    <h2>Analyse Comparative</h2>
  </div>
  <div class="section-body">
    <p>La comparaison des indicateurs WASH entre les deux d&eacute;partements r&eacute;v&egrave;le des profils relativement similaires, avec des &eacute;carts faibles sur la plupart des dimensions. Le d&eacute;partement du Goh pr&eacute;sente l&eacute;g&egrave;rement plus de cas de maladies et une plus forte proportion d&rsquo;eau trait&eacute;e, tandis que Loh Djiboua montre un meilleur acc&egrave;s en moins de 30 minutes.</p>

    <div class="chart-full">
      <img src="data:image/png;base64,""" + charts['comparaison'] + """" alt="Comparaison des indicateurs WASH">
      <div class="chart-caption">Fig. 13 &ndash; Comparaison des indicateurs WASH cl&eacute;s par d&eacute;partement</div>
    </div>

    <div class="table-wrap">
      <table class="comparison-table">
        <caption>Tableau 15 &ndash; Tableau comparatif complet des indicateurs WASH</caption>
        <thead>
          <tr>
            <th>Indicateur</th>
            <th>Ensemble (N=1000)</th>
            <th>Goh (N=570)</th>
            <th>Loh Djiboua (N=430)</th>
            <th>Ecart</th>
          </tr>
        </thead>
        <tbody>
          <tr><td><strong>Eau bonne qualit&eacute;</strong></td><td>20.8%</td><td>20.7%</td><td>20.9%</td><td><span class="badge badge-green">+0.2 pts</span></td></tr>
          <tr><td><strong>Eau mauvaise qualit&eacute;</strong></td><td>40.0%</td><td>40.5%</td><td>39.3%</td><td><span class="badge badge-yellow">-1.2 pts</span></td></tr>
          <tr class="comparison-highlight"><td><strong>Acc&egrave;s &lt;30 min</strong></td><td>38.6%</td><td>37.5%</td><td>40.0%</td><td><span class="badge badge-teal">+2.5 pts</span></td></tr>
          <tr><td><strong>Acc&egrave;s &gt;60 min</strong></td><td>27.3%</td><td>28.1%</td><td>26.3%</td><td><span class="badge badge-green">-1.8 pts</span></td></tr>
          <tr class="comparison-highlight"><td><strong>Eau trait&eacute;e</strong></td><td>29.1%</td><td>30.0%</td><td>28.1%</td><td><span class="badge badge-orange">-1.9 pts</span></td></tr>
          <tr><td><strong>Maladies hydriques</strong></td><td>27.1%</td><td>26.7%</td><td>27.7%</td><td><span class="badge badge-yellow">+1.0 pts</span></td></tr>
          <tr class="comparison-highlight"><td><strong>Sans latrine (DAL)</strong></td><td>33.6%</td><td>33.3%</td><td>33.3%</td><td><span class="badge badge-green">0.0 pts</span></td></tr>
          <tr><td><strong>Pollution atmosph&eacute;rique</strong></td><td>49.4%</td><td>50.0%</td><td>48.6%</td><td><span class="badge badge-yellow">-1.4 pts</span></td></tr>
          <tr class="comparison-highlight"><td><strong>Ressources rares</strong></td><td>38.4%</td><td>~39.0%</td><td>~37.5%</td><td><span class="badge badge-yellow">-1.5 pts</span></td></tr>
          <tr><td><strong>D&eacute;gradation ressources</strong></td><td>57.7%</td><td>~58.0%</td><td>~57.2%</td><td><span class="badge badge-yellow">-0.8 pts</span></td></tr>
          <tr><td><strong>Eau suffisante</strong></td><td>58.0%</td><td>~58.5%</td><td>~57.2%</td><td><span class="badge badge-yellow">-1.3 pts</span></td></tr>
          <tr><td><strong>Expos&eacute;s aux d&eacute;chets</strong></td><td>52.3%</td><td>~53.0%</td><td>~51.4%</td><td><span class="badge badge-yellow">-1.6 pts</span></td></tr>
          <tr><td><strong>Nuisances sonores</strong></td><td>57.3%</td><td>~57.9%</td><td>~56.5%</td><td><span class="badge badge-yellow">-1.4 pts</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</section>

<!-- ===== RECOMMANDATIONS ===== -->
<section id="recommandations" class="section">
  <div class="section-header">
    <div class="section-num">Section 9</div>
    <h2>Recommandations</h2>
  </div>
  <div class="section-body">
    <p>Sur la base des r&eacute;sultats de cette enqu&ecirc;te, les recommandations suivantes sont formul&eacute;es, class&eacute;es par ordre de priorit&eacute; selon l&rsquo;urgence et l&rsquo;impact potentiel sur la sant&eacute; publique et l&rsquo;environnement.</p>

    <div class="reco-grid">
      <div class="reco-card">
        <div class="reco-card-header priority-1">
          <div class="priority-num">1</div>
          <div class="reco-title">Qualit&eacute; et traitement de l&rsquo;eau</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Campagnes massives sur le traitement domestique de l&rsquo;eau (chloration, filtration, &eacute;bullition)</li>
            <li>Construction de forages suppl&eacute;mentaires dans les zones &agrave; risque</li>
            <li>Tests r&eacute;guliers de la qualit&eacute; des sources d&rsquo;eau</li>
            <li>R&eacute;habilitation des puits traditionnels existants</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-2">
          <div class="priority-num">2</div>
          <div class="reco-title">Assainissement &amp; Latrines</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Programme d&rsquo;&eacute;limination de la d&eacute;f&eacute;cation &agrave; l&rsquo;air libre (FDAL)</li>
            <li>Subventions pour la construction de latrines am&eacute;lior&eacute;es</li>
            <li>Sensibilisation aux maladies hydriques li&eacute;es au manque d&rsquo;assainissement</li>
            <li>Formation des ma&ccedil;ons locaux &agrave; la construction de latrines</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-3">
          <div class="priority-num">3</div>
          <div class="reco-title">Gestion des d&eacute;chets</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Mise en place de syst&egrave;mes de collecte structur&eacute;s dans les zones d&eacute;ficitaires</li>
            <li>Promotion du compostage et de la valorisation des d&eacute;chets organiques</li>
            <li>Interdiction et sanctions du br&ucirc;lage &agrave; ciel ouvert</li>
            <li>D&eacute;p&ocirc;ts de transit r&eacute;glement&eacute;s et s&eacute;curis&eacute;s</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-4">
          <div class="priority-num">4</div>
          <div class="reco-title">Pollution atmospherique</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Cartographie et surveillance des sources de pollution</li>
            <li>R&eacute;glementation des &eacute;missions industrielles et artisanales</li>
            <li>Promotion de foyeers am&eacute;lior&eacute;s pour r&eacute;duire la fum&eacute;e domestique</li>
            <li>Zones tampon autour des agglom&eacute;rations</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-5">
          <div class="priority-num">5</div>
          <div class="reco-title">Ressources naturelles</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Plans locaux de gestion durable des ressources naturelles</li>
            <li>Reboisement et restauration des terres d&eacute;grad&eacute;es</li>
            <li>Sensibilisation communautaire sur la pr&eacute;servation de l&rsquo;environnement</li>
            <li>Syst&egrave;mes de suivi participatif des ressources</li>
          </ul>
        </div>
      </div>
      <div class="reco-card">
        <div class="reco-card-header priority-6">
          <div class="priority-num">6</div>
          <div class="reco-title">Suivi &amp; Gouvernance</div>
        </div>
        <div class="reco-card-body">
          <ul>
            <li>Syst&egrave;mes de suivi-&eacute;valuation des indicateurs WASH</li>
            <li>Renforcement des capacit&eacute;s des autorit&eacute;s locales</li>
            <li>Implication des communaut&eacute;s dans la gestion des infrastructures</li>
            <li>R&eacute;p&eacute;tition de l&rsquo;enqu&ecirc;te tous les 2 ans pour mesurer les progr&egrave;s</li>
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
    <p>Cette enqu&ecirc;te WASH conduite aupr&egrave;s de <strong>1 000 m&eacute;nages</strong> dans les d&eacute;partements du Goh et de Loh Djiboua en C&ocirc;te d&rsquo;Ivoire met en lumi&egrave;re des d&eacute;fis consid&eacute;rables en mati&egrave;re d&rsquo;eau potable, d&rsquo;assainissement et de gestion environnementale.</p>

    <p>Les r&eacute;sultats les plus pr&eacute;occupants concernent la <strong>qualit&eacute; de l&rsquo;eau</strong> (40% de mauvaise qualit&eacute;), le <strong>faible taux de traitement domestique</strong> (29,1%), et la <strong>pr&eacute;valence de la d&eacute;f&eacute;cation &agrave; l&rsquo;air libre</strong> (33,6%). Ces facteurs combin&eacute;s expliquent en grande partie les 27,1% de m&eacute;nages d&eacute;clarant des maladies li&eacute;es &agrave; l&rsquo;eau.</p>

    <p>Par rapport aux <strong>Objectifs de D&eacute;veloppement Durable (ODD 6)</strong> &mdash; garantir l&rsquo;acc&egrave;s &agrave; l&rsquo;eau potable et &agrave; l&rsquo;assainissement &agrave; tous &mdash; les deux d&eacute;partements se trouvent encore loin des cibles. Des investissements cibl&eacute;s et des programmes de changement de comportement sont indispensables pour r&eacute;duire les in&eacute;galit&eacute;s d&rsquo;acc&egrave;s et prot&eacute;ger la sant&eacute; des populations.</p>

    <p>Les profils des deux d&eacute;partements &eacute;tant tr&egrave;s similaires, les interventions peuvent &ecirc;tre planifi&eacute;es de mani&egrave;re coordonn&eacute;e, tout en tenant compte des sp&eacute;cificit&eacute;s locales de chaque sous-pr&eacute;fecture. Une <strong>r&eacute;p&eacute;tition de l&rsquo;enqu&ecirc;te dans deux ans</strong> permettra de mesurer les progr&egrave;s accomplis et d&rsquo;adapter les strat&eacute;gies d&rsquo;intervention.</p>

    <div class="kpi-grid" style="margin-top:24px">
      <div class="kpi-card red">
        <div class="kpi-num">3</div>
        <div class="kpi-label">Alertes critiques</div>
        <div class="kpi-sub">Eau, assainissement, sant&eacute;</div>
      </div>
      <div class="kpi-card orange">
        <div class="kpi-num">6</div>
        <div class="kpi-label">Recommandations</div>
        <div class="kpi-sub">Class&eacute;es par priorit&eacute;</div>
      </div>
      <div class="kpi-card teal">
        <div class="kpi-num">8</div>
        <div class="kpi-label">Sous-pr&eacute;fectures</div>
        <div class="kpi-sub">2 d&eacute;partements couverts</div>
      </div>
      <div class="kpi-card green">
        <div class="kpi-num">166</div>
        <div class="kpi-label">Variables analys&eacute;es</div>
        <div class="kpi-sub">Couverture exhaustive</div>
      </div>
    </div>
  </div>
</section>

<!-- FOOTER -->
<footer class="report-footer">
  <strong>Rapport WASH &ndash; C&ocirc;te d&rsquo;Ivoire 2026</strong> &bull;
  Enqu&ecirc;te WASH &bull; R&eacute;gions Goh &amp; Loh Djiboua &bull;
  N = 1 000 m&eacute;nages &bull; Janvier&ndash;F&eacute;vrier 2026 &bull;
  G&eacute;n&eacute;r&eacute; le 28 mars 2026
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

out_path = r'C:\Users\ITrebi\Music\PortFolio Upwork\Excel Dashboard\WASH_Rapport.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print("HTML report saved to: " + out_path)
