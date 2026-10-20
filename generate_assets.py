# -*- coding: utf-8 -*-
"""Generate SVG mockups for Microsoft Office Portfolio items."""

excel_svg = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 380" width="100%" height="100%">
  <defs>
    <linearGradient id="exBg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#107C41" />
      <stop offset="100%" stop-color="#0a522a" />
    </linearGradient>
  </defs>
  <!-- Window frame -->
  <rect width="600" height="380" rx="16" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2"/>
  <!-- Ribbon header -->
  <rect width="600" height="44" rx="16" fill="url(#exBg)"/>
  <rect y="24" width="600" height="20" fill="url(#exBg)"/>
  <!-- Window controls -->
  <circle cx="22" cy="22" r="5" fill="#ef4444"/>
  <circle cx="38" cy="22" r="5" fill="#f59e0b"/>
  <circle cx="54" cy="22" r="5" fill="#10b981"/>
  <text x="80" y="27" fill="#ffffff" font-family="'Plus Jakarta Sans', sans-serif" font-size="13" font-weight="600">Microsoft Excel - Business_Financial_Dashboard.xlsx</text>
  
  <!-- Formula bar -->
  <rect x="16" y="54" width="568" height="28" rx="6" fill="#ffffff" stroke="#cbd5e1"/>
  <text x="26" y="72" fill="#107C41" font-family="monospace" font-size="12" font-weight="bold">fx</text>
  <line x1="46" y1="58" x2="46" y2="78" stroke="#e2e8f0"/>
  <text x="54" y="72" fill="#334155" font-family="monospace" font-size="11">=XLOOKUP(Kategori_Produk, Table_Sales[ID], Table_Sales[Revenue]) * 1.15</text>
  
  <!-- KPI Cards -->
  <!-- Card 1 -->
  <rect x="16" y="92" width="132" height="68" rx="8" fill="#ffffff" stroke="#e2e8f0"/>
  <text x="26" y="110" fill="#64748b" font-size="10" font-family="sans-serif">Total Revenue</text>
  <text x="26" y="134" fill="#0f172a" font-size="15" font-weight="bold" font-family="sans-serif">Rp 148.5M</text>
  <rect x="26" y="142" width="45" height="12" rx="4" fill="#dcfce7"/>
  <text x="30" y="151" fill="#16a34a" font-size="9" font-weight="bold" font-family="sans-serif">+14.2% YoY</text>

  <!-- Card 2 -->
  <rect x="160" y="92" width="132" height="68" rx="8" fill="#ffffff" stroke="#e2e8f0"/>
  <text x="170" y="110" fill="#64748b" font-size="10" font-family="sans-serif">Target Tercapai</text>
  <text x="170" y="134" fill="#0f172a" font-size="15" font-weight="bold" font-family="sans-serif">112.4%</text>
  <rect x="170" y="142" width="52" height="12" rx="4" fill="#f3e8ff"/>
  <text x="174" y="151" fill="#7c3aed" font-size="9" font-weight="bold" font-family="sans-serif">EXCEEDED</text>

  <!-- Card 3 -->
  <rect x="304" y="92" width="132" height="68" rx="8" fill="#ffffff" stroke="#e2e8f0"/>
  <text x="314" y="110" fill="#64748b" font-size="10" font-family="sans-serif">Gross Profit</text>
  <text x="314" y="134" fill="#0f172a" font-size="15" font-weight="bold" font-family="sans-serif">Rp 62.8M</text>
  <rect x="314" y="142" width="55" height="12" rx="4" fill="#e0f2fe"/>
  <text x="318" y="151" fill="#0284c7" font-size="9" font-weight="bold" font-family="sans-serif">42.3% Margin</text>

  <!-- Card 4 -->
  <rect x="448" y="92" width="136" height="68" rx="8" fill="#ffffff" stroke="#e2e8f0"/>
  <text x="458" y="110" fill="#64748b" font-size="10" font-family="sans-serif">Efisiensi Operasional</text>
  <text x="458" y="134" fill="#0f172a" font-size="15" font-weight="bold" font-family="sans-serif">94.8%</text>
  <rect x="458" y="142" width="46" height="12" rx="4" fill="#dcfce7"/>
  <text x="462" y="151" fill="#16a34a" font-size="9" font-weight="bold" font-family="sans-serif">A+ Grade</text>

  <!-- Chart & Pivot Area -->
  <rect x="16" y="170" width="320" height="190" rx="8" fill="#ffffff" stroke="#e2e8f0"/>
  <text x="28" y="192" fill="#0f172a" font-size="12" font-weight="bold" font-family="sans-serif">Trend Penjualan Bulanan (Dynamic Chart)</text>
  
  <!-- Bar chart bars -->
  <rect x="36" y="270" width="28" height="65" rx="4" fill="#8b5cf6"/>
  <rect x="74" y="250" width="28" height="85" rx="4" fill="#107C41"/>
  <rect x="112" y="235" width="28" height="100" rx="4" fill="#8b5cf6"/>
  <rect x="150" y="215" width="28" height="120" rx="4" fill="#107C41"/>
  <rect x="188" y="225" width="28" height="110" rx="4" fill="#8b5cf6"/>
  <rect x="226" y="200" width="28" height="135" rx="4" fill="#107C41"/>
  <rect x="264" y="190" width="28" height="145" rx="4" fill="#a855f7"/>

  <line x1="28" y1="340" x2="315" y2="340" stroke="#cbd5e1" stroke-width="1.5"/>
  <text x="42" y="353" fill="#64748b" font-size="9" font-family="sans-serif">Jan</text>
  <text x="80" y="353" fill="#64748b" font-size="9" font-family="sans-serif">Feb</text>
  <text x="118" y="353" fill="#64748b" font-size="9" font-family="sans-serif">Mar</text>
  <text x="156" y="353" fill="#64748b" font-size="9" font-family="sans-serif">Apr</text>
  <text x="194" y="353" fill="#64748b" font-size="9" font-family="sans-serif">Mei</text>
  <text x="232" y="353" fill="#64748b" font-size="9" font-family="sans-serif">Jun</text>
  <text x="270" y="353" fill="#64748b" font-size="9" font-family="sans-serif">Jul</text>

  <!-- Pivot Table Mini Preview -->
  <rect x="348" y="170" width="236" height="190" rx="8" fill="#ffffff" stroke="#e2e8f0"/>
  <rect x="348" y="170" width="236" height="26" rx="8" fill="#f8fafc"/>
  <text x="358" y="188" fill="#0f172a" font-size="11" font-weight="bold" font-family="sans-serif">PivotTable: Analisis Saluran Penjualan</text>

  <text x="358" y="210" fill="#64748b" font-size="9" font-weight="bold">Channel</text>
  <text x="440" y="210" fill="#64748b" font-size="9" font-weight="bold">Volume</text>
  <text x="510" y="210" fill="#64748b" font-size="9" font-weight="bold">Share %</text>
  <line x1="355" y1="216" x2="575" y2="216" stroke="#e2e8f0"/>

  <text x="358" y="232" fill="#334155" font-size="9.5">E-Commerce</text>
  <text x="440" y="232" fill="#334155" font-size="9.5">4,210</text>
  <text x="510" y="232" fill="#16a34a" font-size="9.5" font-weight="600">48.2%</text>

  <text x="358" y="252" fill="#334155" font-size="9.5">Direct Sales</text>
  <text x="440" y="252" fill="#334155" font-size="9.5">2,850</text>
  <text x="510" y="252" fill="#16a34a" font-size="9.5" font-weight="600">32.6%</text>

  <text x="358" y="272" fill="#334155" font-size="9.5">Marketplace</text>
  <text x="440" y="272" fill="#334155" font-size="9.5">1,680</text>
  <text x="510" y="272" fill="#16a34a" font-size="9.5" font-weight="600">19.2%</text>

  <line x1="355" y1="282" x2="575" y2="282" stroke="#cbd5e1" stroke-dasharray="3,3"/>
  <text x="358" y="298" fill="#0f172a" font-size="9.5" font-weight="bold">Total Ringkasan</text>
  <text x="440" y="298" fill="#0f172a" font-size="9.5" font-weight="bold">8,740</text>
  <text x="510" y="298" fill="#7c3aed" font-size="9.5" font-weight="bold">100.0%</text>

  <!-- Tag Pills -->
  <rect x="358" y="320" width="62" height="22" rx="6" fill="#f3e8ff"/>
  <text x="365" y="335" fill="#7c3aed" font-size="8.5" font-weight="bold">VLOOKUP</text>
  <rect x="426" y="320" width="64" height="22" rx="6" fill="#dcfce7"/>
  <text x="433" y="335" fill="#15803d" font-size="8.5" font-weight="bold">XLOOKUP</text>
  <rect x="496" y="320" width="75" height="22" rx="6" fill="#fef3c7"/>
  <text x="503" y="335" fill="#b45309" font-size="8.5" font-weight="bold">PIVOT TABLE</text>
</svg>'''

ppt_svg = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 380" width="100%" height="100%">
  <defs>
    <linearGradient id="pptBg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#D24726" />
      <stop offset="100%" stop-color="#A22D11" />
    </linearGradient>
    <linearGradient id="slideGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#2e1065" />
      <stop offset="100%" stop-color="#581c87" />
    </linearGradient>
  </defs>
  <!-- Window frame -->
  <rect width="600" height="380" rx="16" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2"/>
  <!-- Header -->
  <rect width="600" height="44" rx="16" fill="url(#pptBg)"/>
  <rect y="24" width="600" height="20" fill="url(#pptBg)"/>
  <circle cx="22" cy="22" r="5" fill="#ef4444"/>
  <circle cx="38" cy="22" r="5" fill="#f59e0b"/>
  <circle cx="54" cy="22" r="5" fill="#10b981"/>
  <text x="80" y="27" fill="#ffffff" font-family="'Plus Jakarta Sans', sans-serif" font-size="13" font-weight="600">Microsoft PowerPoint - Business_Pitch_Deck_Strategy.pptx</text>
  
  <!-- Slide Stage -->
  <rect x="24" y="58" width="552" height="265" rx="12" fill="url(#slideGrad)" stroke="#9333ea" stroke-width="1.5"/>
  
  <!-- Slide Content -->
  <rect x="44" y="78" width="105" height="24" rx="12" fill="#8b5cf6" fill-opacity="0.4" stroke="#c084fc" stroke-width="1"/>
  <text x="52" y="94" fill="#e9d5ff" font-size="10" font-weight="bold" font-family="sans-serif">EXECUTIVE DECK</text>
  
  <text x="44" y="132" fill="#ffffff" font-size="22" font-weight="800" font-family="'Plus Jakarta Sans', sans-serif">Digital Business Acceleration 2026</text>
  <text x="44" y="156" fill="#c084fc" font-size="13" font-weight="500" font-family="sans-serif">Strategi Pertumbuhan Berbasis Analisis Data &amp; Desain Interaktif</text>
  
  <!-- 3 Presentation Feature Cards -->
  <rect x="44" y="180" width="155" height="120" rx="10" fill="#3b0764" fill-opacity="0.75" stroke="#a855f7" stroke-width="1"/>
  <circle cx="66" cy="202" r="12" fill="#8b5cf6"/>
  <text x="62" y="206" fill="#ffffff" font-size="12" font-weight="bold">01</text>
  <text x="86" y="206" fill="#f5d0fe" font-size="12" font-weight="bold">Analisis Pasar</text>
  <text x="54" y="228" fill="#d8b4fe" font-size="9">Riset demografi &amp; segmen konsumen dengan akurasi teruji.</text>
  <rect x="54" y="270" width="135" height="18" rx="4" fill="#581c87"/>
  <text x="60" y="282" fill="#e9d5ff" font-size="8.5" font-weight="bold">Proyeksi: +250% Growth</text>

  <rect x="214" y="180" width="155" height="120" rx="10" fill="#3b0764" fill-opacity="0.75" stroke="#a855f7" stroke-width="1"/>
  <circle cx="236" cy="202" r="12" fill="#ec4899"/>
  <text x="232" y="206" fill="#ffffff" font-size="12" font-weight="bold">02</text>
  <text x="256" y="206" fill="#fbcfe8" font-size="12" font-weight="bold">Visual Infografis</text>
  <text x="224" y="228" fill="#f472b6" font-size="9">Slide Master kustom, Morph transition &amp; visual storytelling.</text>
  <rect x="224" y="270" width="135" height="18" rx="4" fill="#581c87"/>
  <text x="230" y="282" fill="#fbcfe8" font-size="8.5" font-weight="bold">Aesthetic Presentation</text>

  <rect x="384" y="180" width="175" height="120" rx="10" fill="#3b0764" fill-opacity="0.75" stroke="#a855f7" stroke-width="1"/>
  <circle cx="406" cy="202" r="12" fill="#06b6d4"/>
  <text x="402" y="206" fill="#ffffff" font-size="12" font-weight="bold">03</text>
  <text x="426" y="206" fill="#cffafe" font-size="12" font-weight="bold">Eksekusi Tim</text>
  <text x="394" y="228" fill="#67e8f9" font-size="9">Roadmap terstruktur, KPI terukur, &amp; komunikasi lintas tim.</text>
  <rect x="394" y="270" width="155" height="18" rx="4" fill="#581c87"/>
  <text x="400" y="282" fill="#cffafe" font-size="8.5" font-weight="bold">Siap Implementasi &amp; Pitch</text>

  <!-- Slide filmstrip at bottom -->
  <rect x="24" y="332" width="552" height="38" rx="6" fill="#e2e8f0"/>
  <rect x="32" y="336" width="45" height="30" rx="4" fill="#581c87" stroke="#8b5cf6" stroke-width="2"/>
  <text x="48" y="354" fill="#ffffff" font-size="10" font-weight="bold">1</text>
  <rect x="84" y="336" width="45" height="30" rx="4" fill="#ffffff" stroke="#cbd5e1"/>
  <text x="102" y="354" fill="#64748b" font-size="10">2</text>
  <rect x="136" y="336" width="45" height="30" rx="4" fill="#ffffff" stroke="#cbd5e1"/>
  <text x="154" y="354" fill="#64748b" font-size="10">3</text>
  <rect x="188" y="336" width="45" height="30" rx="4" fill="#ffffff" stroke="#cbd5e1"/>
  <text x="206" y="354" fill="#64748b" font-size="10">4</text>
  <text x="250" y="355" fill="#475569" font-size="11" font-weight="600" font-family="sans-serif">Slide 1 dari 18 • Dirancang dengan Master Slide &amp; Tipografi Elegan</text>
</svg>'''

word_svg = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 600 380" width="100%" height="100%">
  <defs>
    <linearGradient id="wdBg" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#185ABD" />
      <stop offset="100%" stop-color="#103f82" />
    </linearGradient>
  </defs>
  <!-- Window frame -->
  <rect width="600" height="380" rx="16" fill="#f8fafc" stroke="#cbd5e1" stroke-width="2"/>
  <!-- Header -->
  <rect width="600" height="44" rx="16" fill="url(#wdBg)"/>
  <rect y="24" width="600" height="20" fill="url(#wdBg)"/>
  <circle cx="22" cy="22" r="5" fill="#ef4444"/>
  <circle cx="38" cy="22" r="5" fill="#f59e0b"/>
  <circle cx="54" cy="22" r="5" fill="#10b981"/>
  <text x="80" y="27" fill="#ffffff" font-family="'Plus Jakarta Sans', sans-serif" font-size="13" font-weight="600">Microsoft Word - Laporan_Studi_Kelayakan_Bisnis.docx</text>
  
  <!-- Ribbon Toolbar -->
  <rect x="16" y="52" width="568" height="26" rx="6" fill="#ffffff" stroke="#e2e8f0"/>
  <text x="26" y="68" fill="#185ABD" font-size="11" font-weight="bold">Heading 1</text>
  <text x="96" y="68" fill="#64748b" font-size="11">Heading 2</text>
  <text x="166" y="68" fill="#64748b" font-size="11">Daftar Isi Otomatis</text>
  <text x="280" y="68" fill="#64748b" font-size="11">Mail Merge</text>
  <text x="360" y="68" fill="#64748b" font-size="11">Format Sitasi &amp; Referensi</text>

  <!-- Document Page Mockup -->
  <rect x="80" y="86" width="440" height="280" rx="4" fill="#ffffff" stroke="#cbd5e1" stroke-width="1"/>
  
  <!-- Doc Header -->
  <text x="110" y="116" fill="#185ABD" font-size="16" font-weight="800" font-family="'Plus Jakarta Sans', sans-serif">LAPORAN ANALISIS STUDI KELAYAKAN</text>
  <text x="110" y="132" fill="#475569" font-size="11" font-weight="600">Program Inkubasi Bisnis Digital &amp; Optimasi Operasional</text>
  <line x1="110" y1="140" x2="490" y2="140" stroke="#185ABD" stroke-width="2"/>
  
  <!-- Table of Contents badge -->
  <rect x="110" y="150" width="380" height="65" rx="6" fill="#f8fafc" stroke="#e2e8f0"/>
  <text x="122" y="168" fill="#1e293b" font-size="10" font-weight="bold">DAFTAR ISI (AUTOMATED)</text>
  
  <text x="122" y="184" fill="#334155" font-size="9">1. Ringkasan Eksekutif ................................................................ hal 02</text>
  <text x="122" y="198" fill="#334155" font-size="9">2. Metodologi Analisis &amp; Pengumpulan Data .......................... hal 07</text>
  <text x="122" y="210" fill="#334155" font-size="9">3. Evaluasi Finansial &amp; Proyeksi Arus Kas ................................ hal 14</text>
  
  <!-- Paragraph lines -->
  <text x="110" y="235" fill="#0f172a" font-size="11" font-weight="bold">1. Ringkasan Eksekutif</text>
  <rect x="110" y="245" width="380" height="6" rx="3" fill="#e2e8f0"/>
  <rect x="110" y="256" width="360" height="6" rx="3" fill="#e2e8f0"/>
  <rect x="110" y="267" width="340" height="6" rx="3" fill="#e2e8f0"/>
  <rect x="110" y="278" width="375" height="6" rx="3" fill="#e2e8f0"/>
  
  <rect x="110" y="295" width="130" height="20" rx="4" fill="#eff6ff" stroke="#bfdbfe"/>
  <text x="116" y="308" fill="#1d4ed8" font-size="8.5" font-weight="bold">Standar Formal Dikti / AP</text>
  
  <rect x="246" y="295" width="130" height="20" rx="4" fill="#f3e8ff" stroke="#e9d5ff"/>
  <text x="252" y="308" fill="#7e22ce" font-size="8.5" font-weight="bold">SOP &amp; Proposal Bisnis Rapi</text>
</svg>'''

with open('assets/excel_dashboard.svg', 'w', encoding='utf-8') as f:
    f.write(excel_svg)

with open('assets/ppt_deck.svg', 'w', encoding='utf-8') as f:
    f.write(ppt_svg)

with open('assets/word_report.svg', 'w', encoding='utf-8') as f:
    f.write(word_svg)

print("All 3 SVG mockups generated successfully!")
