# -*- coding: utf-8 -*-
# generate_charts_en.py
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import io, base64

PALETTE = ['#0054A6','#00B0B9','#70AD47','#ED7D31','#C00000','#FFC000','#2E75B6','#7030A0','#BDD7EE','#595959']
plt.rcParams['font.family'] = 'Segoe UI'
plt.rcParams['axes.spines.top'] = False
plt.rcParams['axes.spines.right'] = False

def to_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    data = base64.b64encode(buf.read()).decode()
    plt.close(fig)
    return data

def to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    plt.close(fig)
    return buf

def chart_sources():
    labels = ['Trad. well','Improv. well','Borehole','River','Other','SODECI (piped)']
    vals = [275,200,196,167,82,80]
    colors = PALETTE[:6]
    fig,ax = plt.subplots(figsize=(5,4))
    wedges,texts,autotexts = ax.pie(vals, labels=None, colors=colors, autopct='%1.1f%%',
                                     startangle=140, pctdistance=0.75,
                                     wedgeprops=dict(width=0.55, edgecolor='white', linewidth=2))
    for at in autotexts: at.set_fontsize(7); at.set_color('white'); at.set_fontweight('bold')
    ax.legend(wedges, [f'{l} ({v})' for l,v in zip(labels,vals)],
              loc='lower center', bbox_to_anchor=(0.5,-0.15), ncol=2, fontsize=7)
    ax.set_title("Water Supply Sources", fontsize=11, fontweight='bold', pad=10, color='#1F3864')
    return fig

def chart_qualite_dept():
    depts = ['Goh (570)','Loh Djiboua (430)']
    bonne = [118,90]; accept = [221,171]; mauvaise = [231,169]
    x = np.arange(len(depts)); w=0.25
    fig,ax = plt.subplots(figsize=(5,4))
    b1=ax.bar(x-w, bonne,   w, label='Good',       color='#70AD47', edgecolor='white')
    b2=ax.bar(x,   accept,  w, label='Acceptable',  color='#FFC000', edgecolor='white')
    b3=ax.bar(x+w, mauvaise,w, label='Poor',        color='#C00000', edgecolor='white')
    for bars in [b1,b2,b3]:
        for bar in bars:
            h=bar.get_height()
            ax.text(bar.get_x()+bar.get_width()/2, h+3, str(int(h)), ha='center', va='bottom', fontsize=7)
    ax.set_xticks(x); ax.set_xticklabels(depts, fontsize=9)
    ax.set_ylabel('Number of households'); ax.set_ylim(0,280)
    ax.legend(fontsize=8, loc='upper right')
    ax.set_title("Water Quality by Department", fontsize=11, fontweight='bold', color='#1F3864')
    ax.yaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    return fig

def chart_acces():
    cats = ['< 30 min\n(386)', '30-60 min\n(341)', '> 60 min\n(273)']
    vals = [386,341,273]
    colors = ['#70AD47','#FFC000','#C00000']
    fig,ax = plt.subplots(figsize=(5,3.5))
    bars=ax.barh(cats, vals, color=colors, edgecolor='white', height=0.5)
    for bar,v in zip(bars,vals):
        ax.text(bar.get_width()+5, bar.get_y()+bar.get_height()/2,
                f'{v} ({v/10:.1f}%)', va='center', fontsize=8)
    ax.set_xlim(0,480); ax.set_xlabel('Number of households')
    ax.set_title("Water Access Time", fontsize=11, fontweight='bold', color='#1F3864')
    ax.xaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    return fig

def chart_traitement():
    labels = ['Untreated\n(709)','Treated\n(291)']
    vals = [709,291]; colors = ['#C00000','#70AD47']
    fig,ax = plt.subplots(figsize=(4,3.5))
    wedges,texts,autotexts = ax.pie(vals, colors=colors, autopct='%1.1f%%',
                                     startangle=90, pctdistance=0.65,
                                     wedgeprops=dict(edgecolor='white',linewidth=2))
    for at in autotexts: at.set_fontsize(10); at.set_color('white'); at.set_fontweight('bold')
    ax.legend(wedges, labels, loc='lower center', bbox_to_anchor=(0.5,-0.1), ncol=2, fontsize=8)
    ax.set_title("Household Water\nTreatment", fontsize=11, fontweight='bold', color='#1F3864')
    return fig

def chart_maladies():
    cats = ['Diarrhea','Typhoid','Hepatitis A','Cholera']
    g_vals = [115,72,60,50]; l_vals = [87,55,43,34]
    x=np.arange(len(cats)); w=0.35
    fig,ax = plt.subplots(figsize=(5,3.5))
    b1=ax.bar(x-w/2, g_vals, w, label='Goh', color='#0054A6', edgecolor='white')
    b2=ax.bar(x+w/2, l_vals, w, label='Loh Djiboua', color='#00B0B9', edgecolor='white')
    for bars in [b1,b2]:
        for bar in bars:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1, str(int(bar.get_height())),
                    ha='center', va='bottom', fontsize=7)
    ax.set_xticks(x); ax.set_xticklabels(cats, fontsize=8)
    ax.set_ylabel('Number of reported cases')
    ax.legend(fontsize=8); ax.yaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    ax.set_title('Reported Waterborne Diseases', fontsize=11, fontweight='bold', color='#1F3864')
    return fig

def chart_latrines():
    cats = ['Simple latrine','No latrine (OD)','Other','Improv. latrine','Modern toilet']
    vals = [424,336,92,90,58]
    colors = ['#FFC000','#C00000','#595959','#00B0B9','#70AD47']
    fig,ax = plt.subplots(figsize=(5,3.5))
    bars=ax.barh(cats, vals, color=colors, edgecolor='white', height=0.55)
    for bar,v in zip(bars,vals):
        ax.text(bar.get_width()+3, bar.get_y()+bar.get_height()/2,
                f'{v} ({v/10:.1f}%)', va='center', fontsize=7.5)
    ax.set_xlim(0,520); ax.set_xlabel('Number of households')
    ax.set_title("Sanitation Infrastructure Types", fontsize=11, fontweight='bold', color='#1F3864')
    ax.xaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    return fig

def chart_dechets_satisfaction():
    cats = ['Somewhat sat.','Satisfied','Not satisfied','Very satisfied']
    vals = [378,271,242,109]
    colors = ['#ED7D31','#70AD47','#C00000','#00B0B9']
    fig,ax = plt.subplots(figsize=(5,3.5))
    bars=ax.barh(cats, vals, color=colors, edgecolor='white', height=0.5)
    for bar,v in zip(bars,vals):
        ax.text(bar.get_width()+3, bar.get_y()+bar.get_height()/2,
                f'{v} ({v/10:.1f}%)', va='center', fontsize=7.5)
    ax.set_xlim(0,470); ax.set_xlabel('Number of households')
    ax.set_title('Waste Management Satisfaction', fontsize=11, fontweight='bold', color='#1F3864')
    ax.xaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    return fig

def chart_gestion_modes():
    cats = ['Burning','Collection/Rec.','Open dumping','Reuse']
    vals = [499,515,501,287]
    colors = ['#C00000','#70AD47','#ED7D31','#00B0B9']
    fig,ax = plt.subplots(figsize=(5,3.5))
    bars=ax.barh(cats, vals, color=colors, edgecolor='white', height=0.5)
    for bar,v in zip(bars,vals):
        ax.text(bar.get_width()+3, bar.get_y()+bar.get_height()/2,
                f'{v} ({v/10:.1f}%)', va='center', fontsize=7.5)
    ax.set_xlim(0,640); ax.set_xlabel('Number of households')
    ax.set_title('Waste Disposal Methods', fontsize=11, fontweight='bold', color='#1F3864')
    ax.xaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    return fig

def chart_pollution_atmo():
    depts = ['Goh','Loh Djiboua']
    oui_v=[285,209]; non_v=[285,221]
    x=np.arange(len(depts)); w=0.35
    fig,ax = plt.subplots(figsize=(4.5,3.5))
    b1=ax.bar(x-w/2, oui_v, w, label='Pollution reported', color='#C00000', edgecolor='white')
    b2=ax.bar(x+w/2, non_v, w, label='No pollution', color='#70AD47', edgecolor='white')
    for bars in [b1,b2]:
        for bar in bars:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+2,
                    str(int(bar.get_height())), ha='center', va='bottom', fontsize=8)
    ax.set_xticks(x); ax.set_xticklabels(depts)
    ax.set_ylabel('Households'); ax.legend(fontsize=8)
    ax.yaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    ax.set_title('Atmospheric Pollution', fontsize=11, fontweight='bold', color='#1F3864')
    return fig

def chart_qualite_air():
    cats = ['Poor\n(364)','Average\n(361)','Very poor\n(152)','Good\n(123)']
    vals = [364,361,152,123]
    colors = ['#C00000','#ED7D31','#7030A0','#70AD47']
    fig,ax = plt.subplots(figsize=(4.5,3.5))
    bars=ax.bar(cats, vals, color=colors, edgecolor='white', width=0.55)
    for bar,v in zip(bars,vals):
        ax.text(bar.get_x()+bar.get_width()/2, v+4, f'{v/10:.1f}%',
                ha='center', va='bottom', fontsize=8, fontweight='bold')
    ax.set_ylabel('Households'); ax.set_ylim(0,420)
    ax.yaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    ax.set_title("Perceived Air Quality", fontsize=11, fontweight='bold', color='#1F3864')
    return fig

def chart_ressources_dispo():
    cats = ['Scarce\n(384)','Moderately\navailable (335)','Near\ndepletion (160)','Abundant\n(121)']
    vals = [384,335,160,121]
    colors = ['#ED7D31','#FFC000','#C00000','#70AD47']
    fig,ax = plt.subplots(figsize=(5,3.5))
    bars=ax.bar(cats, vals, color=colors, edgecolor='white', width=0.55)
    for bar,v in zip(bars,vals):
        ax.text(bar.get_x()+bar.get_width()/2, v+4, f'{v/10:.1f}%',
                ha='center', va='bottom', fontsize=8, fontweight='bold')
    ax.set_ylabel('Households'); ax.set_ylim(0,440)
    ax.yaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    ax.set_title('Natural Resources Availability', fontsize=11, fontweight='bold', color='#1F3864')
    return fig

def chart_comparaison():
    indicators = ['Good\nwater qual.','Access\n<30min','Treated\nwater','Waterborne\ndis.','No\nlatrine','Atmo.\npollution']
    goh_pct   = [20.7, 37.5, 30.0, 26.7, 33.3, 50.0]
    ldj_pct   = [20.9, 40.0, 28.1, 27.7, 33.3, 48.6]
    x=np.arange(len(indicators)); w=0.35
    fig,ax = plt.subplots(figsize=(7,4))
    b1=ax.bar(x-w/2, goh_pct, w, label='Goh', color='#0054A6', edgecolor='white', alpha=0.9)
    b2=ax.bar(x+w/2, ldj_pct, w, label='Loh Djiboua', color='#00B0B9', edgecolor='white', alpha=0.9)
    ax.set_xticks(x); ax.set_xticklabels(indicators, fontsize=8)
    ax.set_ylabel('Percentage (%)'); ax.set_ylim(0,65)
    ax.legend(fontsize=9); ax.yaxis.grid(True, alpha=0.3); ax.set_axisbelow(True)
    ax.set_title('WASH Indicators Comparison by Department', fontsize=11, fontweight='bold', color='#1F3864')
    for bars in [b1,b2]:
        for bar in bars:
            ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.5,
                    f'{bar.get_height():.1f}%', ha='center', va='bottom', fontsize=6.5)
    return fig

def chart_departements():
    labels=['Goh\n(570 - 57%)','Loh Djiboua\n(430 - 43%)']
    vals=[570,430]; colors=['#0054A6','#00B0B9']
    fig,ax=plt.subplots(figsize=(4,3.5))
    wedges,texts,autotexts=ax.pie(vals,colors=colors,autopct='%1.1f%%',
                                   startangle=90,pctdistance=0.65,
                                   wedgeprops=dict(edgecolor='white',linewidth=3))
    for at in autotexts: at.set_fontsize(12); at.set_color('white'); at.set_fontweight('bold')
    ax.legend(wedges,labels,loc='lower center',bbox_to_anchor=(0.5,-0.08),ncol=1,fontsize=9)
    ax.set_title('Sample Distribution by Department', fontsize=11, fontweight='bold', color='#1F3864')
    return fig

CHARTS = {
    'sources':       chart_sources,
    'qualite_dept':  chart_qualite_dept,
    'acces':         chart_acces,
    'traitement':    chart_traitement,
    'maladies':      chart_maladies,
    'latrines':      chart_latrines,
    'dechets_sat':   chart_dechets_satisfaction,
    'gestion_modes': chart_gestion_modes,
    'pollution':     chart_pollution_atmo,
    'qualite_air':   chart_qualite_air,
    'ressources':    chart_ressources_dispo,
    'comparaison':   chart_comparaison,
    'departements':  chart_departements,
}

if __name__ == '__main__':
    print("Testing chart generation...")
    for name, fn in CHARTS.items():
        fig = fn()
        b64 = to_b64(fig)
        print(f"  Chart '{name}' OK ({len(b64)} bytes b64)")
    print("All charts generated successfully.")
