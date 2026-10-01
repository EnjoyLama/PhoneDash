"""
Dashboard — Impact Écologique des Smartphones
v3.0 : nouveau fichier Excel, onglets réordonnés, graphes corrigés,
        flèches carte améliorées, onglet Sources & Méthodologie
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os, re, math
from shiny import App, ui, render, reactive
from shinywidgets import output_widget, render_widget

# ╔══════════════════════════════════════════════════════════════════╗
# ║  CONSTANTES                                                      ║
# ╚══════════════════════════════════════════════════════════════════╝
CO2_VOITURE_G_KM   = 130
CO2_AVION_G_KM     = 180
CO2_STREAMING_G_H  = 36
CO2_ARBRE_G_AN     = 22_000
ASSEMBLY_LAT, ASSEMBLY_LON = 22.54, 114.06   # Shenzhen (Foxconn)
ASSEMBLY_LABEL = "Shenzhen, Chine"

# ╔══════════════════════════════════════════════════════════════════╗
# ║  PARSING                                                         ║
# ╚══════════════════════════════════════════════════════════════════╝
def parse_to_g(val):
    """Convertit texte → grammes (float). Gère mg/g/kg, ~, >, plages."""
    if pd.isna(val): return None
    s = str(val).lower().strip()
    if s in ("nulle","inconnu","nan","none",""): return None
    s = s.replace("~","").replace(">","").replace("<","")
    s = re.sub(r"\([^)]+\)", " ", s)
    pairs = re.findall(r"([\d]+(?:[.,]\d+)?)\s*(mg|kg|g\b)", s)
    if pairs:
        vals_g = []
        for num_str, unit in pairs:
            n = float(num_str.replace(",","."))
            if unit == "mg":   vals_g.append(n / 1_000)
            elif unit == "kg": vals_g.append(n * 1_000)
            else:              vals_g.append(n)
        all_nums = re.findall(r"[\d]+(?:[.,]\d+)?", s)
        matched  = {p[0] for p in pairs}
        last_u   = pairs[-1][1]
        for raw in all_nums:
            if raw not in matched:
                n = float(raw.replace(",","."))
                if last_u == "mg":   vals_g.append(n / 1_000)
                elif last_u == "kg": vals_g.append(n * 1_000)
                else:                vals_g.append(n)
        return sum(vals_g) / len(vals_g)
    nums = re.findall(r"[\d]+(?:[.,]\d+)?", s)
    if not nums: return None
    n = sum(float(x.replace(",",".")) for x in nums) / len(nums)
    if "kg" in s: return n * 1_000
    if "mg" in s: return n / 1_000
    return n

# ╔══════════════════════════════════════════════════════════════════╗
# ║  CHARGEMENT DES DONNÉES                                          ║
# ╚══════════════════════════════════════════════════════════════════╝
# Colonnes du nouveau fichier Créa_dig_llama_end.xlsx (0-indexé) :
# 0  Element            – nom
# 1  Proportion (%)
# 2  Grammage (g)       – g de cet élément dans un tél 200g
# 3  Roche / 1g         – col intermédiaire, ignorée
# 4  Roche totale (g)   – déjà calculé : kg roche × 1 000 000 × grammage
# 5  CO2 raf (g)        – TOTAL pour le grammage du tél (pas per-g)
# 6  CO2 ext (g)        – TOTAL pour le grammage du tél
# 7  CO2 asm/lifecycle (g) – TOTAL = facteur 200 kg/kg × grammage
# 8  Pays extraction
# 10 Nombre mines
# 12 Raffinerie
# 13 Nombre raffineries
# 15 Partie téléphone
# 17 Référence

def load_data():
    base = os.path.dirname(os.path.abspath(__file__))
    fp   = os.path.join(base, "Créa_dig_llama_end.xlsx")
    try:
        df = pd.read_excel(fp, engine="openpyxl")
    except Exception as e:
        print(f"[ERREUR] {e}")
        return pd.DataFrame()

    df.columns = [str(c).strip() for c in df.columns]
    # Renomme les colonnes par position pour robustesse
    col_map = {
        df.columns[0]:  "Element",
        df.columns[1]:  "Proportion_pct",
        df.columns[2]:  "Grammage_g",
        df.columns[3]:  "_roche_1g_skip",
        df.columns[4]:  "Roche_totale_g",    # déjà calculé en grammes
        df.columns[5]:  "_co2_raf_raw",
        df.columns[6]:  "_co2_ext_raw",
        df.columns[7]:  "_co2_asm_raw",
        df.columns[8]:  "Pays_extraction",
        df.columns[9]:  "Production_tan",
        df.columns[10]: "Nb_mines",
        df.columns[11]: "Methode_extraction",
        df.columns[12]: "Raffinerie",
        df.columns[13]: "Nb_raffineries",
        df.columns[14]: "Autres",
        df.columns[15]: "Partie_telephone",
        df.columns[16]: "_asm_eq_co2_kg",
        df.columns[17]: "Reference",
    }
    df = df.rename(columns=col_map)

    # Filtrage
    df = df.dropna(subset=["Element"])
    df["Element"] = df["Element"].astype(str).str.strip()
    df = df[~df["Element"].str.lower().isin(["nan","none","","nat"])].copy()

    # Numériques directs
    df["Proportion_pct"] = pd.to_numeric(df["Proportion_pct"], errors="coerce").fillna(0)
    df["Grammage_g"]     = pd.to_numeric(df["Grammage_g"],     errors="coerce").fillna(0)

    # Roche totale (g → kg)
    df["Roche_totale_g"]  = pd.to_numeric(df["Roche_totale_g"], errors="coerce").fillna(0)
    df["Roche_totale_kg"] = df["Roche_totale_g"] / 1_000
    # Roche par gramme d'élément (kg/g) pour le scatter
    df["Roche_par_g_kg"] = df.apply(
        lambda r: r["Roche_totale_kg"] / r["Grammage_g"] if r["Grammage_g"] > 0 else None,
        axis=1
    )

    # CO₂ : les valeurs Excel sont déjà les TOTAUX pour le grammage du tél
    for new_col, raw_col in [
        ("CO2_ext_g",  "_co2_ext_raw"),
        ("CO2_raf_g",  "_co2_raf_raw"),
        ("CO2_asm_g",  "_co2_asm_raw"),
    ]:
        df[new_col] = df[raw_col].apply(parse_to_g)

    df["CO2_total_g"] = df[["CO2_ext_g","CO2_raf_g","CO2_asm_g"]].sum(axis=1, skipna=True)

    return df


df = load_data()

elements_uniques = sorted([
    str(x).strip() for x in df["Element"].unique()
    if str(x).strip().lower() not in ["nan","none","","nat"]
]) if not df.empty else []

# ╔══════════════════════════════════════════════════════════════════╗
# ║  GÉOGRAPHIE                                                      ║
# ╚══════════════════════════════════════════════════════════════════╝
PAYS_FR_EN = {
    "Australie":"Australia","Chili":"Chile","Argentine":"Argentina","Chine":"China",
    "Brésil":"Brazil","Russie":"Russia","Bolivie":"Bolivia","Pérou":"Peru","Perou":"Peru",
    "Mexique":"Mexico","Afrique du Sud":"South Africa","États-Unis":"United States",
    "USA":"United States","Canada":"Canada","Indonésie":"Indonesia","Inde":"India",
    "Japon":"Japan","Corée du Sud":"South Korea","RD Congo":"Democratic Republic of the Congo",
    "République Démocratique du Congo":"Democratic Republic of the Congo",
    "Congo":"Democratic Republic of the Congo","Zambie":"Zambia","Maroc":"Morocco",
    "Turquie":"Turkey","Kazakhstan":"Kazakhstan","Ouzbékistan":"Uzbekistan",
    "Allemagne":"Germany","France":"France","Philippines":"Philippines",
    "Vietnam":"Vietnam","Myanmar":"Myanmar","Birmanie":"Myanmar",
    "Thaïlande":"Thailand","Madagascar":"Madagascar","Mali":"Mali",
    "Guinée":"Guinea","Ghana":"Ghana","Zimbabwe":"Zimbabwe","Belgique":"Belgium",
    "Suisse":"Switzerland","Nouvelle-Calédonie":"New Caledonia",
    "Arabie Saoudite":"Saudi Arabia","Mongolie":"Mongolia","Namibie":"Namibia",
    "Norvège":"Norway","Finlande":"Finland","Suède":"Sweden",
    "Bermude":"Bermuda","Bermudes":"Bermuda","Colombie":"Colombia",
    "Papouasie-Nouvelle-Guinée":"Papua New Guinea","Pérou":"Peru",
    "Péru":"Peru",   # variante orthographique présente dans le fichier source (Sn)
}
COUNTRY_COORDS = {
    "Argentina":(-38.42,-63.62),"Australia":(-25.27,133.78),"Belgium":(50.50,4.47),
    "Bolivia":(-16.29,-63.59),"Brazil":(-14.24,-51.93),"Canada":(56.13,-106.35),
    "Chile":(-35.68,-71.54),"China":(35.86,104.20),"Colombia":(4.57,-74.30),
    "Democratic Republic of the Congo":(-4.04,21.76),"Ecuador":(-1.83,-78.18),
    "Finland":(61.92,25.75),"France":(46.23,2.21),"Germany":(51.17,10.45),
    "Ghana":(7.95,-1.02),"Guinea":(11.80,-15.18),"India":(20.59,78.96),
    "Indonesia":(-0.79,113.92),"Japan":(36.20,138.25),"Kazakhstan":(48.02,66.92),
    "Madagascar":(-18.77,46.87),"Malaysia":(4.21,101.98),"Mali":(17.57,-3.99),
    "Mexico":(23.63,-102.55),"Mongolia":(46.86,103.85),"Morocco":(31.79,-7.09),
    "Myanmar":(21.92,95.96),"Namibia":(-22.96,18.49),"New Caledonia":(-20.90,165.62),
    "Norway":(60.47,8.47),"Papua New Guinea":(-6.31,143.96),"Peru":(-9.19,-75.02),
    "Philippines":(12.88,121.77),"Russia":(61.52,105.32),"Saudi Arabia":(23.89,45.08),
    "South Africa":(-30.56,22.94),"South Korea":(35.91,127.77),"Sweden":(60.13,18.64),
    "Switzerland":(46.82,8.23),"Thailand":(15.87,100.99),"Turkey":(38.96,35.24),
    "United States":(37.09,-95.71),"Uzbekistan":(41.38,64.59),"Vietnam":(14.06,108.28),
    "Zambia":(-13.13,27.85),"Zimbabwe":(-19.02,29.15),"Bermuda":(32.32,-64.76),
}

# ╔══════════════════════════════════════════════════════════════════╗
# ║  CSS                                                             ║
# ╚══════════════════════════════════════════════════════════════════╝
CSS = """
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Space+Grotesk:wght@400;600&display=swap');
*,*::before,*::after{box-sizing:border-box;}
body{font-family:'DM Sans',sans-serif;background:#0f1117;color:#e4e4f0;margin:0;}
h2.app-title{font-family:'Space Grotesk',sans-serif;font-size:1.45rem;font-weight:600;
  color:#f0f0ff;padding:14px 20px 10px;border-bottom:1px solid #1e2030;margin:0;background:#12141f;}
.sidebar{background:#12141f !important;border-right:1px solid #1e2030 !important;}
.bslib-sidebar-layout>.main{background:#0f1117 !important;}

label,.form-label,.control-label{color:#c8c8e0 !important;font-size:0.88em !important;font-weight:500 !important;}
.selectize-control.single .selectize-input,.selectize-control.multi .selectize-input{
  background:#1a1d2e !important;border:1px solid #2e3150 !important;border-radius:8px !important;
  color:#e0e0f4 !important;box-shadow:none !important;padding:7px 10px !important;}
.selectize-control.single .selectize-input.focus{border-color:#5a3e8a !important;}
.selectize-dropdown,.selectize-dropdown.single{background:#1a1d2e !important;
  border:1px solid #2e3150 !important;border-radius:8px !important;color:#e0e0f4 !important;}
.selectize-dropdown .option{color:#c8c8e0 !important;padding:7px 12px !important;}
.selectize-dropdown .option:hover,.selectize-dropdown .option.active{background:#2a2d40 !important;color:#f0f0ff !important;}
.selectize-control.single .selectize-input:after{border-top-color:#8080a0 !important;}

.nav-link{color:#8888aa !important;font-weight:500;}
.nav-link:hover{color:#c8c8e0 !important;background:#1a1d2e !important;}
.nav-link.active{color:#f0f0ff !important;background:#1e2236 !important;border-bottom:2px solid #7c5cbf !important;}
.card,.bslib-card{background:#12141f !important;border:1px solid #1e2030 !important;}
.card-header{background:#12141f !important;border-color:#1e2030 !important;color:#c8c8e0 !important;}
hr{border-color:#1e2030 !important;}
::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:#0f1117;}
::-webkit-scrollbar-thumb{background:#2e3150;border-radius:3px;}

.kpi-wrap{display:flex;flex-direction:column;gap:8px;margin-top:4px;}
.kpi-card{border-radius:10px;padding:12px 14px;display:flex;flex-direction:column;gap:2px;}
.kpi-val{font-family:'Space Grotesk',sans-serif;font-size:1.35em;font-weight:600;color:#fff;}
.kpi-lbl{font-size:0.72em;color:rgba(255,255,255,0.82);letter-spacing:0.02em;}
.kpi-co2 {background:linear-gradient(135deg,#5a3e8a,#7c5cbf);}
.kpi-rock{background:linear-gradient(135deg,#8a4e1a,#c47c30);}
.kpi-car {background:linear-gradient(135deg,#1a6b5a,#2eaf8e);}

.equiv-row{display:flex;gap:10px;flex-wrap:wrap;margin:12px 0;}
.equiv-card{flex:1;min-width:130px;border-radius:12px;padding:18px 12px;
  text-align:center;display:flex;flex-direction:column;align-items:center;gap:4px;}
.equiv-icon{font-size:1.9em;}
.equiv-val{font-family:'Space Grotesk',sans-serif;font-size:1.25em;font-weight:600;color:#fff;}
.equiv-lbl{font-size:0.72em;color:rgba(255,255,255,0.88);line-height:1.3;}

.section-note{color:#8888aa;font-style:italic;font-size:0.85em;margin-bottom:12px;}
.nav-note{background:#1a1d2e;border-left:3px solid #5a3e8a;padding:8px 14px;
  border-radius:0 8px 8px 0;font-size:0.83em;color:#b0b0cc;margin-bottom:14px;}
.map-legend{background:#1a1d2e;border:1px solid #2e3150;border-radius:10px;
  padding:10px 14px;font-size:0.82em;color:#c8c8e0;margin-top:8px;line-height:1.8;}
.map-legend b{color:#e4e4f0;}

/* Sources tab */
.src-section{margin-bottom:22px;}
.src-section h4{font-family:'Space Grotesk',sans-serif;font-size:1.0em;font-weight:600;
  color:#c8a0ff;border-bottom:1px solid #2a2d40;padding-bottom:6px;margin-bottom:10px;}
.src-ref{font-size:0.82em;color:#a0a0c0;line-height:1.6;margin:4px 0 4px 12px;padding-left:8px;
  border-left:2px solid #2e3150;}
.src-ref a{color:#7c9fff;}
.meth-box{background:#1a1d2e;border-radius:10px;padding:14px 18px;margin-bottom:14px;font-size:0.85em;}
.meth-box h4{color:#f0c060;font-size:0.95em;font-weight:600;margin:0 0 8px 0;}
.meth-box p,.meth-box ul{color:#c8c8e0;line-height:1.65;margin:4px 0;}
.meth-box code{background:#12141f;padding:2px 6px;border-radius:4px;
  font-family:monospace;color:#a0f0b0;font-size:0.9em;}
"""

# ╔══════════════════════════════════════════════════════════════════╗
# ║  CONTENU SOURCES (depuis PDF méthodologie)                       ║
# ╚══════════════════════════════════════════════════════════════════╝
SOURCES_HTML = """
<div style='padding:4px 8px;'>

<div class='meth-box'>
  <h4>📱 Téléphone de référence</h4>
  <p>Smartphone fictif de <b>200 g</b>. Proportions massiques issues principalement de
  <i>Gómez et al. (2023)</i>, complétées par <i>Jenness et al. (2016)</i> et
  <i>Compound Interest (2014)</i>. Ce smartphone et ces études sont basées sur le Fairphone 5
  et l'iPhone 15 Pro.</p>
  <ul>
    <li>Plastique (PC/ABS) : 35 % — 70 g</li>
    <li>Cuivre (Cu) : 12 % — 24 g</li>
    <li>Cobalt (Co) &amp; Lithium (Li) : 2,5 % chacun — 5 g</li>
    <li>Nickel (Ni) : 2 % — 4 g · Étain (Sn) : 0,8 % — 1,6 g</li>
    <li>Zinc (Zn) : 0,4 % — 0,8 g · Terres rares (REE) : 0,1 % — 0,2 g</li>
    <li>Indium (In) : 0,006 % — 0,012 g · Or (Au) : 0,017 % — 0,034 g</li>
    <li>Gallium (Ga) : 0,001 % — 0,002 g</li>
  </ul>
</div>

<div class='meth-box'>
  <h4>📐 Indicateurs et équations</h4>
  <p><b>CO₂ extraction &amp; raffinage</b> — intensité carbone × grammage dans le téléphone :<br>
  <code>CO₂_i (g) = Intensité_i (kg CO₂/kg) × grammage_i (g)</code></p>
  <p><b>CO₂ cycle de vie complet (colonne assemblage)</b> — facteur Fairphone 5 LCA :<br>
  <code>CO₂_lifecycle_i (g) = 200 (kg CO₂/kg) × grammage_i (g)</code><br>
  Source : Fairphone (2024) — 42,1 kg CO₂eq pour 212 g ≈ 198,6 kg/kg.</p>
  <p><b>Roche extraite par élément</b> :<br>
  <code>Roche_i (g) = rapport_stérilité_i (g/g) × grammage_i (g)</code></p>
  <p><b>Rapport de stérilité</b> : g de roche extraites par gramme de métal utile récupéré.
  Exemple : Or → ~1,5 t roche / g Au, soit 1 500 000 g/g.</p>
</div>

<div class='meth-box'>
  <h4>⚠️ Approximations et limites</h4>
  <ul>
    <li>Mélange saumure/roche dure pour Li : <b>moyenne</b> des deux voies.</li>
    <li>Les terres rares (REE, 17 éléments) sont agrégées en un seul indicateur.</li>
    <li>Assemblage (lieu) : Shenzhen simplifié comme point unique de fabrication mondiale.</li>
    <li><b>Pollutions non couvertes :</b> seules les émissions atmosphériques (CO₂) et l'impact
    solide (roche extraite) sont quantifiés ici. Les pollutions liquides et chimiques (rejets
    d'acides, métaux lourds dans l'eau, drainage minier acide, etc.) ne sont <b>pas</b> prises en
    compte, alors qu'elles constituent souvent un impact environnemental majeur de l'extraction
    minière.</li>
    <li><b>Transport entre étapes non modélisé :</b> aucun terme de CO₂ n'est calculé pour le
    trajet entre les sites d'extraction, de raffinage et d'assemblage (pas de distance × facteur
    d'émission par étape). La carte « Chaîne de Valeur » illustre l'origine géographique des
    matériaux, mais les flèches ne représentent pas une quantité de CO₂ liée au transport.
    Les valeurs de CO₂ utilisées proviennent d'intensités carbone par matériau tirées de la
    littérature (kg CO₂/kg), dont le périmètre exact (incluant ou non une part de logistique
    interne au site) dépend de chaque étude source — voir bibliographie.</li>
  </ul>
</div>

<div class='src-section'>
  <h4>📚 Général — Composition et ACV smartphones</h4>
  <p class='src-ref'>Alonso, E., Brioche, A. S., Schulte, R. F., Trimmer, L. M., Kim, J.-E., Gulley, A. L., &amp; Pineault, D. G. (2025). <i>World minerals outlook — Cobalt, gallium, helium, lithium, magnesium, palladium, platinum, and titanium through 2029</i> (Scientific Investigations Report 2025–5021). U.S. Geological Survey.
  <a href='https://doi.org/10.3133/sir20255021' target='_blank'>doi.org</a></p>
  <p class='src-ref'>Apple Inc. (2023, 12 septembre). <i>Product environmental report: iPhone 15 Pro and iPhone 15 Pro Max.</i>
  <a href='https://www.apple.com/environment/pdf/products/iphone/iPhone_15_Pro_and_iPhone_15_Pro_Max_Sept2023.pdf' target='_blank'>apple.com</a></p>
  <p class='src-ref'>Chung, J., Xun, S., &amp; Textoris, S. D. (2025). <i>Global maps of critical mineral production in 2023</i> (Fact Sheet 2025–3038). U.S. Geological Survey.
  <a href='https://doi.org/10.3133/fs20253038' target='_blank'>doi.org</a></p>
  <p class='src-ref'>Compound Interest. (2014, 19 février). The chemical elements of a smartphone.
  <a href='https://www.compoundchem.com/2014/02/19/the-chemical-elements-of-a-smartphone/' target='_blank'>compoundchem.com</a></p>
  <p class='src-ref'>European Commission, Joint Research Centre. (2020). <i>Critical raw materials for strategic technologies and sectors in the EU: A foresight study.</i>
  <a href='https://rmis.jrc.ec.europa.eu/uploads/CRMs_for_Strategic_Technologies_and_Sectors_in_the_EU_2020.pdf' target='_blank'>jrc.ec.europa.eu</a></p>
  <p class='src-ref'>Fairphone. (2024). <i>Fairphone 5 life cycle assessment report.</i>
  <a href='https://www.fairphone.com/wp-content/uploads/2024/09/Fairphone5_LCA_Report_2024.pdf' target='_blank'>fairphone.com</a></p>
  <p class='src-ref'>Gómez, M., Grimes, S., Qian, Y., Feng, Y., &amp; Fowler, G. (2023). Critical and strategic metals in mobile phones: A detailed characterisation of multigenerational waste mobile phones and the economic drivers for recovery of metal value. <i>Journal of Cleaner Production, 419</i>, 138099.
  <a href='https://doi.org/10.1016/j.jclepro.2023.138099' target='_blank'>doi.org</a></p>
  <p class='src-ref'>International Energy Agency. (2023). <i>Critical minerals market review 2023.</i>
  <a href='https://www.iea.org/reports/critical-minerals-market-review-2023' target='_blank'>iea.org</a></p>
  <p class='src-ref'>Jenness, J. E., Ober, J. A., Wilkins, A. M., &amp; Gambogi, J. (2016). <i>A world of minerals in your mobile device</i> (General Information Product 167). U.S. Geological Survey.
  <a href='https://doi.org/10.3133/gip167' target='_blank'>doi.org</a></p>
  <p class='src-ref'>L'Élémentarium. (s.d.).
  <a href='https://lelementarium.fr/' target='_blank'>lelementarium.fr</a></p>
  <p class='src-ref'>Nuss, P., &amp; Eckelman, M. J. (2014). Life cycle assessment of metals: A scientific synthesis. <i>PLOS ONE, 9</i>(7), e101298.
  <a href='https://doi.org/10.1371/journal.pone.0101298' target='_blank'>doi.org</a></p>
  <p class='src-ref'>SFA Oxford. (2026). Critical minerals in smartphones.
  <a href='https://www.sfa-oxford.com/knowledge-and-insights/critical-minerals-in-low-carbon-and-future-technologies/critical-minerals-in-electronics/critical-minerals-in-smartphones/' target='_blank'>sfa-oxford.com</a></p>
  <p class='src-ref'>U.S. Geological Survey. (2024). <i>Mineral commodity summaries 2024.</i> U.S. Department of the Interior.
  <a href='https://pubs.usgs.gov/periodicals/mcs2024/mcs2024.pdf' target='_blank'>usgs.gov</a></p>
  <p class='src-ref'>U.S. Geological Survey. (2025). <i>Mineral commodity summaries 2025.</i>
  <a href='https://doi.org/10.3133/mcs2025' target='_blank'>doi.org</a></p>
</div>

<div class='src-section'>
  <h4>🔘 Indium (In)</h4>
  <p class='src-ref'>MineralInfo. (2024, 29 mars). Indium (In).
  <a href='https://www.mineralinfo.fr/fr/substance/indium' target='_blank'>mineralinfo.fr</a></p>
  <p class='src-ref'>SCRREEN. (2020). <i>Indium: Critical raw material factsheet.</i>
  <a href='https://scrreen.eu/wp-content/uploads/2023/01/INDIUM_CRM_2020_Factsheets_critical_Final.pdf' target='_blank'>scrreen.eu</a></p>
</div>

<div class='src-section'>
  <h4>🔋 Lithium (Li)</h4>
  <p class='src-ref'>Arbor. (2025, 23 mai). Lithium environmental impact: Calculated and explained.
  <a href='https://www.arbor.eco/blog/lithium-environmental-impact' target='_blank'>arbor.eco</a></p>
  <p class='src-ref'>Filipenco, D. (2025, 29 avril). Global lithium industry: Five major lithium producing countries. Development Aid.
  <a href='https://www.developmentaid.org/news-stream/post/170661/five-major-lithium-producing-countries' target='_blank'>developmentaid.org</a></p>
  <p class='src-ref'>InvestChile. (2023, 3 janvier). Chilean lithium: The smallest carbon footprint on the planet.
  <a href='https://blog.investchile.gob.cl/chilean-lithium-smallest-carbon-footprint-on-the-planet' target='_blank'>investchile.gob.cl</a></p>
  <p class='src-ref'>Lagos, G., Cifuentes, L., Peters, D., Castro, L., &amp; Valdés, J. M. (2024). Carbon footprint and water inventory of the production of lithium in the Atacama Salt Flat, Chile. <i>Environmental Challenges, 16</i>, 100962.
  <a href='https://doi.org/10.1016/j.envc.2024.100962' target='_blank'>doi.org</a></p>
</div>

<div class='src-section'>
  <h4>🥇 Or (Au)</h4>
  <p class='src-ref'>Académie AuCOFFRE. (2022, 28 septembre). Extraction de l'or : un modèle à bout de souffle.
  <a href='https://www.aucoffre.com/academie/extraction-or-modele-a-bout-de-souffle/' target='_blank'>aucoffre.com</a></p>
  <p class='src-ref'>Santika, E. F. (2025, 4 février). Indonesia to join ranks of world's largest gold producers in 2024.
  <a href='https://databoks.katadata.co.id/en/market/statistics/67a256ff68874/indonesia-to-join-ranks-of-worlds-largest-gold-producers-in-2024' target='_blank'>databoks.katadata.co.id</a></p>
  <p class='src-ref'>MKS PAMP. (2024). <i>Carbon footprints of large gold bars and 1kg gold bars</i> — Provenance Gold, ESG Report 2023–2024. Tableau 8 : Raw materials (Gold) = 4 201 kg CO₂eq/kg.
  <a href='https://www.mkspamp.com/sites/mksandpamp/files/2024-10/MKS%20PAMP%2023-24%20-%20Provenance%20Gold%20Kilo%20and%20Large%20bars%20-%20PER_0.pdf' target='_blank'>mkspamp.com</a></p>
  <p class='src-ref'>World Gold Council. (2026, 9 juin). Gold production by country.
  <a href='https://www.gold.org/goldhub/data/gold-production-by-country' target='_blank'>gold.org</a></p>
</div>

<div class='src-section'>
  <h4>⚪ Étain (Sn)</h4>
  <p class='src-ref'>Santika, E. F. (2024, 4 avril). Indonesia ranks among the world's largest tin producers in 2023.
  <a href='https://databoks.katadata.co.id/en/mining/statistics/f4aa1d72410710d/indonesia-ranks-among-the-worlds-largest-tin-producers-in-2023' target='_blank'>databoks.katadata.co.id</a></p>
  <p class='src-ref'>Nasdaq. (2024, 7 juin). Tin stocks: 10 biggest producers 2024.
  <a href='https://www.nasdaq.com/articles/tin-stocks-10-biggest-producers-2024' target='_blank'>nasdaq.com</a></p>
  <p class='src-ref'>U.S. Geological Survey. (s.d.). Tin statistics and information. National Minerals Information Center.
  <a href='https://www.usgs.gov/centers/national-minerals-information-center/tin-statistics-and-information' target='_blank'>usgs.gov</a></p>
</div>

<div class='src-section'>
  <h4>🔘 Zinc (Zn)</h4>
  <p class='src-ref'>Nexa Resources. (2022). <i>Greenhouse gas emissions inventory 2020.</i>
  <a href='https://www.nexaresources.com/wp-content/uploads/2022/07/Inventario_GEE_2020-en.pdf' target='_blank'>nexaresources.com</a></p>
  <p class='src-ref'>Ressources naturelles Canada. (2024). Faits sur le zinc. Gouvernement du Canada.
  <a href='https://ressourcesnaturelles.canada.ca/mineraux-exploitation-miniere/donnees-statistiques-analyses-exploitation-miniere/faits-mineraux-metaux/faits-zinc' target='_blank'>ressourcesnaturelles.canada.ca</a></p>
</div>

<div class='src-section'>
  <h4>🔩 Cobalt (Co)</h4>
  <p class='src-ref'>Cobalt Institute. (s.d.). Données GWP Crude Cobalt Hydroxide (4,1 kg CO₂/kg extraction).
  <a href='https://www.cobaltinstitute.org/' target='_blank'>cobaltinstitute.org</a></p>
  <p class='src-ref'>Santika, E. F. (2025, 3 février). Indonesia to become one of the world's top cobalt producers in 2024.
  <a href='https://databoks.katadata.co.id/en/mining/statistics/67a0837bd5cc3/indonesia-to-become-one-of-the-worlds-top-cobalt-producers-in-2024' target='_blank'>databoks.katadata.co.id</a></p>
</div>

<div class='src-section'>
  <h4>🌿 Terres rares (REE)</h4>
  <p class='src-ref'>European Commission, Joint Research Centre. (2020). <i>Critical Raw Materials for Strategic Technologies and Sectors in the EU.</i> (REE : 19 kg CO₂/kg extraction, 44 kg CO₂/kg raffinage — mines Bayan Obo + argiles ioniques.)</p>
  <p class='src-ref'>Filipenco, D. (2026, 29 mai). Top 10 countries by rare earth elements production. Development Aid.
  <a href='https://www.developmentaid.org/news-stream/post/193928/top-countries-by-rare-earth-elements-production' target='_blank'>developmentaid.org</a></p>
  <p class='src-ref'>Pistilli, M. (2025, 25 mai). Top 10 countries by rare earth metal production. Investing News Network.
  <a href='https://investingnews.com/daily/resource-investing/critical-metals-investing/rare-earth-investing/rare-earth-metal-production/' target='_blank'>investingnews.com</a></p>
</div>

<div class='src-section'>
  <h4>🔌 Nickel (Ni)</h4>
  <p class='src-ref'>Monica, L., &amp; Priambodo, D. (2025, 19 décembre). Indonesia to cut nickel production in 2026 to maintain price stability. IDN Financials.
  <a href='https://www.idnfinancials.com/news/59725/indonesia-to-cut-nickel-production-in-2026-to-maintain-price-stability' target='_blank'>idnfinancials.com</a></p>
  <p class='src-ref'>Institute for Energy Economics and Financial Analysis. (2024, octobre). <i>Indonesia's nickel companies need renewable energy amid increasing production.</i>
  <a href='https://ieefa.org/sites/default/files/2024-10/IEEFA%20Report%20-%20Indonesia%27s%20nickel%20companies%20need%20RE_Oct2024.pdf' target='_blank'>ieefa.org</a></p>
  <p class='src-ref'>International Nickel Study Group. (2024). <i>The world nickel factbook 2024.</i>
  <a href='https://insg.org/wp-content/uploads/2024/09/publist_The-World-Nickel-Factbook-2024.pdf' target='_blank'>insg.org</a></p>
  <p class='src-ref'>Leandro, A. (2024, 24 avril). Achieving the transition to net zero in Australia (OECD Economics Department Working Papers).
  <a href='https://doi.org/10.1787/9a56c9d2-en' target='_blank'>doi.org</a></p>
  <p class='src-ref'>Rahmaditio, R., Wang, K., Purba, D. C., &amp; Zuhaira, N. (2026, 6 janvier). Nickel's climate cost and Indonesia's push for cleaner production. World Resources Institute.
  <a href='https://www.wri.org/technical-perspectives/nickels-climate-cost-and-indonesias-push-cleaner-production' target='_blank'>wri.org</a></p>
</div>

<div class='src-section'>
  <h4>💎 Gallium (Ga)</h4>
  <p class='src-ref'>Luo, H., Huang, T., Wu, X., &amp; Zhao, F. (2025). Life cycle assessment for primary gallium production at industrial-scale. <i>The International Journal of Life Cycle Assessment, 30</i>, 1545–1559.
  <a href='https://doi.org/10.1007/s11367-025-02492-1' target='_blank'>doi.org</a></p>
  <p class='src-ref'>SCRREEN2. (2020). <i>Factsheet update based on EU factsheets 2020: Gallium.</i> European Commission Horizon 2020.
  <a href='https://scrreen.eu/wp-content/uploads/2023/03/SCRREEN2_factsheets_GALLIIUM.pdf' target='_blank'>scrreen.eu</a></p>
</div>

<div class='src-section'>
  <h4>🪨 Cuivre (Cu)</h4>
  <p class='src-ref'>International Copper Association. (2023, avril). <i>Copper — Pathway to net zero, regional focus: Latin America.</i>
  <a href='https://internationalcopper.org/wp-content/uploads/2023/04/20230419-Pathway-to-Net-Zero-LatAm-ENGLISH.pdf' target='_blank'>internationalcopper.org</a></p>
  <p class='src-ref'>Liu, L., Xiang, D., Cao, H., &amp; Li, P. (2022). Life cycle energy consumption and GHG emissions of the copper production in China and the influence of main factors on the above performance. <i>Processes, 10</i>(12), 2715.
  <a href='https://doi.org/10.3390/pr10122715' target='_blank'>doi.org</a></p>
  <p class='src-ref'>Liwanag, A. (2025, 11 septembre). Beyond the grid — Copper's decarbonization focus shifts toward mines. S&amp;P Global.
  <a href='https://www.spglobal.com/market-intelligence/en/news-insights/research/2025/09/beyond-the-grid-coppers-decarbonization-focus-shifts-toward-mines' target='_blank'>spglobal.com</a></p>
</div>

<div class='src-section'>
  <h4>🧴 Plastique</h4>
  <p class='src-ref'>Green Stars Project. (2025, 24 août). The astonishing carbon footprint of plastic production.
  <a href='https://greenstarsproject.org/2025/08/24/the-astonishing-carbon-footprint-of-plastic-production/' target='_blank'>greenstarsproject.org</a></p>
  <p class='src-ref'>Matt. (2021, 25 août). MRP 119: Carbon intensity of US oil production compared to OPEC and Russia. Mineral Rights Podcast.
  <a href='https://mineralrightspodcast.com/mrp-119-carbon-intensity-of-us-oil-production-compared-to-opec-and-russia/' target='_blank'>mineralrightspodcast.com</a></p>
  <p class='src-ref'>Saudi Aramco. (2018, 7 octobre). Study shows record low carbon intensity of Saudi crude oil.
  <a href='https://www.aramco.com/en/news-media/news/2018/study-shows-record-low-carbon-intensity-of-saudi-crude-oil' target='_blank'>aramco.com</a></p>
</div>

<div class='src-section'>
  <h4>🔄 Équivalences pédagogiques (CO₂)</h4>
  <p class='src-ref'>ADEME. (s.d.). <i>Évolution du taux moyen d'émissions de CO₂ des véhicules neufs en France.</i> Car Labelling ADEME (130 g CO₂/km : valeur retenue pour l'équivalence voiture, proche de l'objectif réglementaire européen 2015 ; la moyenne du parc neuf est depuis descendue à ~111 g/km en 2025).
  <a href='https://carlabelling.ademe.fr/chiffrescles/r/evolutionTauxCo2' target='_blank'>carlabelling.ademe.fr</a></p>
  <p class='src-ref'>ADEME. (s.d.). <i>Avion trajet moyen-courrier.</i> Impact CO₂ (180 g CO₂e/km/passager : valeur retenue pour l'équivalence avion, vol 1 000–3 500 km avec traînées de condensation ; les vols court-courriers sont sensiblement plus émetteurs par km, ≈230 g/km, du fait du décollage).
  <a href='https://impactco2.fr/outils/transport/avion-moyencourrier' target='_blank'>impactco2.fr</a></p>
  <p class='src-ref'>Flint. (s.d.). Quel est l'effet du streaming vidéo sur l'environnement ? (rapporte l'estimation de l'Agence internationale de l'énergie, IEA : 36 g CO₂eq/heure — valeur retenue pour l'équivalence streaming, généralement considérée comme une borne basse ; d'autres études, dont l'ADEME et The Shift Project, avancent des valeurs très variables, de 6 à 64 g CO₂e/h selon l'appareil, la qualité vidéo et le réseau).
  <a href='https://flint.media/posts/161-quel-est-l-effet-du-streaming-video-sur-l-environnement' target='_blank'>flint.media</a></p>
  <p class='src-ref'>EcoTree. (s.d.). Combien de CO₂ un arbre absorbe-t-il ? (22 kg CO₂/an : ordre de grandeur médian retenu pour l'équivalence arbre, conforme aux données ONF/GIEC ; la fourchette publiée va de 10 à 50 kg CO₂/an selon l'espèce, l'âge et le climat).
  <a href='https://ecotree.green/en/how-much-co2-does-a-tree-absorb' target='_blank'>ecotree.green</a></p>
</div>
</div>
"""

# ╔══════════════════════════════════════════════════════════════════╗
# ║  TITRE DE GRAPHE : message pédagogique + sous-titre technique    ║
# ╚══════════════════════════════════════════════════════════════════╝
def chart_title(message, subtitle, size=14.5):
    """Construit un titre Plotly à deux niveaux : un message clair en
    1/2 phrase (ce que le graphe communique), et la légende technique
    précise en plus petit, juste dessous."""
    return dict(
        text=(f"{message}<br>"
              f"<sup style='font-size:10.5px;font-weight:400;color:#8888aa'>{subtitle}</sup>"),
        font=dict(size=size, color="#e4e4f0"),
    )

# ╔══════════════════════════════════════════════════════════════════╗
# ║  UI                                                              ║
# ╚══════════════════════════════════════════════════════════════════╝
app_ui = ui.page_fluid(
    ui.tags.head(ui.tags.style(CSS)),
    ui.tags.h2("📱 Impact Écologique des Smartphones", class_="app-title"),

    ui.layout_sidebar(
        ui.sidebar(
            ui.h4("🔍 Filtres", style="color:#c8c8e0;margin-bottom:10px;"),
            ui.input_select("element","Élément :",
                            choices=["Tous les éléments"] + elements_uniques),
            ui.hr(),
            ui.h5("📊 Indicateurs", style="color:#9090b0;font-size:.82em;text-transform:uppercase;letter-spacing:.06em;"),
            ui.output_ui("kpi_sidebar"),
            ui.hr(),
            ui.p("ℹ️ Ordres de grandeur d'après USGS MCS 2024-25 et rapports LCA publics. "
                 "Les plages sont représentées par leur moyenne.",
                 style="font-size:.73em;color:#666688;line-height:1.5;"),
            width=270, style="padding:14px;"
        ),

        ui.navset_card_tab(

            # ── 1. Vue d'ensemble ──────────────────────────────────
            ui.nav_panel("📊 Vue d'Ensemble",
                ui.row(
                    ui.column(5, output_widget("plot_treemap")),
                    ui.column(7, output_widget("plot_radar")),
                ),
                ui.row(ui.column(12, output_widget("plot_sunburst"))),
            ),

            # ── 2. Impact Physique ─────────────────────────────────
            ui.nav_panel("🪨 Impact Physique",
                ui.div("Pour fabriquer un smartphone, des quantités massives de roche sont extraites. "
                       "L'échelle logarithmique permet de comparer les ordres de grandeur ; "
                       "l'échelle linéaire rend l'écart visuellement immédiat.",
                       class_="nav-note"),
                # Ligne 1 : bar chart avec toggle + Sankey côte à côte
                ui.row(
                    ui.column(6,
                        ui.div(
                            ui.input_radio_buttons(
                                "roche_scale", "Échelle du graphe :",
                                choices={"log": "Logarithmique", "lin": "Linéaire"},
                                selected="log", inline=True,
                            ),
                            style="margin-bottom:4px;"
                        ),
                        output_widget("plot_roche_bar"),
                    ),
                    ui.column(6, output_widget("plot_sankey_norm")),
                ),
                # Ligne 2 : scatter
                ui.row(ui.column(12, output_widget("plot_scatter_effort"))),
            ),

            # ── 3. Empreinte CO₂ ───────────────────────────────────
            ui.nav_panel("🌡️ Empreinte CO₂",
                ui.div("CO₂ extraction &amp; raffinage : intensité carbone × grammage. "
                       "Colonne assemblage = facteur cycle de vie complet Fairphone 5 (≈200 kg CO₂/kg). "
                       "⚠️ In, Au, Ga : valeurs extraction sous-estimées dans le fichier source.",
                       class_="nav-note"),
                ui.row(
                    ui.column(6, output_widget("plot_co2_heatmap")),
                    ui.column(6, output_widget("plot_co2_stacked")),
                ),
                ui.hr(style="border-color:#1e2030;margin:14px 0;"),
                ui.h5("🔄 Équivalences concrètes", style="color:#c8c8e0;margin-bottom:8px;"),
                ui.output_ui("co2_equiv_cards"),
            ),

            # ── 4. Chaîne de Valeur ────────────────────────────────
            ui.nav_panel("🌍 Chaîne de Valeur",
                ui.p("Sélectionner un élément pour voir la chaîne Extraction → Raffinage → Assemblage.",
                     class_="section-note"),
                output_widget("map_chaine"),
                ui.output_ui("map_legend_ui"),
            ),

            # ── 5. Sources & Méthodologie ──────────────────────────
            ui.nav_panel("📖 Sources & Méthodologie",
                ui.div(
                    ui.HTML(SOURCES_HTML),
                    style="max-height:70vh;overflow-y:auto;padding:4px 8px;"
                )
            ),
        )
    )
)

# ╔══════════════════════════════════════════════════════════════════╗
# ║  SERVER                                                          ║
# ╚══════════════════════════════════════════════════════════════════╝
def server(input, output, session):

    @reactive.Calc
    def fdata():
        sel = input.element()
        if sel == "Tous les éléments":
            return df.copy()
        return df[df["Element"].astype(str).str.strip() == sel].copy()

    # ── KPI Sidebar ───────────────────────────────────────────────
    @render.ui
    def kpi_sidebar():
        d = fdata()
        co2_kg  = d["CO2_total_g"].sum() / 1_000 if "CO2_total_g" in d.columns else 0
        roche   = d["Roche_totale_kg"].sum()      if "Roche_totale_kg" in d.columns else 0
        km_voit = co2_kg * 1_000 / CO2_VOITURE_G_KM
        roche_s = f"{roche:.1f} kg" if roche < 1_000 else f"{roche/1_000:.2f} t"
        return ui.div(
            ui.div(ui.div(f"{co2_kg:.1f} kg CO₂", class_="kpi-val"),
                   ui.div("Empreinte carbone totale", class_="kpi-lbl"),
                   class_="kpi-card kpi-co2"),
            ui.div(ui.div(roche_s, class_="kpi-val"),
                   ui.div("Roche extraite (estimation)", class_="kpi-lbl"),
                   class_="kpi-card kpi-rock"),
            ui.div(ui.div(f"≈ {km_voit:.0f} km", class_="kpi-val"),
                   ui.div("Équiv. voiture (130 g/km)", class_="kpi-lbl"),
                   class_="kpi-card kpi-car"),
            class_="kpi-wrap")

    # ── Vue d'ensemble : Treemap ──────────────────────────────────
    @render_widget
    def plot_treemap():
        d = fdata()
        need = ["Element","Proportion_pct","CO2_total_g"]
        if not all(c in d.columns for c in need): return go.Figure()
        pd_ = d[need + (["Partie_telephone"] if "Partie_telephone" in d.columns else [])].copy()
        pd_ = pd_.fillna({"Proportion_pct":0,"CO2_total_g":0})
        pd_ = pd_[pd_["Proportion_pct"] > 0]
        if "Partie_telephone" in pd_.columns:
            pd_["Partie_telephone"] = pd_["Partie_telephone"].fillna("Autre")
            path = ["Partie_telephone","Element"]
        else:
            path = ["Element"]
        fig = px.treemap(pd_, path=path, values="Proportion_pct", color="CO2_total_g",
                         color_continuous_scale="RdYlGn_r",
                         labels={"CO2_total_g":"CO₂ (g)","Proportion_pct":"%"},
                         hover_data={"Proportion_pct":":.2f","CO2_total_g":":.1f"})
        fig.update_traces(textinfo="label+percent entry")
        fig.update_layout(
            title=chart_title(
                "Le poids ne dit pas tout : masse et pollution ne vont pas de pair",
                "🗺️ Proportion massique (surface) & intensité CO₂ (couleur)"),
            height=420, paper_bgcolor="#12141f", font_color="#c8c8e0",
            coloraxis_colorbar=dict(title="CO₂ (g)", tickfont_color="#c8c8e0"),
            margin=dict(t=58,l=10,r=10,b=10))
        return fig

    # ── Vue d'ensemble : Radar ────────────────────────────────────
    @render_widget
    def plot_radar():
        d = fdata()
        metrics = {
            "CO₂ extract.":    "CO2_ext_g",
            "CO₂ raffin.":     "CO2_raf_g",
            "CO₂ lifecycle":   "CO2_asm_g",
            "Roche extraite":  "Roche_totale_kg",
            "Proportion (%)":  "Proportion_pct",
        }
        avail = {k: v for k, v in metrics.items() if v in d.columns}
        if len(avail) < 3: return go.Figure(layout=dict(title="Données insuffisantes"))
        pd_ = d[["Element"] + list(avail.values())].dropna().copy()
        # Normalisation : % de contribution de l'élément au TOTAL de chaque axe
        # → Aucun élément forcé à 100 %, la somme par axe = 100 %
        for col in avail.values():
            total = df[col].sum() if col in df.columns else pd_[col].sum()
            if total > 0: pd_[col] = pd_[col] / total * 100
        cats    = list(avail.keys())
        palette = px.colors.qualitative.Set2
        fig     = go.Figure()
        for i, (_, row) in enumerate(pd_.iterrows()):
            vals  = [row[v] for v in avail.values()]
            color = palette[i % len(palette)]
            # Valeurs absolues pour le hover
            abs_row = df[df["Element"] == row["Element"]].iloc[0] if any(df["Element"] == row["Element"]) else {}
            ht_parts = []
            for k, v in avail.items():
                raw_val = abs_row[v] if hasattr(abs_row, '__getitem__') and v in df.columns else "?"
                try:
                    ht_parts.append(f"{k}: {float(raw_val):.2f} ({vals[list(avail.keys()).index(k)]:.1f}%)")
                except Exception:
                    ht_parts.append(f"{k}: ?")
            ht = "<br>".join(ht_parts)
            fig.add_trace(go.Scatterpolar(
                r=vals + [vals[0]], theta=cats + [cats[0]],
                fill="toself", name=row["Element"], opacity=0.55,
                line=dict(color=color, width=2), fillcolor=color,
                hovertemplate=f"<b>{row['Element']}</b><br>{ht}<extra></extra>",
            ))
        fig.update_layout(
            polar=dict(
                radialaxis=dict(visible=True, range=[0, 100],
                                ticksuffix="%", gridcolor="#2a2d40",
                                tickfont_color="#9090b0"),
                angularaxis=dict(tickfont_color="#c8c8e0"),
                bgcolor="#12141f",
            ),
            title=chart_title(
                "Quelques éléments concentrent l'essentiel de l'impact",
                "🕸️ Part de chaque élément dans chaque indicateur (% du total — somme par axe = 100 %)",
                size=13.5),
            height=420, paper_bgcolor="#12141f", font_color="#c8c8e0",
            legend=dict(orientation="h", y=-0.22, font_size=11),
        )
        return fig

    # ── Vue d'ensemble : Sunburst ─────────────────────────────────
    @render_widget
    def plot_sunburst():
        d = fdata()
        if "Proportion_pct" not in d.columns: return go.Figure()
        pd_ = d.copy()
        pd_["Proportion_pct"] = pd_["Proportion_pct"].fillna(0)
        pd_ = pd_[pd_["Proportion_pct"] > 0]
        if "Partie_telephone" in pd_.columns:
            pd_["Partie_telephone"] = pd_["Partie_telephone"].fillna("Autre")
            color_col = "CO2_total_g" if "CO2_total_g" in pd_.columns else "Proportion_pct"
            fig = px.sunburst(pd_, path=["Partie_telephone","Element"],
                              values="Proportion_pct", color=color_col,
                              color_continuous_scale="RdYlGn_r")
        else:
            fig = px.pie(pd_, names="Element", values="Proportion_pct", hole=0.4)
        fig.update_layout(
            title=chart_title(
                "Peu de matière, mais une grande diversité de métaux",
                "☀️ Composition du smartphone par partie — couleur = CO₂ total"),
            height=460, paper_bgcolor="#12141f", font_color="#c8c8e0",
            coloraxis_colorbar=dict(title="CO₂ (g)", tickfont_color="#c8c8e0"))
        return fig

    # ── Impact Physique : Roche bar (toggle log/lin) ──────────────
    @render_widget
    def plot_roche_bar():
        d     = fdata()
        scale = input.roche_scale()   # "log" ou "lin"
        if "Roche_totale_g" not in d.columns: return go.Figure()
        pd_ = d[["Element","Grammage_g","Roche_totale_g"]].dropna()
        pd_ = pd_[pd_["Roche_totale_g"] > 0].sort_values("Roche_totale_g", ascending=False)

        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="Roche extraite (g)", x=pd_["Element"], y=pd_["Roche_totale_g"],
            marker_color="#e05c4a", opacity=0.88,
            hovertemplate="<b>%{x}</b><br>Roche : %{y:,.0f} g (%{customdata:.2f} kg)<extra></extra>",
            customdata=pd_["Roche_totale_g"] / 1000,
        ))
        fig.add_trace(go.Bar(
            name="Métal utile (g)", x=pd_["Element"], y=pd_["Grammage_g"],
            marker_color="#2ecc71", opacity=0.88,
            hovertemplate="<b>%{x}</b><br>Métal utile : %{y:.4f} g<extra></extra>",
        ))
        y_type = "log" if scale == "log" else "linear"
        y_title = "Quantité (g, échelle log)" if scale == "log" else "Quantité (g, échelle linéaire)"
        title_suffix = "éch. log" if scale == "log" else "éch. linéaire"
        fig.update_layout(
            title=chart_title(
                "Extraire un peu de métal exige d'énormes quantités de roche",
                f"🪨 Roche extraite vs métal utile (g) — {title_suffix}"),
            yaxis=dict(type=y_type, title=y_title, gridcolor="#1e2030"),
            xaxis=dict(gridcolor="#1e2030"),
            barmode="group",
            legend=dict(orientation="h", y=1.1, font_size=11),
            height=420, plot_bgcolor="#12141f", paper_bgcolor="#12141f",
            font_color="#c8c8e0",
        )
        return fig

    # ── Impact Physique : Sankey normalisé (gangue vs téléphone) ──
    @render_widget
    def plot_sankey_norm():
        d = fdata()
        if "Roche_totale_kg" not in d.columns: return go.Figure()
        pd_ = d[["Element","Roche_totale_kg","Grammage_g"]].dropna()
        pd_ = pd_[pd_["Roche_totale_kg"] > 0].reset_index(drop=True)
        n   = len(pd_)
        MIN_PCT = 0.5   # % minimum visible pour la bande "in phone"

        palette = px.colors.qualitative.Pastel
        node_labels = pd_["Element"].tolist() + ["♻️ Gangue / stériles", "📱 Dans votre téléphone"]
        node_colors = [palette[i % len(palette)] for i in range(n)] + ["#555577","#2ecc71"]

        sources, targets, values, link_colors, link_hover = [], [], [], [], []
        for i, (_, row) in enumerate(pd_.iterrows()):
            r_kg    = row["Roche_totale_kg"]
            m_kg    = row["Grammage_g"] / 1_000
            pct_tel = max(m_kg / r_kg * 100, MIN_PCT)
            pct_gan = 100 - pct_tel
            real_pct = m_kg / r_kg * 100  # vraie valeur pour le hover

            sources  += [i, i]
            targets  += [n, n+1]
            values   += [round(pct_gan, 4), round(pct_tel, 4)]
            link_colors += ["rgba(100,100,140,0.35)", "rgba(46,204,113,0.55)"]
            link_hover  += [
                f"{row['Element']} → Gangue : {pct_gan:.2f}% ({r_kg*(pct_gan/100):.2f} kg)",
                f"{row['Element']} → Tél. : {real_pct:.4f}% réel ({m_kg*1000:.3f} g)",
            ]

        fig = go.Figure(go.Sankey(
            arrangement="snap",
            node=dict(pad=18, thickness=20,
                      line=dict(color="#1a1d2e", width=0.5),
                      label=node_labels, color=node_colors),
            link=dict(source=sources, target=targets,
                      value=values, color=link_colors,
                      customdata=link_hover,
                      hovertemplate="%{customdata}<extra></extra>"),
        ))
        fig.update_layout(
            title=chart_title(
                "La quasi-totalité de la roche extraite finit en déchet",
                "🔄 Flux normalisé : part de la roche extraite par élément (100 % = roche totale "
                "extraite pour cet élément) — bande verte amplifiée à 0,5 % min., valeur réelle au survol"),
            height=440, font=dict(size=11, color="#c8c8e0"),
            paper_bgcolor="#12141f",
        )
        return fig

    # ── Impact Physique : Scatter effort ──────────────────────────
    @render_widget
    def plot_scatter_effort():
        d = fdata()
        need = ["Roche_par_g_kg","Proportion_pct","CO2_total_g","Element"]
        if not all(c in d.columns for c in need): return go.Figure()
        pd_ = d[need].dropna().copy()
        pd_["CO2_total_g"] = pd_["CO2_total_g"].clip(lower=0.1)
        fig = px.scatter(
            pd_, x="Proportion_pct", y="Roche_par_g_kg",
            size="CO2_total_g", text="Element", log_y=True,
            size_max=55, color="CO2_total_g", color_continuous_scale="OrRd",
            labels={"Proportion_pct":"% dans le smartphone",
                    "Roche_par_g_kg":"Roche / g métal (kg, éch. log)",
                    "CO2_total_g":"CO₂ total (g)"},
        )
        fig.update_traces(textposition="top center", marker_opacity=0.8)
        fig.update_layout(
            title=chart_title(
                "Les métaux les plus rares sont aussi les plus coûteux à extraire",
                "⚖️ % dans le smartphone vs roche nécessaire par gramme (taille = CO₂ total)"),
            height=380, plot_bgcolor="#12141f", paper_bgcolor="#12141f",
            font_color="#c8c8e0",
            coloraxis_colorbar=dict(title="CO₂ (g)", tickfont_color="#c8c8e0"),
            yaxis=dict(gridcolor="#1e2030"), xaxis=dict(gridcolor="#1e2030"))
        return fig

    # ── CO₂ Heatmap ───────────────────────────────────────────────
    @render_widget
    def plot_co2_heatmap():
        d = fdata()
        cmap = {"Extraction":"CO2_ext_g","Raffinage":"CO2_raf_g","Cycle de vie":"CO2_asm_g"}
        avail = {k: v for k, v in cmap.items() if v in d.columns}
        if not avail: return go.Figure()
        hd = d[["Element"] + list(avail.values())].set_index("Element").fillna(0)
        hd.columns = list(avail.keys())
        fig = px.imshow(hd.T, color_continuous_scale="YlOrRd", aspect="auto",
                        labels=dict(x="Élément", y="Étape", color="g CO₂"), text_auto=".0f")
        fig.update_layout(
            title=chart_title(
                "Le cycle de vie complet pèse souvent plus que l'extraction",
                "🌡️ CO₂ par étape et par élément (g)"),
            height=320, paper_bgcolor="#12141f", font_color="#c8c8e0",
            xaxis=dict(tickfont_color="#c8c8e0"),
            yaxis=dict(tickfont_color="#c8c8e0"),
            coloraxis_colorbar=dict(tickfont_color="#c8c8e0"))
        return fig

    # ── CO₂ Stacked Bar ───────────────────────────────────────────
    @render_widget
    def plot_co2_stacked():
        d = fdata()
        cfg = [("Extraction","CO2_ext_g","#e05c4a"),
               ("Raffinage","CO2_raf_g","#e8a030"),
               ("Cycle de vie","CO2_asm_g","#8b5cf6")]
        avail = [(lbl,col,c) for lbl,col,c in cfg if col in d.columns]
        if not avail: return go.Figure()
        d = d.copy()
        d["_tot"] = d[[col for _,col,_ in avail]].fillna(0).sum(axis=1)
        d = d.sort_values("_tot", ascending=False)
        fig = go.Figure()
        for lbl, col, color in avail:
            fig.add_trace(go.Bar(name=lbl, x=d["Element"], y=d[col].fillna(0),
                                  marker_color=color,
                                  hovertemplate=f"<b>%{{x}}</b><br>{lbl}: %{{y:.1f}} g CO₂<extra></extra>"))
        fig.update_layout(
            title=chart_title(
                "La fabrication (cycle de vie) domine l'empreinte carbone totale",
                "📊 CO₂ par étape (g) pour le grammage dans 1 téléphone"),
            yaxis_title="g CO₂", barmode="stack",
            legend=dict(orientation="h", y=1.12),
            height=330, plot_bgcolor="#12141f", paper_bgcolor="#12141f",
            font_color="#c8c8e0",
                          yaxis=dict(gridcolor="#1e2030"), xaxis=dict(gridcolor="#1e2030"))
        return fig

    # ── CO₂ Équivalences ──────────────────────────────────────────
    @render.ui
    def co2_equiv_cards():
        d     = fdata()
        co2_g = d["CO2_total_g"].sum() if "CO2_total_g" in d.columns else 0
        equivs = [
            ("🚗", f"{co2_g/CO2_VOITURE_G_KM:.0f} km",   "En voiture (130 g CO₂/km)", "#2980b9"),
            ("✈️", f"{co2_g/CO2_AVION_G_KM:.0f} km",     "En avion (180 g CO₂/km)",   "#c0392b"),
            ("📺", f"{co2_g/CO2_STREAMING_G_H:.0f} h",    "Streaming vidéo (36 g/h)",  "#8e44ad"),
            ("🌳", f"{co2_g/CO2_ARBRE_G_AN:.1f} ans",    "Arbre adulte (~22 kg/an)",   "#27ae60"),
        ]
        cards = [ui.div(ui.div(ic,class_="equiv-icon"),ui.div(v,class_="equiv-val"),
                        ui.div(l,class_="equiv-lbl"),class_="equiv-card",
                        style=f"background:{c};") for ic,v,l,c in equivs]
        return ui.div(
            ui.p(f"CO₂ total estimé pour « {input.element()} » : {co2_g/1_000:.2f} kg CO₂",
                 style="font-weight:600;color:#c8c8e0;margin-bottom:10px;"),
            ui.div(*cards, class_="equiv-row"))

    # ── Carte : légende ───────────────────────────────────────────
    @render.ui
    def map_legend_ui():
        is_single = input.element() != "Tous les éléments"
        if is_single:
            return ui.div(ui.HTML(
                "<b>Lecture :</b> &nbsp;"
                "<span style='color:#e05c4a'>●</span> Extraction &nbsp;|&nbsp; "
                "<span style='color:#4a90e0'>◆</span> Raffinage &nbsp;|&nbsp; "
                "<span style='color:#2ecc71'>★</span> Assemblage (Shenzhen) — "
                "Lignes orange : Extraction→Raffinage &nbsp;|&nbsp; "
                "Lignes vertes pointillées : Raffinage→Assemblage — "
                "Les flèches indiquent la direction du flux."
            ), class_="map-legend")
        return ui.div(ui.HTML(
            "<b>Vue globale :</b> Taille = nb d'éléments traités dans ce pays. "
            "Survoler pour voir les éléments. "
            "Sélectionner un élément pour afficher les flux."
        ), class_="map-legend")

    # ── Carte : map principale ────────────────────────────────────
    @render_widget
    def map_chaine():
        d         = fdata()
        is_single = input.element() != "Tous les éléments"

        def parse_geo(names_raw, counts_raw):
            s = str(names_raw).replace(" et ", ",")
            countries = [p.strip() for p in s.split(",")
                         if p.strip().lower() not in ("nulle","inconnu","nan","none","")]
            c_str = str(counts_raw).strip().lower()
            out   = {c: 1 for c in countries}
            if c_str not in ("nulle","inconnu","nan","none",""):
                for pays, nb in re.findall(r"([a-zA-ZÀ-ÿ\s\-]+)\s+(\d+)", str(counts_raw)):
                    pc = pays.strip(); v = max(1, int(nb)); matched = False
                    for bc in list(out):
                        if pc.lower() in bc.lower() or bc.lower() in pc.lower():
                            out[bc] = v; matched = True; break
                    if not matched and pc.lower() not in ("nulle","inconnu","nan","none",""):
                        out[pc] = v
            return out

        def azimuth(lat1, lon1, lat2, lon2):
            r = list(map(math.radians, [lat1,lon1,lat2,lon2]))
            dl = r[3]-r[1]
            x  = math.sin(dl)*math.cos(r[2])
            y  = math.cos(r[0])*math.sin(r[2]) - math.sin(r[0])*math.cos(r[2])*math.cos(dl)
            return (math.degrees(math.atan2(x,y)) + 360) % 360

        def great_circle_path(lat1, lon1, lat2, lon2, n=24):
            """Échantillonne n points le long du grand cercle reliant les deux
            points (interpolation sphérique / slerp). Plotly relie ensuite ces
            points par des segments, ce qui reproduit fidèlement la courbe
            géodésique réellement affichée sur la carte — contrairement à une
            interpolation linéaire naïve en (lat, lon), qui peut s'écarter de
            dizaines de degrés de la courbe réelle sur de longues distances."""
            phi1, lam1 = math.radians(lat1), math.radians(lon1)
            phi2, lam2 = math.radians(lat2), math.radians(lon2)
            x1,y1,z1 = math.cos(phi1)*math.cos(lam1), math.cos(phi1)*math.sin(lam1), math.sin(phi1)
            x2,y2,z2 = math.cos(phi2)*math.cos(lam2), math.cos(phi2)*math.sin(lam2), math.sin(phi2)
            dot = max(-1.0, min(1.0, x1*x2 + y1*y2 + z1*z2))
            d = math.acos(dot)
            if d < 1e-9:
                return [(lat1, lon1)] * n
            pts = []
            for i in range(n):
                t = i / (n - 1)
                A = math.sin((1-t)*d) / math.sin(d)
                B = math.sin(t*d) / math.sin(d)
                x, y, z = A*x1+B*x2, A*y1+B*y2, A*z1+B*z2
                lat = math.degrees(math.atan2(z, math.sqrt(x*x + y*y)))
                lon = math.degrees(math.atan2(y, x))
                pts.append((lat, lon))
            return pts

        # Collecte
        rows = []; el_flows = {}
        for _, row in d.iterrows():
            el  = row["Element"]
            ext = parse_geo(row.get("Pays_extraction",""), row.get("Nb_mines",""))
            raf = parse_geo(row.get("Raffinerie",""),      row.get("Nb_raffineries",""))
            el_flows[el] = {
                "ext": [PAYS_FR_EN.get(k,k) for k in ext],
                "raf": [PAYS_FR_EN.get(k,k) for k in raf],
            }
            for pays_fr,t in ext.items():
                rows.append({"Pays_EN":PAYS_FR_EN.get(pays_fr,pays_fr),"Pays_FR":pays_fr,
                              "Type":"Extraction","Taille":t,"Element":el})
            for pays_fr,t in raf.items():
                rows.append({"Pays_EN":PAYS_FR_EN.get(pays_fr,pays_fr),"Pays_FR":pays_fr,
                              "Type":"Raffinage","Taille":t,"Element":el})

        if not rows:
            return go.Figure(layout=dict(title="Aucune donnée géographique",
                                          paper_bgcolor="#12141f",font_color="#c8c8e0"))

        gdf = pd.DataFrame(rows)
        agg = (gdf.groupby(["Pays_EN","Pays_FR","Type"])
                  .agg(Elements=("Element", lambda x: "<br>".join(sorted(set(x)))),
                       Count=("Element","count")).reset_index())
        agg["Lat"] = agg["Pays_EN"].map(lambda x: COUNTRY_COORDS.get(x,(None,None))[0])
        agg["Lon"] = agg["Pays_EN"].map(lambda x: COUNTRY_COORDS.get(x,(None,None))[1])
        agg = agg.dropna(subset=["Lat","Lon"])

        fig = go.Figure()

        # Flux (élément unique)
        if is_single:
            el   = input.element()
            flow = el_flows.get(el, {"ext":[],"raf":[]})

            def add_arrow_line(lat1, lon1, lat2, lon2, color, dash="solid"):
                if None in (lat1,lon1,lat2,lon2): return
                if abs(lat1-lat2) < 0.3 and abs(lon1-lon2) < 0.3: return
                path = great_circle_path(lat1, lon1, lat2, lon2, n=24)
                fig.add_trace(go.Scattergeo(
                    lat=[p[0] for p in path], lon=[p[1] for p in path],
                    mode="lines", showlegend=False, hoverinfo="skip",
                    line=dict(width=2.5, color=color, dash=dash)))
                # La flèche est placée sur un point RÉEL du tracé (index 80%),
                # donc toujours exactement sur la courbe affichée — plus aucun
                # flottement, et l'orientation suit la tangente locale au point.
                i = max(1, int(0.80 * (len(path) - 1)))
                a_lat, a_lon = path[i]
                p_lat, p_lon = path[i-1]
                brg = azimuth(p_lat, p_lon, a_lat, a_lon)
                fig.add_trace(go.Scattergeo(
                    lat=[a_lat], lon=[a_lon], mode="markers",
                    showlegend=False, hoverinfo="skip",
                    marker=dict(size=10, color=color, symbol="arrow",
                                angle=brg, line=dict(color="white",width=0.8))))

            EXT_COLOR = "#f7a030"   # orange : Extraction → Raffinage
            ASM_COLOR = "#2ecc71"   # vert   : Raffinage → Assemblage

            # Extraction → Raffinage
            for ext_en in flow["ext"]:
                ec = COUNTRY_COORDS.get(ext_en)
                if not ec: continue
                for raf_en in flow["raf"]:
                    rc = COUNTRY_COORDS.get(raf_en)
                    if not rc or ext_en == raf_en: continue
                    add_arrow_line(ec[0],ec[1],rc[0],rc[1], EXT_COLOR)

            # Raffinage → Assemblage (TOUS les pays de raffinage, y compris Chine)
            ac = (ASSEMBLY_LAT, ASSEMBLY_LON)
            for raf_en in flow["raf"]:
                rc = COUNTRY_COORDS.get(raf_en)
                if not rc: continue
                add_arrow_line(rc[0],rc[1],ac[0],ac[1], ASM_COLOR, dash="dot")

            # Légende fantôme
            for name, col, dash in [
                ("→ Extraction → Raffinage", EXT_COLOR, "solid"),
                ("→ Raffinage → Assemblage", ASM_COLOR, "dot"),
            ]:
                fig.add_trace(go.Scattergeo(lat=[None],lon=[None],mode="lines",
                                             line=dict(width=2.5,color=col,dash=dash),name=name))

        # Markers agrégés (refining second = on top, mais on dessine extraction après)
        for type_lbl, color, symbol, icon in [
            ("Raffinage",  "#4a90e0", "square",  "🏭"),
            ("Extraction", "#e05c4a", "circle",  "⛏️"),
        ]:
            sub = agg[agg["Type"] == type_lbl]
            if sub.empty: continue
            fig.add_trace(go.Scattergeo(
                lat=sub["Lat"], lon=sub["Lon"],
                mode="markers+text",
                marker=dict(size=sub["Count"]*8+14, color=color, symbol=symbol,
                            opacity=0.92, line=dict(color="white",width=1.5)),
                text=sub["Count"].apply(lambda n: str(n) if n > 1 else ""),
                textfont=dict(color="white",size=9,family="Space Grotesk"),
                textposition="middle center",
                customdata=list(zip(sub["Pays_FR"],sub["Elements"],sub["Count"])),
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    f"<b>{icon} {type_lbl}</b><br>"
                    "Éléments : %{customdata[1]}<br>"
                    "Nb éléments : %{customdata[2]}<extra></extra>"
                ),
                name=f"{icon} {type_lbl}",
            ))

        # Marker assemblage
        if is_single:
            el = input.element()
            fig.add_trace(go.Scattergeo(
                lat=[ASSEMBLY_LAT], lon=[ASSEMBLY_LON], mode="markers",
                marker=dict(size=22, color="#2ecc71", symbol="star",
                            opacity=0.95, line=dict(color="white",width=1.5)),
                hovertemplate=(f"<b>{ASSEMBLY_LABEL}</b><br>🔩 Assemblage final<br>"
                               f"Élément : {el}<extra></extra>"),
                name="🔩 Assemblage",
            ))

        fig.update_geos(showcountries=True, countrycolor="#3a3d55",
                        showland=True, landcolor="#1a1d2e",
                        showocean=True, oceancolor="#0d1020",
                        showcoastlines=True, coastlinecolor="#2a2d42",
                        projection_type="natural earth")
        fig.update_layout(
            title=chart_title(
                "Un smartphone dépend de matières premières venues de plusieurs continents",
                f"🌍 Chaîne de valeur — {input.element()}"),
            margin=dict(r=0,t=52,l=0,b=0), height=530,
            paper_bgcolor="#12141f", font_color="#c8c8e0",
            legend=dict(bgcolor="#1a1d2e",bordercolor="#2e3150",borderwidth=1,
                        font=dict(color="#c8c8e0",size=11),
                        orientation="h", y=-0.06, x=0),
        )
        return fig


app = App(app_ui, server)