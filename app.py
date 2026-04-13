import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import requests
from geopy.geocoders import Nominatim
import folium
from streamlit_folium import st_folium

# Importiamo LE TUE classi originali e intatte
from motore_h2fast import Analisi_tecnica, Analisi_finanziaria

st.set_page_config(page_title="H2FAsT Simulator", layout="wide")
st.title("Simulatore H2FAsT - Dashboard Interattiva 🏭")

# ==========================================
# FUNZIONE EOLICA
# ==========================================
def calcola_potenza_eolica_kw(v_vento_ms, p_nominale_kw):
    if p_nominale_kw == 0: return 0.0
    p_nominale_w = p_nominale_kw * 1000.0
    v_cut_in = 3.0    
    v_rated = 12.0    
    v_cut_out = 25.0  
    
    if v_vento_ms < v_cut_in or v_vento_ms > v_cut_out:
        return 0.0
    elif v_cut_in <= v_vento_ms < v_rated:
        return (p_nominale_w * ((v_vento_ms**3 - v_cut_in**3) / (v_rated**3 - v_cut_in**3))) / 1000.0
    else:
        return p_nominale_kw

# ==========================================
# BARRA LATERALE: PARAMETRI
# ==========================================
st.sidebar.header("📁 Preferenze")
lingua = st.sidebar.selectbox("Lingua output", ["ITA", "ENG"])
n_progetti = st.sidebar.number_input("Mostra i Top N Progetti", value=5, min_value=1)

st.sidebar.header("⚙️ Parametri Tecnici")
p_PV = st.sidebar.number_input("Fotovoltaico (kWp)", value=1000.0)
p_Wind = st.sidebar.number_input("Eolico (kW)", value=0.0)
dP_el = st.sidebar.number_input("Electrolyzer range limit", value=1.0)
min_elet = st.sidebar.number_input("% Electrolyzer power cutoff", value=0.2)
batteria = st.sidebar.selectbox("Presence of battery in the system", ["SI", "NO"])
dP_bat = st.sidebar.number_input("Battery range limit", value=2.0) if batteria == "SI" else 0.0
min_batt = st.sidebar.number_input("Cutoff SoC [%]", value=0.2)
max_batteria = st.sidebar.number_input("Maximum energy % deliverable in 1h", value=0.5)
eff_batt = st.sidebar.number_input("% Charge/discharge battery efficiency", value=0.95)

st.sidebar.header("💶 CAPEX & Costi Impianto")
Terr = st.sidebar.number_input("Land [€]", value=0.0)
OpeE = st.sidebar.number_input("Building works [€]", value=0.0)
StazzRif = st.sidebar.number_input("Filling station [€]", value=0.0)
ImpPV1eurokW = st.sidebar.number_input("PV system [€/kWp]", value=800.0)
ImpWind1eurokW = st.sidebar.number_input("Wind system [€/kW]", value=1000.0)
ImpExtra1eurokW = st.sidebar.number_input("Fonte Extra (es. Idro) [€/kW]", value=2000.0) # Costo Fonte Extra
EletteuroKW = st.sidebar.number_input("Electrolyzer [€/kW]", value=1650.0)
CompreuroKW = st.sidebar.number_input("Compressor [€/kW]", value=4000.0)
AccuEeurokW = st.sidebar.number_input("Batteries [€/kWh]", value=200.0)
costounitariostoccaggio = st.sidebar.number_input("Storage [€/kg]", value=1200.0)

st.sidebar.header("📊 Parametri Finanziari")
prezzoindrogeno = st.sidebar.number_input("Hydrogen sale price [€/kg]", value=20.0)
tassoVAN = st.sidebar.number_input("Discount rate for NPV", value=0.07)
inflazione = st.sidebar.number_input("General inflation rate", value=0.01)
DurPianEcon = st.sidebar.number_input("Economic plan duration [years]", value=20, step=1)

# ==========================================
# SCHERMATA PRINCIPALE: STEP 1
# ==========================================
st.header("📍 Step 1: Sorgente Dati Rinnovabili")
metodo_dati = st.radio("Come vuoi inserire i dati Base (Sole/Vento)?", ["Mappa Interattiva (Download da PVGIS & Open-Meteo)", "Carica file CSV (Manuale)"], horizontal=True)

lat, lon = None, None
file_csv = None
tipo_file_passato = "SI"

if metodo_dati == "Mappa Interattiva (Download da PVGIS & Open-Meteo)":
    col_testo, col_mappa = st.columns([1, 2])
    with col_testo:
        st.markdown("### 1. Trova la zona (Sole + Vento)")
        luogo = st.text_input("Cerca una città per centrare la mappa:", "Roma, Italia")
        st.markdown("### 2. Piazza il Pin")
        st.info("👉 **Clicca fisicamente un punto sulla mappa** a destra.")
        
        geolocator = Nominatim(user_agent="h2fast_simulator")
        start_loc = [41.89, 12.49]
        zoom = 5
        if luogo:
            try:
                loc = geolocator.geocode(luogo)
                if loc:
                    start_loc = [loc.latitude, loc.longitude]
                    zoom = 10
            except: pass

    with col_mappa:
        m = folium.Map(location=start_loc, zoom_start=zoom)
        map_data = st_folium(m, height=350, width=700)

    if map_data and map_data.get("last_clicked"):
        lat = map_data["last_clicked"]["lat"]
        lon = map_data["last_clicked"]["lng"]
        st.success(f"✅ Coordinate confermate: Latitudine **{lat:.4f}**, Longitudine **{lon:.4f}**")
else:
    file_csv = st.file_uploader("Carica il tuo file Base di serie oraria CSV", type=["csv"])
    tipo_file_passato = st.selectbox("Il file proviene da PVGIS?", ["SI", "NO"])

st.divider()
st.markdown("### ➕ 3. Integrazione Fonti Esterne (Opzionale)")
usa_extra = st.checkbox("Aggiungi un profilo di produzione utente (es. Idroelettrico, Biomasse, Misurazioni Reali)")
file_extra = None
p_Extra = 0.0

if usa_extra:
    col_ex1, col_ex2 = st.columns(2)
    with col_ex1:
        file_extra = st.file_uploader("Carica file CSV (1 colonna numerica, 8760 valori in kW)", type=["csv", "txt"])
    with col_ex2:
        p_Extra = st.number_input("Potenza Nominale della fonte extra (kW)", value=100.0, help="Serve per dimensionare correttamente la taglia massima dell'impianto ibrido")

# ==========================================
# AVVIO SIMULAZIONE E MOTORE
# ==========================================
st.divider()
st.header("⚙️ Step 2: Esecuzione")

if st.button("🚀 AVVIA SIMULAZIONE COMPLETA", use_container_width=True, type="primary"):
    temp_filename = "temp_input_orario"

    # --- FASE 1: CREAZIONE VETTORE BASE ---
    if metodo_dati == "Mappa Interattiva (Download da PVGIS & Open-Meteo)":
        if lat is None or lon is None:
            st.error("❌ Devi cliccare un punto sulla mappa prima di avviare!")
            st.stop()
        with st.spinner('📡 Download e Fusione dei dati (Sole + Vento) in corso...'):
            try:
                url_pv = f"https://re.jrc.ec.europa.eu/api/v5_2/seriescalc?lat={lat}&lon={lon}&startyear=2019&endyear=2019&pvcalculation=1&peakpower={p_PV}&loss=14&outputformat=json"
                resp_pv = requests.get(url_pv).json()
                df_pv = pd.DataFrame(resp_pv['outputs']['hourly'])
                p_solare_kw = (df_pv['P'] / 1000.0).values
                
                url_wind = "https://archive-api.open-meteo.com/v1/archive"
                params_wind = {
                    "latitude": lat, "longitude": lon,
                    "start_date": "2019-01-01", "end_date": "2019-12-31",
                    "hourly": "windspeed_100m", "wind_speed_unit": "ms", "timezone": "UTC"
                }
                resp_wind = requests.get(url_wind, params=params_wind).json()
                v_vento = np.array(resp_wind['hourly']['windspeed_100m'])
                
                p_eolico_kw = np.array([calcola_potenza_eolica_kw(v, p_Wind) for v in v_vento])
                min_len = min(len(p_solare_kw), len(p_eolico_kw))
                p_totale_kw = p_solare_kw[:min_len] + p_eolico_kw[:min_len]

                df_export = pd.DataFrame({'P_kW': p_totale_kw})
                df_export.to_csv(temp_filename + ".csv", index=False, header=['P_kW'], sep=';', decimal=',')
                tipo_file_passato = "NO"
            except Exception as e:
                st.error(f"❌ Errore nel download dei dati: {e}")
                st.stop()
    else:
        if file_csv is None:
            st.error("❌ Devi caricare un CSV Base!")
            st.stop()
        with open(temp_filename + ".csv", "wb") as f:
            f.write(file_csv.getbuffer())

    # --- FASE 2: MOTORE MATEMATICO E INIEZIONE EXTRA ---
    with st.spinner('Calcolo in corso nel motore matematico... L\'ottimizzazione è attiva.'):
        
        potenza_nominale_totale = p_PV + p_Wind + p_Extra
        
        # Inizializziamo il motore per fargli leggere il file Base
        analisi1 = Analisi_tecnica(
            file_csv=temp_filename, tipo_file=tipo_file_passato, p_PV=potenza_nominale_totale, dP_el=dP_el, 
            batteria=batteria, dP_bat=dP_bat, min_batt=min_batt, max_batteria=max_batteria, 
            lingua=lingua, eff_batt=eff_batt, min_elet=min_elet
        )
        
        # MAGIC TRICK: Se l'utente ha inserito un file Extra, lo fondiamo con l'array del motore prima di partire!
        if usa_extra and file_extra is not None:
            try:
                df_extra = pd.read_csv(file_extra, sep=None, engine='python')
                # Estraiamo la prima colonna forzandola in numero
                extra_vals = pd.to_numeric(df_extra.iloc[:, 0].astype(str).str.replace(',', '.'), errors='coerce').fillna(0).values
                
                # Sommiamo l'array base con l'array extra
                min_l = min(len(analisi1.E_PV), len(extra_vals))
                analisi1.E_PV[:min_l] += extra_vals[:min_l]
                analisi1.e_pv = np.sum(analisi1.E_PV) # Aggiorniamo il totale energia
            except Exception as e:
                st.error(f"❌ Errore nella lettura del file Extra: {e}")
                st.stop()

        analisi1.progress_bar = st.progress(0)
        analisi1.status_text = st.empty()
        
        # Avviamo il motore tecnico sui dati fusi!
        analisi1.run_analysis()
        if os.path.exists(temp_filename + ".csv"): os.remove(temp_filename + ".csv")
            
        top_progetti_tuples = []
        dati_scatter = []
        
        def salva_memoria(an_fin, andamenti_tecnici, val_batteria, cur_id):
            van = an_fin.VAN if an_fin.VAN is not None else -float('inf')
            lcoh_calc = an_fin.investimento / (an_fin.ProdAnnuaIdrogkg * DurPianEcon) if an_fin.ProdAnnuaIdrogkg > 0 else 0
            
            dati_scatter.append({
                'ID_Progetto': cur_id,
                'VAN [€]': van, 
                'TIR [%]': an_fin.TIR, 
                'LCOH Semplificato [€/kg]': lcoh_calc,
                'Investimento CAPEX [€]': an_fin.investimento,
                'Produzione Idrogeno [kg]': an_fin.ProdAnnuaIdrogkg, 
                'Taglia Elettrolizzatore [kW]': an_fin.PotEle,
                'Taglia Batteria [kWh]': val_batteria, 
                'Energia Sprecata (Curtailment) [kWh]': an_fin.ProdElettVend, 
                'Capacity Factor [%]': an_fin.CapFac,
            })
            top_progetti_tuples.append((van, an_fin, andamenti_tecnici))
            top_progetti_tuples.sort(key=lambda x: x[0], reverse=True)
            return top_progetti_tuples[:n_progetti]

        prog_bar_fin = st.progress(0)
        status_fin = st.empty()
        tot_progetti = len(analisi1.P_elc) if batteria == "NO" else len(analisi1.potenza_batt)

        # Analisi Finanziaria
        if batteria == "NO":
            for index, (el1, el2, el3, el4, el5, el6, el7) in enumerate(zip(analisi1.CF, analisi1.P_elc, analisi1.E_H2, analisi1.E_im, analisi1.M_H2, analisi1.andamenti, analisi1.spegn_giorn)):
                # ATTENZIONE: Passiamo a ImpPV1 SOLO la vera potenza solare per il calcolo dei costi corretti
                an_fin = Analisi_finanziaria(Terr=Terr, OpeE=OpeE, ImpPV1=p_PV, ImpPV1eurokW=ImpPV1eurokW, EletteuroKW=EletteuroKW, CompreuroKW=CompreuroKW, AccuE=0, AccuEeurokW=AccuEeurokW, idrogstocperc=0.1, StazzRif=StazzRif, SpeTOpere=0, BombSto=0, LavoImp=0, CarrEll=0, CapFac=el1, PotEle=el2, tassoDEN=0.005, ProdAnnuaIdrogkg=el5, bar=500, costlitroacqua=0.025, costounitariostoccaggio=costounitariostoccaggio, PercEserImp=0.005, Percentimpianti=0.0025, PercentOpeEd=0.0005, SpesAmmGen=2000, Affitto=0, CostiPersonal=0, AltriCost=0, IVAsualtriCost="NO", DurPianEcon=DurPianEcon, inflazione=inflazione, inflazionePrezzoElet=0.015, inflazioneIdrog=0.01, tassoVAN=tassoVAN, incentpubb=2.0, duratincentpubb=20, prezzoindrogeno=prezzoindrogeno, ProdElettVend=el4, EnergiaAutocons=el3, prezzoElett=0.1, ContrPubb=0.9, DebitoSenior=0.0, DurDebitoSenior=10, tassoDebito=0.05, FreqPagamenti=1, tassoPonte=0, DurataPonte=0, aliquoMedia=0.275, MaxInterssDed=0.3, lingua=lingua, Perciva=0.22, spegn_giorn=el7)
                # Aggiungiamo i costi delle altre fonti
                an_fin.investimento += (p_Wind * ImpWind1eurokW) + (p_Extra * ImpExtra1eurokW)
                an_fin.RUN()
                top_progetti_tuples = salva_memoria(an_fin, el6, 0, index)
                prog_bar_fin.progress((index + 1) / tot_progetti)
                status_fin.text(f"Analisi Finanziaria: {((index + 1) / tot_progetti)*100:.1f}%")
                
        elif batteria == "SI":
            for index, (el1, el2, el3, el4, el5, el6, el7, el8) in enumerate(zip(analisi1.potenza_batt, analisi1.CF, analisi1.potenza_elett, analisi1.E_H2, analisi1.E_im, analisi1.M_H2, analisi1.andamenti, analisi1.spegn_giorn)):
                an_fin = Analisi_finanziaria(Terr=Terr, OpeE=OpeE, ImpPV1=p_PV, ImpPV1eurokW=ImpPV1eurokW, EletteuroKW=EletteuroKW, CompreuroKW=CompreuroKW, AccuE=el1, AccuEeurokW=AccuEeurokW, idrogstocperc=0.1, StazzRif=StazzRif, SpeTOpere=0, BombSto=0, LavoImp=0, CarrEll=0, CapFac=el2, PotEle=el3, tassoDEN=0.005, ProdAnnuaIdrogkg=el6, bar=500, costlitroacqua=0.025, costounitariostoccaggio=costounitariostoccaggio, PercEserImp=0.005, Percentimpianti=0.0025, PercentOpeEd=0.0005, SpesAmmGen=2000, Affitto=0, CostiPersonal=0, AltriCost=0, IVAsualtriCost="NO", DurPianEcon=DurPianEcon, inflazione=inflazione, inflazionePrezzoElet=0.015, inflazioneIdrog=0.01, tassoVAN=tassoVAN, incentpubb=2.0, duratincentpubb=20, prezzoindrogeno=prezzoindrogeno, ProdElettVend=el5, EnergiaAutocons=el4, prezzoElett=0.1, ContrPubb=0.9, DebitoSenior=0.0, DurDebitoSenior=10, tassoDebito=0.05, FreqPagamenti=1, tassoPonte=0, DurataPonte=0, aliquoMedia=0.275, MaxInterssDed=0.3, lingua=lingua, Perciva=0.22, spegn_giorn=el8)
                an_fin.investimento += (p_Wind * ImpWind1eurokW) + (p_Extra * ImpExtra1eurokW)
                an_fin.RUN()
                top_progetti_tuples = salva_memoria(an_fin, el7, el1, index)
                prog_bar_fin.progress((index + 1) / tot_progetti)
                status_fin.text(f"Analisi Finanziaria: {((index + 1) / tot_progetti)*100:.1f}%")

        prog_bar_fin.empty()
        status_fin.empty()
        analisi1.progress_bar.empty()
        analisi1.status_text.empty()
        
        top_progetti = []
        andamenti_top = []
        for item in top_progetti_tuples:
            an_fin_obj = item[1]
            an_fin_obj.costruzione_tabelle()
            top_progetti.append(an_fin_obj)
            andamenti_top.append(item[2])

        st.session_state['simulazione_completata'] = True
        st.session_state['df_tutti'] = pd.DataFrame(dati_scatter).dropna()
        st.session_state['top_progetti'] = top_progetti
        st.session_state['andamenti_top'] = andamenti_top
        st.session_state['batteria_usata'] = batteria

# ==========================================
# DASHBOARD DEI RISULTATI
# ==========================================
if st.session_state.get('simulazione_completata', False):
    st.success("✅ Dati pronti per l'esplorazione!")
    
    df_tutti = st.session_state['df_tutti']
    top_progetti = st.session_state['top_progetti']
    andamenti_top = st.session_state['andamenti_top']
    batteria_usata = st.session_state['batteria_usata']

    st.divider()
    tab1, tab2, tab3 = st.tabs(["📊 Tabelle Finanziarie", "⚡ Flussi Energetici (8760 ore)", "🎯 Frontiera di Pareto INTERATTIVA"])

    with tab1:
        st.subheader("Sommario Migliori Configurazioni (Per VAN)")
        df_sommario = pd.DataFrame({
            'VAN [€]': [p.VAN for p in top_progetti],
            'TIR [%]': [p.TIR for p in top_progetti],
            'PAYBACK [Anni]': [p.PAYBACK for p in top_progetti],
            'Taglia Elettrolizzatore [kW]': [p.PotEle for p in top_progetti],
            'Taglia Batteria [kWh]': [p.AccuE for p in top_progetti] if batteria_usata == "SI" else [0]*len(top_progetti),
            'Produzione Idrogeno [kg/anno]': [p.ProdAnnuaIdrogkg for p in top_progetti]
        }).T
        df_sommario.columns = [f"Progetto {i+1}" for i in range(len(top_progetti))]
        st.dataframe(df_sommario, use_container_width=True)

    with tab2:
        st.subheader("Andamento Orario dei Flussi (Intero Anno - 8760 Ore)")
        scelta_ene = st.selectbox("Seleziona Progetto per Grafico Flussi:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
        andamenti = andamenti_top[scelta_ene]
        
        ore = np.arange(len(andamenti[0]))
        fig_flussi = go.Figure()
        
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[3], name='Rinnovabili Prodotte', line=dict(color='#92D050')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[2], name='Energia Sprecata (Curtailment)', line=dict(color='#FFC000')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[0], name='In Elettrolizzatore', line=dict(color='#00AF50')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[5], name='Taglia Elettrolizzatore', line=dict(color='gray', dash='dash')))
        if batteria_usata == "SI":
            fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[6], name='Livello Batteria', line=dict(color='orange')))
            
        fig_flussi.update_layout(xaxis=dict(rangeslider=dict(visible=True), type="-"))
        st.plotly_chart(fig_flussi, use_container_width=True)

    with tab3:
        st.subheader("Esploratore Interattivo di Pareto")
        st.markdown("👇 **Clicca su un pallino qualsiasi nel grafico** per vederne i dettagli tecnici ed economici.")
        
        col_x, col_y, col_color = st.columns(3)
        colonne_disp = [c for c in df_tutti.columns if c != 'ID_Progetto']
        
        x_var = col_x.selectbox("Asse X (Obiettivo 1)", colonne_disp, index=colonne_disp.index('Investimento CAPEX [€]'))
        y_var = col_y.selectbox("Asse Y (Obiettivo 2)", colonne_disp, index=colonne_disp.index('VAN [€]'))
        color_var = col_color.selectbox("Colore punti", colonne_disp, index=colonne_disp.index('Taglia Elettrolizzatore [kW]'))

        max_x = False if any(word in x_var.lower() for word in ['costo', 'lcoh', 'investimento']) else True
        max_y = False if any(word in y_var.lower() for word in ['costo', 'lcoh', 'investimento']) else True

        df_sorted = df_tutti.sort_values(by=x_var, ascending=not max_x)
        pareto_front = []
        best_y = -float('inf') if max_y else float('inf')
        
        for _, row in df_sorted.iterrows():
            current_y = row[y_var]
            if (max_y and current_y >= best_y) or (not max_y and current_y <= best_y):
                pareto_front.append(row)
                best_y = current_y
        df_pareto = pd.DataFrame(pareto_front)

        fig_sa = px.scatter(
            df_tutti, x=x_var, y=y_var, color=color_var,
            color_continuous_scale=["#07d5df", "#f407fe"],
            custom_data=['ID_Progetto']
        )
        
        if not df_pareto.empty:
            df_pareto = df_pareto.sort_values(by=x_var)
            fig_sa.add_trace(go.Scatter(
                x=df_pareto[x_var], y=df_pareto[y_var],
                mode='lines', line=dict(color='red', width=3, dash='dash'),
                name='Frontiera di Pareto', hoverinfo='skip'
            ))

        fig_sa.update_traces(marker=dict(size=9, opacity=0.8), selector=dict(mode='markers'))
        
        selezione = st.plotly_chart(fig_sa, use_container_width=True, on_select="rerun", selection_mode="points")

        if selezione and selezione.selection.points:
            id_selezionato = selezione.selection.points[0]["customdata"][0]
            dettagli = df_tutti[df_tutti['ID_Progetto'] == id_selezionato].iloc[0]
            
            st.success("🎯 **Progetto Selezionato dal Grafico:**")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("💰 VAN", f"€ {dettagli['VAN [€]']:,.0f}")
            c2.metric("⚖️ LCOH", f"€ {dettagli['LCOH Semplificato [€/kg]']:.2f} /kg")
            c3.metric("⚡ Elettrolizzatore", f"{dettagli['Taglia Elettrolizzatore [kW]']:,.0f} kW")
            c4.metric("🔋 Batteria", f"{dettagli['Taglia Batteria [kWh]']:,.0f} kWh")
            
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("💨 Produzione H2", f"{dettagli['Produzione Idrogeno [kg]']:,.0f} kg/anno")
            c6.metric("🗑️ Energia Sprecata", f"{dettagli['Energia Sprecata (Curtailment) [kWh]']:,.0f} kWh")
            c7.metric("💶 Investimento", f"€ {dettagli['Investimento CAPEX [€]']:,.0f}")
            c8.metric("📈 Capacity Factor", f"{dettagli['Capacity Factor [%]']:.1f} %")
        else:
            st.info("👆 Clicca su un pallino per vedere i dettagli esatti di quella configurazione.")
