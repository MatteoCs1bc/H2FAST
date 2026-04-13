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
# BARRA LATERALE: PARAMETRI
# ==========================================
st.sidebar.header("📁 Preferenze")
lingua = st.sidebar.selectbox("Lingua output", ["ITA", "ENG"])
n_progetti = st.sidebar.number_input("Mostra i Top N Progetti", value=5, min_value=1)

st.sidebar.header("⚙️ Parametri Tecnici")
p_PV = st.sidebar.number_input("PV system peak power [kWp]", value=1000.0)
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
# SCHERMATA PRINCIPALE: STEP 1 (SORGENTE DATI)
# ==========================================
st.header("📍 Step 1: Sorgente Dati Fotovoltaico")
metodo_dati = st.radio("Come vuoi inserire i dati di produzione solare?", ["Mappa Interattiva (Download da PVGIS)", "Carica file CSV (Manuale)"], horizontal=True)

lat, lon = None, None
file_csv = None
tipo_file = "SI"

if metodo_dati == "Mappa Interattiva (Download da PVGIS)":
    col_testo, col_mappa = st.columns([1, 2])
    with col_testo:
        st.markdown("### 1. Trova la zona")
        luogo = st.text_input("Cerca una città per centrare la mappa:", "Roma, Italia")
        st.markdown("### 2. Piazza il Pin")
        st.info("👉 **Clicca fisicamente un punto sulla mappa** a destra per selezionare le coordinate esatte del tuo impianto.")
        
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
        st.warning("⚠️ In attesa di selezione: Clicca sulla mappa prima di avviare la simulazione.")

else:
    st.markdown("### Caricamento Manuale")
    file_csv = st.file_uploader("Carica il tuo file di serie oraria CSV", type=["csv"])
    tipo_file = st.selectbox("Il file proviene da PVGIS?", ["SI", "NO"])


# ==========================================
# AVVIO SIMULAZIONE E MOTORE
# ==========================================
st.divider()
st.header("⚙️ Step 2: Esecuzione")

if st.button("🚀 AVVIA SIMULAZIONE COMPLETA", use_container_width=True, type="primary"):
    temp_filename = "temp_input_orario"

    if metodo_dati == "Mappa Interattiva (Download da PVGIS)":
        if lat is None or lon is None:
            st.error("❌ Devi cliccare un punto sulla mappa prima di avviare!")
            st.stop()
        with st.spinner('📡 Download dei dati solari dai server PVGIS in corso...'):
            url = f"https://re.jrc.ec.europa.eu/api/v5_2/seriescalc?lat={lat}&lon={lon}&startyear=2019&endyear=2019&pvcalculation=1&peakpower={p_PV}&loss=14&outputformat=csv"
            response = requests.get(url)
            if response.status_code == 200:
                with open(temp_filename + ".csv", "wb") as f:
                    f.write(response.content)
            else:
                st.error("❌ Errore nel download dei dati da PVGIS.")
                st.stop()
    else:
        if file_csv is None:
            st.error("❌ Devi caricare un file CSV prima di avviare!")
            st.stop()
        with open(temp_filename + ".csv", "wb") as f:
            f.write(file_csv.getbuffer())

    with st.spinner('Calcolo in corso nel motore matematico... L\'ottimizzazione è attiva.'):
        
        analisi1 = Analisi_tecnica(
            file_csv=temp_filename, tipo_file=tipo_file, p_PV=p_PV, dP_el=dP_el, 
            batteria=batteria, dP_bat=dP_bat, min_batt=min_batt, max_batteria=max_batteria, 
            lingua=lingua, eff_batt=eff_batt, min_elet=min_elet
        )
        
        # Leghiamo la barra di caricamento a Streamlit
        analisi1.progress_bar = st.progress(0)
        analisi1.status_text = st.empty()
        
        analisi1.run_analysis()
        
        if os.path.exists(temp_filename + ".csv"):
            os.remove(temp_filename + ".csv")
            
        top_progetti_tuples = []
        dati_scatter = []
        
        def salva_memoria(an_fin, andamenti_tecnici, val_batteria):
            van = an_fin.VAN if an_fin.VAN is not None else -float('inf')
            # Calcolo un LCOH semplificato (CAPEX / Kg di H2 Prodotti nella vita utile) per l'analisi
            lcoh_calc = an_fin.investimento / (an_fin.ProdAnnuaIdrogkg * DurPianEcon) if an_fin.ProdAnnuaIdrogkg > 0 else 0
            
            dati_scatter.append({
                'VAN [€]': van, 
                'TIR [%]': an_fin.TIR, 
                'LCOH Semplificato [€/kg]': lcoh_calc,
                'Investimento CAPEX [€]': an_fin.investimento,
                'Produzione Idrogeno [kg]': an_fin.ProdAnnuaIdrogkg, 
                'Taglia Elettrolizzatore [kW]': an_fin.PotEle,
                'Taglia Batteria [kWh]': val_batteria, 
                'Capacity Factor [%]': an_fin.CapFac,
                'Spegnimenti [n]': an_fin.spegn_giorn
            })
            top_progetti_tuples.append((van, an_fin, andamenti_tecnici))
            top_progetti_tuples.sort(key=lambda x: x[0], reverse=True)
            return top_progetti_tuples[:n_progetti]

        prog_bar_fin = st.progress(0)
        status_fin = st.empty()
        tot_progetti = len(analisi1.P_elc) if batteria == "NO" else len(analisi1.potenza_batt)

        if batteria == "NO":
            for index, (el1, el2, el3, el4, el5, el6, el7) in enumerate(zip(analisi1.CF, analisi1.P_elc, analisi1.E_H2, analisi1.E_im, analisi1.M_H2, analisi1.andamenti, analisi1.spegn_giorn)):
                an_fin = Analisi_finanziaria(Terr=Terr, OpeE=OpeE, ImpPV1=p_PV, ImpPV1eurokW=ImpPV1eurokW, EletteuroKW=EletteuroKW, CompreuroKW=CompreuroKW, AccuE=0, AccuEeurokW=AccuEeurokW, idrogstocperc=0.1, StazzRif=StazzRif, SpeTOpere=0, BombSto=0, LavoImp=0, CarrEll=0, CapFac=el1, PotEle=el2, tassoDEN=0.005, ProdAnnuaIdrogkg=el5, bar=500, costlitroacqua=0.025, costounitariostoccaggio=costounitariostoccaggio, PercEserImp=0.005, Percentimpianti=0.0025, PercentOpeEd=0.0005, SpesAmmGen=2000, Affitto=0, CostiPersonal=0, AltriCost=0, IVAsualtriCost="NO", DurPianEcon=DurPianEcon, inflazione=inflazione, inflazionePrezzoElet=0.015, inflazioneIdrog=0.01, tassoVAN=tassoVAN, incentpubb=2.0, duratincentpubb=20, prezzoindrogeno=prezzoindrogeno, ProdElettVend=el4, EnergiaAutocons=el3, prezzoElett=0.1, ContrPubb=0.9, DebitoSenior=0.0, DurDebitoSenior=10, tassoDebito=0.05, FreqPagamenti=1, tassoPonte=0, DurataPonte=0, aliquoMedia=0.275, MaxInterssDed=0.3, lingua=lingua, Perciva=0.22, spegn_giorn=el7)
                an_fin.RUN()
                top_progetti_tuples = salva_memoria(an_fin, el6, 0)
                prog_bar_fin.progress((index + 1) / tot_progetti)
                status_fin.text(f"Analisi Finanziaria: {((index + 1) / tot_progetti)*100:.1f}%")
                
        elif batteria == "SI":
            for index, (el1, el2, el3, el4, el5, el6, el7, el8) in enumerate(zip(analisi1.potenza_batt, analisi1.CF, analisi1.potenza_elett, analisi1.E_H2, analisi1.E_im, analisi1.M_H2, analisi1.andamenti, analisi1.spegn_giorn)):
                an_fin = Analisi_finanziaria(Terr=Terr, OpeE=OpeE, ImpPV1=p_PV, ImpPV1eurokW=ImpPV1eurokW, EletteuroKW=EletteuroKW, CompreuroKW=CompreuroKW, AccuE=el1, AccuEeurokW=AccuEeurokW, idrogstocperc=0.1, StazzRif=StazzRif, SpeTOpere=0, BombSto=0, LavoImp=0, CarrEll=0, CapFac=el2, PotEle=el3, tassoDEN=0.005, ProdAnnuaIdrogkg=el6, bar=500, costlitroacqua=0.025, costounitariostoccaggio=costounitariostoccaggio, PercEserImp=0.005, Percentimpianti=0.0025, PercentOpeEd=0.0005, SpesAmmGen=2000, Affitto=0, CostiPersonal=0, AltriCost=0, IVAsualtriCost="NO", DurPianEcon=DurPianEcon, inflazione=inflazione, inflazionePrezzoElet=0.015, inflazioneIdrog=0.01, tassoVAN=tassoVAN, incentpubb=2.0, duratincentpubb=20, prezzoindrogeno=prezzoindrogeno, ProdElettVend=el5, EnergiaAutocons=el4, prezzoElett=0.1, ContrPubb=0.9, DebitoSenior=0.0, DurDebitoSenior=10, tassoDebito=0.05, FreqPagamenti=1, tassoPonte=0, DurataPonte=0, aliquoMedia=0.275, MaxInterssDed=0.3, lingua=lingua, Perciva=0.22, spegn_giorn=el8)
                an_fin.RUN()
                top_progetti_tuples = salva_memoria(an_fin, el7, el1)
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

    st.success("✅ Calcolo completato!")
    df_tutti = pd.DataFrame(dati_scatter).dropna()

    # ==========================================
    # DASHBOARD DEI RISULTATI
    # ==========================================
    st.divider()
    tab1, tab2, tab3 = st.tabs(["📊 Tabelle Finanziarie", "⚡ Flussi Energetici (8760 ore)", "🎯 Frontiera di Pareto"])

    with tab1:
        st.subheader("Sommario Migliori Configurazioni (Per VAN)")
        df_sommario = pd.DataFrame({
            'VAN [€]': [p.VAN for p in top_progetti],
            'TIR [%]': [p.TIR for p in top_progetti],
            'PAYBACK [Anni]': [p.PAYBACK for p in top_progetti],
            'Taglia Elettrolizzatore [kW]': [p.PotEle for p in top_progetti],
            'Taglia Batteria [kWh]': [p.AccuE for p in top_progetti] if batteria == "SI" else [0]*len(top_progetti),
            'Produzione Idrogeno [kg/anno]': [p.ProdAnnuaIdrogkg for p in top_progetti]
        }).T
        df_sommario.columns = [f"Progetto {i+1}" for i in range(len(top_progetti))]
        st.dataframe(df_sommario, use_container_width=True)

        col_sel1, col_sel2 = st.columns(2)
        scelta_eco = col_sel1.selectbox("Conto Economico:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
        st.dataframe(top_progetti[scelta_eco].dfContoEconomico, use_container_width=True)
        
        scelta_cassa = col_sel2.selectbox("Flussi di Cassa:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
        st.dataframe(top_progetti[scelta_cassa].dfFlussiMonetari, use_container_width=True)

    with tab2:
        st.subheader("Andamento Orario dei Flussi (Intero Anno - 8760 Ore)")
        st.info("💡 Usa la barra grigia sotto il grafico per **zoomare** su specifici giorni o mesi!")
        scelta_ene = st.selectbox("Seleziona Progetto per Grafico Flussi:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
        andamenti = andamenti_top[scelta_ene]
        
        ore = np.arange(len(andamenti[0]))
        fig_flussi = go.Figure()
        
        # Usiamo scattergl (WebGL) per gestire migliaia di punti fluidamente
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[3], name='Prodotta FV', line=dict(color='#92D050')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[2], name='Immessa in Rete', line=dict(color='#FFC000')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[0], name='In Elettrolizzatore', line=dict(color='#00AF50')))
        fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[5], name='Taglia Elettrolizzatore', line=dict(color='gray', dash='dash')))
        if batteria == "SI":
            fig_flussi.add_trace(go.Scattergl(x=ore, y=andamenti[6], name='Livello Batteria', line=dict(color='orange')))
            
        fig_flussi.update_layout(
            xaxis_title="Ora dell'anno", 
            yaxis_title="Energia [kWh] / Potenza [kW]",
            xaxis=dict(rangeslider=dict(visible=True), type="-") # Aggiunge la barra dello zoom!
        )
        st.plotly_chart(fig_flussi, use_container_width=True)

    with tab3:
        st.subheader("Analisi di Sensitività e Frontiera di Pareto")
        st.markdown("Scegli due obiettivi in contrasto (es. Investimento vs VAN). La **linea rossa** ti mostrerà i progetti che offrono il miglior compromesso in assoluto, scartando quelli inefficienti.")
        
        col_x, col_y, col_color = st.columns(3)
        colonne_disp = df_tutti.columns.tolist()
        
        x_var = col_x.selectbox("Asse X (Obiettivo 1)", colonne_disp, index=colonne_disp.index('Investimento CAPEX [€]'))
        y_var = col_y.selectbox("Asse Y (Obiettivo 2)", colonne_disp, index=colonne_disp.index('VAN [€]'))
        color_var = col_color.selectbox("Colore punti", colonne_disp, index=colonne_disp.index('Taglia Elettrolizzatore [kW]'))

        # Logica per capire se l'utente vuole massimizzare o minimizzare quella variabile
        # (Se è Costo/LCOH/Investimento di solito si vuole minimizzare, altrimenti massimizzare)
        max_x = False if any(word in x_var.lower() for word in ['costo', 'lcoh', 'investimento']) else True
        max_y = False if any(word in y_var.lower() for word in ['costo', 'lcoh', 'investimento']) else True

        # Algoritmo di Frontiera di Pareto
        df_sorted = df_tutti.sort_values(by=x_var, ascending=not max_x)
        pareto_front = []
        best_y = -float('inf') if max_y else float('inf')
        
        for _, row in df_sorted.iterrows():
            current_y = row[y_var]
            if max_y:
                if current_y >= best_y:
                    pareto_front.append(row)
                    best_y = current_y
            else:
                if current_y <= best_y:
                    pareto_front.append(row)
                    best_y = current_y
                    
        df_pareto = pd.DataFrame(pareto_front)

        # Grafico
        fig_sa = px.scatter(
            df_tutti, x=x_var, y=y_var, color=color_var,
            color_continuous_scale=["#07d5df", "#f407fe"],
            hover_data=['Taglia Elettrolizzatore [kW]', 'Taglia Batteria [kWh]' if batteria=="SI" else 'Capacity Factor [%]']
        )
        
        # Disegniamo la linea rossa di Pareto sopra i puntini
        if not df_pareto.empty:
            df_pareto = df_pareto.sort_values(by=x_var)
            fig_sa.add_trace(go.Scatter(
                x=df_pareto[x_var], y=df_pareto[y_var],
                mode='lines', line=dict(color='red', width=3, dash='dash'),
                name='Frontiera di Pareto'
            ))

        fig_sa.update_traces(marker=dict(size=8, opacity=0.8), selector=dict(mode='markers'))
        fig_sa.update_layout(title=f"Trade-off: {x_var} vs {y_var}")
        st.plotly_chart(fig_sa, use_container_width=True)
