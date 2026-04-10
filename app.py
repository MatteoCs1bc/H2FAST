import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import requests
from geopy.geocoders import Nominatim

# Importiamo LE TUE classi originali e intatte dal tuo file!
from motore_h2fast import Analisi_tecnica, Analisi_finanziaria

st.set_page_config(page_title="H2FAsT Simulator", layout="wide")
st.title("Simulatore H2FAsT - Dashboard Interattiva 🏭")

# --- BARRA LATERALE ---
st.sidebar.header("🌍 Sorgente Dati Solari")
metodo_dati = st.sidebar.radio("Come vuoi inserire i dati del fotovoltaico?", ["Scarica da PVGIS (Mappa)", "Carica file CSV (Manuale)"])

if metodo_dati == "Scarica da PVGIS (Mappa)":
    luogo = st.sidebar.text_input("Inserisci il luogo (es. Roma, Italia)", value="Roma, Italia")
    geolocator = Nominatim(user_agent="h2fast_simulator")
    try:
        location = geolocator.geocode(luogo)
        if location:
            lat, lon = location.latitude, location.longitude
            st.sidebar.success(f"📍 Trovato: {location.address}")
            # Mostriamo la mappa interattiva!
            df_mappa = pd.DataFrame({'lat': [lat], 'lon': [lon]})
            st.sidebar.map(df_mappa, zoom=5)
            tipo_file = "SI" # Diciamo al tuo motore che è formato PVGIS
        else:
            st.sidebar.error("❌ Luogo non trovato. Riprova.")
            lat, lon = None, None
    except:
        st.sidebar.error("❌ Errore di connessione al servizio mappe.")
        lat, lon = None, None
else:
    file_csv = st.sidebar.file_uploader("Carica file serie oraria CSV", type=["csv"])
    tipo_file = st.sidebar.selectbox("Il file proviene da PVGIS?", ["SI", "NO"])

st.sidebar.header("📁 Preferenze Lingua")
lingua = st.sidebar.selectbox("Lingua", ["ITA", "ENG"])

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
n_progetti = st.sidebar.number_input("Mostra i Top N Progetti", value=5, min_value=1)

# --- AVVIO PROCESSO ---
if st.button("🚀 Avvia Simulazione", use_container_width=True):
    temp_filename = "temp_input_orario"

    # SCARICAMENTO AUTOMATICO O CARICAMENTO MANUALE
    if metodo_dati == "Scarica da PVGIS (Mappa)":
        if lat is None or lon is None:
            st.warning("⚠️ Inserisci un luogo valido prima di avviare.")
            st.stop()
            
        with st.spinner('📡 Download dei dati solari dai server europei PVGIS in corso...'):
            # Chiediamo a PVGIS l'anno 2019 (per avere esattamente 8760 ore, non bisestile)
            url = f"https://re.jrc.ec.europa.eu/api/v5_2/seriescalc?lat={lat}&lon={lon}&startyear=2019&endyear=2019&pvcalculation=1&peakpower={p_PV}&loss=14&outputformat=csv"
            response = requests.get(url)
            if response.status_code == 200:
                with open(temp_filename + ".csv", "wb") as f:
                    f.write(response.content)
            else:
                st.error("❌ Errore nel download dei dati da PVGIS. Riprova più tardi.")
                st.stop()
    else:
        if file_csv is None:
            st.warning("⚠️ Carica prima un file CSV per poter procedere.")
            st.stop()
        with open(temp_filename + ".csv", "wb") as f:
            f.write(file_csv.getbuffer())

    # INIZIO CALCOLO NEL MOTORE ORIGINALE
    with st.spinner('⚙️ Calcolo in corso nel motore matematico... (L\'ottimizzazione memoria è attiva)'):
        
        # 1. CHIAMIAMO LA TUA ANALISI TECNICA
        analisi1 = Analisi_tecnica(
            file_csv=temp_filename, tipo_file=tipo_file, p_PV=p_PV, dP_el=dP_el, 
            batteria=batteria, dP_bat=dP_bat, min_batt=min_batt, max_batteria=max_batteria, 
            lingua=lingua, eff_batt=eff_batt, min_elet=min_elet
        )
        analisi1.run_analysis()
        
        if os.path.exists(temp_filename + ".csv"):
            os.remove(temp_filename + ".csv")
            
        # 2. CHIAMIAMO L'ANALISI FINANZIARIA
        top_progetti_tuples = []
        dati_scatter = []
        
        def salva_memoria(an_fin, andamenti_tecnici, val_batteria):
            van = an_fin.VAN if an_fin.VAN is not None else -float('inf')
            dati_scatter.append({
                'VAN [€]': van, 'TIR [%]': an_fin.TIR, 'Taglia Elettrolizzatore [kW]': an_fin.PotEle,
                'Taglia Batteria [kWh]': val_batteria, 'Capacity Factor [%]': an_fin.CapFac,
                'Produzione Idrogeno [kg]': an_fin.ProdAnnuaIdrogkg, 'Investimento [€]': an_fin.investimento,
                'Spegnimenti [n]': an_fin.spegn_giorn
            })
            top_progetti_tuples.append((van, an_fin, andamenti_tecnici))
            top_progetti_tuples.sort(key=lambda x: x[0], reverse=True)
            return top_progetti_tuples[:n_progetti]

        prog_bar = st.progress(0)
        tot_progetti = len(analisi1.P_elc) if batteria == "NO" else len(analisi1.potenza_batt)

        if batteria == "NO":
            for index, (el1, el2, el3, el4, el5, el6, el7) in enumerate(zip(analisi1.CF, analisi1.P_elc, analisi1.E_H2, analisi1.E_im, analisi1.M_H2, analisi1.andamenti, analisi1.spegn_giorn)):
                an_fin = Analisi_finanziaria(Terr=Terr, OpeE=OpeE, ImpPV1=p_PV, ImpPV1eurokW=ImpPV1eurokW, EletteuroKW=EletteuroKW, CompreuroKW=CompreuroKW, AccuE=0, AccuEeurokW=AccuEeurokW, idrogstocperc=0.1, StazzRif=StazzRif, SpeTOpere=0, BombSto=0, LavoImp=0, CarrEll=0, CapFac=el1, PotEle=el2, tassoDEN=0.005, ProdAnnuaIdrogkg=el5, bar=500, costlitroacqua=0.025, costounitariostoccaggio=costounitariostoccaggio, PercEserImp=0.005, Percentimpianti=0.0025, PercentOpeEd=0.0005, SpesAmmGen=2000, Affitto=0, CostiPersonal=0, AltriCost=0, IVAsualtriCost="NO", DurPianEcon=DurPianEcon, inflazione=inflazione, inflazionePrezzoElet=0.015, inflazioneIdrog=0.01, tassoVAN=tassoVAN, incentpubb=2.0, duratincentpubb=20, prezzoindrogeno=prezzoindrogeno, ProdElettVend=el4, EnergiaAutocons=el3, prezzoElett=0.1, ContrPubb=0.9, DebitoSenior=0.0, DurDebitoSenior=10, tassoDebito=0.05, FreqPagamenti=1, tassoPonte=0, DurataPonte=0, aliquoMedia=0.275, MaxInterssDed=0.3, lingua=lingua, Perciva=0.22, spegn_giorn=el7)
                an_fin.RUN()
                top_progetti_tuples = salva_memoria(an_fin, el6, 0)
                prog_bar.progress((index + 1) / tot_progetti)
                
        elif batteria == "SI":
            for index, (el1, el2, el3, el4, el5, el6, el7, el8) in enumerate(zip(analisi1.potenza_batt, analisi1.CF, analisi1.potenza_elett, analisi1.E_H2, analisi1.E_im, analisi1.M_H2, analisi1.andamenti, analisi1.spegn_giorn)):
                an_fin = Analisi_finanziaria(Terr=Terr, OpeE=OpeE, ImpPV1=p_PV, ImpPV1eurokW=ImpPV1eurokW, EletteuroKW=EletteuroKW, CompreuroKW=CompreuroKW, AccuE=el1, AccuEeurokW=AccuEeurokW, idrogstocperc=0.1, StazzRif=StazzRif, SpeTOpere=0, BombSto=0, LavoImp=0, CarrEll=0, CapFac=el2, PotEle=el3, tassoDEN=0.005, ProdAnnuaIdrogkg=el6, bar=500, costlitroacqua=0.025, costounitariostoccaggio=costounitariostoccaggio, PercEserImp=0.005, Percentimpianti=0.0025, PercentOpeEd=0.0005, SpesAmmGen=2000, Affitto=0, CostiPersonal=0, AltriCost=0, IVAsualtriCost="NO", DurPianEcon=DurPianEcon, inflazione=inflazione, inflazionePrezzoElet=0.015, inflazioneIdrog=0.01, tassoVAN=tassoVAN, incentpubb=2.0, duratincentpubb=20, prezzoindrogeno=prezzoindrogeno, ProdElettVend=el5, EnergiaAutocons=el4, prezzoElett=0.1, ContrPubb=0.9, DebitoSenior=0.0, DurDebitoSenior=10, tassoDebito=0.05, FreqPagamenti=1, tassoPonte=0, DurataPonte=0, aliquoMedia=0.275, MaxInterssDed=0.3, lingua=lingua, Perciva=0.22, spegn_giorn=el8)
                an_fin.RUN()
                top_progetti_tuples = salva_memoria(an_fin, el7, el1)
                prog_bar.progress((index + 1) / tot_progetti)

        prog_bar.empty()
        
        top_progetti = []
        andamenti_top = []
        for item in top_progetti_tuples:
            an_fin_obj = item[1]
            an_fin_obj.costruzione_tabelle()
            top_progetti.append(an_fin_obj)
            andamenti_top.append(item[2])

    st.success("✅ Calcolo completato!")
    df_tutti = pd.DataFrame(dati_scatter).dropna()

    # --- DASHBOARD INTERATTIVA ---
    tab1, tab2, tab3 = st.tabs(["📊 Tabelle Finanziarie", "⚡ Flussi Energetici", "📈 Analisi di Sensitività e Pareto"])

    with tab1:
        st.subheader("Sommario Migliori Configurazioni")
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
        st.subheader("Andamento Orario dei Flussi (Prime 200 ore)")
        scelta_ene = st.selectbox("Progetto per Grafico Flussi:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
        andamenti = andamenti_top[scelta_ene]
        
        ore = np.arange(200)
        fig_flussi = go.Figure()
        fig_flussi.add_trace(go.Scatter(x=ore, y=andamenti[3][:200], name='Prodotta FV', line=dict(color='#92D050')))
        fig_flussi.add_trace(go.Scatter(x=ore, y=andamenti[2][:200], name='Immessa in Rete', line=dict(color='#FFC000')))
        fig_flussi.add_trace(go.Scatter(x=ore, y=andamenti[0][:200], name='In Elettrolizzatore', line=dict(color='#00AF50')))
        fig_flussi.add_trace(go.Scatter(x=ore, y=andamenti[5][:200], name='Taglia Elettrolizzatore', line=dict(color='gray', dash='dash')))
        if batteria == "SI":
            fig_flussi.add_trace(go.Scatter(x=ore, y=andamenti[6][:200], name='Livello Batteria', line=dict(color='orange')))
            
        fig_flussi.update_layout(xaxis_title="Ora dell'anno", yaxis_title="Energia [kWh] / Potenza [kW]")
        st.plotly_chart(fig_flussi, use_container_width=True)

    with tab3:
        st.subheader("Analisi di Sensitività e Frontiera di Pareto")
        st.markdown("Usa le tendine per esplorare visivamente come la taglia dell'elettrolizzatore o della batteria impattano sul VAN o sulla Produzione di Idrogeno di **tutti gli scenari simulati**.")
        col_x, col_y, col_color = st.columns(3)
        colonne_disp = df_tutti.columns.tolist()
        
        x_var = col_x.selectbox("Asse X", colonne_disp, index=colonne_disp.index('Taglia Elettrolizzatore [kW]'))
        y_var = col_y.selectbox("Asse Y", colonne_disp, index=colonne_disp.index('VAN [€]'))
        color_var = col_color.selectbox("Colore", colonne_disp, index=colonne_disp.index('Taglia Batteria [kWh]') if batteria=="SI" else 0)

        fig_sa = px.scatter(
            df_tutti, x=x_var, y=y_var, color=color_var,
            color_continuous_scale=["#07d5df", "#f407fe"],
            hover_data=['VAN [€]', 'TIR [%]', 'Capacity Factor [%]']
        )
        fig_sa.update_traces(marker=dict(size=8, opacity=0.8))
        fig_sa.update_layout(title=f"{x_var} vs {y_var}")
        st.plotly_chart(fig_sa, use_container_width=True)
