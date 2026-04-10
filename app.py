import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os

# Importiamo LE TUE classi originali e intatte dal tuo file!
from motore_h2fast import Analisi_tecnica, Analisi_finanziaria

st.set_page_config(page_title="H2FAsT Simulator", layout="wide")
st.title("Simulatore H2FAsT - Dashboard Interattiva 🏭")

# --- BARRA LATERALE ---
st.sidebar.header("📁 Dati Input")
lingua = st.sidebar.selectbox("Lingua", ["ITA", "ENG"])
file_csv = st.sidebar.file_uploader("Carica file serie oraria CSV", type=["csv"])
tipo_file = st.sidebar.selectbox("Il file proviene da PVGIS?", ["SI", "NO"])

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
    if file_csv is None:
        st.warning("⚠️ Carica prima un file CSV per poter procedere.")
    else:
        # Salviamo temporaneamente il file per il tuo motore
        temp_filename = "temp_input_orario"
        with open(temp_filename + ".csv", "wb") as f:
            f.write(file_csv.getbuffer())

        with st.spinner('Calcolo in corso nel motore originale. Il server sta ottimizzando la memoria...'):
            
            # 1. CHIAMIAMO LA TUA ANALISI TECNICA
            analisi1 = Analisi_tecnica(
                file_csv=temp_filename, tipo_file=tipo_file, p_PV=p_PV, dP_el=dP_el, 
                batteria=batteria, dP_bat=dP_bat, min_batt=min_batt, max_batteria=max_batteria, 
                lingua=lingua, eff_batt=eff_batt, min_elet=min_elet
            )
            analisi1.run_analysis()
            
            if os.path.exists(temp_filename + ".csv"):
                os.remove(temp_filename + ".csv")
                
            # 2. CHIAMIAMO L'ANALISI FINANZIARIA (Gestione Memoria Ottimizzata)
            top_progetti_tuples = []
            dati_scatter = [] # Dizionario leggero per i grafici di sensitivity
            
            # Funzione interna per salvare solo i progetti migliori e svuotare la RAM
            def salva_memoria(an_fin, andamenti_tecnici, val_batteria):
                van = an_fin.VAN if an_fin.VAN is not None else -float('inf')
                # Dati leggeri per i grafici globali
                dati_scatter.append({
                    'VAN [€]': van, 'TIR [%]': an_fin.TIR, 'Taglia Elettrolizzatore [kW]': an_fin.PotEle,
                    'Taglia Batteria [kWh]': val_batteria, 'Capacity Factor [%]': an_fin.CapFac,
                    'Produzione Idrogeno [kg]': an_fin.ProdAnnuaIdrogkg, 'Investimento [€]': an_fin.investimento,
                    'Spegnimenti [n]': an_fin.spegn_giorn
                })
                # Salvataggio Top N
                top_progetti_tuples.append((van, an_fin, andamenti_tecnici))
                top_progetti_tuples.sort(key=lambda x: x[0], reverse=True)
                return top_progetti_tuples[:n_progetti] # Svuota i progetti in eccesso!

            # Avvio ciclo progetti
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
            
            # Estraiamo i Top N definitivi
            top_progetti = []
            andamenti_top = []
            for item in top_progetti_tuples:
                an_fin_obj = item[1]
                an_fin_obj.costruzione_tabelle() # Creiamo i dataframe pandas solo per i migliori!
                top_progetti.append(an_fin_obj)
                andamenti_top.append(item[2])

        st.success("✅ Calcolo completato!")

        # Creiamo il dataframe leggero per i grafici globali
        df_tutti = pd.DataFrame(dati_scatter).dropna()

        # =================================================================
        # DASHBOARD INTERATTIVA (TABS)
        # =================================================================
        tab1, tab2, tab3 = st.tabs(["📊 Tabelle Finanziarie", "⚡ Flussi Energetici (Grafico)", "📈 Analisi Relazioni (Scatter & SA)"])

        # TAB 1: TABELLE
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
            st.dataframe(df_sommario.style.format("{:,.2f}"), use_container_width=True)

            col_sel1, col_sel2 = st.columns(2)
            scelta_eco = col_sel1.selectbox("Seleziona Progetto per Conto Economico:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
            st.write("**Conto Economico**")
            st.dataframe(top_progetti[scelta_eco].dfContoEconomico, use_container_width=True)
            
            scelta_cassa = col_sel2.selectbox("Seleziona Progetto per Flussi di Cassa:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
            st.write("**Flussi di Cassa**")
            st.dataframe(top_progetti[scelta_cassa].dfFlussiMonetari, use_container_width=True)

        # TAB 2: FLUSSI ENERGETICI (Line Chart Plotly)
        with tab2:
            st.subheader("Andamento Orario dei Flussi (Prime 200 ore per fluidità)")
            scelta_ene = st.selectbox("Seleziona Progetto per Grafico Flussi:", range(len(top_progetti)), format_func=lambda x: f"Progetto {x+1}")
            andamenti = andamenti_top[scelta_ene]
            
            # andamenti: 0=AutoH2, 1=Massa, 2=Immessa, 3=PV, 4=PotMinEl, 5=PotMaxEl (6=Batt, 7=MaxBatt se presenti)
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

        # TAB 3: SCATTER E SENSITIVITY ANALYSIS (Plotly Express)
        with tab3:
            st.subheader("Grafici di Relazione Personalizzabili (Analisi su TUTTI i progetti)")
            st.markdown("Crea la tua **Frontiera di Pareto** o **Analisi di Sensitività** selezionando le variabili. I colori creano l'effetto gradiente in automatico.")
            
            col_x, col_y, col_color = st.columns(3)
            colonne_disp = df_tutti.columns.tolist()
            
            x_var = col_x.selectbox("Asse X", colonne_disp, index=colonne_disp.index('Taglia Elettrolizzatore [kW]'))
            y_var = col_y.selectbox("Asse Y", colonne_disp, index=colonne_disp.index('VAN [€]'))
            color_var = col_color.selectbox("Colore (Sensitività)", colonne_disp, index=colonne_disp.index('Taglia Batteria [kWh]') if batteria=="SI" else 0)

            fig_sa = px.scatter(
                df_tutti, x=x_var, y=y_var, color=color_var,
                color_continuous_scale=["#07d5df", "#f407fe"], # I tuoi colori azzurro-magenta
                hover_data=['VAN [€]', 'TIR [%]', 'Capacity Factor [%]']
            )
            fig_sa.update_traces(marker=dict(size=8, opacity=0.8))
            fig_sa.update_layout(title=f"Analisi: {x_var} vs {y_var} (Colore: {color_var})")
            
            st.plotly_chart(fig_sa, use_container_width=True)
