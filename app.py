import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import re
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
    
# Configuração da página
st.set_page_config(page_title="Laboratório Escola", layout="wide")

# ==========================================
# DICIONÁRIO MESTRE DE REFERÊNCIAS
# ==========================================
REFERENCIAS_MASTER = {
    "Hemácias": "4.5 - 5.9", 
    "Hemoglobina": "13.5 - 17.5", 
    "Hematócrito": "41 - 53", 
    "VCM": "80 - 100", 
    "HCM": "27 - 32", 
    "CHCM": "32 - 36", 
    "RDW": "11.5 - 14.5", 
    "Leucócitos Totais": "4000 - 11000", 
    "Bastonetes": "0 - 5", 
    "Segmentados": "45 - 70", 
    "Eosinófilos": "1 - 4", 
    "Basófilos": "0 - 1", 
    "Linfócitos": "20 - 45", 
    "Monócitos": "2 - 10", 
    "Plaquetas": "150000 - 450000",
    "Glicose": "70 - 99", 
    "Colesterol": "0 - 200", 
    "HDL": "40 - 100",
    "LDL": "0 - 130", 
    "Triglicerídeos": "0 - 150", 
    "Ureia": "10 - 50",
    "Creatinina": "0.6 - 1.3",
    "TGO": "0 - 40",
    "TGP": "0 - 41",
    "Albumina": "3.5 - 5.2",
    "Bilirrubinas Totais": "0 - 1.2",
    "Bilirrubina Direta": "0 - 0.3",
    "Bilirrubina Indireta": "0 - 0.9",
    "VHS": "0 - 20", 
    "Reticulócitos": "0.5 - 2.0",
    "Cultura": "Ausência de crescimento", 
    "HIV": "Não reagente", 
    "VDRL": "Não reagente",
    "Beta HCG": "Negativo", 
    "PCR": "< 6 mg/L", 
    "Tipagem Sanguínea": "ABO / Fator Rh",
    "Tempo de Sangria": "1 a 3 minutos",
    "Tempo de Coagulação": "5 a 10 minutos",
    "TP": "10 a 14 segundos",
    "KTTP": "24 a 36 segundos",
    "ASLO": "< 200 UI/mL",
    "Fator Reumatoide": "< 8 UI/mL",
    "HBsAg": "Não Reagente"
}

def verificar_alteracao(exame, valor_str):
    try:
        if exame not in REFERENCIAS_MASTER: return False
        valor_limpo = re.sub(r'[^\d.]', '', str(valor_str).replace(',', '.'))
        if not valor_limpo: return False
        valor = float(valor_limpo)
        ref = REFERENCIAS_MASTER[exame]
        limites = re.findall(r"[\d.]+", ref.replace(',', '.'))
        if len(limites) == 2:
            return valor < float(limites[0]) or valor > float(limites[1])
        elif len(limites) == 1:
            return valor > float(limites[0])
        return False
    except: return False


SETOR_EXAMES = {
    "Bioquímica": ["Glicose", "Colesterol", "HDL", "LDL", "Triglicerídeos", "Ureia", "Creatinina", "TGO", "TGP", "Albumina", "Bilirrubinas"],
    "Hematologia": ["Hemograma", "Plaquetas", "VHS", "Reticulócitos", "Tipagem Sanguínea", "Tempo de Sangria", "Tempo de Coagulação", "TP", "KTTP"],
    "Uroanálise": ["EQU"],
    "Microbiologia": ["Cultura", "Antibiograma", "Coloração de Gram", "Pesquisa de BK"],
    "Imunologia": ["HIV", "VDRL", "Beta HCG", "PCR", "ASLO", "Fator Reumatoide", "HBsAg"],
    "Parasitologia": ["EPF", "Pesquisa de sangue oculto", "Pesquisa de leucócitos fecais", "Coprocultura"]
}

try:
    SENHA_PROFESSOR = st.secrets["senha_professor"]
except:
    st.error("Erro: Configure 'senha_professor' nos Secrets.")
    st.stop()

def salvar_registro(dados):
    arquivo = "laudos_registro.xlsx"

    # Transformar resultados em texto legível
    detalhes_lista = []
    for r in dados["Resultados"]:
        detalhes_lista.append(f"{r[0]}# {r[1]}# {r[2]}")

    dados["Resultados"] = " | ".join(detalhes_lista)

    # Criar arquivo se não existir
    if os.path.exists(arquivo):
        df = pd.read_excel(arquivo)
    else:
        df = pd.DataFrame(columns=[
            "Data","Paciente","Sexo","Setores","Aluno","Supervisor","Resultados"
        ])

    df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)

    df.to_excel(arquivo, index=False)
    
menu = st.sidebar.selectbox("Navegação", ["Emitir Laudo", "Área do Professor"])

if menu == "Emitir Laudo":
    st.title("🧪 Laboratório Escola - Laudos 🧪")
    
    col_n, col_i, col_s = st.columns([3, 1, 1])
    nome_p = col_n.text_input("Paciente")
    idade_p = col_i.number_input("Idade", 0, 120)
    sexo_p = col_s.selectbox("Sexo", ["Masculino", "Feminino"])

    setores_sel = st.multiselect("Selecione os Setores:", list(SETOR_EXAMES.keys()))
    resultados_finais = []

    for setor in setores_sel:
        st.divider()
        st.subheader(f"📍 Setor: {setor}")

        if setor == "Hematologia":
            ex_h = st.multiselect("Exames Hematologia:", SETOR_EXAMES[setor], key="sh")
            
            if "Hemograma" in ex_h:
                with st.expander("🩸 Hemograma Completo", expanded=True):
                    
                    st.markdown("### 1. Série Vermelha")
                    c1, c2, c3, c4 = st.columns(4)
                    he = c1.number_input("Hemácias", format="%.2f", key="he")
                    hb = c2.number_input("Hemoglobina", format="%.1f", key="hb")
                    ht = c3.number_input("Hematócrito", format="%.1f", key="ht")
                    rdw = c4.text_input("RDW (%)", "13.0", key="rdw")
                    
                    vcm = (ht / he * 10) if he > 0 else 0
                    hcm = (hb / he * 10) if he > 0 else 0
                    chcm = (hb / ht * 100) if ht > 0 else 0
                    st.write(f"📊 Índices: VCM: {vcm:.1f} | HCM: {hcm:.1f} | CHCM: {chcm:.1f}")

                    st.markdown("### 2. Série Branca")
                    
                    l1, l2, l3 = st.columns(3)
                    leuc = l1.text_input("Leucócitos Totais", key="leuc")
                    bast = l2.text_input("Bastonetes (%)", key="bast")
                    seg = l3.text_input("Segmentados (%)", key="seg")

                    d4, d5, d6, d7 = st.columns(4)
                    eos = d4.text_input("Eosinófilos (%)", key="h_eos")
                    bas = d5.text_input("Basófilos (%)", key="h_bas")
                    lin = d6.text_input("Linfócitos (%)", key="h_lin")
                    mon = d7.text_input("Monócitos (%)", key="h_mon")

                    obs_hem = st.text_input("Obs Hemograma:", key="obs_hem")
                    if st.checkbox("Confirmar Hemograma no Laudo"):
                        def f_res(n, v): return f"{v} (*)" if verificar_alteracao(n, v) else f"{v}"
                        resultados_finais.append(["--- HEMOGRAMA ---", "", ""])
                        resultados_finais.append(["Hemoglobina", f_res("Hemoglobina", hb), REFERENCIAS_MASTER["Hemoglobina"]])
                        resultados_finais.append(["VCM / HCM", f"{vcm:.1f} / {hcm:.1f}", "80-100 / 27-32"])
                        resultados_finais.append(["Leucócitos", f_res("Leucócitos Totais", leuc), REFERENCIAS_MASTER["Leucócitos Totais"]])
                        if obs_hem: resultados_finais.append(["Obs Hemato", obs_hem, "-"])

            # Exames Avulsos da Hematologia
            outros_hem = [e for e in ex_h if e != "Hemograma"]
            for ex in outros_hem:
                if ex == "Tipagem Sanguínea":
                    with st.expander("🅰️🅱️🅾️ Tipagem Sanguínea", expanded=True):
                        

                        c_abo, c_rh = st.columns(2)
                        g_abo = c_abo.selectbox("Grupo ABO", ["A", "B", "AB", "O"], key="g_abo")
                        f_rh = c_rh.selectbox("Fator Rh", ["Positivo (+)", "Negativo (-)"], key="f_rh")
                        if st.checkbox(f"Confirmar {ex}"):
                            resultados_finais.append(["Tipagem Sanguínea", f"Grupo {g_abo}, Fator {f_rh}", "ABO/Rh"])
                
                elif ex in ["TP", "KTTP"]:
                    with st.expander(f"⏱️ {ex}", expanded=True):
                        


                        c_t1, c_t2 = st.columns(2)
                        t_seg = c_t1.text_input(f"Tempo ({ex}) seg:", key=f"t_{ex}")
                        isi = c_t2.text_input("Relação/INR:", key=f"r_{ex}")
                        if t_seg: resultados_finais.append([ex, f"{t_seg}s (R: {isi})", REFERENCIAS_MASTER.get(ex, "Consultar")])
                
                else:
                    ref = REFERENCIAS_MASTER.get(ex, "Consultar")
                    cr, co = st.columns(2)
                    with cr:
                        v_ex = st.text_input(f"{ex}:", key=f"v_{ex}")
                        st.markdown(f"<small>Ref: {ref}</small>", unsafe_allow_html=True)
                    with co:
                        o_ex = st.text_input(f"Obs {ex}:", key=f"o_{ex}")
                    if v_ex:
                        res = f"{v_ex} (*)" if verificar_alteracao(ex, v_ex) else v_ex
                        if o_ex: res += f" [{o_ex}]"
                        resultados_finais.append([ex, res, ref])

        # --- BLOCO PARASITOLOGIA ---
        elif setor == "Parasitologia":
    
            ex_p = st.multiselect("Selecione os exames realizados:", SETOR_EXAMES["Parasitologia"], key="sel_para")
            
            for ep in ex_p:
                with st.expander(f"Análise: {ep}", expanded=True):
                    if ep == "EPF":
                        
                        c1, c2, c3 = st.columns(3)
                        consist = c1.selectbox("Consistência:", ["Sólida", "Pastosa", "Muco-sanguinolenta", "Líquida"], key="p_cons")
                        metodo = c2.selectbox("Método:", ["Hoffman-Pons-Janer", "Faust", "Direto"], key="p_met")
                        res_epf = c3.selectbox("Resultado:", ["Negativo", "Positivo"], key="p_res_epf")
                        
                        achados = ""
                        if res_epf == "Positivo":
                            achados = st.text_area("Descreva os parasitas encontrados:", placeholder="Ex: Cistos de Giardia lamblia e ovos de Ascaris lumbricoides", key="p_achados")
                        
                        if st.checkbox("Confirmar EPF no Laudo", key="chk_epf"):
                            val_final = "Negativo (Não foram encontrados cistos, ovos ou larvas na amostra analisada.)" if res_epf == "Negativo" else f"POSITIVO: {achados}"
                            resultados_finais.append([f"{ep} ({metodo})", val_final, "Ausência de parasitas"])

                    elif ep == "Coprocultura":
                        
                        c_copro = st.columns(1)[0]
                        res_copro = c_copro.selectbox("Resultado da Cultura:", ["Flora normal (Ausência de patógenos)", "Presença de microrganismo patogênico"], key="p_copro")
                        
                        if res_copro == "Presença de microrganismo patogênico":
                            isolado = st.text_input("Microrganismo isolado:", key="p_iso")
                            if st.checkbox(f"Confirmar {ep}"):
                                resultados_finais.append([ep, f"ISOLADO: {isolado}", "Flora normal"])
                        else:
                            if st.checkbox(f"Confirmar {ep}"):
                                resultados_finais.append([ep, "Flora fecal normal", "Flora normal"])

                    else:
                        # Para Sangue Oculto, Leucócitos Fecais, etc.
                        c_r, c_o = st.columns(2)
                        v_p = c_r.selectbox(f"Resultado {ep}:", ["Negativo", "Positivo"], key=f"vp_{ep}")
                        o_p = c_o.text_input("Observações/Notas:", key=f"op_{ep}")
                        
                        if st.checkbox(f"Confirmar {ep}", key=f"chk_{ep}"):
                            ref_para = "Negativo"
                            res_f = v_p
                            if o_p: res_f += f" ({o_p})"
                            if v_p == "Positivo": res_f += " (*)"
                            resultados_finais.append([ep, res_f, ref_para])

        # --- BLOCO IMUNOLOGIA ---
        elif setor == "Imunologia":
            ex_i = st.multiselect("Exames de Imunologia:", SETOR_EXAMES["Imunologia"], key="si_input")
            for ei in ex_i:
                
                c_i1, c_i2 = st.columns(2)
                v_i = c_i1.selectbox(f"Resultado {ei}:", ["Não Reagente", "Reagente", "Negativo", "Positivo"], key=f"vi_{ei}")
                o_i = c_i2.text_input(f"Título/Valor {ei}:", key=f"ti_{ei}")
                
                if st.checkbox(f"Confirmar {ei}", key=f"chk_i_{ei}"):
                    ref_i = REFERENCIAS_MASTER.get(ei, "Não Reagente")
                    res_i = v_i + (f" [{o_i}]" if o_i else "")
                    if v_i in ["Reagente", "Positivo"]: res_i += " (*)"
                    resultados_finais.append([ei, res_i, ref_i])
                    
        elif setor == "Uroanálise":
            with st.expander("📋 EQU - Exame Qualitativo de Urina", expanded=True):
                
                st.markdown("### 1. Exame Físico")
                u1, u2, u3 = st.columns(3)
                cor_u = u1.selectbox("Cor", ["Amarelo Citrino", "Amarelo claro", "Âmbar", "Avermelhada", "Laranja"], key="cor_u")
                asp_u = u2.selectbox("Aspecto", ["Límpido", "Levemente Turvo", "Turvo"], key="asp_u")
                den_u = u3.text_input("Densidade", "1.020", key="dens_u")
                
                st.markdown("#### Fita Reagente")
                
                st.markdown("### 2. Exame Químico (Fita Reativa)")
                q1, q2, q3, q4 = st.columns(4)
                dens_u = q1.text_input("Densidade", "1.020")
                ph_u = q2.selectbox("pH", ["5.0", "5.5", "6.0", "6.5", "7.0", "7.5", "8.0", "8.5"])
                leuc_u = q3.selectbox("Leucócitos (Esterase)", ["Negativo", "Traços", "75 leu/uL (+)", "125 leu/uL (++)", "500 leu/uL (+++)"])
                nitrito_u = q4.selectbox("Nitrito", ["Negativo", "Positivo"])
                
                q5, q6, q7, q8 = st.columns(4)
                prot_u = q5.selectbox("Proteína", ["Negativo", "Traços", "30 mg/dL (+)", "100 mg/dL (++)", "500 mg/dL (+++)"])
                glic_u = q6.selectbox("Glicose", ["Normal", "50 mg/dL", "100 mg/dL", "250 mg/dL", "500 mg/dL"])
                cetonas_u = q7.selectbox("Corpos Cetónicos", ["Negativo", "Traços", "15 mg/dL (+)", "40 mg/dL (++)", "80 mg/dL (+++)"])
                urobil_u = q8.selectbox("Urobilinogénio", ["Normal", "2 mg/dL", "4 mg/dL", "8 mg/dL", "12 mg/dL"])
                
                q9, q10 = st.columns(2)
                bilirr_u = q9.selectbox("Bilirrubina", ["Negativo", "+", "++", "+++"])
                sangue_u = q10.selectbox("Sangue/Hemoglobina", ["Negativo", "Traços", "10 ery/uL (+)", "25 ery/uL (++)", "50 ery/uL (+++)"])

                st.markdown("### 3. Sedimentoscopia (Microscopia)")
                s1, s2, s3, s4 = st.columns(4)
                cel_ep = s1.text_input("Células Epiteliais", "Raras")
                leuc_sed = s2.text_input("Leucócitos (p/ campo)", "0-2")
                hem_sed = s3.text_input("Hemácias (p/ campo)", "0-1")
                cristais_u = s4.text_input("Cristais", "Ausentes")
                cilindros_u = st.text_input("Cilindros", "Ausentes")
                flora_u = st.selectbox("Flora Bacteriana", ["Normal", "Aumentada", "Muito Aumentada"])

                if st.checkbox("Incluir EQU no Laudo"):
                    resultados_finais.append(["EQU - COR/ASPECTO", f"{cor_u} / {asp_u}", "Am. Citrino / Límpido"])
                    resultados_finais.append(["EQU - DENSIDADE/PH", f"{dens_u} / {ph_u}", "1.005-1.030 / 5.0-8.0"])
                    resultados_finais.append(["EQU - LEUCÓCITOS (FITA)", leuc_u, "Negativo"])
                    resultados_finais.append(["EQU - NITRITO", nitrito_u, "Negativo"])
                    resultados_finais.append(["EQU - PROTEÍNA", prot_u, "Negativo"])
                    resultados_finais.append(["EQU - GLICOSE", glic_u, "Normal"])
                    resultados_finais.append(["EQU - BILIRRUBINA/CETONA", f"{bilirr_u} / {cetonas_u}", "Negativo"])
                    resultados_finais.append(["EQU - SEDIMENTOSCOPIA", f"Leu: {leuc_sed} | Hem: {hem_sed} | Cel: {cel_ep}", "Raros"])
                    
        elif setor == "Microbiologia":
            ex_m = st.multiselect("Exames:", SETOR_EXAMES[setor], key="sm")
            for e_m in ex_m:
                if e_m == "Antibiograma":
                    with st.expander("💊 Antibiograma", expanded=True):
                        
                        qtd = st.number_input("Qtd Antibióticos:", 1, 15, 5, key="atb_q")
                        for i in range(qtd):
                            ca1, ca2 = st.columns([3, 1])
                            n_atb = ca1.text_input(f"ATB {i+1}", key=f"an{i}")
                            r_atb = ca2.selectbox("Perfil", ["S", "I", "R"], key=f"ar{i}")
                            if n_atb: resultados_finais.append([f"ATB: {n_atb}", r_atb, "S/I/R"])
                else:
                    val = st.text_input(f"{e_m}:", key=f"vm_{e_m}")
                    resultados_finais.append([e_m, val, REFERENCIAS_MASTER.get(e_m, "Consultar")])

        elif setor == "Bioquímica":
            ex_b = st.multiselect("Exames Bioquímica:", SETOR_EXAMES["Bioquímica"] + ["Perfil Lipídico"], key="sb")
            
            for e_b in ex_b:
                # --- Caso 1: Perfil Lipídico (Cálculo Automático de LDL e VLDL) ---
                if e_b == "Perfil Lipídico":
                    with st.expander(" lipidogram - Perfil Lipídico (Friedewald)", expanded=True):
                        
                        col_l1, col_l2, col_l3 = st.columns(3)
                        ct = col_l1.number_input("Colesterol Total", min_value=0.0, format="%.1f", key="ct")
                        hdl = col_l2.number_input("HDL", min_value=0.0, format="%.1f", key="hdl")
                        tg = col_l3.number_input("Triglicerídeos", min_value=0.0, format="%.1f", key="tg")
                        
                        vldl = tg / 5 if tg < 400 else 0
                        ldl = ct - hdl - vldl if tg < 400 else 0
                        
                        if tg >= 400:
                            st.warning("⚠️ Triglicerídeos > 400 mg/dL: A fórmula de Friedewald não é recomendada.")
                        else:
                            st.info(f"Cálculos: VLDL = {vldl:.1f} | LDL = {ldl:.1f}")
                        
                        obs_lip = st.text_input("Obs Lipídico:", key="obs_lip")
                        
                        if st.checkbox("Confirmar Perfil Lipídico no Laudo"):
                            resultados_finais.append(["--- PERFIL LIPÍDICO ---", "", ""])
                            resultados_finais.append(["Colesterol Total", f"{ct}", REFERENCIAS_MASTER["Colesterol"]])
                            resultados_finais.append(["HDL", f"{hdl}", REFERENCIAS_MASTER["HDL"]])
                            resultados_finais.append(["Triglicerídeos", f"{tg}", REFERENCIAS_MASTER["Triglicerídeos"]])
                            if tg < 400:
                                flag_ldl = " (*)" if verificar_alteracao("LDL", str(ldl)) else ""
                                resultados_finais.append(["LDL (Calculado)", f"{ldl:.1f}{flag_ldl}", REFERENCIAS_MASTER["LDL"]])
                                resultados_finais.append(["VLDL (Calculado)", f"{vldl:.1f}", "Até 30.0"])
                            if obs_lip: resultados_finais.append(["Obs Lipídico", obs_lip, "-"])

                # --- Caso 2: Bilirrubinas (Cálculo de Indireta) ---
                elif e_b == "Bilirrubinas":
                    with st.expander("🟡 Bilirrubinas Totais e Frações", expanded=True):
                        col_b1, col_b2 = st.columns(2)
                        bt = col_b1.number_input("Bilirrubina Total", format="%.2f", key="bt")
                        bd = col_b2.number_input("Bilirrubina Direta", format="%.2f", key="bd")
                        bi = bt - bd if bt >= bd else 0.0
                        st.write(f"Indireta: **{bi:.2f}**")
                        
                        if st.checkbox("Confirmar Bilirrubinas no Laudo"):
                            resultados_finais.append(["Bilirrubina Total", f"{bt}", REFERENCIAS_MASTER["Bilirrubinas Totais"]])
                            resultados_finais.append(["Bilirrubina Direta", f"{bd}", REFERENCIAS_MASTER["Bilirrubina Direta"]])
                            resultados_finais.append(["Bilirrubina Indireta", f"{bi:.2f}", REFERENCIAS_MASTER["Bilirrubina Indireta"]])

                # --- Caso 3: Demais Exames (Glicose, Ureia, Creatinina...) ---
                else:
                    ref_b = REFERENCIAS_MASTER.get(e_b, "Consultar")
                    cr, co = st.columns(2)
                    v_b = cr.text_input(f"{e_b} (Ref: {ref_b}):", key=f"vb_{e_b}")
                    o_b = co.text_input(f"Obs {e_b}:", key=f"ob_{e_b}")
                    
                    if v_b:
                        flag = " (*)" if verificar_alteracao(e_b, v_b) else ""
                        res_f = f"{v_b}{flag}"
                        if o_b: res_f += f" [{o_b}]"
                        resultados_finais.append([e_b, res_f, ref_b])
    st.divider()
    
# ==========================================
    # SEÇÃO DE FINALIZAÇÃO (DENTRO DO IF EMITIR LAUDO)
    # ==========================================
    st.divider()
    st.subheader("📝 Finalização do Laudo")
    
    cf1, cf2 = st.columns(2)
    # Variáveis que capturam o que foi digitado agora
    aluno_f = cf1.text_input("Nome do Aluno:", key="aluno_f_final")
    supervisor_f = cf2.text_input("Professor Supervisor:", key="super_f_final")

    if st.button("🚀 Gerar PDF e Registrar"):
        # Verificamos se os campos básicos e a lista de exames estão preenchidos
        if nome_p and resultados_finais and aluno_f:
            try:
                # 1. Preparação do PDF em memória
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=A4)
                styles = getSampleStyleSheet()
                elementos = []
                
                # --- CABEÇALHO DO PDF ---
                # Usamos as variáveis DIRETAS da tela (nome_p, idade_p, aluno_f)
                elementos.append(Paragraph(f"<b>LABORATÓRIO ESCOLA - LAUDO OFICIAL</b>", styles["Title"]))
                elementos.append(Spacer(1, 12))
                elementos.append(Paragraph(f"<b>Paciente:</b> {nome_p.upper()}", styles["Normal"]))
                elementos.append(Paragraph(f"<b>Idade:</b> {idade_p} anos | <b>Sexo:</b> {sexo_p}", styles["Normal"]))
                elementos.append(Paragraph(f"<b>Data:</b> {datetime.now().strftime('%d/%m/%Y')}", styles["Normal"]))
                elementos.append(Paragraph(f"<b>Responsável pela Análise:</b> {aluno_f}", styles["Normal"]))
                elementos.append(Spacer(1, 15))
                
                # --- TABELA DE RESULTADOS ---
                # Image of a clinical lab report table showing parameters, results, and reference values
                
                cabecalho_tabela = [["EXAME", "RESULTADO", "VALORES DE REFERÊNCIA"]]
                dados_tabela = cabecalho_tabela + resultados_finais
                
                tabela = Table(dados_tabela, colWidths=[180, 180, 140])
                tabela.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                ]))
                elementos.append(tabela)
                
                # --- ASSINATURA ---
                elementos.append(Spacer(1, 40))
                elementos.append(Paragraph("________________________________________________", styles["Normal"]))
                assinatura = supervisor_f if supervisor_f else "Professor Responsável"
                elementos.append(Paragraph(f"Assinatura do Supervisor: {assinatura}", styles["Normal"]))

                # Constrói o PDF
                doc.build(elementos)
                
                # 2. SALVAMENTO NO EXCEL (Aqui enviamos os dados para a função)
                salvar_registro({
                    "Data": datetime.now().strftime('%d/%m/%Y'),
                    "Paciente": nome_p,
                    "Sexo": sexo_p,
                    "Setores": ", ".join(setores_sel) if setores_sel else "Não informado",
                    "Aluno": aluno_f,
                    "Supervisor": supervisor_f,
                    "Resultados": resultados_finais
                })
                
                st.success("✅ Laudo gerado com sucesso!")
                st.download_button("📥 Baixar Laudo (PDF)", buffer.getvalue(), f"laudo_{nome_p}.pdf", "application/pdf")
                
            except Exception as e:
                st.error(f"Erro técnico ao gerar o documento: {e}")
        else:
            st.error("❌ Preencha o Paciente, seu nome (Aluno) e confirme ao menos um exame acima.")

# ==========================================
# ÁREA DO PROFESSOR
# ==========================================
elif menu == "Área do Professor":
    st.title("🔐 Área do Professor")
    senha = st.text_input("Senha de Acesso:", type="password", key="senha_prof_final")
    
    if senha == SENHA_PROFESSOR:
        st.success("Acesso Autorizado")
        if os.path.exists("laudos_registro.xlsx"):
            df_reg = pd.read_excel("laudos_registro.xlsx")
            if not df_reg.empty:
                st.subheader("📋 Resumo de Atividades")
                cols = [c for c in ["Data", "Paciente", "Aluno", "Nota"] if c in df_reg.columns]
                st.dataframe(df_reg[cols], use_container_width=True)
                
                st.divider()
                st.subheader("🔍 Auditoria Detalhada")
                opcoes = [f"{i} | {row['Paciente']} - Aluno: {row['Aluno']}" for i, row in df_reg.iterrows()]
                escolha = st.selectbox("Selecione um laudo para conferir:", opcoes)
                
                if escolha:
                    idx = int(escolha.split(" | ")[0])
                    detalhes = df_reg.iloc[idx]
                    
                    if pd.notna(detalhes["Resultados"]):
                        lista_bruta = str(detalhes["Resultados"]).split(" | ")
                        audit_data = []
                        for item in lista_bruta:
                            partes = item.split("# ")
                            if len(partes) == 3:
                                audit_data.append({"Exame": partes[0], "Resultado": partes[1], "Ref": partes[2]})
                        st.table(pd.DataFrame(audit_data))
                    
                    with st.form("form_nota"):
                        c1, c2 = st.columns([1, 3])
                        n_nota = c1.text_input("Nota:", value=str(detalhes.get('Nota', '-')))
                        n_feed = c2.text_area("Feedback:", value=str(detalhes.get('Feedback', '-')))
                        if st.form_submit_button("✅ Salvar"):
                            df_reg.at[idx, 'Nota'] = n_nota
                            df_reg.at[idx, 'Feedback'] = n_feed
                            df_reg.to_excel("laudos_registro.xlsx", index=False)
                            st.rerun()

















