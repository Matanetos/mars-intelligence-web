import io
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

EXPECTED_COLUMNS = [
    "LOJA", "PRODUTO", "CATEGORIA", "TIPO PRODUTO", "LICENCIADO",
    "DESIGN ORIGINAL?", "LICENÇA COMERCIAL?", "PREÇO (€)", "PORTES (€)",
    "FAVORITOS", "AVALIAÇÕES", "CLASSIFICAÇÃO", "BESTSELLER", "ETSY PICK",
    "VÍDEO", "Nº FOTOS", "PERSONALIZAÇÃO", "LINK", "RISCO PI",
    "DATA ANÁLISE", "PAÍS", "MATERIAL", "COR", "OBSERVAÇÕES"
]

HIGH_RISK = ["disney","marvel","pokemon","pokémon","star wars","hulk","spiderman","spider-man","batman","superman","mario","zelda","stitch","goofy","bob marley","asterix","popeye","brutus","duffy","looney"]
MED_RISK = ["apple","airpod","airpods","magsafe","dyson","playstation","ps5","xbox","nintendo","steelseries","shokz","alexa","echo dot"]

st.set_page_config(page_title="MARS Intelligence", page_icon="🪐", layout="wide")

def clean_price(value):
    if pd.isna(value):
        return None
    text = str(value).replace("€", "").replace(",", ".").strip()
    if text.lower() in ["grátis", "gratis", "free"]:
        return 0
    try:
        return float(text)
    except ValueError:
        return None

def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[EXPECTED_COLUMNS]

def auto_risk(text):
    t = str(text).lower()
    if any(w in t for w in HIGH_RISK):
        return "Alto"
    if any(w in t for w in MED_RISK):
        return "Médio"
    return "Baixo"

def mars_score(row):
    score = 50
    risco = str(row.get("RISCO PI", "")).strip().lower()
    if risco == "baixo": score += 15
    elif risco in ["médio", "medio"]: score += 5
    elif risco == "alto": score -= 20
    if str(row.get("DESIGN ORIGINAL?", "")).strip().lower() in ["sim", "provável", "provavel"]: score += 10
    if str(row.get("PERSONALIZAÇÃO", "")).strip().lower() == "sim": score += 10
    if str(row.get("VÍDEO", "")).strip().lower() == "sim": score += 5
    if str(row.get("BESTSELLER", "")).strip().lower() == "sim": score += 8
    if str(row.get("ETSY PICK", "")).strip().lower() == "sim": score += 5
    price = row.get("PREÇO LIMPO")
    if pd.notna(price):
        if price >= 25: score += 10
        elif price < 10: score -= 5
    return max(0, min(100, int(score)))

def prepare_data(df):
    df = normalize_columns(df)
    df["PREÇO LIMPO"] = df["PREÇO (€)"].apply(clean_price)
    df["PORTES LIMPO"] = df["PORTES (€)"].apply(clean_price)
    df["LOJA"] = df["LOJA"].astype(str).str.strip()
    df["PRODUTO"] = df["PRODUTO"].astype(str).str.strip()
    df["LINK"] = df["LINK"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["LOJA", "PRODUTO", "LINK"], keep="first")
    df["MARS SCORE"] = df.apply(mars_score, axis=1)
    df["VALE ESTUDAR?"] = df["MARS SCORE"].apply(lambda x: "Sim" if x >= 65 else "Talvez" if x >= 50 else "Não")
    return df

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="PRODUTOS_LIMPOS", index=False)
        workbook = writer.book
        ws = writer.sheets["PRODUTOS_LIMPOS"]
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAD3", "border": 1})
        for i, col in enumerate(df.columns):
            ws.write(0, i, col, header_fmt)
            ws.set_column(i, i, 18)
    return output.getvalue()

def load_file(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    else:
        xls = pd.ExcelFile(uploaded_file)
        sheet = "PRODUTOS" if "PRODUTOS" in xls.sheet_names else xls.sheet_names[0]
        raw = pd.read_excel(uploaded_file, sheet_name=sheet)
    return prepare_data(raw)

def template_excel():
    return to_excel(pd.DataFrame(columns=EXPECTED_COLUMNS))

st.sidebar.title("🪐 MARS Intelligence")
st.sidebar.caption("v0.3")
page = st.sidebar.radio("Menu", ["🏠 Dashboard", "📦 Produtos", "🏪 Lojas", "🔎 Pesquisa Etsy", "💰 Preços", "⚠️ Risco PI", "💡 Oportunidades", "📤 Exportar"])
uploaded = st.sidebar.file_uploader("Carregar Excel/CSV", type=["xlsx", "xls", "csv"])

if uploaded:
    df = load_file(uploaded)
    st.session_state["df"] = df
else:
    df = st.session_state.get("df")

if df is not None:
    with st.sidebar:
        st.divider(); st.subheader("Filtros")
        categorias = sorted([c for c in df["CATEGORIA"].dropna().unique() if str(c).strip()])
        riscos = sorted([r for r in df["RISCO PI"].dropna().unique() if str(r).strip()])
        estudar = sorted(df["VALE ESTUDAR?"].dropna().unique())
        fc = st.multiselect("Categoria", categorias)
        fr = st.multiselect("Risco PI", riscos)
        fe = st.multiselect("Vale estudar?", estudar, default=["Sim", "Talvez"] if "Sim" in estudar else estudar)
    filtered = df.copy()
    if fc: filtered = filtered[filtered["CATEGORIA"].isin(fc)]
    if fr: filtered = filtered[filtered["RISCO PI"].isin(fr)]
    if fe: filtered = filtered[filtered["VALE ESTUDAR?"].isin(fe)]
else:
    filtered = None

if page == "🔎 Pesquisa Etsy":
    st.title("🔎 Pesquisa Etsy assistida")
    st.info("A v0.3 ainda não recolhe dados automaticamente. Ajuda-te a criar linhas normalizadas a partir de dados recolhidos no Etsy/EverBee.")
    c1, c2 = st.columns(2)
    with c1:
        loja = st.text_input("Loja")
        produto = st.text_input("Produto")
        categoria = st.selectbox("Categoria", ["Headphone Stand", "Controller Holder", "Desk Organizer", "Smart Home Holder", "Home Decor", "Gaming Decor", "Phone Holder", "Planter", "Outro"])
        tipo = st.text_input("Tipo produto", value="Gaming / Desk")
        preco = st.text_input("Preço (€)")
        portes = st.text_input("Portes (€)")
        link = st.text_input("Link Etsy")
    with c2:
        licenciado = st.selectbox("Licenciado", ["Não", "Sim", "A confirmar"])
        design = st.selectbox("Design original?", ["Provável", "Sim", "Não", "A confirmar"])
        licenca = st.selectbox("Licença comercial?", ["Não", "Sim", "Desconhecida", "A confirmar"])
        personalizacao = st.selectbox("Personalização", ["Não", "Sim", "A confirmar"])
        video = st.selectbox("Vídeo", ["A confirmar", "Sim", "Não"])
        pais = st.text_input("País")
        material = st.text_input("Material", value="A confirmar")
    obs = st.text_area("Observações")
    risco_sug = auto_risk(f"{produto} {categoria} {obs}")
    risco = st.selectbox("Risco PI", ["Baixo", "Médio", "Alto", "A confirmar"], index=["Baixo", "Médio", "Alto", "A confirmar"].index(risco_sug))
    if st.button("Gerar linha"):
        row = pd.DataFrame([{
            "LOJA": loja, "PRODUTO": produto, "CATEGORIA": categoria, "TIPO PRODUTO": tipo,
            "LICENCIADO": licenciado, "DESIGN ORIGINAL?": design, "LICENÇA COMERCIAL?": licenca,
            "PREÇO (€)": preco, "PORTES (€)": portes, "FAVORITOS": "", "AVALIAÇÕES": "",
            "CLASSIFICAÇÃO": "", "BESTSELLER": "A confirmar", "ETSY PICK": "A confirmar",
            "VÍDEO": video, "Nº FOTOS": "A confirmar", "PERSONALIZAÇÃO": personalizacao,
            "LINK": link, "RISCO PI": risco, "DATA ANÁLISE": datetime.now().strftime("%d/%m/%Y"),
            "PAÍS": pais, "MATERIAL": material, "COR": "", "OBSERVAÇÕES": obs
        }])
        row = prepare_data(row)
        st.dataframe(row, use_container_width=True)
        st.download_button("⬇️ Exportar esta linha", data=to_excel(row), file_name="MARS_nova_linha_produto.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("⬇️ Descarregar template Excel", data=template_excel(), file_name="MARS_template_produtos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif filtered is None:
    st.title("🪐 MARS Intelligence"); st.info("Carrega o teu ficheiro Excel/CSV na barra lateral para começar.")
elif page == "🏠 Dashboard":
    st.title("🏠 Dashboard MARS Intelligence")
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Produtos", len(filtered)); col2.metric("Lojas", filtered["LOJA"].nunique())
    col3.metric("Preço médio", f"{filtered['PREÇO LIMPO'].mean():.2f} €" if filtered["PREÇO LIMPO"].notna().any() else "—")
    col4.metric("Score médio", f"{filtered['MARS SCORE'].mean():.1f}"); col5.metric("Vale estudar", int((filtered["VALE ESTUDAR?"] == "Sim").sum()))
    st.divider(); c1, c2 = st.columns(2)
    with c1:
        cat = filtered["CATEGORIA"].value_counts().reset_index(); cat.columns = ["Categoria", "Produtos"]
        if not cat.empty: st.plotly_chart(px.bar(cat.head(15), x="Categoria", y="Produtos", title="Top categorias"), use_container_width=True)
    with c2: st.plotly_chart(px.histogram(filtered, x="MARS SCORE", nbins=20, title="Distribuição MARS SCORE"), use_container_width=True)
    st.subheader("Top oportunidades"); top = filtered.sort_values("MARS SCORE", ascending=False).head(20)
    st.dataframe(top[["LOJA", "PRODUTO", "CATEGORIA", "PREÇO (€)", "RISCO PI", "MARS SCORE", "VALE ESTUDAR?"]], use_container_width=True)
elif page == "📦 Produtos":
    st.title("📦 Produtos"); st.dataframe(filtered, use_container_width=True, height=650)
elif page == "🏪 Lojas":
    st.title("🏪 Lojas")
    lojas = filtered.groupby("LOJA").agg(Produtos=("PRODUTO","count"), Score_Médio=("MARS SCORE","mean"), Preço_Médio=("PREÇO LIMPO","mean")).reset_index().sort_values("Score_Médio", ascending=False)
    st.dataframe(lojas, use_container_width=True, height=650)
elif page == "💰 Preços":
    st.title("💰 Preços"); price_df = filtered.dropna(subset=["PREÇO LIMPO"]); c1, c2 = st.columns(2)
    with c1:
        avg = price_df.groupby("CATEGORIA", as_index=False)["PREÇO LIMPO"].mean().sort_values("PREÇO LIMPO", ascending=False).head(20)
        if not avg.empty: st.plotly_chart(px.bar(avg, x="CATEGORIA", y="PREÇO LIMPO", title="Preço médio por categoria"), use_container_width=True)
    with c2:
        if not price_df.empty: st.plotly_chart(px.histogram(price_df, x="PREÇO LIMPO", nbins=25, title="Distribuição de preços"), use_container_width=True)
    st.dataframe(price_df.sort_values("PREÇO LIMPO", ascending=False), use_container_width=True)
elif page == "⚠️ Risco PI":
    st.title("⚠️ Risco de Propriedade Intelectual"); risco = filtered["RISCO PI"].value_counts().reset_index(); risco.columns = ["Risco", "Produtos"]
    if not risco.empty: st.plotly_chart(px.pie(risco, names="Risco", values="Produtos"), use_container_width=True)
    st.subheader("Produtos de risco alto"); st.dataframe(filtered[filtered["RISCO PI"].astype(str).str.lower() == "alto"], use_container_width=True)
elif page == "💡 Oportunidades":
    st.title("💡 Oportunidades MARS 3D")
    oportunidades = filtered[(filtered["MARS SCORE"] >= 65) & (filtered["RISCO PI"].astype(str).str.lower().isin(["baixo", "médio", "medio"]))].sort_values("MARS SCORE", ascending=False)
    st.dataframe(oportunidades[["LOJA", "PRODUTO", "CATEGORIA", "PREÇO (€)", "RISCO PI", "MARS SCORE", "LINK", "OBSERVAÇÕES"]], use_container_width=True, height=650)
elif page == "📤 Exportar":
    st.title("📤 Exportar dados")
    st.download_button("⬇️ Exportar Excel filtrado", data=to_excel(filtered), file_name=f"MARS_Intelligence_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
