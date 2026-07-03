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


st.set_page_config(
    page_title="MARS Intelligence",
    page_icon="🪐",
    layout="wide"
)


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


def mars_score(row):
    score = 50

    risco = str(row.get("RISCO PI", "")).strip().lower()
    if risco == "baixo":
        score += 15
    elif risco in ["médio", "medio"]:
        score += 5
    elif risco == "alto":
        score -= 20

    design = str(row.get("DESIGN ORIGINAL?", "")).strip().lower()
    if design in ["sim", "provável", "provavel"]:
        score += 10

    personalizado = str(row.get("PERSONALIZAÇÃO", "")).strip().lower()
    if personalizado == "sim":
        score += 10

    video = str(row.get("VÍDEO", "")).strip().lower()
    if video == "sim":
        score += 5

    bestseller = str(row.get("BESTSELLER", "")).strip().lower()
    if bestseller == "sim":
        score += 8

    etsy_pick = str(row.get("ETSY PICK", "")).strip().lower()
    if etsy_pick == "sim":
        score += 5

    price = row.get("PREÇO LIMPO")
    if pd.notna(price):
        if price >= 25:
            score += 10
        elif price < 10:
            score -= 5

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
    df["VALE ESTUDAR?"] = df["MARS SCORE"].apply(
        lambda x: "Sim" if x >= 65 else "Talvez" if x >= 50 else "Não"
    )

    return df


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="PRODUTOS_LIMPOS", index=False)

        workbook = writer.book
        worksheet = writer.sheets["PRODUTOS_LIMPOS"]

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",
            "border": 1
        })

        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 18)

    return output.getvalue()


def load_file(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        raw_df = pd.read_csv(uploaded_file)
    else:
        xls = pd.ExcelFile(uploaded_file)
        sheet = "PRODUTOS" if "PRODUTOS" in xls.sheet_names else xls.sheet_names[0]
        raw_df = pd.read_excel(uploaded_file, sheet_name=sheet)

    return prepare_data(raw_df)


st.sidebar.title("🪐 MARS Intelligence")
st.sidebar.caption("v0.2")

page = st.sidebar.radio(
    "Menu",
    [
        "🏠 Dashboard",
        "📦 Produtos",
        "🏪 Lojas",
        "💰 Preços",
        "⚠️ Risco PI",
        "💡 Oportunidades",
        "📤 Exportar",
    ],
)

uploaded_file = st.sidebar.file_uploader(
    "Carregar Excel/CSV",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    df = load_file(uploaded_file)
    st.session_state["df"] = df
else:
    df = st.session_state.get("df")

if df is None:
    st.title("🪐 MARS Intelligence")
    st.info("Carrega o teu ficheiro Excel/CSV na barra lateral para começar.")
    st.stop()


with st.sidebar:
    st.divider()
    st.subheader("Filtros")

    categorias = sorted([c for c in df["CATEGORIA"].dropna().unique() if str(c).strip()])
    riscos = sorted([r for r in df["RISCO PI"].dropna().unique() if str(r).strip()])
    estudar = sorted(df["VALE ESTUDAR?"].dropna().unique())

    filtro_categoria = st.multiselect("Categoria", categorias)
    filtro_risco = st.multiselect("Risco PI", riscos)
    filtro_estudar = st.multiselect(
        "Vale estudar?",
        estudar,
        default=["Sim", "Talvez"] if "Sim" in estudar else estudar
    )

filtered = df.copy()

if filtro_categoria:
    filtered = filtered[filtered["CATEGORIA"].isin(filtro_categoria)]
if filtro_risco:
    filtered = filtered[filtered["RISCO PI"].isin(filtro_risco)]
if filtro_estudar:
    filtered = filtered[filtered["VALE ESTUDAR?"].isin(filtro_estudar)]


if page == "🏠 Dashboard":
    st.title("🏠 Dashboard MARS Intelligence")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Produtos", len(filtered))
    col2.metric("Lojas", filtered["LOJA"].nunique())
    col3.metric("Preço médio", f"{filtered['PREÇO LIMPO'].mean():.2f} €" if filtered["PREÇO LIMPO"].notna().any() else "—")
    col4.metric("Score médio", f"{filtered['MARS SCORE'].mean():.1f}")
    col5.metric("Vale estudar", int((filtered["VALE ESTUDAR?"] == "Sim").sum()))

    st.divider()

    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Top categorias")
        cat = filtered["CATEGORIA"].value_counts().reset_index()
        cat.columns = ["Categoria", "Produtos"]
        if not cat.empty:
            fig = px.bar(cat.head(15), x="Categoria", y="Produtos")
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Distribuição do MARS SCORE")
        if not filtered.empty:
            fig = px.histogram(filtered, x="MARS SCORE", nbins=20)
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("Top oportunidades")
    top = filtered.sort_values("MARS SCORE", ascending=False).head(20)
    st.dataframe(top[["LOJA", "PRODUTO", "CATEGORIA", "PREÇO (€)", "RISCO PI", "MARS SCORE", "VALE ESTUDAR?"]], use_container_width=True)


elif page == "📦 Produtos":
    st.title("📦 Produtos")
    st.dataframe(filtered, use_container_width=True, height=650)


elif page == "🏪 Lojas":
    st.title("🏪 Lojas")
    lojas = (
        filtered.groupby("LOJA")
        .agg(
            Produtos=("PRODUTO", "count"),
            Score_Médio=("MARS SCORE", "mean"),
            Preço_Médio=("PREÇO LIMPO", "mean"),
            Categorias=("CATEGORIA", lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip()]))[:5])),
        )
        .reset_index()
        .sort_values("Score_Médio", ascending=False)
    )
    st.dataframe(lojas, use_container_width=True, height=650)


elif page == "💰 Preços":
    st.title("💰 Preços")

    price_df = filtered.dropna(subset=["PREÇO LIMPO"])

    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Preço médio por categoria")
        avg = price_df.groupby("CATEGORIA", as_index=False)["PREÇO LIMPO"].mean()
        avg = avg.sort_values("PREÇO LIMPO", ascending=False).head(20)
        if not avg.empty:
            fig = px.bar(avg, x="CATEGORIA", y="PREÇO LIMPO")
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Distribuição de preços")
        if not price_df.empty:
            fig = px.histogram(price_df, x="PREÇO LIMPO", nbins=25)
            st.plotly_chart(fig, use_container_width=True)

    st.dataframe(price_df.sort_values("PREÇO LIMPO", ascending=False), use_container_width=True)


elif page == "⚠️ Risco PI":
    st.title("⚠️ Risco de Propriedade Intelectual")

    risco = filtered["RISCO PI"].value_counts().reset_index()
    risco.columns = ["Risco", "Produtos"]

    if not risco.empty:
        fig = px.pie(risco, names="Risco", values="Produtos")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Produtos de risco alto")
    alto = filtered[filtered["RISCO PI"].astype(str).str.lower() == "alto"]
    st.dataframe(alto, use_container_width=True)


elif page == "💡 Oportunidades":
    st.title("💡 Oportunidades MARS 3D")

    oportunidades = filtered[
        (filtered["MARS SCORE"] >= 65) &
        (filtered["RISCO PI"].astype(str).str.lower().isin(["baixo", "médio", "medio"]))
    ].sort_values("MARS SCORE", ascending=False)

    st.caption("Produtos com bom score e risco PI controlado.")
    st.dataframe(
        oportunidades[["LOJA", "PRODUTO", "CATEGORIA", "PREÇO (€)", "RISCO PI", "MARS SCORE", "LINK", "OBSERVAÇÕES"]],
        use_container_width=True,
        height=650
    )


elif page == "📤 Exportar":
    st.title("📤 Exportar dados")
    st.write("Exporta os dados filtrados já com MARS SCORE e classificação.")

    st.download_button(
        "⬇️ Exportar Excel filtrado",
        data=to_excel(filtered),
        file_name=f"MARS_Intelligence_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
