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
    elif risco == "médio" or risco == "medio":
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


st.set_page_config(
    page_title="MARS Intelligence",
    page_icon="🪐",
    layout="wide"
)

st.title("🪐 MARS Intelligence")
st.caption("Análise de mercado Etsy para produtos físicos impressos em 3D")

uploaded_file = st.file_uploader(
    "Carrega o teu ficheiro Excel ou CSV",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        raw_df = pd.read_csv(uploaded_file)
    else:
        xls = pd.ExcelFile(uploaded_file)
        sheet = "PRODUTOS" if "PRODUTOS" in xls.sheet_names else xls.sheet_names[0]
        raw_df = pd.read_excel(uploaded_file, sheet_name=sheet)

    df = prepare_data(raw_df)

    st.success(f"Ficheiro carregado com sucesso: {len(df)} produtos únicos analisados.")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Produtos", len(df))
    col2.metric("Lojas", df["LOJA"].nunique())
    col3.metric("Preço médio", f"{df['PREÇO LIMPO'].mean():.2f} €" if df["PREÇO LIMPO"].notna().any() else "—")
    col4.metric("Score médio", f"{df['MARS SCORE'].mean():.1f}")
    col5.metric("Vale estudar", int((df["VALE ESTUDAR?"] == "Sim").sum()))

    st.divider()

    with st.sidebar:
        st.header("Filtros")
        categorias = sorted([c for c in df["CATEGORIA"].dropna().unique() if str(c).strip()])
        riscos = sorted([r for r in df["RISCO PI"].dropna().unique() if str(r).strip()])
        estudar = sorted(df["VALE ESTUDAR?"].dropna().unique())

        filtro_categoria = st.multiselect("Categoria", categorias)
        filtro_risco = st.multiselect("Risco PI", riscos)
        filtro_estudar = st.multiselect("Vale estudar?", estudar, default=["Sim", "Talvez"] if "Sim" in estudar else estudar)

    filtered = df.copy()
    if filtro_categoria:
        filtered = filtered[filtered["CATEGORIA"].isin(filtro_categoria)]
    if filtro_risco:
        filtered = filtered[filtered["RISCO PI"].isin(filtro_risco)]
    if filtro_estudar:
        filtered = filtered[filtered["VALE ESTUDAR?"].isin(filtro_estudar)]

    st.subheader("Produtos filtrados")
    st.dataframe(filtered, use_container_width=True, height=420)

    st.download_button(
        "⬇️ Exportar Excel limpo",
        data=to_excel(filtered),
        file_name=f"MARS_Intelligence_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()

    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Produtos por categoria")
        if not filtered.empty:
            cat = filtered["CATEGORIA"].value_counts().reset_index()
            cat.columns = ["Categoria", "Produtos"]
            fig = px.bar(cat.head(15), x="Categoria", y="Produtos")
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Distribuição do MARS SCORE")
        if not filtered.empty:
            fig = px.histogram(filtered, x="MARS SCORE", nbins=20)
            st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)

    with c3:
        st.subheader("Preço médio por categoria")
        price_df = filtered.dropna(subset=["PREÇO LIMPO"])
        if not price_df.empty:
            avg = price_df.groupby("CATEGORIA", as_index=False)["PREÇO LIMPO"].mean()
            avg = avg.sort_values("PREÇO LIMPO", ascending=False).head(15)
            fig = px.bar(avg, x="CATEGORIA", y="PREÇO LIMPO")
            st.plotly_chart(fig, use_container_width=True)

    with c4:
        st.subheader("Risco de propriedade intelectual")
        risco = filtered["RISCO PI"].value_counts().reset_index()
        risco.columns = ["Risco", "Produtos"]
        if not risco.empty:
            fig = px.pie(risco, names="Risco", values="Produtos")
            st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Carrega o teu Excel para começar. A aplicação espera uma folha chamada PRODUTOS.")
