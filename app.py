
import io
import sqlite3
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="MARS Intelligence", page_icon="🪐", layout="wide")

COLS = [
    "LOJA","PRODUTO","CATEGORIA","TIPO PRODUTO","LICENCIADO","DESIGN ORIGINAL?",
    "LICENÇA COMERCIAL?","PREÇO (€)","PORTES (€)","FAVORITOS","AVALIAÇÕES",
    "CLASSIFICAÇÃO","BESTSELLER","ETSY PICK","VÍDEO","Nº FOTOS","PERSONALIZAÇÃO",
    "LINK","RISCO PI","DATA ANÁLISE","PAÍS","MATERIAL","COR","OBSERVAÇÕES"
]

HIGH = ["disney","marvel","pokemon","pokémon","star wars","hulk","spiderman","batman",
        "superman","mario","zelda","stitch","goofy","bob marley","asterix","popeye",
        "brutus","duffy","looney"]
MED = ["apple","airpod","magsafe","dyson","playstation","ps5","xbox","nintendo",
       "steelseries","shokz","alexa","echo dot","ikea"]

def price(v):
    if pd.isna(v) or str(v).strip()=="":
        return None
    t=str(v).replace("€","").replace(",",".").strip()
    if t.lower() in ["grátis","gratis","free"]:
        return 0
    try:
        return float(t)
    except Exception:
        return None

def normalize(df):
    df=df.copy()
    df.columns=[str(c).strip().upper() for c in df.columns]
    for c in COLS:
        if c not in df.columns:
            df[c]=""
    return df[COLS]

def auto_risk(txt):
    t=str(txt).lower()
    if any(w in t for w in HIGH):
        return "Alto"
    if any(w in t for w in MED):
        return "Médio"
    return "Baixo"

def score(row):
    s=50
    r=str(row.get("RISCO PI","")).lower()
    if r=="baixo":
        s+=15
    elif r in ["médio","medio"]:
        s+=5
    elif r=="alto":
        s-=20
    if str(row.get("DESIGN ORIGINAL?","")).lower() in ["sim","provável","provavel"]:
        s+=10
    if str(row.get("PERSONALIZAÇÃO","")).lower()=="sim":
        s+=10
    if str(row.get("VÍDEO","")).lower()=="sim":
        s+=5
    if str(row.get("BESTSELLER","")).lower()=="sim":
        s+=8
    if str(row.get("ETSY PICK","")).lower()=="sim":
        s+=5
    p=row.get("PREÇO LIMPO")
    if pd.notna(p):
        if p>=25:
            s+=10
        elif p<10:
            s-=5
    return max(0,min(100,int(s)))

def prepare(df):
    df=normalize(df)
    df["PREÇO LIMPO"]=df["PREÇO (€)"].apply(price)
    df["PORTES LIMPO"]=df["PORTES (€)"].apply(price)
    df=df.drop_duplicates(subset=["LOJA","PRODUTO","LINK"], keep="first")
    df["MARS SCORE"]=df.apply(score, axis=1)
    df["VALE ESTUDAR?"]=df["MARS SCORE"].apply(lambda x: "Sim" if x>=65 else "Talvez" if x>=50 else "Não")
    return df

def to_excel(df):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="PRODUTOS_LIMPOS", index=False)
        wb=writer.book
        ws=writer.sheets["PRODUTOS_LIMPOS"]
        fmt=wb.add_format({"bold":True,"bg_color":"#D9EAD3","border":1})
        for i,c in enumerate(df.columns):
            ws.write(0,i,c,fmt)
            ws.set_column(i,i,18)
    return bio.getvalue()

def load_file(up):
    if up.name.endswith(".csv"):
        raw=pd.read_csv(up)
    else:
        xls=pd.ExcelFile(up)
        sh="PRODUTOS" if "PRODUTOS" in xls.sheet_names else xls.sheet_names[0]
        raw=pd.read_excel(up, sheet_name=sh)
    return prepare(raw)

def export_sqlite(df):
    bio=io.BytesIO()
    conn=sqlite3.connect(":memory:")
    df.to_sql("produtos", conn, if_exists="replace", index=False)

    lojas = df.groupby("LOJA").agg(
        produtos=("PRODUTO","count"),
        score_medio=("MARS SCORE","mean"),
        preco_medio=("PREÇO LIMPO","mean")
    ).reset_index()
    lojas.to_sql("lojas", conn, if_exists="replace", index=False)

    script = "\n".join(conn.iterdump())
    bio.write(script.encode("utf-8"))
    bio.seek(0)
    conn.close()
    return bio.getvalue()

def empty_df():
    return pd.DataFrame(columns=COLS)

if "manual" not in st.session_state:
    st.session_state.manual=empty_df()
if "mars_products" not in st.session_state:
    st.session_state.mars_products=pd.DataFrame(columns=[
        "PRODUTO MARS","ESTADO","CATEGORIA","TEMPO IMPRESSÃO","PESO (g)",
        "CUSTO ESTIMADO (€)","PREÇO VENDA (€)","ETSY","EBAY","WALLAPOP",
        "TIKTOK","OBSERVAÇÕES"
    ])

st.sidebar.title("🪐 MARS Intelligence")
st.sidebar.caption("v0.5")
page=st.sidebar.radio("Menu", [
    "🏠 Dashboard","🌍 Mercado","➕ Novo Produto","📦 Produtos","🏪 Lojas",
    "📦 Produtos MARS","🧩 Roadmap","💰 Preços","⚠️ Risco PI","💡 Oportunidades",
    "📤 Exportar","⚙️ Base de Dados"
])

up=st.sidebar.file_uploader("Carregar Excel/CSV", type=["xlsx","xls","csv"])
if up:
    st.session_state.base=load_file(up)

base=st.session_state.get("base")
manual=prepare(st.session_state.manual) if len(st.session_state.manual) else pd.DataFrame()

if base is not None and len(manual):
    df=prepare(pd.concat([base[COLS], manual[COLS]], ignore_index=True))
elif base is not None:
    df=base
elif len(manual):
    df=manual
else:
    df=None

if df is not None:
    with st.sidebar:
        st.divider()
        st.subheader("Filtros")
        cat=st.multiselect("Categoria", sorted([x for x in df["CATEGORIA"].dropna().unique() if str(x).strip()]))
        risk=st.multiselect("Risco PI", sorted([x for x in df["RISCO PI"].dropna().unique() if str(x).strip()]))
        ve=st.multiselect("Vale estudar?", sorted(df["VALE ESTUDAR?"].dropna().unique()))
    f=df.copy()
    if cat:
        f=f[f["CATEGORIA"].isin(cat)]
    if risk:
        f=f[f["RISCO PI"].isin(risk)]
    if ve:
        f=f[f["VALE ESTUDAR?"].isin(ve)]
else:
    f=None

if page=="🏠 Dashboard":
    st.title("🏠 Dashboard Executivo")
    if f is None:
        st.info("Carrega o Excel/CSV ou adiciona produtos.")
        st.stop()
    c1,c2,c3,c4,c5,c6=st.columns(6)
    c1.metric("Produtos",len(f))
    c2.metric("Lojas",f["LOJA"].nunique())
    c3.metric("Preço médio",f"{f['PREÇO LIMPO'].mean():.2f} €" if f["PREÇO LIMPO"].notna().any() else "—")
    c4.metric("Score médio",f"{f['MARS SCORE'].mean():.1f}")
    c5.metric("Oportunidades",int((f["VALE ESTUDAR?"]=="Sim").sum()))
    c6.metric("Risco alto",int((f["RISCO PI"].astype(str).str.lower()=="alto").sum()))
    st.divider()
    a,b=st.columns(2)
    with a:
        cc=f["CATEGORIA"].value_counts().reset_index()
        cc.columns=["Categoria","Produtos"]
        if len(cc): st.plotly_chart(px.bar(cc.head(15),x="Categoria",y="Produtos",title="Top categorias"),use_container_width=True)
    with b:
        st.plotly_chart(px.histogram(f,x="MARS SCORE",nbins=20,title="Distribuição MARS SCORE"),use_container_width=True)
    st.subheader("Top 20 oportunidades")
    st.dataframe(f.sort_values("MARS SCORE",ascending=False).head(20),use_container_width=True)

elif page=="🌍 Mercado":
    st.title("🌍 Market Intelligence")
    if f is None:
        st.info("Carrega dados para analisar o mercado.")
        st.stop()
    k1,k2,k3=st.columns(3)
    k1.metric("% com vídeo", f"{(f['VÍDEO'].astype(str).str.lower()=='sim').mean()*100:.1f}%")
    k2.metric("% personalizável", f"{(f['PERSONALIZAÇÃO'].astype(str).str.lower()=='sim').mean()*100:.1f}%")
    k3.metric("% risco baixo", f"{(f['RISCO PI'].astype(str).str.lower()=='baixo').mean()*100:.1f}%")
    st.subheader("Resumo por categoria")
    resumo=f.groupby("CATEGORIA").agg(
        Produtos=("PRODUTO","count"),
        Lojas=("LOJA","nunique"),
        Preço_Médio=("PREÇO LIMPO","mean"),
        Score_Médio=("MARS SCORE","mean")
    ).reset_index().sort_values("Score_Médio",ascending=False)
    st.dataframe(resumo,use_container_width=True)

elif page=="➕ Novo Produto":
    st.title("➕ Ficha rápida de análise de produto")
    with st.form("form_produto"):
        c1,c2,c3=st.columns(3)
        with c1:
            loja=st.text_input("Loja")
            produto=st.text_input("Produto")
            categoria=st.selectbox("Categoria", ["Headphone Stand","Controller Holder","Desk Organizer","Smart Home Holder","Home Decor","Gaming Decor","Phone Holder","Planter","IKEA SKÅDIS","Outro"])
            tipo=st.text_input("Tipo produto")
            preco=st.text_input("Preço (€)")
            portes=st.text_input("Portes (€)")
            link=st.text_input("Link")
        with c2:
            lic=st.selectbox("Licenciado",["Não","Sim","A confirmar"])
            design=st.selectbox("Design original?",["Provável","Sim","Não","A confirmar"])
            licenca=st.selectbox("Licença comercial?",["Não","Sim","Desconhecida","A confirmar"])
            pers=st.selectbox("Personalização",["Não","Sim","A confirmar"])
            video=st.selectbox("Vídeo",["A confirmar","Sim","Não"])
            fotos=st.text_input("Nº fotos","A confirmar")
        with c3:
            fav=st.text_input("Favoritos")
            av=st.text_input("Avaliações")
            cla=st.text_input("Classificação")
            best=st.selectbox("Bestseller",["A confirmar","Sim","Não"])
            epick=st.selectbox("Etsy Pick",["A confirmar","Sim","Não"])
            pais=st.text_input("País")
            material=st.text_input("Material","A confirmar")
            cor=st.text_input("Cor")
        obs=st.text_area("Observações")
        sug=auto_risk(produto+" "+categoria+" "+obs)
        risco=st.selectbox("Risco PI",["Baixo","Médio","Alto","A confirmar"], index=["Baixo","Médio","Alto","A confirmar"].index(sug))
        ok=st.form_submit_button("Adicionar produto")
    if ok:
        row=pd.DataFrame([[loja,produto,categoria,tipo,lic,design,licenca,preco,portes,fav,av,cla,best,epick,video,fotos,pers,link,risco,datetime.now().strftime("%d/%m/%Y"),pais,material,cor,obs]], columns=COLS)
        st.session_state.manual=pd.concat([st.session_state.manual,row], ignore_index=True)
        st.success("Produto adicionado.")
    if len(st.session_state.manual):
        pv=prepare(st.session_state.manual)
        st.dataframe(pv,use_container_width=True)
        st.download_button("⬇️ Exportar produtos adicionados", to_excel(pv), "MARS_produtos_adicionados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page=="📦 Produtos":
    st.title("📦 Produtos")
    if f is None: st.info("Sem dados."); st.stop()
    q=st.text_input("Pesquisar produto/loja/link")
    view=f.copy()
    if q:
        ql=q.lower()
        view=view[view.apply(lambda r: ql in str(r.to_dict()).lower(), axis=1)]
    st.dataframe(view,use_container_width=True,height=650)

elif page=="🏪 Lojas":
    st.title("🏪 Lojas")
    if f is None: st.info("Sem dados."); st.stop()
    lojas=f.groupby("LOJA").agg(
        Produtos=("PRODUTO","count"),
        Score_Médio=("MARS SCORE","mean"),
        Preço_Médio=("PREÇO LIMPO","mean"),
        Oportunidades=("VALE ESTUDAR?", lambda x: (x=="Sim").sum())
    ).reset_index().sort_values("Score_Médio",ascending=False)
    st.dataframe(lojas,use_container_width=True,height=650)

elif page=="📦 Produtos MARS":
    st.title("📦 Produtos MARS")
    st.caption("Gestão simples do teu portfólio MARS 3D.")
    with st.form("mars_form"):
        c1,c2,c3=st.columns(3)
        with c1:
            mp=st.text_input("Produto MARS")
            estado=st.selectbox("Estado",["Ideia","Modelação","Protótipo","Teste impressão","Fotografia","Vídeo","Publicado Etsy","Publicado marketplaces","Primeira venda","Otimização"])
            catm=st.text_input("Categoria")
            tempo=st.text_input("Tempo impressão")
        with c2:
            peso=st.text_input("Peso (g)")
            custo=st.text_input("Custo estimado (€)")
            pvenda=st.text_input("Preço venda (€)")
            etsy=st.selectbox("Etsy",["Não","Sim"])
        with c3:
            ebay=st.selectbox("eBay",["Não","Sim"])
            wall=st.selectbox("Wallapop",["Não","Sim"])
            tiktok=st.selectbox("TikTok",["Não","Sim"])
            obsm=st.text_area("Observações MARS")
        add=st.form_submit_button("Adicionar produto MARS")
    if add:
        row=pd.DataFrame([[mp,estado,catm,tempo,peso,custo,pvenda,etsy,ebay,wall,tiktok,obsm]], columns=st.session_state.mars_products.columns)
        st.session_state.mars_products=pd.concat([st.session_state.mars_products,row], ignore_index=True)
        st.success("Produto MARS adicionado.")
    st.dataframe(st.session_state.mars_products,use_container_width=True)
    if len(st.session_state.mars_products):
        st.download_button("⬇️ Exportar Produtos MARS", to_excel(st.session_state.mars_products), "MARS_produtos_proprios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page=="🧩 Roadmap":
    st.title("🧩 Roadmap de desenvolvimento")
    st.write("Funil dos produtos MARS por estado.")
    if len(st.session_state.mars_products):
        r=st.session_state.mars_products["ESTADO"].value_counts().reset_index()
        r.columns=["Estado","Produtos"]
        st.plotly_chart(px.bar(r,x="Estado",y="Produtos"),use_container_width=True)
        st.dataframe(st.session_state.mars_products,use_container_width=True)
    else:
        st.info("Adiciona produtos em 'Produtos MARS' para ver o roadmap.")

elif page=="💰 Preços":
    st.title("💰 Preços")
    if f is None: st.info("Sem dados."); st.stop()
    p=f.dropna(subset=["PREÇO LIMPO"])
    if len(p):
        st.plotly_chart(px.histogram(p,x="PREÇO LIMPO",nbins=25,title="Distribuição de preços"),use_container_width=True)
        st.dataframe(p.sort_values("PREÇO LIMPO",ascending=False),use_container_width=True)

elif page=="⚠️ Risco PI":
    st.title("⚠️ Risco de Propriedade Intelectual")
    if f is None: st.info("Sem dados."); st.stop()
    r=f["RISCO PI"].value_counts().reset_index()
    r.columns=["Risco","Produtos"]
    if len(r): st.plotly_chart(px.pie(r,names="Risco",values="Produtos"),use_container_width=True)
    st.dataframe(f[f["RISCO PI"].astype(str).str.lower()=="alto"],use_container_width=True)

elif page=="💡 Oportunidades":
    st.title("💡 Oportunidades")
    if f is None: st.info("Sem dados."); st.stop()
    o=f[(f["MARS SCORE"]>=65)&(f["RISCO PI"].astype(str).str.lower().isin(["baixo","médio","medio"]))].sort_values("MARS SCORE",ascending=False)
    st.dataframe(o,use_container_width=True,height=650)

elif page=="📤 Exportar":
    st.title("📤 Exportar")
    if f is None: st.info("Sem dados."); st.stop()
    st.download_button("⬇️ Exportar Excel filtrado", to_excel(f), f"MARS_Intelligence_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page=="⚙️ Base de Dados":
    st.title("⚙️ Base de Dados")
    st.write("Exportação experimental da base para SQL/SQLite.")
    if f is None:
        st.info("Sem dados.")
        st.stop()
    st.download_button("⬇️ Exportar dump SQL", export_sqlite(f), "MARS_Intelligence_SQL_Dump.sql", mime="text/plain")
