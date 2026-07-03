
import io
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="MARS Intelligence", page_icon="🪐", layout="wide")

COLS=["LOJA","PRODUTO","CATEGORIA","TIPO PRODUTO","LICENCIADO","DESIGN ORIGINAL?","LICENÇA COMERCIAL?","PREÇO (€)","PORTES (€)","FAVORITOS","AVALIAÇÕES","CLASSIFICAÇÃO","BESTSELLER","ETSY PICK","VÍDEO","Nº FOTOS","PERSONALIZAÇÃO","LINK","RISCO PI","DATA ANÁLISE","PAÍS","MATERIAL","COR","OBSERVAÇÕES"]
MARS_COLS=["PRODUTO MARS","ESTADO","CATEGORIA","TEMPO IMPRESSÃO","PESO (g)","CUSTO ESTIMADO (€)","PREÇO VENDA (€)","LUCRO ESTIMADO (€)","MARGEM (%)","ETSY","EBAY","WALLAPOP","TIKTOK","OBSERVAÇÕES"]
HIGH=["disney","marvel","pokemon","pokémon","star wars","hulk","spiderman","batman","superman","mario","zelda","stitch","goofy","bob marley","asterix","popeye","brutus","duffy","looney"]
MED=["apple","airpod","magsafe","dyson","playstation","ps5","xbox","nintendo","steelseries","shokz","alexa","echo dot","ikea"]

def price(v):
    if pd.isna(v) or str(v).strip()=="": return None
    t=str(v).replace("€","").replace(",",".").strip()
    if t.lower() in ["grátis","gratis","free"]: return 0
    try: return float(t)
    except: return None

def normalize(df):
    df=df.copy(); df.columns=[str(c).strip().upper() for c in df.columns]
    for c in COLS:
        if c not in df.columns: df[c]=""
    return df[COLS]

def auto_risk(txt):
    t=str(txt).lower()
    if any(w in t for w in HIGH): return "Alto"
    if any(w in t for w in MED): return "Médio"
    return "Baixo"

def score(r):
    s=50; risco=str(r.get("RISCO PI","")).lower()
    if risco=="baixo": s+=15
    elif risco in ["médio","medio"]: s+=5
    elif risco=="alto": s-=20
    if str(r.get("DESIGN ORIGINAL?","")).lower() in ["sim","provável","provavel"]: s+=10
    if str(r.get("PERSONALIZAÇÃO","")).lower()=="sim": s+=10
    if str(r.get("VÍDEO","")).lower()=="sim": s+=5
    if str(r.get("BESTSELLER","")).lower()=="sim": s+=8
    if str(r.get("ETSY PICK","")).lower()=="sim": s+=5
    p=r.get("PREÇO LIMPO")
    if pd.notna(p):
        if p>=25: s+=10
        elif p<10: s-=5
    return max(0,min(100,int(s)))

def prepare(df):
    df=normalize(df)
    df["PREÇO LIMPO"]=df["PREÇO (€)"].apply(price)
    df["PORTES LIMPO"]=df["PORTES (€)"].apply(price)
    df=df.drop_duplicates(subset=["LOJA","PRODUTO","LINK"], keep="first")
    df["MARS SCORE"]=df.apply(score,axis=1)
    df["VALE ESTUDAR?"]=df["MARS SCORE"].apply(lambda x:"Sim" if x>=65 else "Talvez" if x>=50 else "Não")
    return df

def to_excel(df, sheet="DADOS"):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine="xlsxwriter") as w:
        df.to_excel(w,sheet_name=sheet[:31],index=False)
        wb=w.book; ws=w.sheets[sheet[:31]]
        fmt=wb.add_format({"bold":True,"bg_color":"#D9EAD3","border":1})
        for i,c in enumerate(df.columns):
            ws.write(0,i,c,fmt); ws.set_column(i,i,18)
    return bio.getvalue()

def load(up):
    if up.name.endswith(".csv"): raw=pd.read_csv(up)
    else:
        xls=pd.ExcelFile(up); sh="PRODUTOS" if "PRODUTOS" in xls.sheet_names else xls.sheet_names[0]
        raw=pd.read_excel(up,sheet_name=sh)
    return prepare(raw)

def empty_market(): return pd.DataFrame(columns=COLS)
def empty_mars(): return pd.DataFrame(columns=MARS_COLS)

if "manual" not in st.session_state: st.session_state.manual=empty_market()
if "mars" not in st.session_state: st.session_state.mars=empty_mars()

st.sidebar.title("🪐 MARS Intelligence"); st.sidebar.caption("v0.6")
page=st.sidebar.radio("Menu",["🏠 Dashboard","🚨 Radar MARS","➕ Novo Produto","📦 Produtos","🏪 Lojas","📦 Produtos MARS","🧩 Roadmap","🧮 Produção/Custos","💰 Preços","⚠️ Risco PI","💡 Oportunidades","📤 Exportar"])
up=st.sidebar.file_uploader("Carregar Excel/CSV",type=["xlsx","xls","csv"])
if up: st.session_state.base=load(up)
base=st.session_state.get("base")
manual=prepare(st.session_state.manual) if len(st.session_state.manual) else pd.DataFrame()
df=prepare(pd.concat([x[COLS] for x in [base,manual] if x is not None and len(x)],ignore_index=True)) if ((base is not None and len(base)) or len(manual)) else None

if df is not None:
    with st.sidebar:
        st.divider(); st.subheader("Filtros")
        cat=st.multiselect("Categoria",sorted([x for x in df["CATEGORIA"].dropna().unique() if str(x).strip()]))
        risk=st.multiselect("Risco PI",sorted([x for x in df["RISCO PI"].dropna().unique() if str(x).strip()]))
        ve=st.multiselect("Vale estudar?",sorted(df["VALE ESTUDAR?"].dropna().unique()))
    f=df.copy()
    if cat: f=f[f["CATEGORIA"].isin(cat)]
    if risk: f=f[f["RISCO PI"].isin(risk)]
    if ve: f=f[f["VALE ESTUDAR?"].isin(ve)]
else: f=None

if page=="🏠 Dashboard":
    st.title("🏠 Dashboard Executivo")
    if f is None: st.info("Carrega dados para começar."); st.stop()
    c=st.columns(6)
    c[0].metric("Produtos",len(f)); c[1].metric("Lojas",f["LOJA"].nunique()); c[2].metric("Categorias",f["CATEGORIA"].nunique())
    c[3].metric("Preço médio",f"{f['PREÇO LIMPO'].mean():.2f} €" if f["PREÇO LIMPO"].notna().any() else "—")
    c[4].metric("Score médio",f"{f['MARS SCORE'].mean():.1f}"); c[5].metric("Oportunidades",int((f["VALE ESTUDAR?"]=="Sim").sum()))
    c2=st.columns(4)
    c2[0].metric("Risco alto",int((f["RISCO PI"].astype(str).str.lower()=="alto").sum()))
    c2[1].metric("Com vídeo",int((f["VÍDEO"].astype(str).str.lower()=="sim").sum()))
    c2[2].metric("Personalizáveis",int((f["PERSONALIZAÇÃO"].astype(str).str.lower()=="sim").sum()))
    c2[3].metric("Premium >40€",int((f["PREÇO LIMPO"]>40).sum()))
    a,b=st.columns(2)
    with a:
        cc=f["CATEGORIA"].value_counts().reset_index(); cc.columns=["Categoria","Produtos"]
        if len(cc): st.plotly_chart(px.bar(cc.head(15),x="Categoria",y="Produtos"),use_container_width=True)
    with b: st.plotly_chart(px.histogram(f,x="MARS SCORE",nbins=20),use_container_width=True)
    st.dataframe(f.sort_values("MARS SCORE",ascending=False).head(20),use_container_width=True)

elif page=="🚨 Radar MARS":
    st.title("🚨 Radar MARS")
    if f is None: st.info("Carrega dados para ativar o radar."); st.stop()
    r=f.copy()
    r["AÇÃO SUGERIDA"]=r.apply(lambda x:"DESENVOLVER / ESTUDAR" if x["MARS SCORE"]>=75 and str(x["RISCO PI"]).lower()!="alto" else ("ANALISAR" if x["MARS SCORE"]>=60 else "EVITAR"),axis=1)
    k=st.columns(3); k[0].metric("Desenvolver",int((r["AÇÃO SUGERIDA"]=="DESENVOLVER / ESTUDAR").sum())); k[1].metric("Analisar",int((r["AÇÃO SUGERIDA"]=="ANALISAR").sum())); k[2].metric("Evitar",int((r["AÇÃO SUGERIDA"]=="EVITAR").sum()))
    st.dataframe(r.sort_values("MARS SCORE",ascending=False)[["LOJA","PRODUTO","CATEGORIA","PREÇO (€)","RISCO PI","MARS SCORE","AÇÃO SUGERIDA","LINK"]],use_container_width=True,height=650)

elif page=="➕ Novo Produto":
    st.title("➕ Novo Produto")
    with st.form("novo"):
        c1,c2,c3=st.columns(3)
        with c1:
            loja=st.text_input("Loja"); produto=st.text_input("Produto")
            categoria=st.selectbox("Categoria",["Headphone Stand","Controller Holder","Desk Organizer","Smart Home Holder","Home Decor","Gaming Decor","Phone Holder","Planter","IKEA SKÅDIS","Outro"])
            tipo=st.text_input("Tipo produto"); preco=st.text_input("Preço (€)"); portes=st.text_input("Portes (€)"); link=st.text_input("Link")
        with c2:
            lic=st.selectbox("Licenciado",["Não","Sim","A confirmar"]); design=st.selectbox("Design original?",["Provável","Sim","Não","A confirmar"])
            licenca=st.selectbox("Licença comercial?",["Não","Sim","Desconhecida","A confirmar"]); pers=st.selectbox("Personalização",["Não","Sim","A confirmar"])
            video=st.selectbox("Vídeo",["A confirmar","Sim","Não"]); fotos=st.text_input("Nº fotos","A confirmar")
        with c3:
            fav=st.text_input("Favoritos"); av=st.text_input("Avaliações"); cla=st.text_input("Classificação")
            best=st.selectbox("Bestseller",["A confirmar","Sim","Não"]); epick=st.selectbox("Etsy Pick",["A confirmar","Sim","Não"])
            pais=st.text_input("País"); material=st.text_input("Material","A confirmar"); cor=st.text_input("Cor")
        obs=st.text_area("Observações"); sug=auto_risk(produto+" "+categoria+" "+obs)
        risco=st.selectbox("Risco PI",["Baixo","Médio","Alto","A confirmar"],index=["Baixo","Médio","Alto","A confirmar"].index(sug))
        ok=st.form_submit_button("Adicionar")
    if ok:
        row=pd.DataFrame([[loja,produto,categoria,tipo,lic,design,licenca,preco,portes,fav,av,cla,best,epick,video,fotos,pers,link,risco,datetime.now().strftime("%d/%m/%Y"),pais,material,cor,obs]],columns=COLS)
        st.session_state.manual=pd.concat([st.session_state.manual,row],ignore_index=True); st.success("Produto adicionado.")
    if len(st.session_state.manual):
        pv=prepare(st.session_state.manual); st.dataframe(pv,use_container_width=True)
        st.download_button("⬇️ Exportar adicionados",to_excel(pv),"MARS_produtos_adicionados.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page=="📦 Produtos":
    st.title("📦 Produtos")
    if f is None: st.info("Sem dados."); st.stop()
    q=st.text_input("Pesquisar")
    view=f if not q else f[f.apply(lambda r:q.lower() in str(r.to_dict()).lower(),axis=1)]
    st.dataframe(view,use_container_width=True,height=650)

elif page=="🏪 Lojas":
    st.title("🏪 Lojas")
    if f is None: st.info("Sem dados."); st.stop()
    lojas=f.groupby("LOJA").agg(Produtos=("PRODUTO","count"),Score_Médio=("MARS SCORE","mean"),Preço_Médio=("PREÇO LIMPO","mean"),Oportunidades=("VALE ESTUDAR?",lambda x:(x=="Sim").sum())).reset_index().sort_values("Score_Médio",ascending=False)
    st.dataframe(lojas,use_container_width=True,height=650)

elif page=="📦 Produtos MARS":
    st.title("📦 Produtos MARS")
    with st.form("mars"):
        c1,c2,c3=st.columns(3)
        with c1:
            mp=st.text_input("Produto MARS"); estado=st.selectbox("Estado",["Ideia","Modelação","Protótipo","Teste impressão","Fotografia","Vídeo","Publicado Etsy","Publicado marketplaces","Primeira venda","Otimização"]); catm=st.text_input("Categoria"); tempo=st.text_input("Tempo impressão")
        with c2:
            peso=st.text_input("Peso (g)"); custo=st.text_input("Custo estimado (€)"); pvenda=st.text_input("Preço venda (€)"); lucro=st.text_input("Lucro estimado (€)"); margem=st.text_input("Margem (%)")
        with c3:
            etsy=st.selectbox("Etsy",["Não","Sim"]); ebay=st.selectbox("eBay",["Não","Sim"]); wall=st.selectbox("Wallapop",["Não","Sim"]); tiktok=st.selectbox("TikTok",["Não","Sim"]); obsm=st.text_area("Observações")
        add=st.form_submit_button("Adicionar")
    if add:
        row=pd.DataFrame([[mp,estado,catm,tempo,peso,custo,pvenda,lucro,margem,etsy,ebay,wall,tiktok,obsm]],columns=MARS_COLS)
        st.session_state.mars=pd.concat([st.session_state.mars,row],ignore_index=True); st.success("Produto MARS adicionado.")
    st.dataframe(st.session_state.mars,use_container_width=True)
    if len(st.session_state.mars): st.download_button("⬇️ Exportar Produtos MARS",to_excel(st.session_state.mars,"PRODUTOS_MARS"),"MARS_produtos_proprios.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page=="🧩 Roadmap":
    st.title("🧩 Roadmap")
    if len(st.session_state.mars):
        rr=st.session_state.mars["ESTADO"].value_counts().reset_index(); rr.columns=["Estado","Produtos"]
        st.plotly_chart(px.bar(rr,x="Estado",y="Produtos"),use_container_width=True); st.dataframe(st.session_state.mars,use_container_width=True)
    else: st.info("Adiciona produtos MARS primeiro.")

elif page=="🧮 Produção/Custos":
    st.title("🧮 Calculadora de Produção e Margem")
    a,b,c=st.columns(3)
    with a: peso=st.number_input("Peso (g)",0.0,value=180.0); fil=st.number_input("Filamento €/kg",0.0,value=18.0); horas=st.number_input("Horas impressão",0.0,value=6.0)
    with b: energia=st.number_input("Energia/desgaste €/h",0.0,value=0.15); emb=st.number_input("Embalagem €",0.0,value=0.80); com=st.number_input("Comissões %",0.0,value=12.0)
    with c: venda=st.number_input("Preço venda €",0.0,value=34.90)
    material=peso/1000*fil; e=horas*energia; fees=venda*com/100; total=material+e+emb+fees; lucro=venda-total; margem=lucro/venda*100 if venda else 0
    m=st.columns(6)
    for col,label,val in zip(m,["Material","Energia","Embalagem","Comissões","Custo total","Lucro"],[material,e,emb,fees,total,lucro]): col.metric(label,f"{val:.2f} €")
    st.metric("Margem",f"{margem:.1f}%")

elif page=="💰 Preços":
    st.title("💰 Preços")
    if f is None: st.info("Sem dados."); st.stop()
    p=f.dropna(subset=["PREÇO LIMPO"])
    if len(p): st.plotly_chart(px.histogram(p,x="PREÇO LIMPO",nbins=25),use_container_width=True)
    st.dataframe(p.sort_values("PREÇO LIMPO",ascending=False),use_container_width=True)

elif page=="⚠️ Risco PI":
    st.title("⚠️ Risco PI")
    if f is None: st.info("Sem dados."); st.stop()
    r=f["RISCO PI"].value_counts().reset_index(); r.columns=["Risco","Produtos"]
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
    st.download_button("⬇️ Exportar Excel",to_excel(f),f"MARS_Intelligence_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
