
import io
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='MARS Intelligence', page_icon='🪐', layout='wide')

COLS = ['LOJA','PRODUTO','CATEGORIA','TIPO PRODUTO','LICENCIADO','DESIGN ORIGINAL?','LICENÇA COMERCIAL?','PREÇO (€)','PORTES (€)','FAVORITOS','AVALIAÇÕES','CLASSIFICAÇÃO','BESTSELLER','ETSY PICK','VÍDEO','Nº FOTOS','PERSONALIZAÇÃO','LINK','RISCO PI','DATA ANÁLISE','PAÍS','MATERIAL','COR','OBSERVAÇÕES']
HIGH = ['disney','marvel','pokemon','pokémon','star wars','hulk','spiderman','batman','superman','mario','zelda','stitch','goofy','bob marley','asterix','popeye','brutus','duffy','looney']
MED = ['apple','airpod','magsafe','dyson','playstation','ps5','xbox','nintendo','steelseries','shokz','alexa','echo dot','ikea']

def price(v):
    if pd.isna(v) or str(v).strip() == '':
        return None
    t = str(v).replace('€','').replace(',','.').strip()
    if t.lower() in ['grátis','gratis','free']:
        return 0
    try:
        return float(t)
    except Exception:
        return None

def auto_risk(txt):
    t = str(txt).lower()
    if any(w in t for w in HIGH):
        return 'Alto'
    if any(w in t for w in MED):
        return 'Médio'
    return 'Baixo'

def normalize(df):
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    for c in COLS:
        if c not in df.columns:
            df[c] = ''
    return df[COLS]

def score(r):
    s = 50
    risco = str(r.get('RISCO PI','')).lower()
    if risco == 'baixo':
        s += 15
    elif risco in ['médio','medio']:
        s += 5
    elif risco == 'alto':
        s -= 20
    if str(r.get('DESIGN ORIGINAL?','')).lower() in ['sim','provável','provavel']:
        s += 10
    if str(r.get('PERSONALIZAÇÃO','')).lower() == 'sim':
        s += 10
    if str(r.get('VÍDEO','')).lower() == 'sim':
        s += 5
    if str(r.get('BESTSELLER','')).lower() == 'sim':
        s += 8
    if str(r.get('ETSY PICK','')).lower() == 'sim':
        s += 5
    p = r.get('PREÇO LIMPO')
    if pd.notna(p):
        if p >= 25:
            s += 10
        elif p < 10:
            s -= 5
    return max(0, min(100, int(s)))

def prepare(df):
    df = normalize(df)
    df['PREÇO LIMPO'] = df['PREÇO (€)'].apply(price)
    df['PORTES LIMPO'] = df['PORTES (€)'].apply(price)
    df['RISCO PI'] = df.apply(lambda r: auto_risk(str(r['PRODUTO']) + ' ' + str(r['OBSERVAÇÕES'])) if str(r['RISCO PI']).strip()=='' else r['RISCO PI'], axis=1)
    df = df.drop_duplicates(subset=['LOJA','PRODUTO','LINK'], keep='first')
    df['MARS SCORE'] = df.apply(score, axis=1)
    df['VALE ESTUDAR?'] = df['MARS SCORE'].apply(lambda x: 'Sim' if x>=65 else 'Talvez' if x>=50 else 'Não')
    return df

def to_excel(df, sheet='DADOS'):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='xlsxwriter') as w:
        df.to_excel(w, sheet_name=sheet[:31], index=False)
        wb=w.book; ws=w.sheets[sheet[:31]]
        fmt=wb.add_format({'bold':True,'bg_color':'#D9EAD3','border':1})
        for i,c in enumerate(df.columns):
            ws.write(0,i,c,fmt); ws.set_column(i,i,18)
    return bio.getvalue()

def load_file(up):
    if up.name.endswith('.csv'):
        raw = pd.read_csv(up)
    else:
        xls = pd.ExcelFile(up)
        sh = 'PRODUTOS' if 'PRODUTOS' in xls.sheet_names else xls.sheet_names[0]
        raw = pd.read_excel(up, sheet_name=sh)
    return prepare(raw)

if 'manual' not in st.session_state:
    st.session_state.manual = pd.DataFrame(columns=COLS)

st.sidebar.title('🪐 MARS Intelligence')
st.sidebar.caption('v0.8')
page = st.sidebar.radio('Menu', ['🏠 Dashboard','🚨 Radar MARS','📥 Importação em Lote','🔎 Ficha Produto','⚖️ Comparador MARS','➕ Novo Produto','📦 Produtos','🏪 Lojas','💰 Preços','⚠️ Risco PI','💡 Oportunidades','📤 Exportar'])

up = st.sidebar.file_uploader('Carregar Excel/CSV', type=['xlsx','xls','csv'])
if up:
    st.session_state.base = load_file(up)

base_df = st.session_state.get('base')
manual = prepare(st.session_state.manual) if len(st.session_state.manual) else pd.DataFrame()
parts = []
if base_df is not None and len(base_df): parts.append(base_df[COLS])
if len(manual): parts.append(manual[COLS])
df = prepare(pd.concat(parts, ignore_index=True)) if parts else None

if df is not None:
    with st.sidebar:
        st.divider()
        st.subheader('Filtros')
        cats = sorted([x for x in df['CATEGORIA'].dropna().unique() if str(x).strip()])
        risks = sorted([x for x in df['RISCO PI'].dropna().unique() if str(x).strip()])
        vals = sorted(df['VALE ESTUDAR?'].dropna().unique())
        fc = st.multiselect('Categoria', cats)
        fr = st.multiselect('Risco PI', risks)
        fv = st.multiselect('Vale estudar?', vals)
    f = df.copy()
    if fc: f = f[f['CATEGORIA'].isin(fc)]
    if fr: f = f[f['RISCO PI'].isin(fr)]
    if fv: f = f[f['VALE ESTUDAR?'].isin(fv)]
else:
    f = None

if page == '🏠 Dashboard':
    st.title('🏠 Dashboard Executivo')
    if f is None:
        st.info('Carrega dados para começar.')
        st.stop()
    c = st.columns(6)
    c[0].metric('Produtos', len(f))
    c[1].metric('Lojas', f['LOJA'].nunique())
    c[2].metric('Categorias', f['CATEGORIA'].nunique())
    c[3].metric('Preço médio', f"{f['PREÇO LIMPO'].mean():.2f} €" if f['PREÇO LIMPO'].notna().any() else '—')
    c[4].metric('Score médio', f"{f['MARS SCORE'].mean():.1f}")
    c[5].metric('Oportunidades', int((f['VALE ESTUDAR?']=='Sim').sum()))
    a,b = st.columns(2)
    with a:
        cc = f['CATEGORIA'].value_counts().reset_index(); cc.columns=['Categoria','Produtos']
        if len(cc): st.plotly_chart(px.bar(cc.head(15), x='Categoria', y='Produtos'), use_container_width=True)
    with b:
        st.plotly_chart(px.histogram(f, x='MARS SCORE', nbins=20), use_container_width=True)
    st.dataframe(f.sort_values('MARS SCORE', ascending=False).head(20), use_container_width=True)

elif page == '🚨 Radar MARS':
    st.title('🚨 Radar MARS')
    if f is None:
        st.info('Carrega dados para ativar o Radar.')
        st.stop()
    r = f.copy()
    r['AÇÃO SUGERIDA'] = r.apply(lambda x: 'DESENVOLVER / ESTUDAR' if x['MARS SCORE']>=75 and str(x['RISCO PI']).lower()!='alto' else ('ANALISAR' if x['MARS SCORE']>=60 else 'EVITAR'), axis=1)
    st.dataframe(r.sort_values('MARS SCORE', ascending=False)[['LOJA','PRODUTO','CATEGORIA','PREÇO (€)','RISCO PI','MARS SCORE','AÇÃO SUGERIDA','LINK']], use_container_width=True, height=650)

elif page == '🔎 Ficha Produto':
    st.title('🔎 Ficha de Produto')
    if f is None:
        st.info('Carrega dados primeiro.')
        st.stop()
    options = (f['PRODUTO'].fillna('') + ' — ' + f['LOJA'].fillna('')).tolist()
    sel = st.selectbox('Escolher produto', options)
    row = f.iloc[options.index(sel)]
    c = st.columns(4)
    c[0].metric('Preço', str(row.get('PREÇO (€)','')))
    c[1].metric('MARS Score', int(row.get('MARS SCORE',0)))
    c[2].metric('Risco PI', str(row.get('RISCO PI','')))
    c[3].metric('Vale estudar?', str(row.get('VALE ESTUDAR?','')))
    st.subheader(row['PRODUTO'])
    st.write(f"**Loja:** {row['LOJA']} | **Categoria:** {row['CATEGORIA']} | **Tipo:** {row['TIPO PRODUTO']}")
    if str(row.get('LINK','')).startswith('http'):
        st.link_button('Abrir anúncio', row['LINK'])
    st.dataframe(pd.DataFrame({'Campo': row.index, 'Valor': row.values}), use_container_width=True)
    st.subheader('Produtos semelhantes na mesma categoria')
    sim = f[(f['CATEGORIA']==row['CATEGORIA']) & (f['PRODUTO']!=row['PRODUTO'])].sort_values('MARS SCORE', ascending=False).head(20)
    st.dataframe(sim[['LOJA','PRODUTO','PREÇO (€)','RISCO PI','MARS SCORE','LINK']], use_container_width=True)

elif page == '⚖️ Comparador MARS':
    st.title('⚖️ Comparador MARS vs Mercado')
    if f is None:
        st.info('Carrega dados primeiro.')
        st.stop()
    categoria = st.selectbox('Categoria de mercado', sorted([x for x in f['CATEGORIA'].dropna().unique() if str(x).strip()]))
    mercado = f[f['CATEGORIA']==categoria].dropna(subset=['PREÇO LIMPO'])
    c1,c2,c3 = st.columns(3)
    meu_preco = c1.number_input('Preço MARS previsto (€)', 0.0, value=29.90)
    meu_custo = c2.number_input('Custo total estimado (€)', 0.0, value=5.00)
    horas = c3.number_input('Horas de impressão', 0.0, value=6.0)
    lucro = meu_preco - meu_custo
    margem = lucro / meu_preco * 100 if meu_preco else 0
    k = st.columns(5)
    k[0].metric('Preço médio mercado', f"{mercado['PREÇO LIMPO'].mean():.2f} €" if len(mercado) else '—')
    k[1].metric('Preço mediano', f"{mercado['PREÇO LIMPO'].median():.2f} €" if len(mercado) else '—')
    k[2].metric('Teu lucro', f"{lucro:.2f} €")
    k[3].metric('Tua margem', f"{margem:.1f}%")
    k[4].metric('Lucro/hora', f"{lucro/horas:.2f} €/h" if horas else '—')
    if len(mercado):
        st.plotly_chart(px.histogram(mercado, x='PREÇO LIMPO', nbins=20, title=f'Distribuição de preços — {categoria}'), use_container_width=True)
        st.dataframe(mercado.sort_values('MARS SCORE', ascending=False)[['LOJA','PRODUTO','PREÇO (€)','RISCO PI','MARS SCORE','LINK']].head(30), use_container_width=True)

elif page == '📥 Importação em Lote':
    st.title('📥 Importação em Lote')
    st.download_button('⬇️ Descarregar template Excel', to_excel(pd.DataFrame(columns=COLS), 'PRODUTOS'), 'MARS_template_produtos.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    txt = st.text_area('Colar tabela TSV/CSV', height=250)
    sep = st.selectbox('Separador', ['Tabulação','Ponto e vírgula','Vírgula'])
    if st.button('Pré-visualizar importação') and txt.strip():
        delimiter = {'Tabulação':'\t','Ponto e vírgula':';','Vírgula':','}[sep]
        try:
            imported = pd.read_csv(io.StringIO(txt), sep=delimiter)
            imported = prepare(imported)
            st.session_state.preview_import = imported
            st.success(f'{len(imported)} produtos preparados.')
        except Exception as e:
            st.error(f'Erro: {e}')
    if 'preview_import' in st.session_state:
        st.dataframe(st.session_state.preview_import, use_container_width=True)
        if st.button('Adicionar à sessão'):
            st.session_state.manual = pd.concat([st.session_state.manual, st.session_state.preview_import[COLS]], ignore_index=True)
            st.success('Produtos adicionados.')
        st.download_button('⬇️ Exportar importação limpa', to_excel(st.session_state.preview_import), 'MARS_importacao_lote_limpa.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

elif page == '➕ Novo Produto':
    st.title('➕ Novo Produto')
    with st.form('novo'):
        c1,c2,c3 = st.columns(3)
        with c1:
            loja = st.text_input('Loja'); produto = st.text_input('Produto')
            categoria = st.selectbox('Categoria',['Headphone Stand','Controller Holder','Desk Organizer','Smart Home Holder','Home Decor','Gaming Decor','Phone Holder','Planter','IKEA SKÅDIS','Outro'])
            tipo = st.text_input('Tipo produto'); preco = st.text_input('Preço (€)'); portes = st.text_input('Portes (€)'); link = st.text_input('Link')
        with c2:
            lic = st.selectbox('Licenciado',['Não','Sim','A confirmar']); design = st.selectbox('Design original?',['Provável','Sim','Não','A confirmar'])
            licenca = st.selectbox('Licença comercial?',['Não','Sim','Desconhecida','A confirmar']); pers = st.selectbox('Personalização',['Não','Sim','A confirmar'])
            video = st.selectbox('Vídeo',['A confirmar','Sim','Não']); fotos = st.text_input('Nº fotos','A confirmar')
        with c3:
            fav = st.text_input('Favoritos'); av = st.text_input('Avaliações'); cla = st.text_input('Classificação')
            best = st.selectbox('Bestseller',['A confirmar','Sim','Não']); epick = st.selectbox('Etsy Pick',['A confirmar','Sim','Não'])
            pais = st.text_input('País'); material = st.text_input('Material','A confirmar'); cor = st.text_input('Cor')
        obs = st.text_area('Observações'); sug = auto_risk(produto+' '+categoria+' '+obs)
        risco = st.selectbox('Risco PI',['Baixo','Médio','Alto','A confirmar'], index=['Baixo','Médio','Alto','A confirmar'].index(sug))
        ok = st.form_submit_button('Adicionar')
    if ok:
        row = pd.DataFrame([[loja,produto,categoria,tipo,lic,design,licenca,preco,portes,fav,av,cla,best,epick,video,fotos,pers,link,risco,datetime.now().strftime('%d/%m/%Y'),pais,material,cor,obs]], columns=COLS)
        st.session_state.manual = pd.concat([st.session_state.manual,row], ignore_index=True); st.success('Produto adicionado.')

elif page == '📦 Produtos':
    st.title('📦 Produtos')
    if f is None: st.info('Sem dados.'); st.stop()
    q = st.text_input('Pesquisar')
    view = f if not q else f[f.apply(lambda r: q.lower() in str(r.to_dict()).lower(), axis=1)]
    st.dataframe(view, use_container_width=True, height=650)

elif page == '🏪 Lojas':
    st.title('🏪 Lojas')
    if f is None: st.info('Sem dados.'); st.stop()
    lojas = f.groupby('LOJA').agg(Produtos=('PRODUTO','count'),Score_Médio=('MARS SCORE','mean'),Preço_Médio=('PREÇO LIMPO','mean'),Oportunidades=('VALE ESTUDAR?',lambda x:(x=='Sim').sum())).reset_index().sort_values('Score_Médio', ascending=False)
    st.dataframe(lojas, use_container_width=True, height=650)

elif page == '💰 Preços':
    st.title('💰 Preços')
    if f is None: st.info('Sem dados.'); st.stop()
    p=f.dropna(subset=['PREÇO LIMPO'])
    if len(p): st.plotly_chart(px.histogram(p,x='PREÇO LIMPO',nbins=25),use_container_width=True)
    st.dataframe(p.sort_values('PREÇO LIMPO',ascending=False),use_container_width=True)

elif page == '⚠️ Risco PI':
    st.title('⚠️ Risco PI')
    if f is None: st.info('Sem dados.'); st.stop()
    r=f['RISCO PI'].value_counts().reset_index(); r.columns=['Risco','Produtos']
    if len(r): st.plotly_chart(px.pie(r,names='Risco',values='Produtos'),use_container_width=True)
    st.dataframe(f[f['RISCO PI'].astype(str).str.lower()=='alto'],use_container_width=True)

elif page == '💡 Oportunidades':
    st.title('💡 Oportunidades')
    if f is None: st.info('Sem dados.'); st.stop()
    o=f[(f['MARS SCORE']>=65)&(f['RISCO PI'].astype(str).str.lower().isin(['baixo','médio','medio']))].sort_values('MARS SCORE',ascending=False)
    st.dataframe(o,use_container_width=True,height=650)

elif page == '📤 Exportar':
    st.title('📤 Exportar')
    if f is None: st.info('Sem dados.'); st.stop()
    st.download_button('⬇️ Exportar Excel',to_excel(f),f"MARS_Intelligence_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
