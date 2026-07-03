# MARS Intelligence Web

Aplicação Web em Streamlit para analisar a base de dados da MARS 3D.

## Como usar online
1. Criar conta em https://streamlit.io
2. Criar conta em https://github.com
3. Criar um repositório chamado `mars-intelligence-web`
4. Carregar estes ficheiros para o repositório
5. No Streamlit Community Cloud, escolher `app.py`
6. Abrir a aplicação no browser

## Como usar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## O que a versão v0.1 faz
- Upload de Excel ou CSV
- Leitura da folha PRODUTOS
- Limpeza básica de dados
- Remoção de duplicados
- Cálculo de MARS SCORE
- Dashboard com KPIs
- Filtros por categoria, risco PI e vale estudar
- Exportação de Excel limpo
