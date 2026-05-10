# 🧭 Bússola do Representante

## Estrutura

```
bussola/
├── hub.py                  ← app principal (sobe este no Streamlit Cloud)
├── requirements.txt
└── pages/
    ├── home.py             ← tela inicial
    ├── radar.py            ← Radar Comercial (seu app.py original)
    ├── cotabot.py          ← CotaBot (seu cotacao.py original)
    └── visitas.py          ← Bússola de Visitas (aguardando código)
```

## Como subir no Streamlit Cloud

1. Crie um repositório no GitHub com esta estrutura
2. Acesse https://share.streamlit.io
3. Conecte o repositório
4. **Main file path:** `hub.py`
5. Deploy

## Para adicionar a Bússola de Visitas

1. Substitua o conteúdo de `pages/visitas.py` pelo código do seu app
2. Remova a linha `st.set_page_config(...)` do arquivo (o hub já cuida disso)
3. Faça commit — o deploy atualiza automaticamente

## Regra para qualquer novo app

- Coloque o arquivo em `pages/`
- Remova o `st.set_page_config` do arquivo
- Adicione a página no `hub.py` com `st.Page(...)`
- Pronto — zero alteração no resto do código
