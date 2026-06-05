# Deploy controlado - Bussola do Representante

Este guia prepara a publicacao online do app sem alterar a logica atual.

Versao protegida antes do deploy:

```text
C:\Users\magno\Downloads\VERSAO_FINAL_LOCAL_FUNCIONANDO.zip
```

## Arquitetura recomendada

Use:

- Frontend: Vercel
- Backend: Render
- Codigo: GitHub

Motivo:

- O frontend ja esta em Next.js dentro de `frontend/`.
- O backend ja esta em FastAPI dentro de `backend/`.
- O app processa uploads e planilhas em Python, entao o backend precisa rodar como servico web separado.
- A Vercel fica responsavel pela interface.
- A Render fica responsavel pela API e processamento dos arquivos.

## Estrutura esperada no GitHub

```text
bussola_do_representante/
  backend/
    app/
    requirements.txt
  frontend/
    app/
    components/
    lib/
    package.json
    .env.example
  README.md
  README_DEPLOY.md
  render.yaml
```

Nao subir:

- `frontend/node_modules/`
- `frontend/.next/`
- `.portal_runtime/`
- `.venv/`
- `venv/`
- `__pycache__/`
- arquivos `.zip`

Esses itens ja estao protegidos no `.gitignore`.

## 1. Subir para o GitHub

Na pasta do projeto:

```powershell
cd C:\Users\magno\Downloads\bussola_do_representante
git init
git add .
git commit -m "VERSAO_FINAL_LOCAL_FUNCIONANDO"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/bussola-do-representante.git
git push -u origin main
```

Se o `git` nao estiver instalado, use uma destas opcoes:

- Instalar Git for Windows.
- Usar GitHub Desktop.
- Usar o controle de versao do VS Code depois de instalar o Git.

## 2. Publicar backend na Render

Opcao A - usando `render.yaml`:

1. Acesse Render.
2. Crie um Blueprint.
3. Conecte o repositorio do GitHub.
4. A Render deve detectar o arquivo `render.yaml`.
5. Confirme o servico `bussola-backend`.
6. Aguarde o build.

Opcao B - configuracao manual:

1. Crie um novo Web Service.
2. Conecte o repositorio do GitHub.
3. Root Directory:

```text
backend
```

4. Build Command:

```text
pip install -r requirements.txt
```

5. Start Command:

```text
python -m uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

6. Environment:

```text
PYTHON_VERSION=3.11.9
```

7. Depois do deploy, teste:

```text
https://SEU-BACKEND.onrender.com/api/health
```

Resposta esperada:

```json
{"status":"ok"}
```

## 3. Publicar frontend na Vercel

1. Acesse Vercel.
2. Importe o mesmo repositorio do GitHub.
3. Configure:

```text
Framework Preset: Next.js
Root Directory: frontend
Build Command: npm run build
Install Command: npm install
```

4. Crie a variavel de ambiente:

```text
NEXT_PUBLIC_API_BASE_URL=https://SEU-BACKEND.onrender.com
```

5. Faça o deploy.
6. Abra o link gerado pela Vercel.

## 4. Variaveis de ambiente

Local:

```text
frontend/.env.local
NEXT_PUBLIC_API_BASE_URL=http://127.0.0.1:8000
```

Exemplo versionado:

```text
frontend/.env.example
NEXT_PUBLIC_API_BASE_URL=http://127.0.0.1:8000
```

Vercel:

```text
NEXT_PUBLIC_API_BASE_URL=https://SEU-BACKEND.onrender.com
```

Render:

```text
PYTHON_VERSION=3.11.9
```

## 5. CORS

O backend atual usa CORS aberto:

```python
allow_origins=["*"]
```

Isso permite:

- `http://127.0.0.1:3000`
- `http://localhost:3000`
- dominio futuro da Vercel

Para um teste controlado, isso evita erro de `Failed to fetch` por CORS.

Depois que o dominio final da Vercel estiver definido, uma melhoria futura sera restringir o CORS apenas para:

```text
http://127.0.0.1:3000
http://localhost:3000
https://SEU-FRONTEND.vercel.app
```

Nao fazer essa restricao antes de validar o deploy.

## 6. Testes pos-deploy

Teste backend:

```text
GET https://SEU-BACKEND.onrender.com/api/health
GET https://SEU-BACKEND.onrender.com/api/importador/status
GET https://SEU-BACKEND.onrender.com/api/cliente360
GET https://SEU-BACKEND.onrender.com/api/plano-ataque
```

Teste frontend:

```text
https://SEU-FRONTEND.vercel.app/
https://SEU-FRONTEND.vercel.app/monitor
https://SEU-FRONTEND.vercel.app/cobranca
https://SEU-FRONTEND.vercel.app/cotabot
https://SEU-FRONTEND.vercel.app/pedidos
```

Teste uploads:

- Campanha: subir planilha e confirmar KPIs, rankings, Raio-X e WhatsApp.
- Cobranca: subir carteira e confirmar filtros, checklist, mensagem e detalhe.
- Cotacao: subir base e cotacao, inspecionar e processar.
- Pedidos: subir pedidos e base, processar, visualizar previas e baixar arquivos.

## 7. O que pode quebrar e como evitar

### Backend dormindo

Em plano gratuito, a Render pode pausar o backend. A primeira chamada pode demorar.

Como evitar:

- Aguardar alguns segundos na primeira abertura.
- Testar `/api/health` antes de testar uploads.
- Em uso real, avaliar plano pago.

### `Failed to fetch`

Causa comum:

- `NEXT_PUBLIC_API_BASE_URL` errado na Vercel.
- Backend ainda nao terminou deploy.
- Backend dormindo.

Como evitar:

- Confirmar a URL do backend na Render.
- Atualizar a variavel na Vercel.
- Fazer redeploy do frontend depois de alterar variavel.

### Upload nao persistente

Hoje os arquivos importados ficam em runtime temporario e parte do estado fica em memoria.

Consequencia:

- Se o backend reiniciar, a base global importada pode ser perdida.

Como evitar no teste controlado:

- Reimportar a planilha depois de reinicio.
- Testar com poucas pessoas ao mesmo tempo.

Melhoria futura:

- Adicionar banco de dados e armazenamento persistente por usuario.

### Arquivo muito grande

Planilhas grandes podem consumir muita memoria.

Como evitar:

- Comecar os testes com arquivos reais, mas moderados.
- Validar Campanha, Cobranca, Cotacao e Pedidos separadamente.

### Dependencias Python

O backend usa `pandas`, `numpy`, `openpyxl`, `pyxlsb`, `xlrd` e `python-multipart`.

Como evitar erro:

- Manter `backend/requirements.txt`.
- Usar Python 3.11.9 no Render.

## 8. Checklist rapido

Antes de publicar:

- [ ] Backup `VERSAO_FINAL_LOCAL_FUNCIONANDO.zip` salvo.
- [ ] Projeto subido no GitHub.
- [ ] Render backend publicado.
- [ ] `/api/health` retorna `{"status":"ok"}`.
- [ ] Vercel frontend publicado.
- [ ] `NEXT_PUBLIC_API_BASE_URL` na Vercel aponta para a Render.
- [ ] Pagina inicial abre.
- [ ] Campanha abre.
- [ ] Cobranca abre.
- [ ] Cotacao abre.
- [ ] Pedidos abre.
- [ ] Upload de Campanha testado.
- [ ] Upload de Cobranca testado.
- [ ] Processamento de Cotacao testado.
- [ ] Processamento de Pedidos testado.

