# Bussola do Representante V1.3

Portal local integrado para representantes, com frontend em Next.js e backend em Python/FastAPI.

## O que existe na V1.3

- Início como tela inicial, com esboço do dia e atalhos para os módulos importantes.
- Cliente360 como motor central de dados consolidados dos clientes.
- Raio-X Cliente integrado à Visão Geral do Cliente como tela operacional por cliente.
- Visão Geral do Cliente como ranking pratico de clientes prioritarios e ações do dia.
- Importacao de bases nas abas operacionais, sem upload de planilhas no Início.
- Campo de busca por cliente no Raio-X.
- Cards de dados cadastrais, campanha, cobranca, oportunidades e acao sugerida.
- Botao para gerar mensagem de WhatsApp no Raio-X Cliente.
- Botao WhatsApp por cliente na Visão Geral do Cliente.
- TOP 3 prioridades do dia a partir da inteligencia comercial.
- Modulos Campanha, Cobrança, Cotação e Pedidos mantidos.
- O Início orienta o representante; as importações acontecem nas abas especificas.
- Visão Geral do Cliente usa CMK e Curva como fontes globais.
- Campanha usa a planilha carregada na propria aba para montar a visao de campanhas.
- Cotação e Pedidos seguem com seus arquivos próprios, sem depender da importação central.
- Backend FastAPI respondendo os endpoints do sistema.
- Estrutura pronta para evoluir, sem login, banco de dados ou SaaS nesta versao.

## Estrutura principal

```text
frontend/  Site em Next.js com multiplas paginas.
backend/   API Python/FastAPI e logicas dos modulos.
```

As regras de negocio ficam em:

```text
backend/app/modules/
```

A camada central Cliente360 fica em:

```text
backend/app/api/cliente360.py
backend/app/schemas/cliente360.py
backend/app/services/cliente360_service.py
```

A camada Visão Geral fica em:

```text
backend/app/api/plano_ataque.py
backend/app/schemas/plano_ataque.py
backend/app/services/plano_ataque_service.py
frontend/app/plano-ataque/page.jsx
frontend/components/module-plano-ataque.jsx
```

A camada Importador fica em:

```text
backend/app/api/importador.py
backend/app/schemas/importador.py
backend/app/services/importador_service.py
backend/app/modules/parsers/
frontend/app/importador/page.jsx
frontend/components/module-importador.jsx
```

## Como rodar o backend

Na raiz do projeto, rode:

```powershell
pip install -r backend/requirements.txt
cd backend
python -m uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Teste a API de saude:

```text
http://127.0.0.1:8000/api/health
```

Resposta esperada:

```json
{"status":"ok"}
```

## Como rodar o frontend

Em outro terminal, rode:

```powershell
cd frontend
copy .env.example .env.local
npm install
npm run dev
```

Abra o site em:

```text
http://127.0.0.1:3000
```

## Como testar /cliente360

Com frontend e backend rodando, acesse:

```text
http://127.0.0.1:3000/cliente360
```

Essa pagina mostra a visao tecnica/consolidada do Cliente360: contrato de dados, endpoint, estrutura entregue e clientes consolidados.

Para testar a API do Cliente360 diretamente, acesse:

```text
http://127.0.0.1:8000/api/cliente360
```

A resposta deve retornar `source`, `total_clientes` e `clientes`.

## Como testar /raiox-cliente

Com frontend e backend rodando, acesse:

```text
http://127.0.0.1:3000/raiox-cliente
```

Use a busca para selecionar um cliente. A tela deve exibir dados cadastrais, campanha, cobranca, oportunidades, acao sugerida e botao para gerar WhatsApp.

## O que e a Visão Geral do Cliente

A Visão Geral do Cliente transforma o Cliente360 em prioridade pratica do dia para o representante.

Ele responde:

- Quem visitar hoje.
- Por que visitar.
- O que vender.
- Qual acao executar.

Os criterios atuais sao:

- Cliente com cobranca pendente e risco de bloqueio.
- Cliente com oportunidade de campanha proxima de fechar POS.
- Cliente com VB faltante em familia importante.
- Cliente que comprava e parou.
- Cliente com muitos dias sem comprar.
- Cliente estrategico por ticket medio.

## Como testar /plano-ataque

Com frontend e backend rodando, acesse:

```text
http://127.0.0.1:3000/plano-ataque
```

A tela aparece no menu como `Visão Geral do Cliente` e deve exibir cards de resumo, filtros por cidade, bandeira, campanha, cobranca e tipo de acao, ranking dos clientes prioritarios, botao WhatsApp e atalho para o Raio-X Cliente.

## Como testar /api/plano-ataque

Com o backend rodando, acesse:

```text
http://127.0.0.1:8000/api/plano-ataque
```

A resposta deve retornar `resumo`, `filtros` e `prioridades`.

## O que e a Central de Importacao

A Central de Importacao e a porta de entrada das bases reais usadas no dia a dia do representante.

Ela aceita:

- CMK.
- Curva.
- Campanha.
- Cobrança.

Formatos aceitos:

- XLSX.
- XLSB.
- CSV.

A tela mostra para cada base:

- Status importado ou pendente.
- Nome do arquivo.
- Data da importacao.
- Quantidade de registros.
- Quantidade de clientes.
- Quantidade de produtos.
- Quantidade de familias.
- Mensagem amigavel em caso de arquivo vazio, planilha invalida, colunas ausentes ou formato nao suportado.

## Como testar /importador

Com frontend e backend rodando, acesse:

```text
http://127.0.0.1:3000/importador
```

Selecione uma base, envie um arquivo XLSX, XLSB ou CSV e confira o status no card da respectiva base.

## Como testar /api/importador/status

Com o backend rodando, acesse:

```text
http://127.0.0.1:8000/api/importador/status
```

A resposta deve retornar `accepted_formats`, `total_bases`, `imported_bases`, `bases` e `flow`.

## Relacao entre Importador, Cliente360, Raio-X Cliente e Visão Geral do Cliente

`Cliente360` e o motor/base geral consolidada do sistema:

```text
GET http://127.0.0.1:8000/api/cliente360
```

`Raio-X Cliente` e a tela operacional por cliente. Ela consome o Cliente360 para mostrar busca, cadastro, campanha, cobranca, oportunidades, acao sugerida e WhatsApp.

`Visão Geral do Cliente` e o ranking de prioridades do dia. Ela transforma os dados do Cliente360 em uma lista pratica de clientes, motivo da visita, produto/familia recomendada e acao comercial.

`Importador` e a entrada das bases reais. Nesta V1.3 ele guarda os arquivos em memoria temporaria e prepara o caminho para alimentar o Cliente360 futuramente.

Fluxo atual de uso das bases:

- CMK + Curva alimentam a Visão Geral do Cliente.
- A aba Campanha carrega a planilha de campanha pela propria area de carregar e filtrar.
- Cotação e Pedidos continuam usando os arquivos enviados nas próprias telas.
- Reimportar uma base na aba operacional atualiza o estado global para as telas que consomem essa base.

## Como o Importador alimentara o Cliente360 futuramente

Na V1.3, o Importador valida e registra o status das bases em memoria temporaria. Na evolucao seguinte, essas bases serao consolidadas pelo Cliente360 para substituir gradualmente os dados mockados por informacoes reais de cadastro, campanha, cobranca, metas, historico de vendas e curva de produtos.

O Cliente360 continuara sendo a unica fonte central para as telas operacionais. O Raio-X Cliente e a Visão Geral do Cliente nao devem ler planilhas diretamente; eles devem consumir os dados ja consolidados pelo Cliente360.

O endpoint da Visão Geral do Cliente e:

```text
GET http://127.0.0.1:8000/api/plano-ataque
```

Esse endpoint e gerado a partir do Cliente360. Portanto, o fluxo da V1.3 fica:

```text
Importador -> Cliente360
Cliente360 -> Raio-X Cliente
Cliente360 -> Visão Geral do Cliente
```

## Paginas principais

- Início: `http://127.0.0.1:3000`
- Visão Geral do Cliente: `http://127.0.0.1:3000/plano-ataque`
- Campanha: `http://127.0.0.1:3000/monitor`
- Cobrança: `http://127.0.0.1:3000/cobranca`
- Cotação: `http://127.0.0.1:3000/cotabot`
- Pedidos: `http://127.0.0.1:3000/pedidos`

Rotas internas acessadas por atalho:

- Importador tecnico: `http://127.0.0.1:3000/importador`
- Raio-X Cliente dentro da Visão Geral do Cliente: `http://127.0.0.1:3000/raiox-cliente`
- Cliente360 como motor tecnico: `http://127.0.0.1:3000/cliente360`

## Checklist de validacao V1.3

Com backend e frontend rodando, essas rotas devem responder status `200`:

- Backend health: `http://127.0.0.1:8000/api/health`
- API Cliente360: `http://127.0.0.1:8000/api/cliente360`
- API Visão Geral do Cliente: `http://127.0.0.1:8000/api/plano-ataque`
- API Importador Status: `http://127.0.0.1:8000/api/importador/status`
- Frontend Início: `http://127.0.0.1:3000`
- Frontend Cliente360 interno: `http://127.0.0.1:3000/cliente360`
- Frontend Raio-X Cliente interno: `http://127.0.0.1:3000/raiox-cliente`
- Frontend Visão Geral do Cliente: `http://127.0.0.1:3000/plano-ataque`
- Frontend Importador interno: `http://127.0.0.1:3000/importador`
- Frontend Campanha: `http://127.0.0.1:3000/monitor`
- Frontend Cobrança: `http://127.0.0.1:3000/cobranca`
- Frontend Cotação: `http://127.0.0.1:3000/cotabot`
- Frontend Pedidos: `http://127.0.0.1:3000/pedidos`

## Observacao importante

Sempre reinicie o backend depois de alterar rotas, services ou schemas Python. Se o backend antigo continuar aberto em outro terminal, ele pode nao carregar mudancas recentes como `/api/cliente360`.

## Arquitetura futura

Esta versao esta pronta para evoluir para SaaS, mas ainda nao implementa:

- Login.
- Banco de dados.
- Multiplos usuarios.
- Permissoes.
- Planos pagos.
- Upload de planilhas por usuario.

Pastas reservadas para evolucao:

- `backend/app/api/`
- `backend/app/core/`
- `backend/app/db/`
- `backend/app/models/`
- `backend/app/schemas/`
- `backend/app/services/`
