# Famossul — Gerador de Romaneios e Etiquetas

MVP web para processar **Pedido (PDF)** + **Usinagem (XLS/XLSX, opcional)**, montar pallets para conferência e gerar **Romaneio XLSX** e **Etiquetas PDF**.

## Stack
- Next.js 15 + TypeScript
- Vercel
- Neon PostgreSQL + Prisma (estrutura pronta para histórico)
- ExcelJS / SheetJS
- pdf-parse
- pdf-lib
- Login administrativo único por variável de ambiente

## 1. Instalação
```bash
npm install
cp .env.example .env.local
npm run admin:hash
```
Cole o hash em `ADMIN_PASSWORD_HASH` e configure `ADMIN_USER`, `AUTH_SECRET` e `DATABASE_URL`.

```bash
npm run db:push
npm run dev
```
Abra `http://localhost:3000`.

## 2. Deploy no Vercel
1. Suba este projeto no GitHub.
2. Importe o repositório no Vercel.
3. Crie um banco Neon e copie a `DATABASE_URL`.
4. Cadastre no Vercel: `DATABASE_URL`, `AUTH_SECRET`, `ADMIN_USER`, `ADMIN_PASSWORD_HASH`.
5. Faça o deploy.

## 3. Regra inicial implementada
O parser foi desenhado a partir do fluxo fornecido da Famossul e do caso de referência **Pedido 26716**:
- identifica Pedido, Cliente e Destino;
- procura Folha de Porta, Marco, Alizar, Fechadura e Dobradiça;
- lê a planilha de usinagem e procura a divisão `X Direitas / Y Esquerdas`;
- cria 4 grupos iniciais: portas, marcos, alizares e ferragens;
- **sempre exibe conferência editável antes da emissão**.

### Importante
O sistema não inventa divisão física de pallets fora das regras cadastradas. Novos formatos de produto devem ser validados com romaneios reais e adicionados ao motor de regras em `lib/parser.ts`.

## 4. Próximas evoluções já previstas
- salvar processamentos e pallets no Neon;
- tela de Histórico e reimpressão;
- múltiplas regras de embalagem por família de produto;
- upload do modelo oficial de romaneio/etiqueta por Configurações;
- impressão direta por formato de etiqueta;
- auditoria de alterações antes da impressão.
