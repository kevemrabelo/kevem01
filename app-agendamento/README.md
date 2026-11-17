# App de Agendamento (ligado à planilha)

App para celular que lê e grava direto na **"Planilha de Agendamento."** do Google Drive,
seguindo o mesmo layout dela:

| A | B | C | D/E | F/G | H/I | J/K | L/M | N/O | P |
|---|---|---|---|---|---|---|---|---|---|
| Cliente (ex.: `Cliente - retorno`) | Telefone | Valor do serviço | Segunda (data/hora) | Terça | Quarta | Quinta | Sexta | Sábado | Serviço |

Quando você agenda pelo app, ele:
- coloca a linha **no bloco do dia certo**, na ordem do horário (antes ou depois do “INTERVALO DE ALMOÇO”);
- grava a data na coluna do dia da semana (ex.: quarta → colunas H/I) e a hora no formato da planilha (`8h30`, `14h`);
- grava o valor como `R$130,00` e o serviço (padrão `HIGIENIZAÇÃO`);
- se o dia ainda não existe, cria o bloco do dia (com a linha de almoço) e, se precisar, o cabeçalho da semana nova;
- copia a formatação (cores, negrito) das linhas que já existem.

No app também dá para **ver a semana**, **editar**, **excluir** e chamar o cliente no **WhatsApp**.

## Como instalar (uma vez só, uns 5 minutos)

1. Abra a planilha **Planilha de Agendamento.** no Google Sheets (pelo computador é mais fácil).
2. Menu **Extensões → Apps Script**.
3. No arquivo `Código.gs` que aparece, apague tudo e cole o conteúdo de [`Code.gs`](Code.gs).
4. Clique no **+** ao lado de “Arquivos” → **HTML**, dê o nome **`Index`** (sem `.html`) e cole o conteúdo de [`Index.html`](Index.html).
5. Clique em **Salvar** (ícone de disquete).
6. Clique em **Implantar → Nova implantação** → engrenagem → **App da Web**:
   - *Executar como*: **Eu**
   - *Quem pode acessar*: **Somente eu** (mais seguro) — ou “Qualquer pessoa com conta do Google” se outra pessoa também for agendar.
7. Clique em **Implantar**, autorize o acesso com sua conta Google
   (se aparecer “O Google não verificou este app”, clique em **Avançado → Acessar (não seguro)** — é o seu próprio script).
8. Copie o **URL do app da Web** (termina em `/exec`) e abra no celular.
9. No celular: no Chrome, menu ⋮ → **Adicionar à tela inicial** (no iPhone, Safari → Compartilhar → **Adicionar à Tela de Início**).
   Pronto, fica com ícone igual a um aplicativo.

> Se depois você alterar o código, use **Implantar → Gerenciar implantações → editar (lápis) → Nova versão**
> para o mesmo link continuar funcionando.

## Ajustes

No topo do `Code.gs`:
- `SHEET_NAME` — nome da aba (hoje `Página1`).
- `SPREADSHEET_ID` — só precisa preencher se o script for criado fora da planilha.
- `ALMOCO_TEXTO` / `ALMOCO_INICIO_MIN` — horário do almoço.

No `Index.html`, as listas `OBS` (atalhos de observação) e `HORAS` (atalhos de horário) podem ser trocadas.

## Observações
- Domingo não tem coluna na planilha, então o app não deixa agendar no domingo.
- Abrir o `Index.html` direto no navegador mostra um **modo demonstração** (não grava nada) — só para ver a tela.
