# App de Agendamento (ligado à planilha)

App para celular que lê e grava direto na **"Planilha de Agendamento."** do Google Drive,
seguindo o mesmo layout dela:

| A | B | C | D/E | F/G | H/I | J/K | L/M | N/O | P |
|---|---|---|---|---|---|---|---|---|---|
| Cliente (ex.: `Cliente - retorno`) | Telefone | Valor do serviço | Segunda (data/hora) | Terça | Quarta | Quinta | Sexta | Sábado | Serviço |

Planilha usada: https://docs.google.com/spreadsheets/d/1EaHMWG1U7BeiJB25ZTMCjAchtUdCIwBamLpiR-M8kws

**Vagas:** cada dia tem os horários **8h30, 10h30, 14h e 16h**, com **2 vagas em cada — uma por equipe**
(a 1ª linha do horário é a Equipe 1, a 2ª é a Equipe 2). Na planilha, a vaga livre é a linha que tem só
`Cliente` e nenhum telefone. No app ela aparece em verde como **Livre** — é só tocar para agendar.

Quando você agenda pelo app, ele:
- **preenche a vaga livre** daquele horário (nome, observação, telefone e serviço) — não cria linha nova;
- se o horário estiver lotado, coloca uma linha extra logo depois dele;
- se o dia ainda não existe na planilha, cria o dia inteiro com as 8 vagas (2 por horário) e a linha de almoço,
  e, se for semana nova, o cabeçalho da semana;
- a data e a hora ficam como estão na planilha; a coluna de valor não é mexida.

**Cancelar** um agendamento faz a linha voltar a ser vaga livre (`Cliente`, sem telefone).

No app também dá para **ver a semana**, **editar**, **cancelar** e chamar o cliente no **WhatsApp**.

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
- `SPREADSHEET_ID` — ID da planilha (já vem preenchido com a planilha atual).
- `HORARIOS_MIN` / `NUM_EQUIPES` — horários fixos e quantas equipes por horário.
- `ALMOCO_TEXTO` / `ALMOCO_INICIO_MIN` — horário do almoço.

No `Index.html`, as listas `OBS` (atalhos de observação) e `HORARIOS` (horários fixos do dia) e `EQUIPES` (nomes das equipes) podem ser trocadas.

## Observações
- Domingo não tem coluna na planilha, então o app não deixa agendar no domingo.
- Abrir o `Index.html` direto no navegador mostra um **modo demonstração** (não grava nada) — só para ver a tela.
