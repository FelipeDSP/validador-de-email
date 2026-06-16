# Validador de E-mails — versão Web (VPS / Coolify)

Versão sem interface gráfica do validador, para rodar na VPS e fazer a
**verificação SMTP** (que não funciona no seu PC porque o provedor bloqueia a
porta 25). Reaproveita toda a lógica já testada do app desktop.

- **MX + sintaxe** em todos os e-mails.
- **SMTP só nos domínios corporativos** — gmail/hotmail/yahoo/outlook etc. são
  aceitos por MX (sondá-los por SMTP não é confiável e queima o IP).
- Saída em **.xlsx no mesmo modelo da planilha original** (todas as colunas e
  cabeçalho). Acima de 1.000.000 de linhas, divide em `_parte2.xlsx`, etc.
- Remove inválidos, armadilhas (postmaster/abuse), opcionalmente automáticos
  (noreply/newsletter) e duplicados. Corrige typos de domínio.
- Protegido por **senha** (HTTP Basic).

---

## ⚠️ Leia antes: escala e tempo

A sondagem SMTP é lenta de propósito (para não bloquear o IP). Estimativa por
arquivo de ~1 milhão de linhas, com ~27% de domínios corporativos
(~270 mil sondas) ao ritmo padrão `SMTP_RATE=300/min`:

> ~270.000 ÷ 300 = **~15 horas por arquivo** (só a parte SMTP).

- Os ~73% de consumo (gmail/hotmail...) são rápidos (só MX, em paralelo).
- Se aumentar `SMTP_RATE`, vai mais rápido, **mas sobe o risco de bloqueio do
  IP**. Como os domínios corporativos são variados (1 sonda por servidor),
  300–600/min costuma ser seguro. Acima disso, monitore.
- **Recomendado:** rode 1 arquivo por vez e deixe processando (a fila do app já
  serializa). Comece com um arquivo pequeno para validar o fluxo.

---

## Deploy no Coolify

1. **Suba este projeto para um repositório Git** (GitHub/GitLab) — a pasta
   inteira `leads/` (precisa de `app-desktop/email_validator_app.py` +
   `app-web/`). O `.dockerignore` já exclui as planilhas pesadas e o `dist/`.

2. No Coolify: **+ New Resource → Application → (seu repositório)**.

3. Configure o build:
   - **Build Pack:** `Dockerfile`
   - **Dockerfile Location:** `app-web/Dockerfile`
   - **Base Directory / Build Context:** `/` (raiz do repo)
   - **Port (Ports Exposes):** `8000`

4. **Variáveis de ambiente** (Environment Variables):
   | Variável         | Valor sugerido          | Para quê |
   |------------------|-------------------------|----------|
   | `APP_PASSWORD`   | *(uma senha forte)*     | **Obrigatória.** Sem ela o app não sobe. |
   | `APP_USER`       | `admin`                 | Usuário do login. |
   | `SMTP_RATE`      | `300`                   | Sondas SMTP por minuto. |
   | `DNS_RATE`       | `50`                    | Consultas DNS por segundo. |
   | `WORKERS`        | `8`                     | Paralelismo (I/O-bound; ok em VPS fraca). |
   | `MAX_UPLOAD_MB`  | `2048`                  | Tamanho máximo de upload. |
   | `DEEP`           | `1`                     | SMTP profundo (detecta catch-all). `0` desliga. |
   | `GREYLIST_RETRIES`| `0`                    | Reenvios em greylisting (4xx). `0` = não trava o lote. |
   | `GREYLIST_DELAY` | `20`                    | Segundos entre reenvios (se `RETRIES`>0). |
   | `SMTP_FROM`      | `verify@SEU-DOMINIO`    | Remetente do envelope SMTP (ver "Identidade SMTP"). |
   | `SMTP_HELO_HOST` | `mail.SEU-DOMINIO`      | Nome de apresentação (EHLO). |
   | `SMTP_STARTTLS`  | `1`                     | Usa TLS quando o servidor oferece. |

5. **Volume persistente** (Storages): monte um volume em **`/data`**
   (uploads, resultados e a **lista de supressão** ficam aqui; sem isso, perde
   tudo — inclusive a supressão — a cada redeploy).

6. **Deploy.** Acesse pela URL/domínio que o Coolify gerar. O navegador vai
   pedir usuário e senha (os de `APP_USER` / `APP_PASSWORD`).

> A porta 25 de saída já foi testada e está **aberta** nesta VPS — requisito
> essencial para o SMTP funcionar.

---

## Identidade SMTP (deixa a verificação corporativa confiável)

Por padrão o app se apresenta como `localhost` / `verify@example.com` — muitos
servidores rejeitam ou mentem na resposta a isso, gerando veredito errado. Para
a checagem corporativa ficar confiável, configure no DNS um domínio/subdomínio
seu (ex.: `mail.seudominio.com` apontando para o IP da VPS) e:

1. **A record:** `mail.seudominio.com` → IP da VPS.
2. **rDNS / PTR:** peça ao provedor da VPS para o IP resolver de volta para
   `mail.seudominio.com` (painel da VPS ou suporte). **É o item mais importante.**
3. **SPF (TXT):** em `seudominio.com` inclua o IP da VPS
   (`v=spf1 ip4:IP_DA_VPS -all`).
4. Defina as variáveis `SMTP_HELO_HOST=mail.seudominio.com` e
   `SMTP_FROM=verify@seudominio.com`.

Sem isso o SMTP ainda funciona, mas com mais "desconhecidos"/falsos negativos.

---

## Lista de supressão (anti-bounce que aprende)

O cartão **"2) Lista de supressão"** na página recebe os e-mails que já
**bouncaram** ou estão em **Do-Not-Contact** (exporte do Mautic). Eles passam a
ser **removidos automaticamente** de toda planilha nova. Fluxo recomendado:

1. Rode a campanha no Mautic.
2. Exporte os bounces / DNC (CSV ou Excel — qualquer coluna com e-mail serve).
3. Suba no cartão de supressão. A cada ciclo a lista aprende e as próximas
   limpezas já saem sem esses endereços (é o único jeito de atacar o resíduo de
   gmail/hotmail, que ninguém consegue verificar por SMTP).

A supressão fica em `/data/suppression.txt` (no volume persistente).

---

## Uso

1. Abra a URL, faça login.
2. Selecione a planilha, marque as opções e clique **Enviar e processar**.
3. A página mostra o progresso e atualiza sozinha. Pode fechar o navegador — o
   processamento continua na VPS.
4. Ao concluir, aparecem os links **baixar parte 1, 2, …** (.xlsx limpo).

---

## Rodar localmente para testar (opcional)

```bash
cd app-web
pip install -r requirements.txt
# Windows PowerShell:
$env:APP_PASSWORD="teste"; $env:DATA_DIR="./_data"; python webapp.py
# Linux/Mac:
APP_PASSWORD=teste DATA_DIR=./_data python webapp.py
```
Acesse http://localhost:8000 (login: admin / teste).
