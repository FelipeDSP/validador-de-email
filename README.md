# Validador de E-mails — projeto

Ferramentas para limpar listas de leads (planilhas grandes de CNPJ/SP):
validação de sintaxe, correção de typos de domínio, registro MX, verificação
SMTP (só corporativa) e classificação B2B, preservando o modelo da planilha.

## Estrutura das pastas

```
leads/
├── app-desktop/        Aplicativo de mesa (Windows, com interface)
│   ├── email_validator_app.py   código-fonte (também é o núcleo reutilizado)
│   ├── ValidadorEmails.spec     receita do PyInstaller
│   ├── dist/ValidadorEmails.exe o programa pronto para usar  ◀── clique aqui
│   └── build/                   arquivos temporários do build
│
├── app-web/            Aplicação web (roda na VPS via Coolify, faz o SMTP)
│   ├── webapp.py                servidor Flask (upload/status/download)
│   ├── validator_core.py        processamento headless (reusa o núcleo)
│   ├── Dockerfile               imagem para o Coolify
│   ├── requirements.txt
│   └── README.md        ◀── passo a passo do deploy no Coolify
│
├── planilhas/          Dados (NÃO entram no Docker nem no Git)
│   ├── originais/               planilhas de entrada (Empresas do SP, testes)
│   └── limpas/                  resultados já validados (_LIMPA)
│
├── .dockerignore       Impede enviar planilhas/exe para o build da imagem
└── README.md           Este arquivo
```

## Qual usar?

- **No seu PC (rápido, sem SMTP):** abra `app-desktop/dist/ValidadorEmails.exe`.
  Faz sintaxe + MX + B2B + dedup e exporta `.xlsx` no mesmo modelo. Não faz SMTP
  porque o seu provedor bloqueia a porta 25.

- **Na VPS (com SMTP corporativo):** a `app-web/` roda no Coolify, onde a porta
  25 está liberada. Veja o passo a passo em `app-web/README.md`.

> As duas compartilham o MESMO núcleo de validação (`email_validator_app.py`):
> a web o importa como biblioteca, então a lógica nunca diverge.

## Rebuildar o aplicativo desktop (quando mexer no código)

```powershell
# feche o ValidadorEmails.exe antes (ele trava o arquivo)
cd app-desktop
python -m PyInstaller --noconfirm ValidadorEmails.spec
# resultado: app-desktop/dist/ValidadorEmails.exe
```
