# -*- coding: utf-8 -*-
"""
Nucleo de processamento para o servidor web (sem interface grafica).

Reaproveita TODA a logica ja testada do app desktop (email_validator_app.py):
sintaxe, correcao de typos, classificacao B2B em 3 niveis, MX com cache,
e SMTP "so corporativo" (o modo 'completo' ja pula gmail/hotmail/yahoo... que
estao em ACCEPT_ALL_DOMAINS e so sonda dominios de empresa).

Acrescenta (reducao maxima de bounce):
  - SMTP PROFUNDO: deteta catch-all (aceita-tudo) e trata greylisting.
  - LISTA DE SUPRESSAO: remove enderecos que ja bouncaram antes (aprende com
    os bounces exportados do Mautic).
  - Null MX / listas ampliadas de descartaveis.
  - Celula com VARIOS e-mails (a@x; b@y) -> uma linha por e-mail valido.
  - Dedup inteligente (gmail ignora pontos e +tag).
  - process_file(): streaming gravando .xlsx (mesmo modelo da planilha
    original), com divisao automatica acima de ~1 milhao de linhas.
"""

import os
import re
import sys
import collections
from concurrent.futures import ThreadPoolExecutor

# Importa o nucleo do app desktop como biblioteca. O __main__ (a janela) NAO
# roda no import; so quando executado diretamente. O import do tkinter e
# opcional no modulo, entao funciona headless (Linux/container) sem tkinter.
# O email_validator_app.py vive em ../app-desktop (mesma estrutura no
# container). Procuramos em alguns lugares para ser robusto local e no Docker.
_HERE = os.path.dirname(os.path.abspath(__file__))          # .../app-web
_ROOT = os.path.dirname(_HERE)                              # .../ (raiz ou /app)
for _cand in (os.path.join(_ROOT, "app-desktop"), _ROOT, _HERE):
    if os.path.isfile(os.path.join(_cand, "email_validator_app.py")):
        sys.path.insert(0, _cand)
        break
import email_validator_app as eva  # noqa: E402


# Margem de seguranca abaixo do teto do XLSX (1.048.576 linhas por aba).
# Inclui o cabecalho, entao usamos 1.000.000 de linhas de DADOS por arquivo.
MAX_DATA_ROWS_PER_FILE = 1_000_000

_GMAIL_DOMAINS = {"gmail.com", "googlemail.com"}
_EMAIL_SPLIT_RE = re.compile(r"[;,/|\s]+")


# --------------------------------------------------------------------------- #
#  Listas ampliadas (descartaveis) carregadas de data/disposable_domains.txt
# --------------------------------------------------------------------------- #
def _load_extra_disposable():
    path = os.path.join(_HERE, "data", "disposable_domains.txt")
    if not os.path.isfile(path):
        return 0
    try:
        with open(path, encoding="utf-8", errors="ignore") as f:
            extra = {ln.strip().lower() for ln in f
                     if ln.strip() and not ln.lstrip().startswith("#")}
        eva.DISPOSABLE_DOMAINS |= extra
        return len(extra)
    except Exception:
        return 0


EXTRA_DISPOSABLE_LOADED = _load_extra_disposable()


# --------------------------------------------------------------------------- #
#  Normalizacao para dedup e supressao
# --------------------------------------------------------------------------- #
def normalize_email(email):
    """
    Chave canonica de um e-mail para dedup/supressao:
      - minusculas;
      - remove sufixo +tag (sub-addressing);
      - no Gmail, remove os pontos do nome (gmail os ignora).
    """
    e = (email or "").strip().lower()
    if "@" not in e:
        return e
    local, _, domain = e.partition("@")
    local = local.split("+", 1)[0]
    if domain in _GMAIL_DOMAINS:
        local = local.replace(".", "")
    return f"{local}@{domain}"


def is_corporate_domain(domain):
    """True se o dominio NAO e um provedor de consumo (gmail/hotmail/...)."""
    return domain.lower().strip() not in eva.ACCEPT_ALL_DOMAINS


def split_emails_in_cell(cell):
    """Extrai os e-mails de uma celula que pode ter varios (a@x; b@y)."""
    if not cell or "@" not in cell:
        return []
    out, seen = [], set()
    for part in _EMAIL_SPLIT_RE.split(cell):
        t = part.strip().strip("<>()[]\"'")
        if "@" in t and t.lower() not in seen:
            seen.add(t.lower())
            out.append(t)
    return out


# --------------------------------------------------------------------------- #
#  Lista de supressao (Do-Not-Contact que aprende com bounces)
# --------------------------------------------------------------------------- #
def load_suppression(path):
    """Le o arquivo de supressao (1 e-mail por linha) -> set normalizado."""
    s = set()
    if path and os.path.isfile(path):
        with open(path, encoding="utf-8", errors="ignore") as f:
            for ln in f:
                e = ln.strip()
                if e and "@" in e:
                    s.add(normalize_email(e))
    return s


def extract_emails_from_file(path):
    """Le um CSV/Excel (export de bounces do Mautic) e devolve os e-mails."""
    found = set()
    try:
        _h, _ci, rows, _t = eva.open_table_stream(path)
        for r in rows:
            for cell in r:
                cs = str(cell)
                if "@" in cs:
                    for tok in split_emails_in_cell(cs):
                        if eva.check_syntax(tok):
                            found.add(tok.strip())
    except Exception:
        pass
    return found


def add_to_suppression(path, emails):
    """Acrescenta e-mails ao arquivo de supressao (sem duplicar). Retorna
    (novos_adicionados, total_no_arquivo)."""
    existing = load_suppression(path)
    novos = 0
    with open(path, "a", encoding="utf-8") as f:
        for e in emails:
            n = normalize_email(e)
            if n and n not in existing:
                existing.add(n)
                f.write(e.strip() + "\n")
                novos += 1
    return novos, len(existing)


# --------------------------------------------------------------------------- #
#  Escrita .xlsx em streaming, com divisao automatica
# --------------------------------------------------------------------------- #
class _XlsxSplitWriter:
    """
    Grava linhas em .xlsx em streaming (write_only) e abre automaticamente um
    novo arquivo (..._parte2.xlsx, _parte3...) quando passa de MAX_DATA_ROWS,
    para nunca estourar o limite do formato e nunca perder o processamento.
    """

    def __init__(self, out_path, header):
        if eva.openpyxl is None:
            raise RuntimeError("openpyxl nao esta instalado para gravar Excel.")
        self.base, self.ext = os.path.splitext(out_path)
        if self.ext.lower() not in (".xlsx", ".xlsm"):
            self.ext = ".xlsx"
        self.header = header
        self.part = 0
        self.files = []
        self.row_in_part = 0
        self.wb = None
        self.ws = None
        self.cur_path = None
        self._open_new()

    def _open_new(self):
        if self.wb is not None:
            self.wb.save(self.cur_path)
        self.part += 1
        if self.part == 1:
            self.cur_path = self.base + self.ext
        else:
            self.cur_path = f"{self.base}_parte{self.part}{self.ext}"
        self.wb = eva.openpyxl.Workbook(write_only=True)
        self.ws = self.wb.create_sheet("Limpa")
        self.row_in_part = 0
        self.files.append(self.cur_path)
        if self.header is not None:
            self.ws.append(eva._sanitize_row(self.header))

    def append(self, row):
        if self.row_in_part >= MAX_DATA_ROWS_PER_FILE:
            self._open_new()
        self.ws.append(eva._sanitize_row(row))
        self.row_in_part += 1

    def close(self):
        if self.wb is not None:
            self.wb.save(self.cur_path)
            self.wb = None


# --------------------------------------------------------------------------- #
#  Processamento principal
# --------------------------------------------------------------------------- #
def process_file(in_path, out_path, column=None, mode="completo",
                 dedup=True, excluir_arriscados=False, deep=True,
                 suppression=None, dns_rate=50, smtp_rate=300, workers=8,
                 greylist_retries=0, greylist_delay=20,
                 progress_cb=None, stop_event=None, every=1000):
    """
    Processa uma planilha (CSV/Excel) em streaming e grava .xlsx limpo com o
    MESMO modelo da original (todas as colunas + cabecalho), e-mails com typo
    corrigidos, so linhas enviaveis.

    mode='completo' -> MX em tudo + SMTP so nos dominios corporativos.
    mode='rapido'   -> so MX + sintaxe (sem SMTP).
    deep=True       -> SMTP profundo: deteta catch-all + greylisting.
    suppression     -> set de e-mails normalizados a remover (bounces antigos).

    greylist_retries=0 (padrao) evita travar o lote esperando 4xx; quem
    greylistou cai como 'arriscado' (nao confirmado) e nao e enviado.

    Devolve um dict de contadores (inclui 'arquivos': lista de .xlsx gerados).
    """
    eva.set_dns_rate_per_sec(max(1, dns_rate))
    eva.set_smtp_rate_per_min(max(1, smtp_rate))
    eva.SMTP_GREYLIST_RETRIES = max(0, greylist_retries)
    eva.SMTP_GREYLIST_DELAY = max(1, greylist_delay)

    sup = suppression or set()
    header, col_idx, rows, approx_total = eva.open_table_stream(in_path, column)

    c = {"processadas": 0, "com_email": 0, "sem_email": 0,
         "corporativos": 0, "catch_all": 0, "seguro": 0, "arriscado": 0,
         "invalido": 0, "corrigido": 0, "duplicados": 0, "suprimidos": 0,
         "mantidas": 0, "total_estimado": approx_total or 0, "arquivos": []}
    seen = set()
    writer = _XlsxSplitWriter(out_path, header)

    def work(row):
        cell = str(row[col_idx]).strip() if col_idx < len(row) else ""
        tokens = split_emails_in_cell(cell)
        results = []
        for tok in tokens:
            res = eva.validate_email(tok, mode, deep=deep)
            final = res.get("email_final", tok)
            domain = final.rsplit("@", 1)[-1].lower() if "@" in final else ""
            results.append((res, domain))
        return (row, results)

    def handle(row, results):
        c["processadas"] += 1
        if not results:
            c["sem_email"] += 1
            return
        for res, domain in results:
            c["com_email"] += 1
            if domain and is_corporate_domain(domain):
                c["corporativos"] += 1
            motivo = res.get("motivo", "")
            if "catch-all" in motivo:
                c["catch_all"] += 1
            risco = res.get("risco", "")
            if risco == "seguro":
                c["seguro"] += 1
            elif risco == "arriscado":
                c["arriscado"] += 1
            else:
                c["invalido"] += 1
            if res.get("corrigido"):
                c["corrigido"] += 1

            final = res.get("email_final", "")
            manter = (res.get("status") == "valido")
            if manter and excluir_arriscados and risco == "arriscado":
                manter = False
            key = normalize_email(final)
            if manter and key in sup:           # ja bouncou antes
                manter = False
                c["suprimidos"] += 1
            if manter and dedup:
                if key in seen:
                    manter = False
                    c["duplicados"] += 1
                else:
                    seen.add(key)
            if manter:
                new_r = list(row)
                if col_idx < len(new_r):
                    new_r[col_idx] = final
                writer.append(new_r)
                c["mantidas"] += 1

    try:
        row_iter = iter(rows)
        inflight = collections.deque()
        with ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
            def submit_next():
                try:
                    row = next(row_iter)
                except StopIteration:
                    return False
                inflight.append(ex.submit(work, row))
                return True

            # janela deslizante: mantem ~4x workers em voo, sem carregar tudo
            for _ in range(max(1, workers) * 4):
                if not submit_next():
                    break

            while inflight:
                if stop_event is not None and stop_event.is_set():
                    break
                fut = inflight.popleft()
                row, results = fut.result()
                handle(row, results)
                if progress_cb is not None and c["processadas"] % every == 0:
                    c["arquivos"] = list(writer.files)
                    progress_cb(dict(c))
                submit_next()
    finally:
        writer.close()

    c["arquivos"] = list(writer.files)
    if progress_cb is not None:
        progress_cb(dict(c))
    return c
