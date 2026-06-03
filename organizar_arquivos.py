"""
Organizador de Arquivos Financeiros
====================================
Renomeia PDFs e planilhas Excel em pastas mensais seguindo o padrão:
  - Excel  → "Composição de Saldos - {Mês} {Ano}.xlsx"
  - Extrato PDF (contém 'extrato' no nome) → "Banco {BANCO} - {Mês} {Ano}.pdf"
  - Consolidado PDF (contém 'consolidado' no nome) → "Aplicação Banco {BANCO} - {Mês} {Ano}.pdf"

"""

import os
import re
import shutil
import argparse
from pathlib import Path

# ─── CONFIGURAÇÕES ────────────────────────────────────────────────────────────

# Caminho padrão (altere aqui se quiser rodar sem argumentos)
PASTA_RAIZ = r"C:\Users\carolinecs\Desktop\teste automação\01"

# Palavras-chave que identificam o tipo de PDF (em minúsculas)
PALAVRAS_EXTRATO      = ["extrato"]
PALAVRAS_CONSOLIDADO  = ["consolidado", "mensal", "aplicacao", "aplicação"]

# Bancos conhecidos — o script tenta detectar automaticamente pelo nome do arquivo.
# Adicione outros bancos aqui se necessário.
BANCOS_CONHECIDOS = [
    "ITAU", "ITAÚ", "BRADESCO", "SANTANDER", "CAIXA",
    "BB", "BANCO DO BRASIL", "SICREDI", "SICOOB", "NUBANK",
    "INTER", "BTG", "XP", "C6", "SAFRA",
]

# Meses em português
MESES_PT = {
    "01": "Janeiro",  "1":  "Janeiro",
    "02": "Fevereiro","2":  "Fevereiro",
    "03": "Março",    "3":  "Março",
    "04": "Abril",    "4":  "Abril",
    "05": "Maio",     "5":  "Maio",
    "06": "Junho",    "6":  "Junho",
    "07": "Julho",    "7":  "Julho",
    "08": "Agosto",   "8":  "Agosto",
    "09": "Setembro", "9":  "Setembro",
    "10": "Outubro",
    "11": "Novembro",
    "12": "Dezembro",
}

# Também aceita meses por extenso no nome do arquivo
MESES_EXTENSO = {v.upper(): v for v in MESES_PT.values()}

# ─── FUNÇÕES AUXILIARES ────────────────────────────────────────────────────────

def detectar_mes_ano(nome: str):
    """
    Tenta extrair mês e ano do nome do arquivo.
    Aceita formatos como: JAN 2026, JANEIRO 2026, 01/2026, 01-2026, 2026-01, etc.
    Retorna (mes_str, ano_str) ou (None, None).
    """
    nome_upper = nome.upper()

    # Primeiro tenta mês por extenso (ex: JANEIRO, FEVEREIRO...)
    for extenso, label in MESES_EXTENSO.items():
        if extenso in nome_upper:
            ano_match = re.search(r"\b(20\d{2})\b", nome)
            ano = ano_match.group(1) if ano_match else None
            return label, ano

    # Tenta abreviações de 3 letras (JAN, FEV, MAR...)
    abreviacoes = {
        "JAN": "Janeiro", "FEV": "Fevereiro", "MAR": "Março",
        "ABR": "Abril",   "MAI": "Maio",      "JUN": "Junho",
        "JUL": "Julho",   "AGO": "Agosto",    "SET": "Setembro",
        "OUT": "Outubro", "NOV": "Novembro",  "DEZ": "Dezembro",
    }
    for abrev, label in abreviacoes.items():
        if abrev in nome_upper:
            ano_match = re.search(r"\b(20\d{2})\b", nome)
            ano = ano_match.group(1) if ano_match else None
            return label, ano

    # Tenta formatos numéricos: MM/YYYY, MM-YYYY, YYYY-MM, YYYYMM
    padroes = [
        r"\b(0?[1-9]|1[0-2])[/\-](20\d{2})\b",   # MM/YYYY ou MM-YYYY
        r"\b(20\d{2})[/\-](0?[1-9]|1[0-2])\b",   # YYYY/MM ou YYYY-MM
    ]
    for padrao in padroes:
        m = re.search(padrao, nome)
        if m:
            g1, g2 = m.group(1), m.group(2)
            if g1.startswith("20"):   # YYYY-MM
                mes_num, ano = g2, g1
            else:                     # MM-YYYY
                mes_num, ano = g1, g2
            return MESES_PT.get(mes_num.zfill(2)), ano

    return None, None


def detectar_banco(nome: str) -> str:
    """Extrai o nome do banco do nome do arquivo."""
    nome_upper = nome.upper()
    for banco in BANCOS_CONHECIDOS:
        if banco.upper() in nome_upper:
            # Normaliza Itaú
            if banco.upper() in ("ITAU", "ITAÚ"):
                return "Itaú"
            return banco.title()

    # Fallback: tenta pegar a primeira palavra em maiúsculas que não seja
    # palavra reservada (EXTRATO, MENSAL, CONSOLIDADO, JANEIRO...)
    palavras_ignorar = set(
        list(MESES_EXTENSO.keys()) +
        ["EXTRATO", "MENSAL", "CONSOLIDADO", "APLICACAO", "APLICAÇÃO",
         "BANCO", "CONTA", "CORRENTE", "POUPANCA", "POUPANÇA"]
    )
    tokens = re.findall(r"[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÇ]{3,}", nome_upper)
    for token in tokens:
        if token not in palavras_ignorar:
            return token.title()

    return "BANCO"


def tipo_pdf(nome: str):
    """
    Retorna 'extrato', 'consolidado' ou None conforme palavras-chave no nome.
    Consolidado tem prioridade sobre extrato (arquivo pode conter ambas as palavras).
    """
    nome_lower = nome.lower()
    for p in PALAVRAS_CONSOLIDADO:
        if p in nome_lower:
            return "consolidado"
    for p in PALAVRAS_EXTRATO:
        if p in nome_lower:
            return "extrato"
    return None


def novo_nome_excel(mes: str, ano: str) -> str:
    sufixo = f" - {mes} {ano}" if ano else f" - {mes}"
    return f"Composição de Saldos{sufixo}.xlsx"


def novo_nome_extrato(banco: str, mes: str, ano: str) -> str:
    sufixo = f" - {mes} {ano}" if ano else f" - {mes}"
    return f"Banco {banco}{sufixo}.pdf"


def novo_nome_consolidado(banco: str, mes: str, ano: str) -> str:
    sufixo = f" - {mes} {ano}" if ano else f" - {mes}"
    return f"Aplicação Banco {banco}{sufixo}.pdf"


def nome_unico(pasta: Path, nome_desejado: str) -> str:
    """Garante que não vai sobrescrever um arquivo existente."""
    caminho = pasta / nome_desejado
    if not caminho.exists():
        return nome_desejado
    stem = Path(nome_desejado).stem
    suffix = Path(nome_desejado).suffix
    contador = 2
    while (pasta / f"{stem} ({contador}){suffix}").exists():
        contador += 1
    return f"{stem} ({contador}){suffix}"


# ─── PROCESSAMENTO PRINCIPAL ──────────────────────────────────────────────────

def processar_pasta(pasta_str: str, dry_run: bool = False):
    pasta = Path(pasta_str)

    if not pasta.exists():
        print(f"[ERRO] Pasta não encontrada: {pasta}")
        return

    arquivos = [f for f in pasta.iterdir() if f.is_file()]
    if not arquivos:
        print("[AVISO] Nenhum arquivo encontrado na pasta.")
        return

    print(f"\n{'='*60}")
    print(f"  Pasta: {pasta}")
    print(f"  Arquivos encontrados: {len(arquivos)}")
    if dry_run:
        print("  MODO SIMULAÇÃO — nenhum arquivo será alterado")
    print(f"{'='*60}\n")

    renomeados   = 0
    ignorados    = 0
    sem_mes      = 0

    for arquivo in sorted(arquivos):
        ext  = arquivo.suffix.lower()
        nome = arquivo.stem   # sem extensão

        # ── Excel ──────────────────────────────────────────────────────────
        if ext in (".xlsx", ".xls", ".xlsm"):
            mes, ano = detectar_mes_ano(nome)
            if not mes:
                print(f"  [SEM MÊS]  {arquivo.name}  →  não foi possível detectar o mês")
                sem_mes += 1
                continue
            destino_nome = nome_unico(pasta, novo_nome_excel(mes, ano))
            destino = pasta / destino_nome
            print(f"  [EXCEL]    {arquivo.name}")
            print(f"         →   {destino_nome}")
            if not dry_run:
                arquivo.rename(destino)
            renomeados += 1

        # ── PDF ────────────────────────────────────────────────────────────
        elif ext == ".pdf":
            tipo = tipo_pdf(nome)
            if tipo is None:
                print(f"  [IGNORADO] {arquivo.name}  →  não identificado como extrato ou consolidado")
                ignorados += 1
                continue

            mes, ano = detectar_mes_ano(nome)
            if not mes:
                print(f"  [SEM MÊS]  {arquivo.name}  →  não foi possível detectar o mês")
                sem_mes += 1
                continue

            banco = detectar_banco(nome)

            if tipo == "extrato":
                destino_nome = nome_unico(pasta, novo_nome_extrato(banco, mes, ano))
            else:
                destino_nome = nome_unico(pasta, novo_nome_consolidado(banco, mes, ano))

            destino = pasta / destino_nome
            print(f"  [PDF/{tipo.upper()[:5]}] {arquivo.name}")
            print(f"         →   {destino_nome}")
            if not dry_run:
                arquivo.rename(destino)
            renomeados += 1

        # ── Outros ─────────────────────────────────────────────────────────
        else:
            print(f"  [PULADO]   {arquivo.name}  →  extensão não tratada ({ext})")
            ignorados += 1

    print(f"\n{'─'*60}")
    print(f"  ✔ Renomeados : {renomeados}")
    print(f"  ✗ Sem mês    : {sem_mes}")
    print(f"  – Ignorados  : {ignorados}")
    print(f"{'─'*60}\n")


# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Renomeia arquivos financeiros (Excel/PDF) por mês."
    )
    parser.add_argument(
        "--pasta", "-p",
        default=PASTA_RAIZ,
        help="Caminho da pasta a processar (padrão: PASTA_RAIZ no script)"
    )
    parser.add_argument(
        "--simular", "-s",
        action="store_true",
        help="Modo dry-run: mostra o que seria feito sem renomear nada"
    )
    args = parser.parse_args()

    processar_pasta(args.pasta, dry_run=args.simular)
