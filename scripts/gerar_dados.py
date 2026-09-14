"""
Le as planilhas e grava dados/dados.json, que e o arquivo consumido pelo painel.

Entrada esperada (pasta entrada/):
  produtividade.xlsx  -> aba "Produt_Montagem-Formas_R02", "BD-Obras", "BD-Custos"
  custos.xlsx         -> aba "Área Forma"   (opcional)

Se algum arquivo faltar, a parte correspondente do JSON fica vazia e o painel
continua funcionando com o que existe.
"""

import datetime as dt
import json
import math
import pathlib
import sys

import pandas as pd
from openpyxl import load_workbook

RAIZ = pathlib.Path(__file__).resolve().parents[1]
ENTRADA = RAIZ / "entrada"
SAIDA = RAIZ / "dados" / "dados.json"

PREFIXO_ABA = "Prod_"         # cada obra tem sua aba: Prod_Serena, Prod_Arte-Penha, ...
ABA_RESUMO = "GERAL"          # quadro consolidado por obra
ABA_ANTIGA = "Produt_Montagem-Formas_R02"   # formato anterior, aba unica
LINHA_CABECALHO = 15          # onde comeca o cabecalho da tabela semanal
LINHA_RESUMO = (7, 12)        # linhas do quadro da aba GERAL

# posicao das colunas no quadro GERAL
COL_RESUMO = {"obra": 2, "efetivo": 4, "efetivoAj": 5,
              "m2med": 6, "m2aj": 7, "m2max": 8,
              "m3med": 9, "m3aj": 10, "m3max": 11,
              "esp": 12, "espaj": 13, "espmax": 14}


def num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f)) else round(f, 4)
    except (TypeError, ValueError):
        return None


def data(v):
    if v is None:
        return None
    try:
        return pd.to_datetime(v).strftime("%Y-%m-%d")
    except Exception:
        return None


def numera_repetidas(cab):
    """As colunas dos dias repetem o mesmo titulo seis vezes. Numera as
    repeticoes como .1, .2 ... para que cada dia possa ser lido pelo nome."""
    vistos, saida = {}, []
    for nome in cab:
        n = vistos.get(nome, 0)
        saida.append(nome if n == 0 else f"{nome}.{n}")
        vistos[nome] = n + 1
    return saida


def abas_de_obra(wb):
    """Abas Prod_<obra>. Se nao houver nenhuma, cai no formato antigo de aba unica."""
    abas = [s for s in wb.sheetnames if s.startswith(PREFIXO_ABA)]
    if abas:
        return abas
    if ABA_ANTIGA in wb.sheetnames:
        print(f"Nenhuma aba '{PREFIXO_ABA}*'; usando '{ABA_ANTIGA}'.")
        return [ABA_ANTIGA]
    sys.exit(f"ERRO: nao encontrei abas '{PREFIXO_ABA}*' nem '{ABA_ANTIGA}'.")


def le_semanal(wb):
    """Junta as abas de obra. As colunas sao lidas pelo nome, entao a ordem
    pode variar de uma aba para outra sem quebrar nada."""
    semanas = []
    for aba in abas_de_obra(wb):
        semanas += le_aba(wb[aba], aba)
    if not semanas:
        sys.exit("ERRO: nenhuma linha de obra encontrada.")
    return semanas


def le_aba(ws, nome_aba):
    linhas = list(ws.iter_rows(min_row=LINHA_CABECALHO, values_only=True))
    if not linhas:
        return []
    cab = numera_repetidas([str(c) for c in linhas[0]])
    df = pd.DataFrame(linhas[1:], columns=cab)
    if "OBRA" not in df.columns:
        print(f"- {nome_aba}: sem coluna OBRA, ignorada.")
        return []
    df = df[df["OBRA"].notna()]

    def col(base, i):
        """Nome da coluna do dia i. O pandas numera as repetidas com .1, .2...
        Quando a coluna aparece uma vez so (caso de TORRE em algumas abas),
        o mesmo valor vale para todos os dias."""
        nome = base if i == 0 else f"{base}.{i}"
        return nome if nome in df.columns else (base if base in df.columns else None)

    semanas = []
    for _, r in df.iterrows():
        detalhe = []
        for i in range(6):                       # colunas DIA 01 a DIA 06
            c_area = col("ÁREA MONTADA (m2)", i)
            area = num(r.get(c_area)) if c_area else None
            if area and area > 0:
                def v(base):
                    c = col(base, i)
                    return r.get(c) if c else None
                detalhe.append({
                    "data": data(v("DATA CONCRET.")),
                    "torre": str(v("TORRE")),
                    "pavto": str(v("PAVTO.")),
                    "trecho": str(v("TRECHO MONTADO")),
                    "area": area,
                    "vol": num(v("VOLUME CONCRETADO (m3)")),
                })
        semanas.append({
            "obra": r["OBRA"],
            "ano": int(num(r["ANO"]) or 0),
            "sem": int(num(r["SEMANA DO ANO"]) or 0),
            "periodo": r["PERÍODO"],
            "ini": data(r.get("DATA CONCRET.")),
            "mes": data(r.get("MÊS")),
            "concr": num(r.get("QTDE. CONCRETAGENS DA TORRE NA SEMANA")),
            "sabado": (str(r.get("EQUIPE TRABALHOU NO SÁBADO?")) or "").strip(),
            "feriados": num(r.get("QTDE. FERIADOS NA SEMANA")),
            "area": num(r.get("SOMA DA ÁREA DE FÔRMA MONTADA NA SEMANA (m2)")) or 0,
            "vol": num(r.get("SOMA DE VOLUME CONCRETADO NA SEMANA (m3)")) or 0,
            "mont": num(r.get("QTDE. MÉDIA DE MONTADORES NA SEMANA")) or 0,
            "dias": num(r.get("DIAS TRABALHADOS NA SEMANA")) or 0,
            "horas": num(r.get("HORAS TRABALHADAS NA SEMANA")) or 0,
            "pm2sem": num(r.get("PRODUTIVIDADE MÉDIA SEMANAL POR MONTADOR  (m2/semana)")),
            "pm2dia": num(r.get("PRODUTIVIDADE MÉDIA DIÁRIA POR MONTADOR (m2/dia)")),
            "pm2h": num(r.get("PRODUTIVIDADE MÉDIA POR HORA POR MONTADOR (m2/h)")),
            "pm3dia": num(r.get("PRODUTIVIDADE MÉDIA DIÁRIA POR MONTADOR (m3/dia)")),
            "pm3h": num(r.get("PRODUTIVIDADE MÉDIA POR HORA POR MONTADOR (m3/h)")),
            "esp": num(r.get("ESPESSURA MÉDIA PRODUZIDA (m2/m3/dia)")),
            "rup": num(r.get("RUP - RAZÃO UNITÁRIA DE PRODUÇÃO (H.h/m²)")),
            "obs": str(r.get("OBSERVAÇÕES")) if r.get("OBSERVAÇÕES") else "",
            "detalhe": detalhe,
        })
    return semanas


def le_resumo(wb):
    """Quadro consolidado por obra. Fica na aba GERAL no formato novo."""
    aba = ABA_RESUMO if ABA_RESUMO in wb.sheetnames else None
    if not aba:
        return []
    ws = wb[aba]
    fora = []
    for r in ws.iter_rows(min_row=LINHA_RESUMO[0], max_row=LINHA_RESUMO[1], values_only=True):
        def pega(chave):
            i = COL_RESUMO[chave]
            return r[i] if i < len(r) else None
        obra = pega("obra")
        if not obra or not isinstance(obra, str):
            continue
        item = {"obra": obra}
        for chave in COL_RESUMO:
            if chave != "obra":
                item[chave] = num(pega(chave))
        fora.append(item)
    return fora


def le_bd_obras(wb):
    if "BD-Obras" not in wb.sheetnames:
        return []
    return [
        {"obra": r[0], "trecho": str(r[1]), "desc": r[2],
         "area": num(r[3]), "vol": num(r[4])}
        for r in wb["BD-Obras"].iter_rows(min_row=4, values_only=True) if r[0]
    ]


def le_bd_custos(wb):
    if "BD-Custos" not in wb.sheetnames:
        return [], []
    ws = wb["BD-Custos"]
    cargos = [
        {"cargo": r[0], "salario": num(r[1]), "comEncargos": num(r[2])}
        for r in ws.iter_rows(min_row=8, max_row=10, values_only=True) if r[0]
    ]
    equip = [
        {"tipo": r[0], "mensal": num(r[1])}
        for r in ws.iter_rows(min_row=18, max_row=19, values_only=True) if r[0]
    ]
    return cargos, equip


def le_area_forma(caminho):
    """Aba 'Área Forma' da planilha de mao de obra. Opcional."""
    if not caminho.exists():
        return []
    wb = load_workbook(caminho, read_only=True, data_only=True)
    aba = next((s for s in wb.sheetnames if s.strip().lower().startswith("área forma")
                or s.strip().lower().startswith("area forma")), None)
    if not aba:
        return []
    campos = ["obra", "quartos", "tipologia", "aptos", "areaFormaSLaje", "areaLaje",
              "indice", "custoParedeMO", "areaMontagem", "rsPorM2", "rsPorApto",
              "custoInst", "instConc60", "instPorM2", "totalPorM2"]
    saida = []
    for r in wb[aba].iter_rows(min_row=11, max_row=40, values_only=True):
        if not r[1] or not isinstance(r[1], str):
            continue
        item = {"obra": r[1], "quartos": r[2], "tipologia": r[3]}
        for i, campo in enumerate(campos[3:], start=4):
            item[campo] = num(r[i])
        if item["areaMontagem"] and item["rsPorM2"]:   # descarta linhas de nota
            saida.append(item)
    return saida


def main():
    prod = ENTRADA / "produtividade.xlsx"
    if not prod.exists():
        sys.exit(f"ERRO: {prod} nao encontrado. Rode baixar.py antes.")

    wb = load_workbook(prod, read_only=True, data_only=True)

    # data de atualizacao anotada na aba GERAL (celula D3)
    atualizacao = None
    for aba in (ABA_RESUMO, ABA_ANTIGA):
        if aba in wb.sheetnames:
            try:
                atualizacao = data(wb[aba].cell(3, 4).value)
            except Exception:
                atualizacao = None
            if atualizacao:
                break

    cargos, equip = le_bd_custos(wb)
    saida = {
        "geradoEm": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "atualizacaoPlanilha": atualizacao,
        "semanas": le_semanal(wb),
        "resumo": le_resumo(wb),
        "trechos": le_bd_obras(wb),
        "cargos": cargos,
        "equip": equip,
        "areaForma": le_area_forma(ENTRADA / "custos.xlsx"),
    }

    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    SAIDA.write_text(json.dumps(saida, ensure_ascii=False, separators=(",", ":")),
                     encoding="utf-8")

    com_producao = sum(1 for s in saida["semanas"] if s["area"] > 0)
    print(f"dados/dados.json gerado: {SAIDA.stat().st_size:,} bytes")
    print(f"  {len(saida['semanas'])} semanas ({com_producao} com producao)")
    for o in sorted(set(s["obra"] for s in saida["semanas"])):
        n = sum(1 for s in saida["semanas"] if s["obra"] == o and s["area"] > 0)
        print(f"    {o}: {n} semanas com producao")
    print(f"  {len(saida['areaForma'])} linhas de orcamento")


if __name__ == "__main__":
    main()
