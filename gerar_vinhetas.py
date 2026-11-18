#!/usr/bin/env python3
"""
Gera vinhetas clínicas padronizadas a partir de uma planilha codificada (modelo do TCC).

COMO USAR
    1. Preencha a aba "dados_codificados" do modelo (uma linha por caso, colunas A a I).
    2. Deixe este arquivo na mesma pasta da planilha e rode:
           python3 gerar_vinhetas.py                    # acha a planilha da pasta sozinho
           python3 gerar_vinhetas.py minha_planilha.xlsx
    3. Saem, na mesma pasta:
           <planilha>_com_vinhetas.xlsx  – cópia da planilha com as colunas "Vinheta" e "Status" preenchidas
           <planilha>_com_vinhetas.txt   – as mesmas vinhetas em texto corrido (comando + vinheta)

O texto é montado por regras fixas a partir do dicionário (aba com Letra/Código/Significado/
Subgrupo), sem IA: o mesmo código gera sempre o mesmo texto. Códigos novos entram sozinhos,
desde que recebam um dos subgrupos já usados (ver aba "Instruções" do modelo).

O comando e a regra da radiografia ficam na aba "configuracao" — não precisa mexer aqui.
Opção de linha de comando (sobrepõe a aba): --sem-radiografia=P04,P20

Requer Python 3 e:  pip install openpyxl
"""
import sys, re
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Falta a biblioteca openpyxl. Instale com:  pip install openpyxl")

# Valores padrão (usados só se a planilha não tiver a aba "configuracao")
COMANDO = "Preciso que, baseado nesses dados, você elabore as principais hipóteses diagnósticas, em ordem de probabilidade:"
INCLUIR_RADIOGRAFIA = True        # acrescenta "Radiografia da região em anexo." quando H = 2
SEM_RADIOGRAFIA_QUALIDADE = set() # IDs com H = 2 cuja radiografia foi descartada (tratados como sem radiografia)

OBRIGATORIAS = "ABCDEFGH"         # variáveis que entram na vinheta (I e J não entram)

# Concordância dos termos de coloração com a palavra "coloração".
COLORACAO_FEM = {"arroxeado": "arroxeada", "azulado": "azulada",
                 "normocrômico": "normocrômica", "amarelado": "amarelada",
                 "eritematoso/avermelhado": "eritematosa/avermelhada"}

FEMININO = {"diabético": "diabética", "diabético descompensado": "diabética descompensada",
            "hipertenso": "hipertensa", "imunossuprimido": "imunossuprimida",
            "etilista crônico": "etilista crônica"}
ADJETIVOS = {"tabagista", "ex-tabagista", "etilista crônico", "gestante", "diabético",
             "diabético descompensado", "hipertenso", "imunossuprimido"}
LOCAL_TEXTO = {"ATM/região pré-auricular": "ATM/pré-auricular"}


class ErroCodigo(ValueError):
    pass


def low(s):
    s = " ".join(str(s).split())   # minúsculas, preservando siglas (HPV, ISRS, ATM, TGI, LP)
    return re.sub(r"\w+", lambda m: m.group() if (m.group().isupper() and len(m.group()) > 1) else m.group().lower(), s)


def lista(itens):
    itens = [i for i in itens if i]
    return itens[0] if len(itens) == 1 else ", ".join(itens[:-1]) + " e " + itens[-1]


def cabecalho(ws):
    return [str(c.value or "").strip() for c in ws[1]]


def achar_abas(wb):
    dic = dados = None
    for ws in wb:
        h = cabecalho(ws)
        if dic is None and {"Letra", "Código", "Significado"} <= set(h):
            dic = ws
        if dados is None and "ID" in h and any(x.startswith("D ") for x in h):   # 1ª aba de dados
            dados = ws
    if not dic or not dados:
        sys.exit("Não encontrei a aba do dicionário (colunas Letra/Código/Significado) "
                 "e/ou a aba de dados (colunas ID, A ..., B ..., etc.).")
    return dic, dados


def ler_dicionario(ws):
    h = cabecalho(ws)
    iL, iC, iS = h.index("Letra"), h.index("Código"), h.index("Significado")
    iG = h.index("Subgrupo") if "Subgrupo" in h else None
    d, letra = {}, None
    for r in ws.iter_rows(min_row=2, values_only=True):
        letra = r[iL] or letra            # linha sem letra herda a de cima
        try:
            cod = int(float(r[iC]))
        except (TypeError, ValueError):
            continue
        if r[iS] is None:
            continue
        d[(str(letra).strip(), cod)] = (low(r[iS]), (r[iG] or "") if iG is not None else "")
    return d


def ler_config(wb):
    """Aba 'configuracao': coluna A = nome, coluna B = valor."""
    cfg = {"comando": COMANDO, "radiografia": INCLUIR_RADIOGRAFIA,
           "sem_rad": set(SEM_RADIOGRAFIA_QUALIDADE)}
    if "configuracao" not in wb.sheetnames:
        return cfg
    for nome, valor, *_ in wb["configuracao"].iter_rows(min_row=1, values_only=True):
        nome = str(nome or "").strip().lower()
        if nome.startswith("comando") and valor:
            cfg["comando"] = " ".join(str(valor).split())
        elif nome.startswith("incluir frase"):
            cfg["radiografia"] = str(valor or "").strip().lower() not in ("não", "nao", "n", "0", "false")
        elif nome.startswith("ids com radiografia"):
            cfg["sem_rad"] = {x.strip().upper() for x in re.split(r"[;,\s]+", str(valor or "")) if x.strip()}
    return cfg


def codigos(v):
    """'3;4' -> [3, 4]. Recusa número quebrado (ex.: 3,4 virou 3.4 no Excel em português)."""
    if v is None or str(v).strip() == "":
        return []
    if isinstance(v, float) and not v.is_integer():
        raise ErroCodigo(f"o valor {v} parece '3;4' digitado com vírgula e convertido em número "
                         "decimal pelo Excel — use ponto e vírgula")
    out = set()
    for x in re.split(r"[;,]", str(v)):
        x = x.strip()
        if not x:
            continue
        try:
            f = float(x)
        except ValueError:
            raise ErroCodigo(f"'{x}' não é um código numérico")
        if not f.is_integer():
            raise ErroCodigo(f"'{x}' não é um código inteiro")
        out.add(int(f))
    return sorted(out)


def vinheta(caso, D, radiografia=None, incluir_radiografia=INCLUIR_RADIOGRAFIA):
    """radiografia: True/False força a frase; None usa H = 2."""
    t = lambda L: [D[(L, c)] for c in caso[L]]          # [(termo, subgrupo)]
    fem = caso["A"] == [1]
    frases = []

    # Localização (E)
    locais = [LOCAL_TEXTO.get(s, s) for s, _ in t("E")]
    difuso = "difuso" in locais
    locais = [l for l in locais if l != "difuso"]
    onde = ("em região de " + lista(locais)) if locais else ""
    if difuso:
        onde = (onde + ", de forma difusa") if onde else "de forma difusa"

    # 1. Identificação + queixa (com o local da queixa, quando houver sintoma)
    sexo = "feminino" if fem else "masculino"
    sint = [s for s, _ in t("C")]
    assintomatico = not sint or caso["C"] == [1]
    # sintoma que já traz o próprio local (ex.: "dor em face") mantém o local no exame
    local_na_queixa = (not assintomatico) and bool(onde) and not any(" em " in s for s in sint)
    if assintomatico:
        queixa = "sem sintomatologia"
    else:
        queixa = "relata " + lista(sint) + (f" {onde}" if local_na_queixa else "")
    frases.append(f"Paciente do sexo {sexo}, {caso['B']} anos, {queixa}.")

    # 2. Exame clínico: achado principal + localização + sinais associados
    exame = "Ao exame clínico da região" if local_na_queixa else "Ao exame clínico"
    onde_ex = "" if local_na_queixa else (f" {onde}" if onde else "")

    d = t("D")
    lesoes = [s for s, g in d if g == "Lesão fundamental"]
    outros = [s for s, g in d if g.startswith(("Alteração", "ATM"))]
    sinais = [s for s, g in d if g.startswith("Sinais")]
    carac = {}                      # subgrupo -> lista de termos (um subgrupo pode ter vários)
    for s, g in d:
        if g.startswith("Característica"):
            carac.setdefault(g.split("–")[-1].strip(), []).append(s)

    if caso["D"] == [1]:
        if local_na_queixa or not locais:
            frases.append(f"{exame}, não foram observadas alterações.")
        else:
            frases.append(f"Ao exame clínico da região de {lista(locais)}, não foram observadas alterações.")
    elif lesoes or outros:
        frase = f"{exame}, observa-se presença de {lista(lesoes + outros)}{onde_ex}"
        frase += f", com {lista(sinais)}" if sinais else ""
        frases.append(frase + ".")
    elif sinais:
        frases.append(f"{exame}, observa-se presença de {lista(sinais)}{onde_ex}.")

    # 3. Características da lesão
    if carac:
        plural = "múltiplos" in carac.get("número", [])
        if lesoes:
            sujeito = "Lesões" if plural else "Lesão"
            if "número" in carac:
                sujeito += " múltiplas" if plural else " única"
        else:
            sujeito = "Quadro"

        def agrupado(chave, substantivo):
            """'superfície estriada' + 'superfície rugosa' -> 'superfície estriada e rugosa'."""
            termos = carac.get(chave, [])
            if not termos:
                return None
            adj = [t[len(substantivo):].strip() if t.lower().startswith(substantivo) else t
                   for t in termos]
            return f"{substantivo} {lista(adj)}"

        meio = []
        for k in ("limites", "base", "consistência"):
            if k in carac:
                meio.append(lista(carac[k]))
        for k, subst in (("superfície", "superfície"), ("bordas", "bordas")):
            p = agrupado(k, subst)
            if p:
                meio.append(p)
        if "coloração" in carac:
            meio.append("coloração " + lista([COLORACAO_FEM.get(c, c)
                                              for c in carac["coloração"]]))

        # a vírgula só cabe quando o sujeito tem qualificador ("Lesão única, de ...");
        # em "Lesão" ou "Quadro" sozinhos ela fica solta
        liga = ", de " if (" " in sujeito and meio) else (" de " if meio else "")
        frase = sujeito + (liga + lista(meio) if meio else "")
        if "tempo de evolução" in carac:
            tempo = carac["tempo de evolução"][0]
            if plural and sujeito.startswith("Lesões"):
                tempo = tempo.replace("presente", "presentes", 1)
            frase += ("," if meio else "") + " " + tempo
        frases.append(frase + ".")

    # 4. Fatores locais
    if caso["F"] and caso["F"] != [1]:
        frases.append("Como fatores locais, apresenta " + lista([s for s, _ in t("F")]) + ".")
    elif caso["F"] == [1]:
        frases.append("Sem fatores locais modificadores.")

    # 5. Fatores sistêmicos: condições → histórico → medicamentos
    if caso["G"] == [1]:
        frases.append("Sem fatores sistêmicos associados.")
    elif caso["G"]:
        g = t("G")
        adj = [FEMININO.get(s, s) if fem else s for s, _ in g if s in ADJETIVOS]
        hist = [s for s, grp in g if s not in ADJETIVOS and not grp.startswith("Medicamento")]
        meds = [re.sub(r"^medicamentos? (da classe d[oa]s |para )", "", s)
                for s, grp in g if grp.startswith("Medicamento")]
        if adj:
            frases.append("Sistemicamente, é " + lista(adj) + ".")
        if hist:
            frases.append(("Apresenta" if not adj else "Apresenta também") + " histórico de " + lista(hist) + ".")
        if meds:
            frases.append("Faz uso de " + lista(meds) + ".")

    # 6. Radiografia
    tem_rad = (caso.get("H") == [2]) if radiografia is None else radiografia
    if incluir_radiografia and tem_rad:
        frases.append("Radiografia da região em anexo.")

    return " ".join(frases)


def processar_linha(pid, r, col, D, cfg):
    """Devolve (vinheta, status). Vinheta vazia quando há erro que impede o texto."""
    caso, problemas = {}, []
    for L in "ACDEFGHIJ":
        if L not in col:
            caso[L] = []
            continue
        try:
            caso[L] = codigos(r[col[L]])
        except ErroCodigo as e:
            return "", f"ERRO em {L}: {e}"
    b = r[col["B"]]
    try:
        caso["B"] = int(float(b)) if b not in (None, "") else None
    except (TypeError, ValueError):
        return "", f"ERRO em B: idade '{b}' não é um número"

    faltando = [f"{L}{c}" for L in "ACDEFGHIJ" for c in caso[L] if (L, c) not in D]
    if faltando:
        return "", "ERRO: código(s) que não existem no dicionário: " + ", ".join(faltando)

    vazias = [L for L in OBRIGATORIAS if not caso[L] and caso.get(L) is not None and L != "B"]
    if caso["B"] is None:
        vazias.insert(1, "B")
        caso["B"] = "?"
    for L in OBRIGATORIAS:
        if L in "AH" and len(caso[L]) > 1:
            problemas.append(f"{L} aceita um código só")
    if vazias:
        problemas.insert(0, "incompleta, falta: " + ", ".join(vazias))

    rad = caso["H"] == [2] and str(pid).upper() not in cfg["sem_rad"]
    if caso["H"] == [2] and not rad:
        problemas.append("radiografia descartada (configuracao)")
    texto = vinheta(caso, D, radiografia=rad, incluir_radiografia=cfg["radiografia"])
    return texto, ("ATENÇÃO: " + "; ".join(problemas)) if problemas else "OK"


def achar_planilha():
    pasta = Path(__file__).resolve().parent
    cands = [p for p in sorted(pasta.glob("*.xlsx"))
             if not p.stem.endswith("_com_vinhetas") and not p.name.startswith("~$")]
    if len(cands) == 1:
        return cands[0]
    if not cands:
        sys.exit(f"Nenhuma planilha .xlsx na pasta {pasta}. Coloque o modelo preenchido ao lado do script "
                 "ou passe o caminho:  python3 gerar_vinhetas.py minha_planilha.xlsx")
    sys.exit("Há mais de uma planilha na pasta; diga qual usar:\n  python3 gerar_vinhetas.py <arquivo.xlsx>\n"
             + "\n".join("  - " + p.name for p in cands))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    opts = [a for a in sys.argv[1:] if a.startswith("--")]
    if args and args[0] in ("-h", "ajuda"):
        sys.exit(__doc__)
    src = Path(args[0]) if args else achar_planilha()
    if not src.exists():
        sys.exit(f"Arquivo não encontrado: {src}")

    valores = openpyxl.load_workbook(src, data_only=True)   # para ler (resultado das fórmulas)
    wb = openpyxl.load_workbook(src)                         # para gravar (preserva as fórmulas)
    cfg = ler_config(valores)
    for o in opts:
        if o.startswith("--sem-radiografia="):
            cfg["sem_rad"] = {x.strip().upper() for x in o.split("=", 1)[1].split(",") if x.strip()}

    wdic, wdat = achar_abas(valores)
    D = ler_dicionario(wdic)
    h = cabecalho(wdat)
    col = {}
    for L in "ABCDEFGHIJ":
        i = next((i for i, x in enumerate(h) if x.startswith(L + " ")), None)
        if i is not None:
            col[L] = i
    sem_coluna = [L for L in OBRIGATORIAS if L not in col]
    if sem_coluna:
        sys.exit("Faltam colunas na aba de dados: " + ", ".join(sem_coluna)
                 + " (o título deve começar pela letra e um espaço, ex.: 'C Sintomatologia').")
    iid = h.index("ID")

    # colunas de saída na própria aba de dados (cria se não existirem)
    ws = wb[wdat.title]
    def coluna_saida(nome):
        for c in ws[1]:
            if str(c.value or "").strip() == nome:
                return c.column
        n = max(len(h), ws.max_column) + 1
        while ws.cell(1, n).value not in (None, ""):
            n += 1
        ws.cell(1, n, nome)
        return n
    cV, cS = coluna_saida("Vinheta"), coluna_saida("Status")

    saida, avisos, n_ok = [], [], 0
    for i, r in enumerate(wdat.iter_rows(min_row=2, values_only=True), start=2):
        pid = r[iid] if iid < len(r) else None
        dados = [r[col[L]] for L in OBRIGATORIAS]
        if pid in (None, "") and all(v in (None, "") for v in dados):
            continue
        if pid in (None, ""):
            pid = f"(linha {i})"
        if all(v in (None, "") for v in dados):
            texto, status = "", "sem dados codificados (linha ignorada)"
        else:
            texto, status = processar_linha(pid, r, col, D, cfg)
        ws.cell(i, cV, texto)
        ws.cell(i, cS, status)
        ws.cell(i, cV).alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")
        if texto:
            saida.append((pid, texto))
        if status == "OK":
            n_ok += 1
        else:
            avisos.append(f"  {pid}: {status}")

    destino = src.with_name(src.stem + "_com_vinhetas.xlsx")
    try:
        wb.save(destino)
    except PermissionError:
        sys.exit(f"Não consegui salvar {destino.name}: feche o arquivo no Excel e rode de novo.")
    with open(src.with_name(src.stem + "_com_vinhetas.txt"), "w", encoding="utf-8") as f:
        for pid, v in saida:
            f.write(f"{pid}\n{cfg['comando']}\n{v}\n\n")

    print(f"{len(saida)} vinhetas geradas ({n_ok} sem pendências) em:\n  {destino}\n  {destino.with_suffix('.txt')}")
    if avisos:
        print("Linhas para conferir (ver coluna Status):\n" + "\n".join(avisos))


if __name__ == "__main__":
    main()
