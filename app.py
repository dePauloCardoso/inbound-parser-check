import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import re
import io
from pathlib import Path

NAMESPACE   = {"ns": "http://www.portalfiscal.inf.br/nfe"}
SKU_PATTERN = re.compile(r'\b[A-Z0-9]{13}\b')
SKU_FILE    = Path(__file__).parent / "skuProd.csv"
ZZSETE_FILE = Path(__file__).parent / "zzsete.xlsx"


# ─── Tabela auxiliar de SKUs ──────────────────────────────────────────────────

@st.cache_data
def carregar_sku_set():
    if not SKU_FILE.exists():
        return set(), pd.DataFrame()
    try:
        df = pd.read_csv(SKU_FILE, dtype=str, sep=None, engine="python")
    except Exception:
        df = pd.read_csv(SKU_FILE, dtype=str, sep=",")
    col_sku = df.columns[0]
    sku_set = set(df[col_sku].dropna().str.strip().str.upper().tolist())
    return sku_set, df


# ─── Tabela auxiliar de Kits (zzsete) ────────────────────────────────────────

@st.cache_data
def carregar_zzsete():
    if not ZZSETE_FILE.exists():
        return pd.DataFrame()
    df = pd.read_excel(ZZSETE_FILE, dtype=str)
    df.columns = df.columns.str.strip()
    for col in ["ZZ7_CODIGO", "ZZ7_CODPAI", "ZZ7_PRODUT", "ZZ7_SRINK"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()
    return df


# ─── Helpers gerais ───────────────────────────────────────────────────────────

def extrair_nf_retorno(infCpl):
    if not infCpl:
        return ""
    padroes = [
        r'RETORNO\s+DA\s+NOTA\s+(\d{1,5})',
        r'NOTA\s+FISCAL\s+(\d{1,5})',
        r'REF\s+NF\s+(\d{1,5})',
        r'NOTA\s+(\d{1,5})',
        r'NF\s+(\d{1,5})',
    ]
    for padrao in padroes:
        match = re.search(padrao, infCpl, re.IGNORECASE)
        if match:
            return match.group(1)
    return ""

def is_sku(code):
    return bool(code and re.fullmatch(r'[A-Z0-9]{13}', str(code).strip().upper()))

def extrair_sku_do_xprod(xprod, sku_set):
    if not xprod:
        return None
    xprod_normalizado = xprod.replace("_", " ").strip().upper()
    tokens = xprod_normalizado.split()
    for token in tokens:
        token_clean = re.sub(r'[^A-Z0-9]', '', token)
        if len(token_clean) == 13 and token_clean in sku_set:
            return token_clean
    for m in SKU_PATTERN.findall(xprod_normalizado):
        if m in sku_set:
            return m
    return None

def normalizar_chave(valor):
    try:
        return str(int(str(valor).strip()))
    except (ValueError, TypeError):
        return str(valor).strip()

def parse_numero_br(valor_str):
    if pd.isna(valor_str) or str(valor_str).strip() == "":
        return None
    s = str(valor_str).strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1:
        parts = s.split(".")
        s = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(s)
    except ValueError:
        return None

def normalizar_kit(val):
    try:
        return str(int(float(str(val).strip())))
    except (ValueError, TypeError):
        return str(val).strip().upper()


# ─── Parse XML ────────────────────────────────────────────────────────────────

def parse_nfe_xml(xml_content, filename="arquivo.xml"):
    root   = ET.fromstring(xml_content)
    infNFe = root.find(".//ns:infNFe", NAMESPACE)

    ide   = infNFe.find("ns:ide",  NAMESPACE)
    emit  = infNFe.find("ns:emit", NAMESPACE)
    dest  = infNFe.find("ns:dest", NAMESPACE)
    total = infNFe.find("ns:total/ns:ICMSTot", NAMESPACE)

    infAdic = infNFe.find("ns:infAdic", NAMESPACE)
    infCpl  = ""
    if infAdic is not None:
        infCpl = infAdic.findtext("ns:infCpl", default="", namespaces=NAMESPACE)

    header = {
        "arquivo":        filename,
        "id_nfe":         infNFe.attrib.get("Id"),
        "numero_nf":      ide.findtext("ns:nNF",   default="", namespaces=NAMESPACE),
        "serie":          ide.findtext("ns:serie",  default="", namespaces=NAMESPACE),
        "data_emissao":   ide.findtext("ns:dhEmi",  default="", namespaces=NAMESPACE),
        "valor_total_nf": total.findtext("ns:vNF",  default="", namespaces=NAMESPACE),
        "emit_cnpj":      emit.findtext("ns:CNPJ",  default="", namespaces=NAMESPACE),
        "emit_nome":      emit.findtext("ns:xNome", default="", namespaces=NAMESPACE),
        "dest_cnpj":      dest.findtext("ns:CNPJ",  default="", namespaces=NAMESPACE),
        "dest_nome":      dest.findtext("ns:xNome", default="", namespaces=NAMESPACE),
        "infCpl":         infCpl,
        "ret_nf":         extrair_nf_retorno(infCpl),
    }

    itens = []
    for det in infNFe.findall("ns:det", NAMESPACE):
        prod = det.find("ns:prod", NAMESPACE)
        item = {
            **header,
            "n_item":   det.attrib.get("nItem"),
            "cProd":    prod.findtext("ns:cProd",    default="", namespaces=NAMESPACE),
            "xProd":    prod.findtext("ns:xProd",    default="", namespaces=NAMESPACE),
            "CFOP":     prod.findtext("ns:CFOP",     default="", namespaces=NAMESPACE),
            "NCM":      prod.findtext("ns:NCM",      default="", namespaces=NAMESPACE),
            "qCom":     int(float(prod.findtext("ns:qCom",    default="0", namespaces=NAMESPACE))),
            "vUnCom":   round(float(prod.findtext("ns:vUnCom", default="0", namespaces=NAMESPACE)), 2),
            "vProd":    round(float(prod.findtext("ns:vProd",  default="0", namespaces=NAMESPACE)), 2),
            "xPed":     prod.findtext("ns:xPed",     default="", namespaces=NAMESPACE),
            "nItemPed": prod.findtext("ns:nItemPed", default="", namespaces=NAMESPACE),
        }
        itens.append(item)

    return itens


# ─── Resolução de SKUs ────────────────────────────────────────────────────────

def resolver_skus(df, sku_set):
    def resolver(row):
        cprod = str(row["cProd"]).strip().upper()
        if is_sku(cprod) and cprod in sku_set:
            return pd.Series({"cProd": cprod, "sku_origem": "cProd"})
        sku_encontrado = extrair_sku_do_xprod(row["xProd"], sku_set)
        if sku_encontrado:
            return pd.Series({"cProd": sku_encontrado, "sku_origem": "xProd"})
        return pd.Series({"cProd": row["cProd"], "sku_origem": "não resolvido"})

    resultado        = df.apply(resolver, axis=1)
    df               = df.copy()
    df["cProd"]      = resultado["cProd"]
    df["sku_origem"] = resultado["sku_origem"]
    return df


# ─── Check unicidade xPed ─────────────────────────────────────────────────────

def check_unicidade_xped(df: pd.DataFrame) -> pd.DataFrame:
    resumo = (
        df[df["xPed"].str.strip() != ""]
        .groupby(["arquivo", "numero_nf"])["xPed"]
        .agg(lambda x: sorted(x.unique().tolist()))
        .reset_index()
        .rename(columns={"xPed": "xPed_valores"})
    )
    resumo["qtd_pedidos_distintos"] = resumo["xPed_valores"].apply(len)
    resumo["xPed_valores"]          = resumo["xPed_valores"].apply(lambda v: " | ".join(v))
    return resumo[resumo["qtd_pedidos_distintos"] > 1].copy()


# ─── Pedidos de Compra ────────────────────────────────────────────────────────

def carregar_pedidos(file) -> pd.DataFrame:
    if file.name.endswith(".csv"):
        df = pd.read_csv(file, dtype=str)
    else:
        df = pd.read_excel(file, dtype=str)

    df.columns = df.columns.str.strip()

    colunas_numericas = [
        "Qtd.pedido",
        "Valor líquido da ordem",
        "Quantidade a ser fornecida",
    ]
    for col in colunas_numericas:
        if col in df.columns:
            df[col] = df[col].apply(parse_numero_br)

    if "Valor líquido da ordem" in df.columns and "Qtd.pedido" in df.columns:
        df["vUnit"] = (
            df["Valor líquido da ordem"].astype(float)
            / df["Qtd.pedido"].astype(float)
        ).round(2)

    for col in ["Pedido de compras", "Item", "Material", "Numero_KIT"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    if "Pedido de compras" in df.columns:
        df["_pedido_norm"] = df["Pedido de compras"].apply(normalizar_chave)
    if "Item" in df.columns:
        df["_item_norm"] = df["Item"].apply(normalizar_chave)

    return df


# ─── Comparativo NF-e x PO ───────────────────────────────────────────────────

def comparar_nfe_po(df_nfe: pd.DataFrame, df_po: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for _, nfe_row in df_nfe.iterrows():
        xped   = normalizar_chave(nfe_row.get("xPed",     ""))
        nitem  = normalizar_chave(nfe_row.get("nItemPed", ""))
        cprod  = str(nfe_row.get("cProd",  "")).strip().upper()
        qcom   = nfe_row.get("qCom",   None)
        vuncom = nfe_row.get("vUnCom", None)

        po_pedido = df_po[df_po["_pedido_norm"] == xped]

        if po_pedido.empty:
            rows.append({
                **nfe_row.to_dict(),
                "po_pedido":        xped,
                "po_item":          nitem,
                "po_material":      "",
                "po_nome_material": "",
                "po_qtd_pedido":    None,
                "po_qtd_fornecida": None,
                "po_vUnit":         None,
                "check_pedido":     "❌ Pedido não encontrado no PO",
                "check_material":   "",
                "check_qtd":        "",
                "check_vunit":      "",
            })
            continue

        po_linha = po_pedido[po_pedido["_item_norm"] == nitem]

        if po_linha.empty:
            itens_disponiveis = po_pedido["Item"].tolist()
            rows.append({
                **nfe_row.to_dict(),
                "po_pedido":        xped,
                "po_item":          nitem,
                "po_material":      "",
                "po_nome_material": "",
                "po_qtd_pedido":    None,
                "po_qtd_fornecida": None,
                "po_vUnit":         None,
                "check_pedido":     "✅ Pedido encontrado",
                "check_material":   f"❌ Item {nitem} não encontrado no PO (disponíveis: {itens_disponiveis})",
                "check_qtd":        "",
                "check_vunit":      "",
            })
            continue

        po = po_linha.iloc[0]

        po_material      = str(po.get("Material",              "")).strip().upper()
        po_nome          = str(po.get("Nome do material",       "")).strip()
        po_qtd           = po.get("Qtd.pedido",                 None)
        po_qtd_fornecida = po.get("Quantidade a ser fornecida", None)
        po_vunit         = po.get("vUnit",                      None)

        try:
            po_qtd           = float(po_qtd)           if po_qtd           is not None else None
            po_qtd_fornecida = float(po_qtd_fornecida) if po_qtd_fornecida is not None else None
            po_vunit         = float(po_vunit)         if po_vunit         is not None else None
        except (ValueError, TypeError):
            po_qtd = po_qtd_fornecida = po_vunit = None

        check_pedido = "✅ Pedido encontrado"

        check_material = (
            "✅ Material OK"
            if cprod == po_material
            else f"❌ Material divergente — NF: {cprod} | PO: {po_material}"
        )

        if qcom is not None and po_qtd is not None:
            qcom_f = float(qcom)
            if qcom_f == po_qtd:
                check_qtd = "✅ Qtd OK"
            elif qcom_f < po_qtd:
                if po_qtd_fornecida is not None and po_qtd_fornecida == 0:
                    check_qtd = (
                        f"✅ Qtd OK (entrega parcial) — "
                        f"NF: {qcom_f:.0f} | Qtd. a fornecer: 0 "
                        f"| Qtd. pedido: {po_qtd:.0f}"
                    )
                elif po_qtd_fornecida is not None and qcom_f <= po_qtd_fornecida:
                    check_qtd = (
                        f"✅ Qtd OK (entrega parcial) — "
                        f"NF: {qcom_f:.0f} | Qtd. a fornecer: {po_qtd_fornecida:.0f} "
                        f"| Qtd. pedido: {po_qtd:.0f}"
                    )
                else:
                    qtd_ref = f"{po_qtd_fornecida:.0f}" if po_qtd_fornecida is not None else "não informada"
                    check_qtd = (
                        f"❌ Qtd divergente — "
                        f"NF: {qcom_f:.0f} | Qtd. a fornecer: {qtd_ref} "
                        f"| Qtd. pedido: {po_qtd:.0f}"
                    )
            else:
                check_qtd = (
                    f"❌ Qtd excede o pedido — "
                    f"NF: {qcom_f:.0f} | Qtd. pedido: {po_qtd:.0f}"
                )
        else:
            check_qtd = "⚠️ Qtd não comparável"

        if vuncom is not None and po_vunit is not None:
            diff_pct = (
                abs(float(vuncom) - po_vunit) / po_vunit * 100
                if po_vunit != 0 else 0.0
            )
            check_vunit = (
                "✅ Valor unit. OK"
                if diff_pct < 1
                else (
                    f"❌ Valor unit. divergente — "
                    f"NF: {float(vuncom):.2f} | PO: {po_vunit:.2f} ({diff_pct:.1f}%)"
                )
            )
        else:
            check_vunit = "⚠️ Valor unit. não comparável"

        rows.append({
            **nfe_row.to_dict(),
            "po_pedido":        xped,
            "po_item":          nitem,
            "po_material":      po_material,
            "po_nome_material": po_nome,
            "po_qtd_pedido":    po_qtd,
            "po_qtd_fornecida": po_qtd_fornecida,
            "po_vUnit":         round(po_vunit, 2) if po_vunit is not None else None,
            "check_pedido":     check_pedido,
            "check_material":   check_material,
            "check_qtd":        check_qtd,
            "check_vunit":      check_vunit,
        })

    df_result = pd.DataFrame(rows)
    for col in ["_pedido_norm", "_item_norm"]:
        if col in df_result.columns:
            df_result.drop(columns=[col], inplace=True)
    return df_result


# ─── Check Composição de Kits (Aba 3) ────────────────────────────────────────

def check_composicao_kits(df_po: pd.DataFrame, df_zzsete: pd.DataFrame) -> pd.DataFrame:
    if df_zzsete.empty:
        return pd.DataFrame([{
            "Numero_KIT":     "—",
            "kit_zzsete":     "—",
            "kit_sugerido":   "—",
            "material_po":    "—",
            "nome_material":  "—",
            "qtd_po":         None,
            "componente_kit": "—",
            "check_kit":      "⚠️ Base zzsete não disponível",
            "check_qtd_kit":  "—",
        }])

    zzsete_index = (
        df_zzsete
        .groupby("ZZ7_CODPAI")["ZZ7_PRODUT"]
        .agg(lambda x: frozenset(x.str.upper().unique()))
        .to_dict()
    )

    def buscar_kit_exato(mats: frozenset) -> str | None:
        for codpai, componentes in zzsete_index.items():
            if componentes == mats:
                return codpai
        return None

    def buscar_melhor_match(mats: frozenset) -> tuple[str | None, int]:
        melhor, melhor_n = None, 0
        for codpai, componentes in zzsete_index.items():
            n = len(componentes & mats)
            if n > melhor_n:
                melhor_n = n
                melhor   = codpai
        return melhor, melhor_n

    df_com_kit = df_po[
        df_po["Numero_KIT"].notna()
        & (df_po["Numero_KIT"].str.strip() != "")
        & (df_po["Numero_KIT"].str.lower() != "nan")
    ].copy()

    if df_com_kit.empty:
        return pd.DataFrame([{
            "Numero_KIT":     "—",
            "kit_zzsete":     "—",
            "kit_sugerido":   "—",
            "material_po":    "—",
            "nome_material":  "—",
            "qtd_po":         None,
            "componente_kit": "—",
            "check_kit":      "⚠️ Nenhum Numero_KIT encontrado no pedido",
            "check_qtd_kit":  "—",
        }])

    df_com_kit["_kit_norm"] = df_com_kit["Numero_KIT"].apply(normalizar_kit)
    df_com_kit["Material"]  = df_com_kit["Material"].astype(str).str.strip().str.upper()

    rows = []

    for kit_id, grupo in df_com_kit.groupby("_kit_norm"):
        materiais_po     = set(grupo["Material"].unique())
        materiais_frozen = frozenset(materiais_po)

        # ── Kit identificado pela composição completa de materiais ─────────────
        kit_por_composicao = buscar_kit_exato(materiais_frozen)

        if kit_por_composicao is not None:
            # Composição exata encontrada
            componentes_ref = zzsete_index[kit_por_composicao]
            if kit_por_composicao == kit_id:
                check_kit    = "✅ Composição OK"
                kit_sugerido = "—"
            else:
                check_kit    = (
                    f"⚠️ Composição OK mas Numero_KIT divergente — "
                    f"Pedido: {kit_id} | zzsete: {kit_por_composicao}"
                )
                kit_sugerido = kit_por_composicao
            kit_ref = kit_por_composicao

        else:
            melhor_match, melhor_n = buscar_melhor_match(materiais_frozen)
            kit_ref         = melhor_match or "—"
            componentes_ref = zzsete_index.get(melhor_match, set())

            if melhor_match and materiais_frozen.issubset(componentes_ref):
                faltando     = sorted(componentes_ref - materiais_frozen)
                check_kit    = (
                    f"⚠️ Kit incompleto no pedido — possível match: {melhor_match} "
                    f"| Faltando: {', '.join(faltando)}"
                )
                kit_sugerido = melhor_match
            else:
                check_kit    = (
                    f"❌ Composição não encontrada na zzsete "
                    f"(melhor match: {kit_ref} — "
                    f"{melhor_n}/{len(materiais_po)} componentes em comum)"
                )
                kit_sugerido = melhor_match or "—"

        # ── Mapa qtd → materiais com essa quantidade ──────────────────────────
        qtd_para_materiais: dict[float, set] = {}
        for _, linha in grupo.iterrows():
            qtd = linha.get("Qtd.pedido", None)
            mat = linha["Material"]
            if qtd is not None:
                try:
                    qtd_f = round(float(qtd), 4)
                    qtd_para_materiais.setdefault(qtd_f, set()).add(mat)
                except (ValueError, TypeError):
                    pass

        # ── Para cada grupo de quantidade, identifica o kit específico ────────
        # Isso evita usar componentes_ref do kit geral para validar subgrupos
        qtd_para_kit_especifico: dict[float, str | None] = {}
        qtd_para_comp_ref: dict[float, frozenset]        = {}
        for qtd_f, mats_qtd in qtd_para_materiais.items():
            kit_qtd = buscar_kit_exato(frozenset(mats_qtd))
            if kit_qtd:
                qtd_para_kit_especifico[qtd_f] = kit_qtd
                qtd_para_comp_ref[qtd_f]       = zzsete_index[kit_qtd]
            else:
                # Usa componentes_ref do kit geral como fallback
                qtd_para_kit_especifico[qtd_f] = None
                qtd_para_comp_ref[qtd_f]       = componentes_ref

        # ── Gera linhas por componente ────────────────────────────────────────
        for _, linha in grupo.iterrows():
            mat = linha["Material"]
            qtd = linha.get("Qtd.pedido", None)

            try:
                qtd_f = round(float(qtd), 4) if qtd is not None else None
            except (ValueError, TypeError):
                qtd_f = None

            if qtd_f is not None:
                mats_com_essa_qtd  = qtd_para_materiais.get(qtd_f, set())
                comp_ref_qtd       = qtd_para_comp_ref.get(qtd_f, componentes_ref)
                kit_especifico_qtd = qtd_para_kit_especifico.get(qtd_f)
                faltando_nessa_qtd = comp_ref_qtd - mats_com_essa_qtd

                if not faltando_nessa_qtd:
                    # Kit completo para essa quantidade
                    if kit_especifico_qtd and kit_especifico_qtd != kit_id:
                        check_qtd_kit = (
                            f"⚠️ Kit completo para qtd {qtd_f:.0f} "
                            f"mas Numero_KIT divergente — "
                            f"Pedido: {kit_id} | zzsete: {kit_especifico_qtd}"
                        )
                    else:
                        check_qtd_kit = f"✅ Kit completo para qtd {qtd_f:.0f}"
                else:
                    check_qtd_kit = (
                        f"⚠️ Kit incompleto para qtd {qtd_f:.0f} "
                        f"— faltando: {', '.join(sorted(faltando_nessa_qtd))}"
                    )
            else:
                check_qtd_kit = "⚠️ Qtd não comparável"

            rows.append({
                "Numero_KIT":     kit_id,
                "kit_zzsete":     kit_ref,
                "kit_sugerido":   kit_sugerido,
                "material_po":    mat,
                "nome_material":  str(linha.get("Nome do material", "")).strip(),
                "qtd_po":         qtd,
                "componente_kit": "✅" if mat in componentes_ref else "⚠️ não está no kit",
                "check_kit":      check_kit,
                "check_qtd_kit":  check_qtd_kit,
            })

        # Componentes ausentes no pedido (apenas quando há match parcial)
        ausentes_po = componentes_ref - materiais_po
        for comp in sorted(ausentes_po):
            rows.append({
                "Numero_KIT":     kit_id,
                "kit_zzsete":     kit_ref,
                "kit_sugerido":   kit_sugerido,
                "material_po":    "—",
                "nome_material":  "—",
                "qtd_po":         None,
                "componente_kit": f"❌ ausente no pedido: {comp}",
                "check_kit":      check_kit,
                "check_qtd_kit":  "❌ Componente ausente no pedido",
            })

    return pd.DataFrame(rows)


# ─── Check Kits na NF-e (Aba 4) ──────────────────────────────────────────────

def check_kits_nfe(df_nfe: pd.DataFrame, df_zzsete: pd.DataFrame) -> pd.DataFrame:
    """
    Para cada item da NF-e:
      1. Consulta ZZ7_SRINK na zzsete pelo cProd (ZZ7_PRODUT)
         - Se ZZ7_SRINK == 'N' → Avulso OK, encerra para esse item
         - Se ZZ7_SRINK == 'S' → continua validação de kit
         - Se não encontrado na zzsete → Avulso (não cadastrado)
      2. Agrupa por NF e verifica se os itens com SRINK=S formam kit completo,
         parcial ou têm componentes faltando
    """
    if df_zzsete.empty:
        return pd.DataFrame([{
            "numero_nf":        "—",
            "arquivo":          "—",
            "n_item":           "—",
            "cProd":            "—",
            "xProd":            "—",
            "qCom":             None,
            "vUnCom":           None,
            "vProd":            None,
            "xPed":             "",
            "nItemPed":         "",
            "ZZ7_SRINK":        "—",
            "kit_identificado": "—",
            "status_kit":       "⚠️ Base zzsete não disponível",
            "observacao":       "",
            "_eh_faltante":     False,
        }])

    # Índice zzsete por produto: ZZ7_PRODUT → {ZZ7_CODPAI, ZZ7_SRINK}
    # Um produto pode aparecer em múltiplos kits — pegamos o primeiro match
    srink_por_produto: dict[str, str] = {}
    for _, row in df_zzsete.iterrows():
        prod = str(row.get("ZZ7_PRODUT", "")).strip().upper()
        if prod and prod != "NAN":
            srink = str(row.get("ZZ7_SRINK", "N")).strip().upper()
            if prod not in srink_por_produto:
                srink_por_produto[prod] = srink

    # Índice zzsete: ZZ7_CODPAI → frozenset(ZZ7_PRODUT)
    zzsete_index: dict[str, frozenset] = (
        df_zzsete[df_zzsete["ZZ7_PRODUT"].notna() & (df_zzsete["ZZ7_PRODUT"] != "NAN")]
        .groupby("ZZ7_CODPAI")["ZZ7_PRODUT"]
        .agg(lambda x: frozenset(x.str.upper().unique()))
        .to_dict()
    )

    # Índice reverso: ZZ7_PRODUT → lista de ZZ7_CODPAI
    produto_para_kits: dict[str, list] = {}
    for codpai, produtos in zzsete_index.items():
        for prod in produtos:
            produto_para_kits.setdefault(prod, []).append(codpai)

    rows = []

    for (arquivo, numero_nf), grupo_nf in df_nfe.groupby(["arquivo", "numero_nf"]):

        # Separa itens por SRINK
        itens_srink_s: dict[str, dict] = {}   # cProd → info, apenas SRINK=S
        itens_srink_n: list            = []    # linhas com SRINK=N ou não cadastrado

        for _, row in grupo_nf.iterrows():
            cp    = str(row["cProd"]).strip().upper()
            srink = srink_por_produto.get(cp, None)

            info = {
                "n_item":   row.get("n_item",   ""),
                "xProd":    row.get("xProd",    ""),
                "qCom":     row.get("qCom",     None),
                "vUnCom":   row.get("vUnCom",   None),
                "vProd":    row.get("vProd",    None),
                "xPed":     row.get("xPed",     ""),
                "nItemPed": row.get("nItemPed", ""),
                "ZZ7_SRINK": srink if srink else "Não cadastrado",
            }

            if srink == "S":
                if cp not in itens_srink_s:
                    itens_srink_s[cp] = info
            else:
                # SRINK=N ou produto não encontrado na zzsete → avulso OK
                itens_srink_n.append((cp, info, srink))

        # ── Linhas para itens SRINK=N ou não cadastrado ───────────────────────
        for cp, info, srink in itens_srink_n:
            if srink == "N":
                status = "⚪ Avulso - OK"
                observ = "Produto não requer composição de kit (ZZ7_SRINK=N)"
            else:
                status = "⚪ Avulso - não cadastrado na zzsete"
                observ = "Produto não encontrado na base zzsete"

            rows.append({
                "numero_nf":        numero_nf,
                "arquivo":          arquivo,
                "n_item":           info["n_item"],
                "cProd":            cp,
                "xProd":            info["xProd"],
                "qCom":             info["qCom"],
                "vUnCom":           info["vUnCom"],
                "vProd":            info["vProd"],
                "xPed":             info["xPed"],
                "nItemPed":         info["nItemPed"],
                "ZZ7_SRINK":        info["ZZ7_SRINK"],
                "kit_identificado": "—",
                "status_kit":       status,
                "observacao":       observ,
                "_eh_faltante":     False,
            })

        # ── Validação de kit para itens SRINK=S ───────────────────────────────
        if not itens_srink_s:
            continue

        produtos_nf_s = frozenset(itens_srink_s.keys())

        # Kits candidatos: qualquer kit que contenha pelo menos 1 produto da NF
        kits_candidatos: dict[str, frozenset] = {}
        for cp in produtos_nf_s:
            for kit in produto_para_kits.get(cp, []):
                kits_candidatos[kit] = zzsete_index[kit]

        kits_completos = {k: v for k, v in kits_candidatos.items() if v.issubset(produtos_nf_s)}
        kits_parciais  = {k: v for k, v in kits_candidatos.items() if k not in kits_completos}

        # Mapeia cada cProd ao melhor kit
        cprod_kit_map: dict[str, str] = {}
        for cp in produtos_nf_s:
            kits_cp_comp = [k for k, v in kits_completos.items() if cp in v]
            kits_cp_parc = [k for k, v in kits_parciais.items()  if cp in v]
            if kits_cp_comp:
                cprod_kit_map[cp] = kits_cp_comp[0]
            elif kits_cp_parc:
                cprod_kit_map[cp] = kits_cp_parc[0]
            else:
                cprod_kit_map[cp] = "SEM_KIT"

        # Linhas dos itens SRINK=S faturados
        itens_processados = set()
        for cp, info in itens_srink_s.items():
            if cp in itens_processados:
                continue
            itens_processados.add(cp)

            kit_id = cprod_kit_map.get(cp, "SEM_KIT")

            if kit_id == "SEM_KIT":
                status = "⚠️ Avulso - requer kit mas sem match na zzsete"
                observ = "Produto tem ZZ7_SRINK=S mas não forma nenhum kit reconhecido"
            elif kit_id in kits_completos:
                status = "✅ Kit completo"
                observ = f"Kit {kit_id} — todos os componentes faturados"
            else:
                faltando = sorted(zzsete_index[kit_id] - produtos_nf_s)
                status   = "⚠️ Possível falta de componentes"
                observ   = f"Kit {kit_id} — faltando: {', '.join(faltando)}"

            rows.append({
                "numero_nf":        numero_nf,
                "arquivo":          arquivo,
                "n_item":           info["n_item"],
                "cProd":            cp,
                "xProd":            info["xProd"],
                "qCom":             info["qCom"],
                "vUnCom":           info["vUnCom"],
                "vProd":            info["vProd"],
                "xPed":             info["xPed"],
                "nItemPed":         info["nItemPed"],
                "ZZ7_SRINK":        "S",
                "kit_identificado": kit_id if kit_id != "SEM_KIT" else "—",
                "status_kit":       status,
                "observacao":       observ,
                "_eh_faltante":     False,
            })

        # Linhas adicionais para componentes faltando
        kits_ja_sinalizados = set()
        for cp, kit_id in cprod_kit_map.items():
            if kit_id in kits_parciais and kit_id not in kits_ja_sinalizados:
                kits_ja_sinalizados.add(kit_id)
                faltando = sorted(zzsete_index[kit_id] - produtos_nf_s)
                for comp in faltando:
                    rows.append({
                        "numero_nf":        numero_nf,
                        "arquivo":          arquivo,
                        "n_item":           "—",
                        "cProd":            comp,
                        "xProd":            "— COMPONENTE NÃO FATURADO —",
                        "qCom":             None,
                        "vUnCom":           None,
                        "vProd":            None,
                        "xPed":             "",
                        "nItemPed":         "",
                        "ZZ7_SRINK":        "S",
                        "kit_identificado": kit_id,
                        "status_kit":       "❌ Componente não faturado",
                        "observacao":       f"Componente do kit {kit_id} ausente na NF-e",
                        "_eh_faltante":     True,
                    })

    return pd.DataFrame(rows)


# ─── Layout ───────────────────────────────────────────────────────────────────

st.set_page_config(page_title="NF-e Parser", layout="wide")
st.title("NF-e Parser")

sku_set, sku_df = carregar_sku_set()
df_zzsete       = carregar_zzsete()

if sku_set:
    st.caption(f"Tabela de SKUs carregada: {len(sku_set)} produtos ({SKU_FILE.name})")
else:
    st.warning("Arquivo skuProd.csv não encontrado. A coluna cProd será mantida como veio no XML.")

if not df_zzsete.empty:
    st.caption(
        f"Base de Kits (zzsete) carregada: "
        f"{df_zzsete['ZZ7_CODPAI'].nunique()} kits / {len(df_zzsete)} componentes."
    )
else:
    st.warning("Arquivo zzsete.xlsx não encontrado. As abas 3 e 4 não estarão disponíveis.")

aba_parser, aba_po, aba_kit, aba_nfe_kit = st.tabs([
    "📄 Parser NF-e",
    "🔍 Comparativo NF-e x PO",
    "📦 Check Composição de Kits",
    "🧩 Check Kits na NF-e",
])


# ══════════════════════════════════════════════════════════════════════════════
# ABA 1 — Parser NF-e
# ══════════════════════════════════════════════════════════════════════════════

with aba_parser:
    st.subheader("Importar XMLs de NF-e")

    uploaded_files = st.file_uploader(
        "Selecione os arquivos XML",
        type=["xml"],
        accept_multiple_files=True,
        key="xml_parser",
    )

    if uploaded_files:
        if st.button("Processar", type="primary", key="btn_parser"):
            all_rows, erros = [], []

            with st.spinner("Processando arquivos..."):
                for f in uploaded_files:
                    try:
                        rows = parse_nfe_xml(f.read(), filename=f.name)
                        all_rows.extend(rows)
                    except Exception as e:
                        erros.append(f"{f.name}: {e}")

            for erro in erros:
                st.error(f"Erro: {erro}")

            if all_rows:
                df = pd.DataFrame(all_rows)

                if sku_set:
                    df = resolver_skus(df, sku_set)
                    nao_resolvidos = (df["sku_origem"] == "não resolvido").sum()
                    if nao_resolvidos:
                        with st.expander(f"⚠️ {nao_resolvidos} item(ns) sem SKU resolvido"):
                            st.dataframe(
                                df[df["sku_origem"] == "não resolvido"][["n_item", "cProd", "xProd"]],
                                use_container_width=True,
                            )

                divergentes_xped = check_unicidade_xped(df)
                if not divergentes_xped.empty:
                    with st.expander(
                        f"⚠️ {len(divergentes_xped)} NF(s) com mais de um xPed distinto",
                        expanded=True,
                    ):
                        st.warning(
                            "As NFs abaixo possuem itens referenciando pedidos de compra diferentes. "
                            "Verifique se isso é esperado antes de prosseguir."
                        )
                        st.dataframe(divergentes_xped, use_container_width=True)
                else:
                    st.success("✅ xPed único em todas as NFs processadas.")

                st.success(
                    f"{len(uploaded_files) - len(erros)} arquivo(s) processado(s) "
                    f"— {len(df)} linha(s) extraída(s)."
                )
                st.dataframe(df, use_container_width=True)

                st.session_state["df_nfe"] = df

                csv_buffer = io.StringIO()
                df.to_csv(csv_buffer, index=False, encoding="utf-8-sig")

                excel_buffer = io.BytesIO()
                df.to_excel(excel_buffer, index=False, engine="openpyxl")

                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="⬇️ Baixar CSV",
                        data=csv_buffer.getvalue().encode("utf-8-sig"),
                        file_name="nfe_itens.csv",
                        mime="text/csv",
                    )
                with col2:
                    st.download_button(
                        label="⬇️ Baixar Excel",
                        data=excel_buffer.getvalue(),
                        file_name="nfe_itens.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


# ══════════════════════════════════════════════════════════════════════════════
# ABA 2 — Comparativo NF-e x PO
# ══════════════════════════════════════════════════════════════════════════════

with aba_po:
    st.subheader("Comparativo NF-e × Pedido de Compras")

    po_file = st.file_uploader(
        "Upload do arquivo pedidosCompra (.xlsx ou .csv)",
        type=["xlsx", "csv"],
        key="po_file",
    )

    st.divider()

    usar_nfe_existente = (
        "df_nfe" in st.session_state
        and not st.session_state["df_nfe"].empty
    )

    if usar_nfe_existente:
        st.info(
            f"Usando NF-e já processada na Aba 1 "
            f"({len(st.session_state['df_nfe'])} linhas). "
            "Ou faça um novo upload abaixo para substituir."
        )

    xml_files_po = st.file_uploader(
        "Upload de XMLs de NF-e (opcional se já processou na Aba 1)",
        type=["xml"],
        accept_multiple_files=True,
        key="xml_po",
    )

    if st.button("Comparar", type="primary", key="btn_po"):

        if xml_files_po:
            all_rows, erros = [], []
            for f in xml_files_po:
                try:
                    rows = parse_nfe_xml(f.read(), filename=f.name)
                    all_rows.extend(rows)
                except Exception as e:
                    st.error(f"Erro ao processar {f.name}: {e}")
            if all_rows:
                df_nfe = pd.DataFrame(all_rows)
                if sku_set:
                    df_nfe = resolver_skus(df_nfe, sku_set)
            else:
                st.error("Nenhum XML válido processado.")
                st.stop()
        elif usar_nfe_existente:
            df_nfe = st.session_state["df_nfe"]
        else:
            st.error("Faça o upload dos XMLs de NF-e na Aba 1 ou aqui.")
            st.stop()

        if not po_file:
            st.error("Faça o upload do arquivo pedidosCompra.")
            st.stop()

        with st.spinner("Carregando pedidos e comparando..."):
            df_po     = carregar_pedidos(po_file)
            df_result = comparar_nfe_po(df_nfe, df_po)

        divergentes_xped = check_unicidade_xped(df_nfe)
        if not divergentes_xped.empty:
            with st.expander(
                f"⚠️ {len(divergentes_xped)} NF(s) com mais de um xPed distinto",
                expanded=True,
            ):
                st.warning(
                    "As NFs abaixo possuem itens referenciando pedidos de compra diferentes. "
                    "O comparativo pode conter resultados mistos para essas NFs."
                )
                st.dataframe(divergentes_xped, use_container_width=True)
        else:
            st.success("✅ xPed único em todas as NFs.")

        st.divider()

        total       = len(df_result)
        ok_material = df_result["check_material"].str.startswith("✅").sum()
        ok_qtd      = df_result["check_qtd"].str.startswith("✅").sum()
        ok_vunit    = df_result["check_vunit"].str.startswith("✅").sum()
        sem_pedido  = df_result["check_pedido"].str.startswith("❌").sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Itens na NF-e",          total)
        c2.metric("Pedidos não encontrados", sem_pedido)
        c3.metric("Material OK",             f"{ok_material}/{total}")
        c4.metric("Qtd OK",                  f"{ok_qtd}/{total}")

        st.divider()

        filtro = st.radio(
            "Mostrar:",
            ["Todos", "Somente divergências", "Somente OK"],
            horizontal=True,
        )

        if filtro == "Somente divergências":
            mask = (
                df_result["check_pedido"].str.startswith("❌")
                | df_result["check_material"].str.startswith("❌")
                | df_result["check_qtd"].str.startswith("❌")
                | df_result["check_vunit"].str.startswith("❌")
            )
            df_view = df_result[mask]
        elif filtro == "Somente OK":
            mask = (
                df_result["check_pedido"].str.startswith("✅")
                & df_result["check_material"].str.startswith("✅")
                & df_result["check_qtd"].str.startswith("✅")
                & df_result["check_vunit"].str.startswith("✅")
            )
            df_view = df_result[mask]
        else:
            df_view = df_result

        st.dataframe(df_view, use_container_width=True)

        csv_buf = io.StringIO()
        df_result.to_csv(csv_buf, index=False, encoding="utf-8-sig")

        xls_buf = io.BytesIO()
        df_result.to_excel(xls_buf, index=False, engine="openpyxl")

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇️ Baixar CSV comparativo",
                data=csv_buf.getvalue().encode("utf-8-sig"),
                file_name="comparativo_nfe_po.csv",
                mime="text/csv",
            )
        with col2:
            st.download_button(
                label="⬇️ Baixar Excel comparativo",
                data=xls_buf.getvalue(),
                file_name="comparativo_nfe_po.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ══════════════════════════════════════════════════════════════════════════════
# ABA 3 — Check Composição de Kits
# ══════════════════════════════════════════════════════════════════════════════

with aba_kit:
    st.subheader("Check Composição de Kits — Pedido x zzsete")
    st.write(
        "Verifica se os itens agrupados por `Numero_KIT` no pedido de compras "
        "correspondem a um kit cadastrado na base zzsete."
    )

    po_file_kit = st.file_uploader(
        "Upload do arquivo pedidosCompra (.xlsx ou .csv)",
        type=["xlsx", "csv"],
        key="po_file_kit",
    )

    if st.button("Verificar Kits", type="primary", key="btn_kit"):

        if not po_file_kit:
            st.error("Faça o upload do arquivo pedidosCompra.")
            st.stop()

        if df_zzsete.empty:
            st.error("Base zzsete.xlsx não encontrada na raiz do projeto.")
            st.stop()

        with st.spinner("Verificando composição dos kits..."):
            df_po_kit  = carregar_pedidos(po_file_kit)
            df_kit_res = check_composicao_kits(df_po_kit, df_zzsete)

        kits_unicos = df_kit_res["Numero_KIT"].nunique()
        ok_kits     = df_kit_res.groupby("Numero_KIT")["check_kit"].first().str.startswith("✅").sum()
        err_kits    = df_kit_res.groupby("Numero_KIT")["check_kit"].first().str.startswith("❌").sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("Kits no pedido",  kits_unicos)
        c2.metric("Composição OK",   ok_kits)
        c3.metric("Com divergência", err_kits)

        st.divider()

        filtro_kit = st.radio(
            "Mostrar:",
            ["Todos", "Somente divergências", "Somente OK"],
            horizontal=True,
            key="filtro_kit",
        )

        if filtro_kit == "Somente divergências":
            kits_div    = df_kit_res[df_kit_res["check_kit"].str.startswith("❌")]["Numero_KIT"].unique()
            df_kit_view = df_kit_res[df_kit_res["Numero_KIT"].isin(kits_div)]
        elif filtro_kit == "Somente OK":
            kits_ok     = df_kit_res[df_kit_res["check_kit"].str.startswith("✅")]["Numero_KIT"].unique()
            df_kit_view = df_kit_res[df_kit_res["Numero_KIT"].isin(kits_ok)]
        else:
            df_kit_view = df_kit_res

        st.dataframe(df_kit_view, use_container_width=True)

        csv_buf_kit = io.StringIO()
        df_kit_res.to_csv(csv_buf_kit, index=False, encoding="utf-8-sig")

        xls_buf_kit = io.BytesIO()
        df_kit_res.to_excel(xls_buf_kit, index=False, engine="openpyxl")

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇️ Baixar CSV kits",
                data=csv_buf_kit.getvalue().encode("utf-8-sig"),
                file_name="check_kits.csv",
                mime="text/csv",
            )
        with col2:
            st.download_button(
                label="⬇️ Baixar Excel kits",
                data=xls_buf_kit.getvalue(),
                file_name="check_kits.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ══════════════════════════════════════════════════════════════════════════════
# ABA 4 — Check Kits na NF-e
# ══════════════════════════════════════════════════════════════════════════════

with aba_nfe_kit:
    st.subheader("Check Kits na NF-e")
    st.write(
        "Analisa os itens faturados na NF-e. Consulta `ZZ7_SRINK` na zzsete: "
        "se `N` → avulso OK. Se `S` → verifica se os componentes formam um kit completo."
    )

    usar_nfe_aba4 = (
        "df_nfe" in st.session_state
        and not st.session_state["df_nfe"].empty
    )

    if usar_nfe_aba4:
        st.info(
            f"Usando NF-e já processada na Aba 1 "
            f"({len(st.session_state['df_nfe'])} linhas). "
            "Ou faça um novo upload abaixo para substituir."
        )

    xml_files_aba4 = st.file_uploader(
        "Upload de XMLs de NF-e (opcional se já processou na Aba 1)",
        type=["xml"],
        accept_multiple_files=True,
        key="xml_aba4",
    )

    if st.button("Analisar Kits na NF-e", type="primary", key="btn_aba4"):

        if df_zzsete.empty:
            st.error("Base zzsete.xlsx não encontrada na raiz do projeto.")
            st.stop()

        if xml_files_aba4:
            all_rows, erros = [], []
            for f in xml_files_aba4:
                try:
                    rows = parse_nfe_xml(f.read(), filename=f.name)
                    all_rows.extend(rows)
                except Exception as e:
                    st.error(f"Erro ao processar {f.name}: {e}")
            if all_rows:
                df_nfe_aba4 = pd.DataFrame(all_rows)
                if sku_set:
                    df_nfe_aba4 = resolver_skus(df_nfe_aba4, sku_set)
            else:
                st.error("Nenhum XML válido processado.")
                st.stop()
        elif usar_nfe_aba4:
            df_nfe_aba4 = st.session_state["df_nfe"]
        else:
            st.error("Faça o upload dos XMLs de NF-e na Aba 1 ou aqui.")
            st.stop()

        with st.spinner("Analisando kits na NF-e..."):
            df_aba4 = check_kits_nfe(df_nfe_aba4, df_zzsete)

        df_itens  = df_aba4[~df_aba4["_eh_faltante"]]
        kits_ok   = (df_itens["status_kit"] == "✅ Kit completo").sum()
        kits_parc = (df_itens["status_kit"] == "⚠️ Possível falta de componentes").sum()
        avulsos_ok = (df_itens["status_kit"] == "⚪ Avulso - OK").sum()
        avulsos_nc = (df_itens["status_kit"] == "⚪ Avulso - não cadastrado na zzsete").sum()
        faltantes  = df_aba4["_eh_faltante"].sum()

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Kits completos",           kits_ok)
        c2.metric("Kits com falta",           kits_parc)
        c3.metric("Avulsos OK (SRINK=N)",     avulsos_ok)
        c4.metric("Avulsos não cadastrados",  avulsos_nc)
        c5.metric("Componentes não faturados",faltantes)

        st.divider()

        filtro_aba4 = st.radio(
            "Mostrar:",
            ["Todos", "Somente problemas", "Somente kits completos", "Somente avulsos OK"],
            horizontal=True,
            key="filtro_aba4",
        )

        if filtro_aba4 == "Somente problemas":
            nfs_prob = df_aba4[
                df_aba4["status_kit"].isin([
                    "⚠️ Possível falta de componentes",
                    "❌ Componente não faturado",
                    "⚠️ Avulso - requer kit mas sem match na zzsete",
                ])
            ]["numero_nf"].unique()
            df_view4 = df_aba4[df_aba4["numero_nf"].isin(nfs_prob)]
        elif filtro_aba4 == "Somente kits completos":
            nfs_ok   = df_itens[df_itens["status_kit"] == "✅ Kit completo"]["numero_nf"].unique()
            df_view4 = df_aba4[df_aba4["numero_nf"].isin(nfs_ok) & ~df_aba4["_eh_faltante"]]
        elif filtro_aba4 == "Somente avulsos OK":
            df_view4 = df_aba4[df_aba4["status_kit"] == "⚪ Avulso - OK"]
        else:
            df_view4 = df_aba4

        colunas_exibir = [c for c in df_view4.columns if c != "_eh_faltante"]
        st.dataframe(df_view4[colunas_exibir], use_container_width=True)

        csv_buf4 = io.StringIO()
        df_aba4[colunas_exibir].to_csv(csv_buf4, index=False, encoding="utf-8-sig")

        xls_buf4 = io.BytesIO()
        df_aba4[colunas_exibir].to_excel(xls_buf4, index=False, engine="openpyxl")

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇️ Baixar CSV",
                data=csv_buf4.getvalue().encode("utf-8-sig"),
                file_name="check_kits_nfe.csv",
                mime="text/csv",
            )
        with col2:
            st.download_button(
                label="⬇️ Baixar Excel",
                data=xls_buf4.getvalue(),
                file_name="check_kits_nfe.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )