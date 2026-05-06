import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import re
import io
from pathlib import Path

NAMESPACE   = {"ns": "http://www.portalfiscal.inf.br/nfe"}
SKU_PATTERN = re.compile(r'\b[A-Z0-9]{13}\b')
SKU_FILE    = Path(__file__).parent / "skuProd.csv"


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
    """
    Procura dentro do texto xProd um token de 13 caracteres presente na tabela auxiliar.
    Substitui underscores por espaços antes de tokenizar para cobrir casos como:
    'Pedido 4500004291_Item 00010_ID PG26LA212SD00_26 LV AL EF1...'
    """
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
    """Remove zeros à esquerda e espaços para comparação segura entre NF-e e PO."""
    try:
        return str(int(str(valor).strip()))
    except (ValueError, TypeError):
        return str(valor).strip()

def parse_numero_br(valor_str):
    """Converte string numérica no padrão brasileiro (1.000,00) para float."""
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
    """
    Verifica se o campo xPed é único dentro de cada NF-e (agrupado por numero_nf).
    Retorna um DataFrame com as NFs que possuem mais de um xPed distinto,
    listando quais pedidos aparecem em cada uma.
    """
    resumo = (
        df[df["xPed"].str.strip() != ""]
        .groupby(["arquivo", "numero_nf"])["xPed"]
        .agg(lambda x: sorted(x.unique().tolist()))
        .reset_index()
        .rename(columns={"xPed": "xPed_valores"})
    )
    resumo["qtd_pedidos_distintos"] = resumo["xPed_valores"].apply(len)
    resumo["xPed_valores"]          = resumo["xPed_valores"].apply(lambda v: " | ".join(v))

    divergentes = resumo[resumo["qtd_pedidos_distintos"] > 1].copy()
    return divergentes


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

    for col in ["Pedido de compras", "Item", "Material"]:
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

        # 1. Localiza o pedido
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

        # 2. Localiza o item dentro do pedido
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

        # ── Check pedido ──────────────────────────────────────────────────────
        check_pedido = "✅ Pedido encontrado"

        # ── Check material ────────────────────────────────────────────────────
        check_material = (
            "✅ Material OK"
            if cprod == po_material
            else f"❌ Material divergente — NF: {cprod} | PO: {po_material}"
        )

        # ── Check quantidade ──────────────────────────────────────────────────
        # qCom == Qtd.pedido                                   → OK
        # qCom <  Qtd.pedido  E  Qtd. a fornecer == 0         → OK (entrega parcial)
        # qCom <  Qtd.pedido  E  qCom <= Qtd. a fornecer      → OK (entrega parcial)
        # qualquer outro caso                                  → ERRO
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

        # ── Check valor unitário ──────────────────────────────────────────────
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


# ─── Layout ───────────────────────────────────────────────────────────────────

st.set_page_config(page_title="NF-e Parser", layout="wide")
st.title("NF-e Parser")

sku_set, sku_df = carregar_sku_set()
if sku_set:
    st.caption(f"Tabela de SKUs carregada: {len(sku_set)} produtos ({SKU_FILE.name})")
else:
    st.warning("Arquivo skuProd.csv não encontrado. A coluna cProd será mantida como veio no XML.")

aba_parser, aba_po = st.tabs(["📄 Parser NF-e", "🔍 Comparativo NF-e x PO"])


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

                # ── Check unicidade xPed ──────────────────────────────────────
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

        # ── Check unicidade xPed na aba 2 ─────────────────────────────────────
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
            st.success("✅ Pedido único em todas as NFs.")

        st.divider()

        # ── Resumo ────────────────────────────────────────────────────────────
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