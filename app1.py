import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import dropbox

# =========================
# CONFIG
# =========================
ABA_ESTOQUE = "Estoque"
ABA_HIST = "Historico"

# O seu shared link (vem do secrets)
DEFAULT_SHARED_LINK = "https://www.dropbox.com/scl/fi/y9cqb8png4s9rnvhi9r0h/estoque_base.xlsx?rlkey=4cjlyvt0m5zzf0iy0n4ci6n1h&dl=0"

st.set_page_config(page_title="Controle de Estoque (Dropbox)", layout="wide")


def get_dbx() -> dropbox.Dropbox:
    """
    Cria o cliente Dropbox usando token guardado em st.secrets.
    """
    token = st.secrets["DROPBOX_ACCESS_TOKEN"]
    return dropbox.Dropbox(token)


def get_shared_link() -> str:
    """
    Pega o shared link do secrets; se não existir, usa o default (o que você mandou).
    """
    return st.secrets.get("DROPBOX_SHARED_LINK", DEFAULT_SHARED_LINK)


def baixar_excel_via_shared_link(dbx: dropbox.Dropbox, shared_link: str) -> tuple[bytes, str]:
    """
    Baixa o arquivo Excel usando o shared link e tenta descobrir o caminho real no Dropbox.
    Retorna:
    - bytes do arquivo
    - path do arquivo no Dropbox (para upload overwrite depois)
    """
    # 1) Pega metadados do shared link (aqui normalmente vem o path_lower/path_display)
    meta = dbx.sharing_get_shared_link_metadata(url=shared_link)

    # tenta obter um caminho “real” no Dropbox (ideal para sobrescrever depois)
    dropbox_path = None
    if hasattr(meta, "path_lower") and meta.path_lower:
        dropbox_path = meta.path_lower
    elif hasattr(meta, "path_display") and meta.path_display:
        dropbox_path = meta.path_display

    # 2) Baixa o conteúdo do arquivo do shared link
    _, res = dbx.sharing_get_shared_link_file(url=shared_link)
    content = res.content

    if not dropbox_path:
        # Se não vier path, ainda dá pra ler, mas pra salvar precisa de um caminho.
        # Nesse caso, vamos salvar em um caminho padrão do app (Apps/...) para garantir upload.
        # (Você pode mudar esse caminho depois se quiser.)
        dropbox_path = "/Apps/streamlit-estoque/estoque_base.xlsx"

    return content, dropbox_path


def ler_dados(excel_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Lê as abas do Excel em memória.
    Espera aba "Estoque" com colunas: Item, Quantidade
    Cria histórico vazio se não existir.
    """
    xls = pd.ExcelFile(BytesIO(excel_bytes), engine="openpyxl")

    if ABA_ESTOQUE not in xls.sheet_names:
        raise ValueError(f"Não encontrei a aba '{ABA_ESTOQUE}' no Excel.")

    df = pd.read_excel(xls, sheet_name=ABA_ESTOQUE)
    # normaliza
    if "Item" not in df.columns or "Quantidade" not in df.columns:
        raise ValueError("A aba 'Estoque' precisa ter colunas: Item e Quantidade.")

    df["Item"] = df["Item"].astype(str).str.strip()
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0).astype(int)

    if ABA_HIST in xls.sheet_names:
        hist = pd.read_excel(xls, sheet_name=ABA_HIST)
        # garante colunas
        for c in ["Data", "Item", "Movimento", "Quantidade", "Estoque_Apos"]:
            if c not in hist.columns:
                hist[c] = None
        hist = hist[["Data", "Item", "Movimento", "Quantidade", "Estoque_Apos"]].copy()
    else:
        hist = pd.DataFrame(columns=["Data", "Item", "Movimento", "Quantidade", "Estoque_Apos"])

    # normaliza histórico
    if not hist.empty:
        hist["Item"] = hist["Item"].fillna("").astype(str).str.strip()
        hist["Movimento"] = hist["Movimento"].fillna("").astype(str).str.strip().str.upper()
        hist["Quantidade"] = pd.to_numeric(hist["Quantidade"], errors="coerce").fillna(0).astype(int)
        hist["Estoque_Apos"] = pd.to_numeric(hist["Estoque_Apos"], errors="coerce").fillna(0).astype(int)

    return df, hist


def gerar_excel_bytes(df: pd.DataFrame, hist: pd.DataFrame) -> bytes:
    """
    Gera um XLSX em memória com abas Estoque e Historico.
    """
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.sort_values("Item").to_excel(writer, sheet_name=ABA_ESTOQUE, index=False)
        hist.to_excel(writer, sheet_name=ABA_HIST, index=False)
    buffer.seek(0)
    return buffer.getvalue()


def registrar_movimento(hist: pd.DataFrame, item: str, movimento: str, qtd: int, estoque_apos: int) -> pd.DataFrame:
    """
    Adiciona uma linha ao histórico.
    """
    linha = {
        "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Item": item,
        "Movimento": movimento,
        "Quantidade": int(qtd),
        "Estoque_Apos": int(estoque_apos),
    }
    return pd.concat([hist, pd.DataFrame([linha])], ignore_index=True)


def upload_overwrite(dbx: dropbox.Dropbox, dropbox_path: str, excel_bytes: bytes):
    """
    Sobrescreve o arquivo no Dropbox.
    """
    dbx.files_upload(
        excel_bytes,
        dropbox_path,
        mode=dropbox.files.WriteMode.overwrite,
        mute=True,
    )


# =========================
# APP
# =========================

st.title("📦 Controle de Estoque (Dropbox)")

dbx = get_dbx()
shared_link = get_shared_link()

# Baixa SEMPRE a versão mais recente do Dropbox
try:
    original_bytes, DROPBOX_FILE_PATH = baixar_excel_via_shared_link(dbx, shared_link)
    df, hist = ler_dados(original_bytes)
except Exception as e:
    st.error("Não consegui abrir o Excel do Dropbox. Verifique o token e o link compartilhado.")
    st.exception(e)
    st.stop()

# Exportar
st.subheader("⬇️ Exportar")
st.download_button(
    "Baixar Excel atualizado",
    data=gerar_excel_bytes(df, hist),
    file_name="estoque_atual.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.caption(f"Arquivo no Dropbox: `{DROPBOX_FILE_PATH}`")
st.divider()

# Sidebar: movimentação
with st.sidebar:
    st.header("⚙️ Entrada / Saída")

    item_sel = st.selectbox("Item", df["Item"].tolist())
    movimento = st.radio("Movimento", ["ENTRADA", "SAIDA"], horizontal=True)
    qtd = st.number_input("Quantidade", min_value=1, step=1, value=1)

    if st.button("Aplicar movimento", use_container_width=True):
        atual = int(df.loc[df["Item"] == item_sel, "Quantidade"].iloc[0])

        if movimento == "SAIDA" and qtd > atual:
            st.error(f"Estoque insuficiente. Atual: {atual}")
        else:
            novo = atual + qtd if movimento == "ENTRADA" else atual - qtd
            df.loc[df["Item"] == item_sel, "Quantidade"] = novo
            hist = registrar_movimento(hist, item_sel, movimento, int(qtd), novo)

            # gera excel e sobe pro Dropbox
            novo_excel = gerar_excel_bytes(df, hist)
            try:
                upload_overwrite(dbx, DROPBOX_FILE_PATH, novo_excel)
                st.success("✅ Atualizado no Dropbox!")
                st.rerun()
            except Exception as e:
                st.error("Atualizei no app, mas falhei ao salvar no Dropbox.")
                st.exception(e)

# Dashboard
st.subheader("📊 Dashboard do item")

qtd_atual = int(df.loc[df["Item"] == item_sel, "Quantidade"].iloc[0])

entradas = hist[(hist["Item"] == item_sel) & (hist["Movimento"] == "ENTRADA")]["Quantidade"].sum()
saidas = hist[(hist["Item"] == item_sel) & (hist["Movimento"] == "SAIDA")]["Quantidade"].sum()

c1, c2, c3 = st.columns(3)
c1.metric("Quantidade atual", qtd_atual)
c2.metric("Total entradas", int(entradas))
c3.metric("Total saídas", int(saidas))

st.caption("Histórico do item")
st.dataframe(
    hist[hist["Item"] == item_sel].sort_values("Data", ascending=False),
    use_container_width=True,
    hide_index=True
)

st.subheader("📋 Estoque completo")
st.dataframe(df.sort_values("Item"), use_container_width=True, hide_index=True)
