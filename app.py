#!/usr/bin/env python3
"""Interface Streamlit do Extrator de Recebiveis PDF para Excel."""

from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
from pathlib import Path

import streamlit as st

from extrair_recebiveis_gemini import (
    DEFAULT_MODEL,
    UserFacingError,
    build_workbook,
    extract_with_gemini,
    validate_pdf,
    verify_workbook,
)


APP_TITLE = "Extrator de Recebíveis"
MAX_MB = 10.0
MAX_PAGES = 10


def configure_page() -> None:
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="📊",
        layout="centered",
        initial_sidebar_state="auto",
    )
    st.markdown(
        """
        <style>
        :root {
          --fc-blue: #00308F;
          --fc-green: #6FC256;
          --fc-navy: #042653;
        }
        .stApp { background: #F5F7FB; }
        .block-container { max-width: 880px; padding-top: 2rem; }
        [data-testid="stHeader"] { background: transparent; }
        .hero {
          background: linear-gradient(135deg, var(--fc-navy), var(--fc-blue));
          border-radius: 18px;
          color: white;
          padding: 28px 30px;
          margin-bottom: 22px;
          box-shadow: 0 10px 28px rgba(0, 48, 143, .16);
        }
        .hero h1 { margin: 0 0 8px 0; font-size: 2rem; }
        .hero p { margin: 0; color: #E8EEF8; }
        .hero h1, .hero p { overflow-wrap: anywhere; }
        @media (max-width: 640px) {
          .block-container { padding: 1rem; }
          .hero { padding: 20px; }
          .hero h1 { font-size: 1.65rem; }
        }
        .privacy-note {
          border-left: 4px solid var(--fc-green);
          background: white;
          border-radius: 8px;
          padding: 12px 16px;
          color: #334155;
          margin: 12px 0 20px 0;
        }
        .stButton > button, .stDownloadButton > button {
          border-radius: 10px;
          min-height: 44px;
          font-weight: 650;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def configure_api_key() -> str:
    """Lê a credencial sem copiá-la para o ambiente global."""
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key:
        try:
            api_key = str(st.secrets["GEMINI_API_KEY"]).strip()
        except Exception:
            # Streamlit usa uma excecao propria quando ainda nao existe secrets.toml.
            api_key = ""
    if not api_key:
        raise UserFacingError(
            "A chave GEMINI_API_KEY não foi configurada nos Secrets do aplicativo."
        )
    return api_key


def configured_model() -> str:
    try:
        return str(st.secrets.get("GEMINI_MODEL", os.getenv("GEMINI_MODEL", DEFAULT_MODEL))).strip() or DEFAULT_MODEL
    except Exception:
        return os.getenv("GEMINI_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL


def validate_upload(file_name: str, content: bytes) -> None:
    if not file_name.lower().endswith(".pdf"):
        raise ValueError("Envie um arquivo com extensão .pdf.")
    if not content.startswith(b"%PDF-"):
        raise ValueError("O conteúdo enviado não possui a assinatura de um PDF válido.")
    size_mb = len(content) / (1024 * 1024)
    if size_mb > MAX_MB:
        raise ValueError(
            f"O PDF possui {size_mb:.2f} MB. O limite desta versão é {MAX_MB:.0f} MB."
        )


def clear_previous_result(file_digest: str) -> None:
    if st.session_state.get("file_digest") != file_digest:
        for key in ("xlsx_bytes", "xlsx_name", "json_bytes", "summary"):
            st.session_state.pop(key, None)
        st.session_state["file_digest"] = file_digest


def process_pdf(file_name: str, content: bytes) -> None:
    api_key = configure_api_key()
    raw_stem = Path(file_name).stem.strip() or "relatorio_recebiveis"
    safe_stem = re.sub(r"[^\w .-]", "_", raw_stem, flags=re.UNICODE).strip(" .")
    safe_stem = (safe_stem or "relatorio_recebiveis")[:100]

    with tempfile.TemporaryDirectory(prefix="recebiveis_") as temp_dir:
        work_dir = Path(temp_dir)
        pdf_path = work_dir / "entrada.pdf"
        output_path = work_dir / f"{safe_stem}_tabelas.xlsx"
        pdf_path.write_bytes(content)

        try:
            page_count = validate_pdf(pdf_path, max_pages=MAX_PAGES, max_mb=MAX_MB)
        except ValueError as exc:
            raise UserFacingError(str(exc)) from None
        extraction = extract_with_gemini(
            pdf_path=pdf_path,
            model=configured_model(),
            expected_pages=page_count,
            attempts=2,
            api_key=api_key,
        )
        # O nome temporario nao deve aparecer como fonte no Excel.
        extraction.source_file = file_name

        build_workbook(extraction, output_path, freeze=True)
        verify_workbook(output_path, len(extraction.tables))

        st.session_state["xlsx_bytes"] = output_path.read_bytes()
        st.session_state["xlsx_name"] = output_path.name
        st.session_state["json_bytes"] = json.dumps(
            extraction.model_dump(), ensure_ascii=False, indent=2
        ).encode("utf-8")
        st.session_state["summary"] = {
            "pages": page_count,
            "tables": len(extraction.tables),
            "warnings": extraction.warnings,
        }


def render_sidebar() -> None:
    with st.sidebar:
        st.markdown("### Como utilizar")
        st.markdown(
            """
            1. Selecione o relatório em PDF.
            2. Clique em **Extrair tabelas**.
            3. Aguarde o processamento.
            4. Baixe o Excel e, se necessário, o JSON de auditoria.
            """
        )
        st.divider()
        st.caption(f"Limites do piloto: {MAX_PAGES} páginas e {MAX_MB:.0f} MB por PDF.")
        st.caption("O sistema extrai somente o conteúdo visível no documento.")


def render_result() -> None:
    summary = st.session_state.get("summary")
    if not summary:
        return

    st.success("Extração concluída com sucesso.")
    col_pages, col_tables = st.columns(2)
    col_pages.metric("Páginas analisadas", summary["pages"])
    col_tables.metric("Tabelas encontradas", summary["tables"])

    for warning in summary["warnings"]:
        st.warning(warning)

    col_excel, col_json = st.columns(2)
    with col_excel:
        st.download_button(
            "Baixar Excel",
            data=st.session_state["xlsx_bytes"],
            file_name=st.session_state["xlsx_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with col_json:
        st.download_button(
            "Baixar auditoria JSON",
            data=st.session_state["json_bytes"],
            file_name=f"{Path(st.session_state['xlsx_name']).stem}_auditoria.json",
            mime="application/json",
            use_container_width=True,
        )


def main() -> None:
    configure_page()
    render_sidebar()

    st.markdown(
        """
        <section class="hero">
          <h1>Extrator de Recebíveis</h1>
          <p>Transforme as tabelas visíveis do relatório PDF em um Excel organizado.</p>
        </section>
        <div class="privacy-note">
          O PDF não é gravado permanentemente pela aplicação. Ele é enviado à
          Gemini API durante a extração e o arquivo temporário local é apagado ao final.
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Relatório de recebíveis",
        type=["pdf"],
        accept_multiple_files=False,
        help=f"PDF com até {MAX_PAGES} páginas e {MAX_MB:.0f} MB.",
    )

    if uploaded_file is None:
        clear_previous_result("")
        st.info("Selecione um PDF para iniciar.")
        return

    content = uploaded_file.getvalue()
    digest = hashlib.sha256(uploaded_file.name.encode() + content).hexdigest()
    clear_previous_result(digest)

    if st.button("Extrair tabelas", type="primary", use_container_width=True):
        for key in ("xlsx_bytes", "xlsx_name", "json_bytes", "summary"):
            st.session_state.pop(key, None)
        try:
            try:
                validate_upload(uploaded_file.name, content)
            except ValueError as exc:
                raise UserFacingError(str(exc)) from None
            with st.spinner("Analisando o PDF e construindo o Excel..."):
                process_pdf(uploaded_file.name, content)
        except UserFacingError as exc:
            st.error(f"Não foi possível processar o relatório: {exc}")
        except Exception:
            st.error("Não foi possível processar o relatório. Verifique o PDF ou procure o administrador.")

    render_result()


if __name__ == "__main__":
    main()
