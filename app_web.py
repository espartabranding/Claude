#!/usr/bin/env python3
"""
Handwriting Recognition to Excel - Web Interface (Streamlit)
=============================================================
Modern web interface for uploading handwritten images and downloading
the recognized content as formatted Excel spreadsheets.

Run with:
    streamlit run app_web.py
"""

import io
import os
import tempfile
import time
import logging
from pathlib import Path

import cv2
import numpy as np
import streamlit as st
from PIL import Image

from handwriting_to_excel import (
    HandwritingToExcel,
    ImagePreprocessor,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Cache: load OCR engine once, reuse across requests
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner="Carregando motores OCR (apenas na primeira vez)...")
def get_pipeline(engine, languages_tuple, enhance_level, mode):
    """Cache the pipeline so OCR models are loaded only once."""
    return HandwritingToExcel(
        engine=engine,
        languages=list(languages_tuple),
        enhance_level=enhance_level,
        mode=mode,
    )

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Reconhecimento de Escrita Manual",
    page_icon="<p>&#x1f4dd;</p>",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------

st.markdown("""
<style>
    /* Main container */
    .main .block-container {
        padding-top: 2rem;
        max-width: 1200px;
    }

    /* Header styling */
    .app-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        color: white;
    }
    .app-header h1 {
        color: white;
        margin: 0 0 0.5rem 0;
        font-size: 2rem;
    }
    .app-header p {
        color: #b8d4e8;
        margin: 0;
        font-size: 1.05rem;
    }

    /* Upload area */
    .upload-section {
        background: #f8fafc;
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 2rem;
        text-align: center;
        transition: border-color 0.3s;
    }
    .upload-section:hover {
        border-color: #2d6a9f;
    }

    /* Result card */
    .result-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }

    /* Stats badges */
    .stat-badge {
        display: inline-block;
        background: #eef2ff;
        color: #3730a3;
        padding: 0.4rem 1rem;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
        margin: 0.25rem;
    }

    /* Success message */
    .success-box {
        background: #ecfdf5;
        border: 1px solid #6ee7b7;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        color: #065f46;
    }

    /* Processing steps */
    .step-indicator {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        padding: 0.5rem 0;
        color: #475569;
    }
    .step-done {
        color: #059669;
        font-weight: 600;
    }

    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Sidebar */
    .css-1d391kg {
        padding-top: 1rem;
    }

    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #10b981 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.1rem;
        font-weight: 600;
        border-radius: 8px;
        width: 100%;
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #047857 0%, #059669 100%);
        color: white;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Sidebar - Settings
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## Configuracoes")
    st.markdown("---")

    engine = st.selectbox(
        "Motor OCR",
        options=["both", "tesseract", "easyocr"],
        format_func=lambda x: {
            "both": "Ambos (Tesseract + EasyOCR)",
            "tesseract": "Tesseract",
            "easyocr": "EasyOCR",
        }[x],
        help="Usar ambos motores da melhor precisao",
    )

    enhance = st.selectbox(
        "Nivel de Melhoria",
        options=["standard", "aggressive"],
        format_func=lambda x: {
            "standard": "Padrao",
            "aggressive": "Agressivo (escrita muito ruim)",
        }[x],
        help="Use 'Agressivo' para escrita muito dificil de ler",
    )

    mode = st.selectbox(
        "Modo de Deteccao",
        options=["auto", "table", "lines"],
        format_func=lambda x: {
            "auto": "Automatico",
            "table": "Tabela (forcado)",
            "lines": "Linhas (sem colunas)",
        }[x],
        help="'Automatico' detecta se o conteudo e tabular ou linear",
    )

    languages = st.multiselect(
        "Idiomas",
        options=["pt", "en", "es", "fr", "de"],
        default=["pt", "en"],
        format_func=lambda x: {
            "pt": "Portugues",
            "en": "Ingles",
            "es": "Espanhol",
            "fr": "Frances",
            "de": "Alemao",
        }[x],
    )

    first_row_header = st.checkbox(
        "Primeira linha e cabecalho",
        value=True,
        help="Formata a primeira linha como cabecalho no Excel",
    )

    show_preview = st.checkbox(
        "Mostrar pre-processamento",
        value=False,
        help="Exibe a imagem apos o pre-processamento",
    )

    st.markdown("---")
    st.markdown(
        "**Dicas para melhores resultados:**\n"
        "- Boa iluminacao, sem sombras\n"
        "- Camera perpendicular ao papel\n"
        "- Alta resolucao (300+ DPI)\n"
        "- Caneta escura em papel claro"
    )


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.markdown("""
<div class="app-header">
    <h1>Reconhecimento de Escrita Manual para Excel</h1>
    <p>Envie uma foto de escrita manual e receba uma planilha Excel organizada.
       Funciona mesmo com escrita desorganizada e dificil de ler.</p>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Main area
# ---------------------------------------------------------------------------

col_upload, col_result = st.columns([1, 1], gap="large")

with col_upload:
    st.markdown("### Enviar Imagem")

    uploaded_file = st.file_uploader(
        "Arraste ou clique para enviar sua imagem",
        type=["jpg", "jpeg", "png", "bmp", "tiff", "tif", "webp"],
        help="Formatos suportados: JPG, PNG, BMP, TIFF, WebP",
    )

    if uploaded_file is not None:
        # Display uploaded image
        image = Image.open(uploaded_file)
        st.image(image, caption=f"{uploaded_file.name} ({image.size[0]}x{image.size[1]}px)", use_container_width=True)

        # File info
        size_mb = uploaded_file.size / (1024 * 1024)
        st.markdown(
            f'<span class="stat-badge">Tamanho: {size_mb:.1f} MB</span>'
            f'<span class="stat-badge">Formato: {uploaded_file.type}</span>',
            unsafe_allow_html=True,
        )

with col_result:
    st.markdown("### Resultado")

    if uploaded_file is None:
        st.info("Envie uma imagem no painel a esquerda para comecar.")
    else:
        # Process button
        process_btn = st.button(
            "Processar Imagem",
            type="primary",
            use_container_width=True,
        )

        if process_btn:
            with st.spinner(""):
                progress = st.progress(0, text="Iniciando...")

                try:
                    # Save uploaded file to temp
                    progress.progress(5, text="Salvando imagem temporaria...")
                    with tempfile.NamedTemporaryFile(
                        suffix=Path(uploaded_file.name).suffix,
                        delete=False,
                    ) as tmp_in:
                        tmp_in.write(uploaded_file.getvalue())
                        tmp_input_path = tmp_in.name

                    tmp_output_path = tempfile.mktemp(suffix=".xlsx")

                    # Step 1: Load and preprocess
                    progress.progress(15, text="Carregando imagem...")
                    preprocessor = ImagePreprocessor(enhance_level=enhance)
                    img = preprocessor.load_image(tmp_input_path)

                    progress.progress(30, text="Pre-processando (corrigindo rotacao, ruido, contraste)...")
                    processed_versions = preprocessor.preprocess_for_ocr(img)

                    # Show preprocessed image if requested
                    if show_preview:
                        st.markdown("**Imagem pre-processada:**")
                        st.image(processed_versions[0], caption="Versao binarizada", use_container_width=True)

                    # Step 2: OCR
                    progress.progress(50, text="Reconhecendo texto (OCR)...")
                    langs = tuple(languages) if languages else ("pt", "en")
                    pipeline = get_pipeline(engine, langs, enhance, mode)

                    # Step 3: Full pipeline
                    progress.progress(70, text="Detectando estrutura da tabela...")
                    table_data = pipeline.convert(
                        image_path=tmp_input_path,
                        output_path=tmp_output_path,
                        first_row_header=first_row_header,
                    )

                    progress.progress(90, text="Gerando arquivo Excel...")

                    # Read the generated Excel file
                    with open(tmp_output_path, "rb") as f:
                        excel_bytes = f.read()

                    progress.progress(100, text="Concluido!")
                    time.sleep(0.3)
                    progress.empty()

                    # Results
                    if table_data and any(any(cell for cell in row) for row in table_data):
                        num_rows = len(table_data)
                        num_cols = max(len(r) for r in table_data) if table_data else 0

                        st.markdown(
                            f'<div class="success-box">'
                            f'Reconhecimento concluido com sucesso!'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        st.markdown("")

                        # Stats
                        st.markdown(
                            f'<span class="stat-badge">{num_rows} linhas</span>'
                            f'<span class="stat-badge">{num_cols} colunas</span>',
                            unsafe_allow_html=True,
                        )
                        st.markdown("")

                        # Preview table
                        st.markdown("**Pre-visualizacao:**")
                        preview_data = []
                        for row in table_data[:20]:
                            preview_data.append(
                                {f"Col {i+1}": cell for i, cell in enumerate(row)}
                            )
                        st.dataframe(preview_data, use_container_width=True)

                        if len(table_data) > 20:
                            st.caption(f"Mostrando 20 de {len(table_data)} linhas")

                        # Download button
                        st.markdown("")
                        output_name = Path(uploaded_file.name).stem + ".xlsx"
                        st.download_button(
                            label=f"Baixar Excel ({output_name})",
                            data=excel_bytes,
                            file_name=output_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    else:
                        st.warning(
                            "Nenhum texto foi detectado na imagem. "
                            "Tente com modo 'Agressivo' ou verifique a qualidade da imagem."
                        )

                except Exception as e:
                    progress.empty()
                    st.error(f"Erro ao processar: {str(e)}")
                    logger.exception("Processing error")

                finally:
                    # Cleanup temp files
                    for path in [tmp_input_path, tmp_output_path]:
                        try:
                            if os.path.exists(path):
                                os.unlink(path)
                        except OSError:
                            pass


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #94a3b8; font-size: 0.85rem;'>"
    "Handwriting Recognition Tool | "
    "Tesseract OCR + EasyOCR | "
    "OpenCV Image Processing"
    "</div>",
    unsafe_allow_html=True,
)
