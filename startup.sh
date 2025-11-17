#!/bin/bash

# Streamlitをポート8000で起動（Azure App Serviceのデフォルト）
python -m streamlit run app.py \
    --server.port=8000 \
    --server.address=0.0.0.0 \
    --server.headless=true \
    --server.enableCORS=false \
    --server.enableXsrfProtection=false
