name: 🔄 Actualizar Datos desde Excel

on:
  push:
    paths:
      - '**.xlsx'
      - '**.xls'
  workflow_dispatch:

permissions:
  contents: write

jobs:
  convert-excel:
    runs-on: ubuntu-latest
    name: Convertir Excel → data.js

    steps:
      - name: 📥 Checkout repositorio
        uses: actions/checkout@v4

      - name: 🐍 Configurar Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: 📦 Instalar dependencias
        run: |
          pip install openpyxl pandas numpy

      - name: 🔄 Convertir Excel a data.js
        run: python convert.py

      - name: 📤 Commit y push data.js actualizado
        run: |
          git config --local user.email "github-actions[bot]@users.noreply.github.com"
          git config --local user.name "github-actions[bot]"
          git add data.js
          if git diff --staged --quiet; then
            echo "No hay cambios en los datos"
          else
            git commit -m "🤖 Auto-update: datos actualizados desde Excel [$(date +'%Y-%m-%d %H:%M')]"
            git push
          fi
