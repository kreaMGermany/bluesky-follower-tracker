name: Bluesky Tracker

on:
  schedule:
    - cron: '0 7 * * *'
  workflow_dispatch:

jobs:
  run:
    runs-on: ubuntu-latest

    permissions:
      contents: read

    steps:
      - name: Checkout repository
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Run tracker
        env:
          TENANT_ID: ${{ secrets.TENANT_ID }}
          CLIENT_ID: ${{ secrets.CLIENT_ID }}
          CLIENT_SECRET: ${{ secrets.CLIENT_SECRET }}
          SENDER_UPN: ${{ secrets.SENDER_UPN }}
          RECIPIENTS: ${{ secrets.RECIPIENTS }}
          ONEDRIVE_FILE_PATH: ${{ secrets.ONEDRIVE_FILE_PATH }}
          ACCOUNTS_JSON: ${{ secrets.ACCOUNTS_JSON }}
        run: python tracker.py
