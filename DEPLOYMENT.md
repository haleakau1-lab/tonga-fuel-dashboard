# Dashboard Deployment Guide

## Option 1: Streamlit Community Cloud (Recommended)

1. Create a GitHub repository and upload the contents of this folder.
2. Confirm these files are included:
   - `dashboard.py`
   - `requirements.txt`
   - `Oil_Data_Consolidated.xlsx`
   - `background.jpg` and `tonga-energy-logo.jpg` (if used)
3. Go to Streamlit Community Cloud: https://share.streamlit.io
4. Click **New app** and select your repository.
5. Set:
   - **Main file path**: `dashboard.py`
   - **Python version**: 3.11 (recommended)
6. Deploy.

### Notes
- This app reads `Oil_Data_Consolidated.xlsx` from the same folder as `dashboard.py`, so keep the Excel file in the repo.
- If your Excel file is large, use Git LFS or a hosted storage source.

## Option 2: Render

1. Push this folder to GitHub.
2. Create a new **Web Service** on Render from your repo.
3. Use:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `streamlit run dashboard.py --server.port $PORT --server.address 0.0.0.0`
4. Deploy and open the generated URL.

## Quick Pre-Deploy Checklist

- App runs locally without errors.
- All required files are committed.
- `requirements.txt` is up to date.
- No local-only absolute paths are required at runtime.
