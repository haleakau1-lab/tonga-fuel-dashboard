# Railway Deployment

## Repository

- GitHub repo: `haleakau1-lab/tonga-fuel-dashboard`

## Railway Setup

1. Go to Railway and create a new project from GitHub.
2. Select the repository `haleakau1-lab/tonga-fuel-dashboard`.
3. Railway will detect the Python app via `requirements.txt`.
4. The app start command is already defined in `railway.json`.

## Runtime Notes

- App entrypoint: `dashboard.py`
- Railway uses the generated `$PORT` environment variable.
- Streamlit binds to `0.0.0.0` for public access.

## Current Start Command

```text
streamlit run dashboard.py --server.port $PORT --server.address 0.0.0.0
```