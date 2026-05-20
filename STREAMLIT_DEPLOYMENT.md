# Streamlit Cloud Deployment Instructions

Use these steps to publish the dashboard online with GitHub and Streamlit Community Cloud.

## 1. Create a GitHub Repository

Recommended repository name:

```text
panynj-port-workforce-analytics
```

On GitHub:

1. Go to `https://github.com/new`.
2. Name the repository `panynj-port-workforce-analytics`.
3. Choose Public or Private.
4. Do not add a README, `.gitignore`, or license because this project already includes them.
5. Click Create repository.

## 2. Push This Project to GitHub

In Terminal:

```bash
cd ~/Desktop/port-authority-workforce-analytics
git remote add origin https://github.com/YOUR-USERNAME/panynj-port-workforce-analytics.git
git branch -M main
git push -u origin main
```

Replace `YOUR-USERNAME` with your GitHub username.

## 3. Deploy on Streamlit Community Cloud

1. Go to `https://share.streamlit.io`.
2. Sign in with GitHub.
3. Click New app.
4. Select the GitHub repository.
5. Set the main file path to:

```text
app.py
```

6. Click Deploy.

Streamlit will install packages from:

```text
requirements.txt
```

## 4. App Entry Point

The dashboard runs with:

```bash
streamlit run app.py
```

## 5. Deployment Notes

- The app does not require paid APIs.
- The sample Excel file is included in `data/`.
- Generated reports are written to `outputs/`.
- The app uses a light Streamlit theme from `.streamlit/config.toml`.
