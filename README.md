# 📊 Exam Results Analyzer

A Streamlit web app that analyzes branch-wise exam results from `.xlsx` files.

## Features

- **Upload any `.xlsx`** result file (Branch | Roll No | Marks format)
- **Auto-detects all branches** from the file
- **Absentee tracking** — marks `AB` are automatically counted with percentage
- **Custom mark ranges** — define your own ranges (e.g. 0-20, 21-40, 51-60...)
- **Visual charts** — stacked bar, heatmap, violin distribution, average scores
- **Summary table** with CSV export

## File Format

Your `.xlsx` file should have these 3 columns (row 1 = header, ignored):

| Branch | Roll Number | Marks |
|--------|------------|-------|
| CSE    | 23101A0101 | 64    |
| ECE    | 23101A0501 | AB    |

- Marks can be a number or `AB` for absent
- Multiple branches in one file are supported

## Local Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect GitHub → select this repo → `app.py`
4. Click **Deploy** — your app is live!
