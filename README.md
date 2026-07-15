# Amazon SP-API RSS Monitor

A Streamlit web application that monitors the Amazon SP-API changelog RSS feed, classifies updates, and provides downloadable Excel reports.

## Features

- 📡 Fetches latest RSS feed from Amazon
- 🏷️ Classifies updates into 9 categories
- 🎯 Rates BP Impact (High/Med/Low/None)
- 📊 Interactive data visualization
- ⬇️ Excel report download with 4 sheets
- 🔄 State tracking to identify new items
- 🎨 Dark theme interface

## Installation

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment (Windows)
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

## Deployment on Streamlit Cloud

1. Push this repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click "New app"
4. Select your repository and branch
5. Set main file path to `app.py`
6. Click "Deploy"

## License

MIT