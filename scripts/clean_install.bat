@echo off
echo 🧹 Cleaning up heavy dependencies...
pip uninstall -y sentence-transformers torch spacy nltk transformers pydantic ftfy unidecode litellm stqdm

echo 📦 Installing lightweight dependencies...
pip install -r requirements.txt

echo ✅ Cleanup complete!
echo 🚀 You can now run the app: streamlit run app.py
pause
