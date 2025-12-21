@echo off
echo 🧹 Cleaning up heavy dependencies...
pip uninstall -y sentence-transformers torch spacy nltk transformers

echo 📦 Installing lightweight dependencies...
pip install -r requirements.txt

echo ✅ Cleanup complete!
echo 🚀 You can now run the app: streamlit run streamlit_app.py
pause
