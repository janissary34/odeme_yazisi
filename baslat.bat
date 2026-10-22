@echo off
cd /d C:\odeme_yazisi
python -m streamlit run app.py --server.address=0.0.0.0 --server.port=8501
pause