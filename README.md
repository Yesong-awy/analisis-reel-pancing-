# Analisis Reel Pancing Shopee

Aplikasi dashboard Streamlit untuk menganalisis data reel pancing dari Shopee.

## Menjalankan secara lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

Aplikasi membaca `shopee-co-id-2026-09-15-8.xlsx` dari folder proyek dan gambar
header dari `assets/reel-fishing-hero.jpg`.

## Deploy di Streamlit Community Cloud

1. Hubungkan akun GitHub ke [Streamlit Community Cloud](https://share.streamlit.io/).
2. Pilih repositori ini, branch `main`, dan file utama `app.py`.
3. Klik **Deploy**. Dependensi akan dipasang dari `requirements.txt`.
