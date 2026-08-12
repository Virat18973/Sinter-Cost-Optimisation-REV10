# Sinter Burden Control V8 — Complete Package

Files:
- app.py — Streamlit dashboard
- optimizer.py — latest optimization backend supplied for this package
- requirements.txt — Python dependencies

Dashboard architecture:
1. Primary raw-material chemistry is loaded from the primary Excel and remains read-only.
2. Primary Price, RM Stock, Availability and Tech Max are editable in the dashboard.
3. Alternative raw-material chemistry can be uploaded through a separate Excel uploader.
4. Alternative chemistry, price, RM Stock, Tech Min/Max and Allow Alternative are editable.
5. Allow Alternative ON makes the alternative eligible; it does not force usage.
6. Manual Burden Control is self-contained and shows adjusted burden composition,
   adjusted cost composition, achieved chemistry and optimized-vs-manual comparison.

Run:
streamlit run app.py
