import streamlit as st
import pandas as pd
from io import BytesIO
import requests
from openpyxl import load_workbook

# === Funktion: Spaltenbreite in Excel automatisch setzen ===
def set_excel_column_width(output_bytesio):
    output_bytesio.seek(0)
    wb = load_workbook(output_bytesio)
    ws = wb.active
    for col in ws.columns:
        max_length = 0
        column_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column_letter].width = max_length + 2
    new_output = BytesIO()
    wb.save(new_output)
    new_output.seek(0)
    return new_output

# === Custom Style (Lila / Schwarz / LED Effekt) ===
st.markdown("""
<style>
body {
    background-color: #0a0a0f;
    color: white;
}
.stApp {
    background: linear-gradient(135deg, #0a0a0f, #1a002b);
    color: white;
    font-family: 'Segoe UI', sans-serif;
}
.title {
    font-size: 40px;
    text-align: center;
    font-weight: bold;
    margin-bottom: 20px;
    color: #d48aff;
    text-shadow: 0 0 15px #ff00ff;
}
.sub {
    text-shadow: 0 0 8px #a64ff7;
    text-align: center;
    font-size: 18px;
}
</style>
""", unsafe_allow_html=True)

# Title
st.markdown("<div class='title'> Artikel-Filter WebApp</div>", unsafe_allow_html=True)

st.write("---")

# === Datei Upload ===
uploaded_file = st.file_uploader("📁 Excel-Datei auswählen", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("Datei wurde geladen!")

    # Filterart auswählen
    filter_mode = st.radio(
        "Wie möchtest du filtern?",
        ["Einen Artikel", "Mehrere Artikel"]
    )

    if filter_mode == "Einen Artikel":
        artikelsuche = st.text_input("Artikelnummer eingeben:")
        artikel_liste = [artikelsuche] if artikelsuche else []

    else:
        artikel_text = st.text_area(
            "Mehrere Artikelnummern (eine pro Zeile):"
        )
        artikel_liste = [x.strip() for x in artikel_text.split("\n") if x.strip()]

    st.write("---")

    if st.button("✨ Filter starten"):
        if not artikel_liste:
            st.error("Bitte mindestens einen Artikel eingeben!")
        else:
            # Nur gewünschte Spalten zulassen
            NEED_COLS = ["Artikelnummer", "Bezeichnung", "Verfügbar (Stück)"]

            for col in NEED_COLS:
                if col not in df.columns:
                    st.error(f"❌ Fehler: Spalte '{col}' fehlt in der Excel-Datei!")
                    st.stop()

            df["Artikelnummer"] = df["Artikelnummer"].astype(str)

            filtered = df[df["Artikelnummer"].isin(artikel_liste)][NEED_COLS]

            if filtered.empty:
                st.warning("⚠️ Keine Treffer gefunden!")
            else:
                st.success("🎉 Artikel erfolgreich gefiltert!")

                # Tabelle im Browser schön anzeigen
                st.markdown("### 📊 Gefilterte Artikel")
                table_html = filtered.to_html(index=False, escape=False)
                st.markdown(f"""
                    <div style="overflow-x:auto; border:2px solid #6a0dad; padding:10px; border-radius:10px;">
                        {table_html}
                    </div>
                """, unsafe_allow_html=True)

                # Excel-Datei vorbereiten
                output = BytesIO()
                filtered.to_excel(output, index=False)
                output.seek(0)
                filtered_bytesio = set_excel_column_width(output)

                # === Download-Button ===
                st.download_button(
                    label="📥 Gefilterte Excel herunterladen",
                    data=filtered_bytesio,
                    file_name="gefilterte_artikel.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.write("---")

                # === SEND TO URL ===
                st.subheader("📡 Datei an einen Link schicken (optional)")
                target_url = st.text_input("Link eingeben (Webhook / API / Upload-Ziel):", key="send_url")

                if st.button("📤 An Link senden"):
                    if not target_url:
                        st.error("Bitte zuerst einen Link eingeben!")
                    else:
                        try:
                            files = {"file": ("gefilterte_artikel.xlsx", filtered_bytesio.getvalue())}
                            r = requests.post(target_url, files=files)
                            if r.status_code < 300:
                                st.success("✔️ Datei erfolgreich gesendet!")
                            else:
                                st.error(f"❌ Fehler beim Senden: Status {r.status_code}")
                        except Exception as e:
                            st.error(f"❌ Fehler: {e}")
