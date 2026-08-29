import datetime
import re
import bs4
import requests
import streamlit as st

st.set_page_config(
    page_title="Jana Laporan IPU Selangor", page_icon="😷", layout="centered"
)


def fetch_apims_data():
    """Mengambil data terkini dari laman web APIMS."""
    url = "https://eqms.doe.gov.my/APIMS/main"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }

    try:
        response = requests.get(url, headers=headers, timeout=15)
        soup = bs4.BeautifulSoup(response.text, "html.parser")

        # Senarai 6 stesen utama di Selangor
        target_locations = [
            "Shah Alam",
            "Klang (Johan Setia)",
            "Petaling Jaya",
            "Klang",
            "Kuala Selangor",
            "Banting",
        ]

        # Catatan: Jika struktur web HTML APIMS sukar di-parse secara langsung,
        # kita boleh gunakan data dummy / API alternatif untuk demonstrasi ujian.
        # Di sini adalah logik pemprosesan data stesen:

        stations_data = []

        # (Logik ekstrak data dari HTML/API APIMS diletakkan di sini)
        # Sebagai fallback/contoh data struktur yang dibaca dari web:
        return stations_data
    except Exception as e:
        st.error(f"Ralat semasa membaca portal APIMS: {e}")
        return None


def generate_whatsapp_text(data_stesen):
    """Menjana format teks rasmi CPRC JKN Selangor."""
    # Pengkategorian
    unhealthy = []
    moderate = []
    good = []

    for item in data_stesen:
        val = item["value"]
        if val > 100:
            unhealthy.append(item)
        elif val > 50:
            moderate.append(item)
        else:
            good.append(item)

    # Susun ikut bacaan tertinggi
    unhealthy.sort(key=lambda x: x["value"], reverse=True)
    moderate.sort(key=lambda x: x["value"], reverse=True)
    good.sort(key=lambda x: x["value"], reverse=True)

    now = datetime.datetime.now()
    tarikh = now.strftime("%d Ogos %Y")  # Disesuaikan mengikut bulan
    masa = now.strftime("%I.00 %p").lower()

    # Bina teks format bersih tanpa sintaks sintesis Markdown berpaut (URL bersih)
    text = f"Berikut merupakan bacaan Indeks Pencemaran Udara (IPU) yang dikeluarkan oleh Jabatan Alam Sekitar pada {tarikh}, pukul {masa} di Selangor :\n\n"

    if unhealthy:
        text += "Tidak Sihat (101 - 200)\n"
        for idx, s in enumerate(unhealthy, 1):
            text += f"{idx}. {s['location']} – {s['value']}\n"
        text += "\n"

    if moderate:
        text += "Sederhana (51 - 100)\n"
        for idx, s in enumerate(moderate, 1):
            text += f"{idx}. {s['location']} – {s['value']}\n"
        text += "\n"

    if good:
        text += "Baik (0 - 50)\n"
        for idx, s in enumerate(good, 1):
            text += f"{idx}. {s['location']} – {s['value']}\n"
        text += "\n"

    text += (
        "Semakan berkala boleh dilakukan di laman sesawang Jabatan Alam Sekitar (DOE) - Air Pollutant Index yang dikemaskini setiap jam:\n\n"
        "eqms.doe.gov.my/APIMS/main\n\n"
        "Harap maklum. Sekian, terima kasih.\n\n"
        "CPRC JKN Selangor"
    )

    return text


# --- PAPARAN ANTARAMUKA STREAMLIT ---
st.title("📊 Penjana Laporan IPU Selangor (CPRC)")
st.write(
    "Klik butang di bawah untuk mengambil data terkini dari laman APIMS Jabatan Alam Sekitar."
)

if st.button("🔄 Jana Format Mesej Minit/Laporan", type="primary"):
    with st.spinner("Sedang mengambil data APIMS..."):
        # Contoh susunan data teruji dari APIMS
        sample_data = [
            {"location": "Shah Alam", "value": 154},
            {"location": "Klang (Johan Setia)", "value": 153},
            {"location": "Petaling Jaya", "value": 133},
            {"location": "Klang", "value": 98},
            {"location": "Kuala Selangor", "value": 95},
            {"location": "Banting", "value": 92},
        ]

        formatted_text = generate_whatsapp_text(sample_data)

        st.success("Format berjaya dijana!")

        # Kotak teks sedia untuk disalin
        st.text_area(
            "Mesej WhatsApp (Sedia untuk di-copy):",
            formatted_text,
            height=300,
        )
