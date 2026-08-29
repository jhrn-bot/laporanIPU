import datetime
from zoneinfo import ZoneInfo
import requests
import streamlit as st

# Setup Konfigurasi Halaman
st.set_page_config(
    page_title="Penjana Mesej IPU Selangor", page_icon="😷", layout="centered"
)


def fetch_apims_live_data():
    """Mengambil data live secara terus daripada API JSON rasmi APIMS DOE."""
    url = "https://eqms.doe.gov.my/APIMS/api/hourly/getHourlyData"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
    }

    selangor_stations = []

    try:
        response = requests.get(url, headers=headers, timeout=12)
        if response.status_code == 200:
            data = response.json()

            for item in data:
                # Filter stesen di Selangor sahaja
                state = str(item.get("state", "")).upper()
                if "SELANGOR" in state:
                    location = item.get("location", "").strip()
                    # Bersihkan nama lokasi dari perkataan ", SELANGOR" jika ada
                    location = location.replace(", SELANGOR", "").strip()

                    # Ambil senarai bacaan jam dan ambil bacaan paling terkini
                    readings = item.get("readings", [])
                    if readings:
                        latest_raw = str(readings[-1])
                        # Ekstrak nombor sahaja (membuang simbol ** atau huruf)
                        val_digits = "".join(filter(str.isdigit, latest_raw))
                        if val_digits:
                            selangor_stations.append(
                                {"location": location, "value": int(val_digits)}
                            )
    except Exception as e:
        st.error(f"Ralat capaian API APIMS: {e}")

    return selangor_stations


def generate_whatsapp_text(data_stesen):
    """Menjana format teks rasmi CPRC JKN Selangor mengikut Zon Masa Malaysia."""

    # 1. Penetapan Waktu Rasmi Malaysia (UTC+8)
    kl_tz = ZoneInfo("Asia/Kuala_Lumpur")
    now = datetime.datetime.now(kl_tz)

    bulan_malay = {
        1: "Januari",
        2: "Februari",
        3: "Mac",
        4: "April",
        5: "Mei",
        6: "Jun",
        7: "Julai",
        8: "Ogos",
        9: "September",
        10: "Oktober",
        11: "November",
        12: "Disember",
    }

    tarikh = f"{now.day} {bulan_malay[now.month]} {now.year}"

    # Format waktu Bahasa Melayu
    hour_12 = now.strftime("%I").lstrip("0")
    am_pm = "pagi" if now.hour < 12 else ("tengah hari" if now.hour == 12 else "petang")
    if now.hour >= 19:
        am_pm = "malam"

    masa = f"{hour_12}.00 {am_pm}"

    # 2. Pengkategorian IPU
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

    # Susun nilai mengikut bacaan tertinggi
    unhealthy.sort(key=lambda x: x["value"], reverse=True)
    moderate.sort(key=lambda x: x["value"], reverse=True)
    good.sort(key=lambda x: x["value"], reverse=True)

    # 3. Pembinaan Teks Mesej Rasmi
    text = f"Berikut merupakan bacaan Indeks Pencemaran Udara (IPU) yang dikeluarkan oleh Jabatan Alam Sekitar pada {tarikh}, pukul {masa} di Selangor :\n\n"

    if unhealthy:
        text += "Unhealthy (101-200)\n"
        for idx, s in enumerate(unhealthy, 1):
            text += f"{idx}. {s['location']} – {s['value']}\n"
        text += "\n"

    if moderate:
        text += "Moderate (51-100)\n"
        for idx, s in enumerate(moderate, 1):
            text += f"{idx}. {s['location']} – {s['value']}\n"
        text += "\n"

    if good:
        text += "Good (0-50)\n"
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


# --- ANTARAMUKA STREAMLIT ---
st.title("😷 Penjana Kenyataan IPU Selangor")
st.caption("Khusus untuk kegunaan CPRC JKN Selangor")

st.markdown("---")

if st.button("🔄 Tarik Data Live & Jana Mesej", type="primary"):
    with st.spinner("Sedang menarik data live dari API APIMS..."):
        stesen_data = fetch_apims_live_data()

        if stesen_data:
            formatted_text = generate_whatsapp_text(stesen_data)
            st.success("Teks berjaya dijana dengan data terkini!")

            # Kotak Paparan Teks Mesej
            st.text_area(
                "Teks Mesej WhatsApp (Sedia Salin):",
                formatted_text,
                height=320,
            )

            # Butang Salin
            st.subheader("Klik salin di penjuru kanan kotak di bawah:")
            st.code(formatted_text, language="text")
        else:
            st.error(
                "Gagal mengambil data dari APIMS. Sila cuba seketika lagi."
            )
