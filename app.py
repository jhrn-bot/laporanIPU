import datetime
import re
from zoneinfo import ZoneInfo
import streamlit as st

st.set_page_config(
    page_title="Penjana Mesej IPU Selangor", page_icon="😷", layout="centered"
)


def parse_raw_apims_text(raw_text):
    """Mengekstrak stesen Selangor dan bacaannya dari teks yang di-paste."""
    selangor_stations = {}

    # Senarai carian lokasi rasmi di Selangor
    target_locations = {
        "SHAH ALAM": "Shah Alam",
        "JOHAN SETIA": "Klang (Johan Setia)",
        "PETALING JAYA": "Petaling Jaya",
        "KLANG": "Klang",
        "KUALA SELANGOR": "Kuala Selangor",
        "BANTING": "Banting",
    }

    lines = raw_text.split("\n")

    for line in lines:
        line_str = line.strip()
        if "SELANGOR" in line_str.upper() or "JOHAN SETIA" in line_str.upper():
            # Regex untuk menangkap nama lokasi dan 2-3 digit nombor bacaan IPU sebelum simbol **
            # Contoh: 14SelangorShah Alam, SELANGORCA20B156** ATAU 20SelangorJOHAN SETIAMCAQM001155**
            match = re.search(r"(?:Selangor)?(.*?)(?:CA|MCAQM)\w+?(\d{1,3})\*\*", line_str, re.IGNORECASE)
            
            if match:
                loc_raw = match.group(1).replace(", SELANGOR", "").strip().upper()
                val = int(match.group(2))

                # Padankan lokasi raw dengan nama standard
                for key, std_name in target_locations.items():
                    if key in loc_raw:
                        selangor_stations[std_name] = val
                        break

    return selangor_stations


def generate_whatsapp_text(data_stesen):
    """Menjana format teks rasmi CPRC JKN Selangor mengikut Zon Masa Malaysia."""
    kl_tz = ZoneInfo("Asia/Kuala_Lumpur")
    now = datetime.datetime.now(kl_tz)

    bulan_malay = {
        1: "Januari", 2: "Februari", 3: "Mac", 4: "April",
        5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos",
        9: "September", 10: "Oktober", 11: "November", 12: "Disember"
    }

    tarikh = f"{now.day} {bulan_malay[now.month]} {now.year}"

    hour_12 = now.strftime("%I").lstrip("0")
    am_pm = "pagi" if now.hour < 12 else ("tengah hari" if now.hour == 12 else "petang")
    if now.hour >= 19:
        am_pm = "malam"

    masa = f"{hour_12}.00 {am_pm}"

    unhealthy, moderate, good = [], [], []

    for loc, val in data_stesen.items():
        item = {"location": loc, "value": val}
        if val > 100:
            unhealthy.append(item)
        elif val > 50:
            moderate.append(item)
        else:
            good.append(item)

    # Susun ikut bacaan tertinggi ke terendah
    unhealthy.sort(key=lambda x: x["value"], reverse=True)
    moderate.sort(key=lambda x: x["value"], reverse=True)
    good.sort(key=lambda x: x["value"], reverse=True)

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

st.subheader("📋 Tampal (Paste) Raw Data dari APIMS di bawah:")
raw_input = st.text_area(
    "Tekan Ctrl+A pada web APIMS, copy dan paste semua kandungan di sini:",
    height=200,
    placeholder="Tampal teks data APIMS di sini..."
)

if st.button("🚀 Ekstrak Selangor & Jana Laporan", type="primary"):
    if raw_input.strip():
        stesen_data = parse_raw_apims_text(raw_input)

        if stesen_data:
            formatted_text = generate_whatsapp_text(stesen_data)
            st.success(f"Berjaya mengesan {len(stesen_data)} stesen di Selangor!")

            st.text_area("Teks Mesej WhatsApp (Sedia Salin):", formatted_text, height=300)
            
            st.subheader("Klik salin di penjuru kanan kotak di bawah:")
            st.code(formatted_text, language="text")
        else:
            st.warning("⚠️ Tiada stesen Selangor dikesan dalam teks yang dimasukkan. Sila pastikan anda copy teks jadual dari APIMS.")
    else:
        st.error("Sila tampal (paste) teks dari laman APIMS terlebih dahulu.")
