import datetime
import re
from zoneinfo import ZoneInfo
import bs4
import requests
import streamlit as st

st.set_page_config(
    page_title="Penjana Mesej IPU Selangor", page_icon="😷", layout="centered"
)


def fetch_apims_direct():
    """Tarik data secara terus dari APIMS."""
    selangor_stations = {}
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        ),
        "Referer": "https://eqms.doe.gov.my/APIMS/main",
    }
    try:
        url = "https://eqms.doe.gov.my/APIMS/main"
        res = requests.get(url, headers=headers, timeout=5)
        soup = bs4.BeautifulSoup(res.text, "html.parser")
        for row in soup.find_all("tr"):
            text = row.get_text(" ", strip=True)
            if "SELANGOR" in text.upper():
                match = re.search(
                    r"(Klang \(Johan Setia\)|Klang|Shah Alam|Petaling Jaya|Kuala Selangor|Banting)(?:.*?)\s(\d{1,3})\*\*",
                    text,
                    re.IGNORECASE,
                )
                if match:
                    selangor_stations[match.group(1).strip()] = int(
                        match.group(2)
                    )
    except Exception:
        pass

    return selangor_stations


def generate_whatsapp_text(data_stesen):
    """Jana teks format rasmi CPRC."""
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


# --- INTERFACE STREAMLIT ---
st.title("😷 Penjana Kenyataan IPU Selangor")
st.caption("Khusus untuk kegunaan CPRC JKN Selangor")

st.markdown("---")

# Pilihan Moda Kaedah Data
mode = st.radio(
    "Pilih Kaedah Masukan Data:",
    ("Automatik (Auto-Fetch)", "Semakan/Kemaskini Manual"),
)

input_data = {}

if mode == "Automatik (Auto-Fetch)":
    if st.button("🔄 Tarik Data APIMS & Jana Mesej", type="primary"):
        with st.spinner("Mengakses APIMS..."):
            input_data = fetch_apims_direct()
            if not input_data:
                st.warning(
                    "⚠️ Sambungan automatik ke APIMS disekat oleh server/firewall. "
                    "Sila tukar ke mod 'Semakan/Kemaskini Manual' di atas untuk jana laporan serta-merta."
                )
else:
    st.subheader("Masukkan Bacaan IPU Stesen Selangor:")
    col1, col2 = st.columns(2)
    with col1:
        v_sa = st.number_input("Shah Alam", value=154)
        v_js = st.number_input("Klang (Johan Setia)", value=153)
        v_pj = st.number_input("Petaling Jaya", value=133)
    with col2:
        v_kl = st.number_input("Klang", value=98)
        v_ks = st.number_input("Kuala Selangor", value=95)
        v_bt = st.number_input("Banting", value=92)

    input_data = {
        "Shah Alam": v_sa,
        "Klang (Johan Setia)": v_js,
        "Petaling Jaya": v_pj,
        "Klang": v_kl,
        "Kuala Selangor": v_ks,
        "Banting": v_bt,
    }

if input_data:
    formatted_text = generate_whatsapp_text(input_data)
    st.success("Teks Kenyataan Berjaya Dijana!")

    st.text_area("Teks Mesej WhatsApp (Sedia Salin):", formatted_text, height=300)
    st.code(formatted_text, language="text")
