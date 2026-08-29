import datetime
import re
from zoneinfo import ZoneInfo
import bs4
import requests
import streamlit as st

st.set_page_config(
    page_title="Penjana Mesej IPU Selangor", page_icon="😷", layout="centered"
)


def fetch_apims_live_data():
    """Mengambil data live APIMS dengan gabungan API & HTML Scraping."""
    selangor_stations = {}

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            " (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Referer": "https://eqms.doe.gov.my/APIMS/main",
        "X-Requested-With": "XMLHttpRequest",
    }

    # KAEDAH 1: Capaian melalui API JSON
    try:
        url_api = "https://eqms.doe.gov.my/APIMS/api/hourly/getHourlyData"
        response = requests.get(url_api, headers=headers, timeout=10)

        if response.status_code == 200 and response.text.strip().startswith(
            "["
        ):
            data = response.json()
            for item in data:
                state = str(item.get("state", "")).upper()
                if "SELANGOR" in state:
                    loc = item.get("location", "").replace(", SELANGOR", "").strip()
                    readings = item.get("readings", [])
                    if readings:
                        # Ambil bacaan paling terkini yang valid (bukan N/A)
                        valid_readings = [
                            r for r in readings if str(r).strip() not in ["", "N/A", "-"]
                        ]
                        if valid_readings:
                            latest_raw = str(valid_readings[-1])
                            val_digits = "".join(filter(str.isdigit, latest_raw))
                            if val_digits:
                                selangor_stations[loc] = int(val_digits)
    except Exception:
        pass

    # KAEDAH 2: HTML Scraping jika API di-block oleh firewall server APIMS
    if not selangor_stations:
        try:
            url_main = "https://eqms.doe.gov.my/APIMS/main"
            resp_main = requests.get(url_main, headers=headers, timeout=10)
            soup = bs4.BeautifulSoup(resp_main.text, "html.parser")

            rows = soup.find_all("tr")
            for row in rows:
                text = row.get_text(strip=True)
                if "SELANGOR" in text.upper():
                    # Format biasa: Kuala Selangor, SELANGORCA18B98**
                    match = re.search(
                        r"(Klang|Shah Alam|Petaling Jaya|Kuala Selangor|Banting)(?:.*?)\D(\d{1,3})\*\*",
                        text,
                        re.IGNORECASE,
                    )
                    if match:
                        loc = match.group(1).strip()
                        val = int(match.group(2))
                        selangor_stations[loc] = val
        except Exception as e:
            st.error(f"Ralat capaian laman web APIMS: {e}")

    # Tukar kepada senarai berformat
    result = [
        {"location": loc, "value": val}
        for loc, val in selangor_stations.items()
    ]
    return result


def generate_whatsapp_text(data_stesen):
    """Menjana format teks rasmi CPRC JKN Selangor mengikut Zon Masa Malaysia."""
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


# --- PAPARAN STREAMLIT ---
st.title("😷 Penjana Kenyataan IPU Selangor")
st.caption("Khusus untuk kegunaan CPRC JKN Selangor")

st.markdown("---")

if st.button("🔄 Tarik Data Live & Jana Mesej", type="primary"):
    with st.spinner("Sedang mengambil bacaan IPU Selangor terkini..."):
        stesen_data = fetch_apims_live_data()

        if stesen_data:
            formatted_text = generate_whatsapp_text(stesen_data)
            st.success("Teks kenyataan berjaya dijana!")

            st.text_area(
                "Teks Mesej WhatsApp (Sedia Salin):",
                formatted_text,
                height=320,
            )

            st.subheader("Klik salin di penjuru kanan kotak di bawah:")
            st.code(formatted_text, language="text")
        else:
            st.error(
                "Talian ke APIMS sibuk/terputus. Sila cuba tekan butang sekali lagi."
            )
