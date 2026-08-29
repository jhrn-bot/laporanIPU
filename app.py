import datetime
import re
from zoneinfo import ZoneInfo
import bs4
import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

st.set_page_config(
    page_title="Penjana Mesej IPU Selangor", page_icon="😷", layout="centered"
)


def get_chrome_driver():
    """Membina Selenium Webdriver yang melepasi Cloudflare/Firewall APIMS."""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    try:
        # Untuk Streamlit Cloud (Linux Container)
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception:
        # Untuk ujian tempatan (Komputer sendiri)
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)

    return driver


def fetch_apims_browser_data():
    """Simulasi pelayar sebenar untuk mengambil bacaan jam terkini dari APIMS."""
    driver = get_chrome_driver()
    url = "https://eqms.doe.gov.my/APIMS/main"

    selangor_stations = {}

    try:
        driver.get(url)
        # Beri masa 5 saat untuk JavaScript APIMS render jadual penuh
        driver.implicitly_wait(5)

        html_content = driver.page_source
        soup = bs4.BeautifulSoup(html_content, "html.parser")

        # Cari semua baris jadual yang mengandungi Selangor
        rows = soup.find_all("tr")
        for row in rows:
            text = row.get_text(" ", strip=True)
            if "SELANGOR" in text.upper():
                # Corak regex mencari Nama Stesen & IPU (cth: Shah Alam 154**)
                match = re.search(
                    r"(Klang \(Johan Setia\)|Klang|Shah Alam|Petaling Jaya|Kuala Selangor|Banting)(?:.*?)\s(\d{1,3})\*\*",
                    text,
                    re.IGNORECASE,
                )
                if match:
                    loc = match.group(1).strip()
                    val = int(match.group(2))
                    selangor_stations[loc] = val

    except Exception as e:
        st.error(f"Ralat semasa membaca APIMS Browser: {e}")
    finally:
        driver.quit()

    return [
        {"location": loc, "value": val}
        for loc, val in selangor_stations.items()
    ]


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

if st.button("🔄 Tarik Data Live (Browser Engine) & Jana Mesej", type="primary"):
    with st.spinner("Membuka browser automatik & mengambil data jam terkini..."):
        stesen_data = fetch_apims_browser_data()

        if stesen_data:
            formatted_text = generate_whatsapp_text(stesen_data)
            st.success("Teks kenyataan berjaya dijana dari laman live!")

            st.text_area(
                "Teks Mesej WhatsApp (Sedia Salin):",
                formatted_text,
                height=320,
            )

            st.subheader("Klik salin di penjuru kanan kotak di bawah:")
            st.code(formatted_text, language="text")
        else:
            st.error(
                "Gagal mendapatkan data stesen. Sila pastikan laman web APIMS boleh diakses dan cuba lagi."
            )
