import datetime
import re
from zoneinfo import ZoneInfo
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Penjana Mesej IPU Selangor", page_icon="😷", layout="centered"
)


def parse_apims_pasted_text(raw_text):
    """Logik mengekstrak stesen Selangor dari sebarang teks copy-paste APIMS."""
    selangor_stations = {}

    stations_map = {
        "SHAH ALAM": "Shah Alam",
        "JOHAN SETIA": "Klang (Johan Setia)",
        "PETALING JAYA": "Petaling Jaya",
        "KLANG": "Klang",
        "KUALA SELANGOR": "Kuala Selangor",
        "BANTING": "Banting",
    }

    cleaned_text = raw_text.replace("\n", " ").replace("\r", " ")

    pattern = r"(Shah Alam|JOHAN SETIA|Petaling Jaya|Kuala Selangor|Banting|Klang)(?:.*?)(?:CA\w+|MCAQM\w+)?\s*(\d{1,3})\*\*"
    matches = re.findall(pattern, cleaned_text, re.IGNORECASE)

    for loc_found, val_str in matches:
        loc_upper = loc_found.strip().upper()

        for key_name, std_name in stations_map.items():
            if key_name in loc_upper:
                if key_name == "KLANG" and "JOHAN SETIA" in loc_upper:
                    continue
                selangor_stations[std_name] = int(val_str)
                break

    return selangor_stations


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


def render_copy_button(text_to_copy):
    """Membina butang khas HTML/JS untuk menyalin teks ke Clipboard."""
    escaped_text = (
        text_to_copy.replace("\\", "\\\\")
        .replace("`", "\\`")
        .replace("\n", "\\n")
    )
    html_code = f"""
    <div style="margin-top: 10px; margin-bottom: 20px;">
        <button id="copyBtn" style="
            background-color: #25D366;
            color: white;
            border: none;
            padding: 12px 24px;
            font-size: 16px;
            font-weight: bold;
            border-radius: 8px;
            cursor: pointer;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
            transition: all 0.2s ease;
        ">📋 Salin Mesej ke WhatsApp</button>
        <span id="copyMsg" style="margin-left: 12px; color: #25D366; font-weight: bold; display: none;">✓ Berjaya Disalin!</span>
    </div>

    <script>
    document.getElementById('copyBtn').addEventListener('click', function() {{
        const text = `{escaped_text}`;
        navigator.clipboard.writeText(text).then(function() {{
            const msg = document.getElementById('copyMsg');
            msg.style.display = 'inline';
            setTimeout(function() {{
                msg.style.display = 'none';
            }}, 2500);
        }}).catch(function(err) {{
            alert('Gagal menyalin teks: ' + err);
        }});
    }});
    </script>
    """
    components.html(html_code, height=70)


# --- PAPARAN STREAMLIT ---
st.title("😷 Penjana Kenyataan IPU Selangor")
st.caption("Khusus untuk kegunaan CPRC JKN Selangor")

st.markdown("---")

st.subheader("📋 Tampal (Paste) Data APIMS di sini:")
raw_input = st.text_area(
    "Copy senarai dari web APIMS (Ctrl+A / highlight) dan tampal di bawah:",
    height=200,
    placeholder="Tampal teks di sini...",
)

if st.button("🚀 Ekstrak Selangor & Jana Laporan", type="primary"):
    if raw_input.strip():
        stesen_data = parse_apims_pasted_text(raw_input)

        if stesen_data:
            formatted_text = generate_whatsapp_text(stesen_data)
            st.success(f"Berjaya mengesan {len(stesen_data)} stesen Selangor!")

            st.text_area(
                "Teks Mesej WhatsApp (Sedia Salin):",
                formatted_text,
                height=280,
            )

            render_copy_button(formatted_text)

        else:
            st.warning(
                "⚠️ Tiada stesen Selangor dikesan. Pastikan teks yang di-copy mengandungi nama stesen Selangor dan bacaan IPU (contoh: 155**)."
            )
    else:
        st.error("Sila tampal (paste) teks terlebih dahulu.")
