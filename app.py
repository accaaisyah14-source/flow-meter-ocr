import streamlit as st
import pandas as pd
import easyocr
import cv2
import os
import re
import time
import numpy as np
import pytz

from PIL import Image
from datetime import datetime

import ssl
ssl._create_default_https_context = ssl._create_unverified_context


# ================= CONFIG =================
st.set_page_config(
    page_title="Flow Meter Recording",
    layout="wide",
    initial_sidebar_state="collapsed"
)

EXCEL_AIR = "meteran_air.xlsx"
EXCEL_LISTRIK = "meteran_listrik.xlsx"
UPLOAD_FOLDER = "uploads"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ================= TIMEZONE =================
tz_jkt = pytz.timezone("Asia/Jakarta")


# ================= OCR =================
@st.cache_resource
def load_reader():
    return easyocr.Reader(
        ['en'],
        gpu=False
    )


reader = load_reader()


def advanced_pre_process(img_np):
    """
    Logika OCR sama seperti laptop lama:
    CLAHE + Gaussian Blur
    """
    gray = cv2.cvtColor(
        img_np,
        cv2.COLOR_RGB2GRAY
    )

    clahe = cv2.createCLAHE(
        clipLimit=5.0,
        tileGridSize=(8, 8)
    )

    enhanced = clahe.apply(gray)

    return cv2.GaussianBlur(
        enhanced,
        (3, 3),
        0
    )


def robust_extract_logic(text_list):
    """
    Logika OCR lama yang paling stabil
    """
    full_text = " ".join(text_list).upper()

    # hapus satuan
    for unit in [
        "KWH",
        "KVARH",
        "M3/H",
        "M3",
        "KVAR"
    ]:
        full_text = full_text.replace(unit, "")

    # mapping karakter OCR error
    mapping = {
        "O": "0",
        "D": "0",
        "Q": "0",
        "B": "8",
        "S": "5",
        "I": "1",
        "L": "1",
        "T": "7",
        "Z": "2",
        "G": "6",
        "A": "4"
    }

    for k, v in mapping.items():
        full_text = full_text.replace(k, v)

    # koma jadi titik
    full_text = full_text.replace(",", ".")

    # regex angka meteran utama
    pattern = re.findall(
        r"\d{5,8}(?:\.\d{1,3})?",
        full_text
    )

    return max(pattern, key=len) if pattern else "Cek Foto"


# ================= SAVE DATA =================
def save_data(df, jenis):
    """
    Excel hanya:
    - Tanggal
    - Nama Meteran
    - Angka Meteran

    Foto tidak masuk Excel
    """
    file_path = (
        EXCEL_AIR
        if jenis == "Air"
        else EXCEL_LISTRIK
    )

    # pastikan hanya kolom excel
    df_save = df.copy()

    if "Foto" not in df_save.columns:
        df_save["Foto"] = ""

    df_excel = df_save[
        [
            "Tanggal",
            "Nama Meteran",
            "Angka Meteran"
        ]
    ].copy()

    with pd.ExcelWriter(
        file_path,
        engine="xlsxwriter"
    ) as writer:
        df_excel.to_excel(
            writer,
            index=False
        )
    st.success(f"Excel berhasil disimpan ke {file_path}")


# ================= LOAD DATA =================
def load_data(jenis):
    """
    Load Excel + pastikan kolom Foto tetap ada
    supaya UI tidak error
    """
    file_path = (
        EXCEL_AIR
        if jenis == "Air"
        else EXCEL_LISTRIK
    )

    if os.path.exists(file_path):
        df = pd.read_excel(
            file_path,
            dtype=str
        )

        # wajib ada kolom Foto
        if "Foto" not in df.columns:
            df["Foto"] = ""

        # wajib ada semua kolom dasar
        for col in [
            "Tanggal",
            "Nama Meteran",
            "Angka Meteran"
        ]:
            if col not in df.columns:
                df[col] = ""

        return df

    else:
        return pd.DataFrame(
            columns=[
                "Tanggal",
                "Nama Meteran",
                "Angka Meteran",
                "Foto"
            ]
        )

# --- 2. Daftar Nama Meteran ---
list_meteran_air = [
    "1. DEEPWELL 1 AQUADUCT SIPA", 
    "2. DEEPWELL 2 FISHPOND SIPA",
    "3. DEEPWELL 3 AWANG-AWANG SIPA",
    "4. STATIC WATER LEVEL (JTL)",
    "5. DYNAMIC WATER LEVEL (AWG)",
    "6. DEEPWELL 4 JATILANGKUNG PIPA",
    "7. STATIC WATER LEVEL (JTL)",
    "8. DYNAMIC WATER LEVEL (JTL)",
    "9. DEEPWELL 5 RIVER SIPA",
    "10. DEEPWELL 6 FRONT OFFICE SIPA",
    "11. DEEPWELL 7 TEMPURAN SIPA",
    "12. STATIC WATER LEVEL (TPR)",
    "13. DYNAMIC WATER LEVEL (TPR)",
    "14. DEEPWELL 1 AQUADUCT MBI",
    "15. DEEPWELL 2 FISHPOND SIPA",
    "16. DEEPWELL 3 AWANG-AWANG SIPA",
    "17. DEEPWELL 4 JATILANGKUNG MBI",
    "18. DEEPWELL 5 RIVER MBI",
    "19. DEEPWELL 6 FRONT OFFICE MBI",
    "20. DEEPWELL 7 TEMPURAN MBI",
    "21. DEEPWELL STORAGE TANK",
    "22. CATION 1",
    "23. CATION 2",
    "24. MM MIX TO PWT",
    "25. FLOW TO PWT",
    "26. AFTER ACF",
    "27. SOFTENER OUTLET",
    "28. MM MIX TO SWT",
    "29. FLOW TO FWT",
    "30. FWT TO CONSUMER",
    "31. SOFT WATER TO CONSUMER",
    "32. DEMIN WATER",
    "33. BOILER MAKE UP WATER",
    "34. CO2 COOLING WATER",
    "35. NH3 EVAP CONDENSOR",
    "36. ETHANOL MIXING WATER",
    "37. CO2 SCRUBBER TANK",
    "38. WWTP WATER CONS",
    "39. SOFT WATER FOR CHAIN LUBRICANT",
    "40. BOTTLE WASHER",
    "41. CRATE WASHER",
    "42. PASTEURIZER",
    "43. FILLERS & VACUUM PUMPS",
    "44. RACKING PLANT",
    "45. SODA STATION",
    "46. PACKAGING MAIN LINE",
    "47. CELLAR FLOOR ",
    "48. CO2 FOAM CATCHER 1",
    "49. CO2 FOAM CATCHER 2",
    "50. FILTRATION & BREWHOUSE MAIN LINE",  
    "51. YEAST TANK",
    "52. PVPP KIESELGUHR TANKS",
    "53. COLD & HOT WATER TANKS",
    "54. MAIN LABORATORY",
    "55. SANITARY WATER MAIN LINE",
    "56. FIRE FIGHTING",
    "57. CORE TEAM WC",
    "58. BOTTLING WC",
    "59. FULL STORE",
    "60. EMPTY STORE",
    "61. FRONT OFFICE",
    "62. PARKING LODGE",
    "63. WELFARE + CLINIC",
    "64. MUSHOLA",
    "65. CANTEEN + PORTER LODGE",
    "66. ENGINE ROOM WC",
    "67. BREWHOUSE WC",
    "68. SILO WC",
    "69. BUNKER FLOW METER",
    "70. BOILER 1 FUEL METER",
    "71. BOILER 2 FUEL METER",
    "72. FORKLIFT FUEL (PICK LIST)",
    "73. CORRECTOR GAS METER",
    "74. CANTEEN GAS METER",
    "75. BOILER 1 STEAM METER",
    "76. BOILER 1 FEED WATER METER",
    "77. BOILER 2 STEAM METER",
    "78. BOILER 2 FEED WATER METER",
    "79. PACKAGING STEAM METER",
    "80. PACKAGING CONDENSATE METER",
    "81. BREWING STEAM METER",
    "82. BREWING CONDENSATE METER",
    "83. BOILER BLOWDOWN METER",
    "84. CO2 RECUPERATION",
    "85. STORAGE TANK 1 VOLUME",
    "86. STORAGE TANK 2 VOLUME",
    "87. PACKAGING TOTALIZER",
    "88. BREWING TOTALIZER",
    "89. CO2 PURCHASE",
    "90. CO2 SOLD",
    "91. CO2 BOTTLED",
    "92. INCOMING CAUSTIC METER",
    "93. PACKAGING CAUSTIC METER",
    "94. BREWING CAUSTIC METER",
    "95. WTP CAUSTIC METER",
    "96. PASTEURIZER STEAM",
    "97. BOTTLE WASHER",
    "98. WK 1 STEAM METER",
    "99. WK 2 STEAM METER",
    "100. WTP STEAM METER",
    "101. RACKING PLANT",
    "102. WTP CAUSTIC METER",
    "103. PDAM MBI",
    "104. PDAM DEPAN",
    "103. ACF 1 CHRIWA",
    "104. ACF 2 CHRIWA",
    "105. FWS CIP",
    "106. FWS PROD",
    "107. STEAM",
    "108. STEAM TO PASTEUR",
    "109. FW TO PASTEUR",
    "110. STEAM CIP",
    "111. STEAM MBI",
    "112. STEAM BECIS",
    "113. FACT CIP BMF",
    "114. PROD ALDOX",
    "115. FACT TANK II",
    "116. SPRAY BALL MASTUN",
    "117. SPRAY BALL WPOOL",
    "118. STEAM BREWING",
    "119. RAW",
    "120. PWT",
    "121. FWT",
    "122. CO2 BALOON",
]

list_meteran_listrik = [
    "1. WELFARE X 30",
    "2. MAIN OFFICE LIGHTING X 30",
    "3. MAIN OFFICE AIR CON X 30",
    "4. PORTER LODGE X 30",
    "5. ENGINE ROOM LIGHTING E12 X 50",
    "6. FIRE FIGHTING W1 X 120",
    "7. BOTTLING LIGHTING / AIR CON B6 X 50",
    "8. WWTP LIGHTING V6 X 50",
    "9. AIR COMP 1 E20 X 30",
    "10. AIR COMP 2 E23 X 30",
    "11. AIR COMP 3 E26 X 30",
    "12. AIR COMP 4",
    "13. NH3 COMP 1 E13 X 60",
    "14. NH3 COMP 2 E14 X 120",
    "15. NH3 COMP 3 E15 X 120",
    "16. NH3 COMP 4",
    "17. NH3 COMP 5",
    "18. COOLING INSTALLATION E16 X 80",
    "19. COOLING INSTASI BARU",
    "20. CO2 RECUPERATION E33 X 120",
    "21. CO2 PLANT",
    "22. BD PLAYER 1",
    "23. BD PLAYER 2",
    "24. STEAM BOILER E10 X 80",
    "25 WATER TREATMENT W3 X 80",
    "26. WORKSHOP S1 X 50",
    "27. WET AREA",
    "28. 8",
    "29. DEEPWELL 1 AQUADUCT X 120",
    "30. DEEPWELL 2 FISHPOND V8 X 30",
    "31. WWTP LINE POWER V7 X 50",
    "32. CAUSTIC RECUPERATION E9 X 50",
    "33. BATTERY CHARGER E34 X 30",
    "34. RACKING PLANT B44 X 30",
    "35. BOTTLING LINE POWER B5 X 120",
    "36. PASTEURIZER B31 X 80",
    "37. MALT / SILO M1 X 50,"
    "38. BREWHOUSE CONTROL Z8 X 120",
    "39. CELLAR POWER Z21 X 50",
    "40. BREWHOUSE POWER Z28 X 80",
    "41. BREWHOUSE POWER Z26 X 50",
    "42. BREWHOUSE POWER Z22 X 50",
    "43. MAIN KWH ENG ROOM E2.2 X 500",
    "44. Cos Enginee Room",
    "45. MAIN KWH BOTTLING B2.2 X 500",
    "46. Cos Bottling Room (min 0.95)",
    "47. MAIN KWH OFFICE T2.3 X 500",
    "48. Cos Office min 0.95",
    "49. MAIN KWH B & C Z2.4 X 500",
    "50. Cos B & C (min 0.95),"
    "51. WWTP V2.1 X 500",
    "52. Cos WWTP (min 0.95),"
    "53. FULL STORE NEW LIGHTING",
    "54. CO2 COMPRESSOR",
    "55. PANEL KUNZEL",
    "56. PANEL BMF",
    "57. PANEL PLN KWH",
    "58. PANEL PLN KVRH",
    "59. KWH PLN MBI KVARH",
    "60. KWH PLN MBI TOTAL",
    "61. KWH PLN MBI LBP",
    "62. KWH PLN MBI BP",
    "63. KWH PLN TPI KVARH",
    "64. KWH PLN TPI TOTAL",
    "65. KWH PLN TPI LBP",
    "66. KWH PLN TPI BP",
    "67. MASTER BREWING",
    "68. MASTER PACKAGING",
    "69. 1.1",
    "70. 2.2",
    "71. 2.3",
    "72. 2.4",
    "73. 3.2",
    "74. 4.1",
    "75. 4.2",
    "76. 4.3",
    "77. 4.4",
    "78. 4.6",
]


# ================= UI =================
st.title("📟 Flow Meter Recording")

col1, col2 = st.columns(2)

with col1:
    jenis_meteran = st.selectbox(
        "Jenis Meteran",
        ["Air", "Listrik"]
    )

with col2:
    if jenis_meteran == "Air":
        nama_meteran = st.selectbox(
            "Nama Meteran",
            list_meteran_air
        )
    else:
        nama_meteran = st.selectbox(
            "Nama Meteran",
            list_meteran_listrik
        )

st.divider()


# ================= INPUT FOTO =================
tab1 = st.tabs(["📁 Upload"])[0]

with tab1:
    files = st.file_uploader(
        "Upload gambar",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )


# ================= PROSES OCR =================
if files:

    if not nama_meteran:
        st.warning("Pilih nama meteran dulu!")
        st.stop()

    now = datetime.now(tz_jkt)

    for f in files:
        try:
            # nama file
            fname = (
                f.name
                if hasattr(f, "name")
                else f"{int(time.time())}.jpg"
            )

            # buka gambar
            img = Image.open(f).convert("RGB")

            # resize stabil seperti laptop lama
            img = img.resize((600, 600))

            # simpan foto
            path = os.path.join(
                UPLOAD_FOLDER,
                fname
            )
            img.save(path)

            # preprocess OCR (sama seperti laptop lama)
            img_np = np.array(img)
            processed = advanced_pre_process(img_np)

            # OCR
            result = reader.readtext(
                processed,
                detail=0
            )

            # ambil angka meteran
            angka = robust_extract_logic(result)

        except Exception as e:
            st.error(f"OCR Error: {e}")
            angka = "Error OCR"

        st.divider()

        colA, colB = st.columns([1, 1])

        with colA:
            st.image(
                img,
                caption="Foto Meteran"
            )

        with colB:
            tgl = st.text_input(
                "Tanggal",
                value=now.strftime("%d-%m-%Y"),
                key=f"tgl_{fname}"
            )

            angka_final = st.text_input(
                "Angka Meteran (Edit jika salah)",
                value=str(angka),
                key=f"angka_{fname}"
            )

            if st.button(
                "✅ KONFIRMASI & SIMPAN",
                key=f"save_{fname}"
            ):

                if angka_final.strip() == "":
                    st.warning("Angka kosong!")
                    st.stop()

                if not re.match(
                    r"^\d+(\.\d+)?$",
                    angka_final
                ):
                    st.warning("Format angka salah!")
                    st.stop()

                data_baru = pd.DataFrame([{
                    "Tanggal": tgl,
                    "Nama Meteran": nama_meteran,
                    "Angka Meteran": angka_final,
                    "Foto": fname
                }])

                df_old = load_data(
                    jenis_meteran
                )

                df = pd.concat(
                    [df_old, data_baru],
                    ignore_index=True
                )

                save_data(
                    df,
                    jenis_meteran
                )

                st.success(
                    f"Data {nama_meteran} berhasil disimpan!"
                )

                st.rerun()


# ================= HISTORI =================
st.divider()
st.subheader("📊 Histori Pencatatan")

df_db = load_data(jenis_meteran)

if not df_db.empty:

    st.dataframe(
        df_db[
            [
                "Tanggal",
                "Nama Meteran",
                "Angka Meteran"
            ]
        ],
        use_container_width=True
    )

    # preview foto sesuai data
    for i, row in df_db.iterrows():
        with st.expander(
            f"{row['Tanggal']} | {row['Nama Meteran']}"
        ):
            st.write(
                f"**Angka Meteran:** {row['Angka Meteran']}"
            )

            foto_file = row.get(
                "Foto",
                ""
            )

            path = os.path.join(
                UPLOAD_FOLDER,
                str(foto_file)
            )

            if (
                foto_file
                and os.path.exists(path)
            ):
                st.image(
                    path,
                    width=300
                )
            else:
                st.info(
                    "Foto tidak tersedia"
                )

    # ================= HAPUS DATA =================
    st.divider()
    st.subheader("🗑️ Hapus Data")

    idx = st.selectbox(
        "Pilih data",
        df_db.index,
        format_func=lambda x:
            f"{df_db.loc[x, 'Tanggal']} | "
            f"{df_db.loc[x, 'Nama Meteran']} | "
            f"{df_db.loc[x, 'Angka Meteran']}"
    )

    if st.button("❌ Hapus Data"):
        df_db = df_db.drop(idx)

        save_data(
            df_db,
            jenis_meteran
        )

        st.warning(
            "Data berhasil dihapus!"
        )

        st.rerun()

    # ================= DOWNLOAD EXCEL =================
    st.divider()

    file_path = (
        EXCEL_AIR
        if jenis_meteran == "Air"
        else EXCEL_LISTRIK
    )

    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            st.download_button(
                "📥 Download Excel",
                data=f,
                file_name=file_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


else:
    st.info(
        "Belum ada data tersimpan."
    )