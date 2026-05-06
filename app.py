import streamlit as st
import json
import base64
import requests
from datetime import datetime
from PIL import Image, ImageEnhance
import io
import openpyxl
import os

# --- إعدادات الصفحة ---
st.set_page_config(page_title="ONCF | Scanner BL", page_icon="🚆", layout="centered")

# --- 🔑 بلاصة API KEY ديالك ---
API_KEY = "AIzaSyBfmryF5dRnwFLtMSJJknFPF3d4Ky_SDzo" 
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key={API_KEY}"

# --- البرومبت ديالك 100% ---
PROMPT = """You are an expert OCR and document analysis system specialized in ONCF "Bordereau de Livraison" (BL).
Return ONLY a valid JSON object.
{
  "documentNumber": null, "date": null, "reference": null,
  "sender": {"name": null, "department": null},
  "recipient": {"name": null, "department": null},
  "deliveryMode": null, "subject": null, "observations": null,
  "items": [{"lineNumber":1,"code":null,"designation":null,"quantity":null,"unit":null,"status":null,"observations":null}],
  "confidenceNote": null
}
Extract now:"""

def process_image(image):
    img = image.convert('L')
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = ImageEnhance.Sharpness(img).enhance(1.5)
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=90)
    return base64.b64encode(img_byte_arr.getvalue()).decode("utf-8")

def extract_data(base64_image):
    payload = {
        "contents": [{"parts": [{"text": PROMPT}, {"inline_data": {"mime_type": "image/jpeg", "data": base64_image}}]}],
        "generationConfig": {"temperature": 0.1, "topP": 0.9}
    }
    response = requests.post(GEMINI_URL, json=payload, headers={"Content-Type": "application/json"})
    if response.status_code == 200:
        raw_text = response.json()["candidates"][0]["content"]["parts"][0]["text"]
        # هادا هو السطر لي تصلح
        clean_json = raw_text.strip().replace("```json", "").replace("```", "").strip()
        return json.loads(clean_json)
    else:
        st.error(f"Erreur API: {response.text}")
        return None

# --- اللوجيك ديالك لدمج البيانات فالملف الأصلي ---
def update_master_excel(data, template_path="BLs Etablissements ONCF 01-01-2026.xlsx"):
    if not os.path.exists(template_path):
        st.error(f"⚠️ الملف الأصلي '{template_path}' ما كاينش! تأكد أنك حطيتيه فـ GitHub.")
        return None

    wb = openpyxl.load_workbook(template_path)

    def get_nom(champ):
        if not champ: return ""
        if isinstance(champ, str): return champ
        if isinstance(champ, dict):
            return champ.get('name') or champ.get('nom') or champ.get('code') or list(champ.values())[0] or ""
        return str(champ)

    exp = get_nom(data.get('sender') or data.get('Expéditeur'))
    dest = get_nom(data.get('recipient') or data.get('Destinataire'))

    expediteur = str(exp).strip().upper() if exp else "INCONNU"
    destinataire = str(dest).strip().upper() if dest else "INCONNU"

    # تقليب على الورقة
    sheet_name1 = f"{expediteur}->{destinataire}"
    sheet_name2 = f"{expediteur}-{destinataire}"
    
    sheet = None
    if sheet_name1 in wb.sheetnames:
        sheet = wb[sheet_name1]
    elif sheet_name2 in wb.sheetnames:
        sheet = wb[sheet_name2]

    if not sheet:
        st.error(f"⚠️ مالقيناش الورقة ديال '{sheet_name1}' ولا '{sheet_name2}' فالملف الأصلي!")
        return None

    date_bl = data.get('date') or data.get('Date') or ""
    num_bl = data.get('reference') or data.get('documentNumber') or ""

    # نقلبو على آخر سطر عامر بصح
    real_last_row = 6
    for i in range(1, sheet.max_row + 1):
        if sheet.cell(row=i, column=1).value or sheet.cell(row=i, column=2).value:
            real_last_row = i

    # نكتبو السلعة مباشرة تحت آخر سطر
    articles = data.get('items') or []
    for item in articles:
        real_last_row += 1
        final_date = item.get('date') or date_bl
        final_num_bl = item.get('reference') or num_bl or item.get('documentNumber') or ""

        sheet.cell(row=real_last_row, column=1, value=final_date)
        sheet.cell(row=real_last_row, column=2, value=item.get('designation', ""))
        sheet.cell(row=real_last_row, column=3, value=item.get('code') or item.get('n_organe') or "")
        sheet.cell(row=real_last_row, column=4, value=item.get('quantity') or 1)
        sheet.cell(row=real_last_row, column=5, value=final_num_bl)
        sheet.cell(row=real_last_row, column=6, value=item.get('observations') or "")

    virtual_workbook = io.BytesIO()
    wb.save(virtual_workbook)
    return virtual_workbook.getvalue()

# --- واجهة المستخدم ---
st.markdown("<h1 style='text-align: center; color: #002E5D;'>🚆 ONCF | Injection BL directe</h1>", unsafe_allow_html=True)
st.divider()

col1, col2 = st.columns(2)
with col1: upload_img = st.file_uploader("📁 Galerie", type=["jpg", "jpeg", "png"])
with col2: camera_img = st.camera_input("📸 Caméra")

image_to_process = camera_img if camera_img else upload_img

if image_to_process:
    st.image(image_to_process, caption="Image prête", use_column_width=True)
    
    if st.button("⚡ Extraire & Injecter dans le fichier de travail", use_container_width=True, type="primary"):
        with st.spinner("⏳ Traitement par IA en cours..."):
            img = Image.open(image_to_process)
            b64_img = process_image(img)
            data = extract_data(b64_img)
            
            if data:
                st.success("✅ Extraction réussie ! Injection en cours...")
                st.json(data) # باش تشوف البيانات لي تجبدات
                
                updated_excel = update_master_excel(data)
                
                if updated_excel:
                    st.download_button(
                        label="📥 Télécharger le fichier de travail mis à jour",
                        data=updated_excel,
                        file_name=f"BLs_Mis_a_jour_{datetime.now().strftime('%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
