import streamlit as st
import json
import base64
import requests
from datetime import datetime
from PIL import Image, ImageEnhance
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# --- إعدادات الصفحة والديزاين ---
st.set_page_config(page_title="ONCF | Scanner BL", page_icon="🚆", layout="centered")

# --- 🔑 بلاصة API KEY ديالك ---
API_KEY = "AIzaSyBfmryF5dRnwFLtMSJJknFPF3d4Ky_SDzo" 
# ------------------------------

GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={API_KEY}"

PROMPT = """You are an expert OCR for ONCF "Bordereau de Livraison" documents.
Return ONLY valid JSON, no markdown. Missing=null. quantity must be JSON number.
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
    # معالجة الصورة (أبيض وأسود + تباين)
    img = image.convert('L')
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = ImageEnhance.Sharpness(img).enhance(1.5)
    
    # تحويل الصورة لـ Base64 باش تمشي لـ Gemini
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
        clean_json = raw_text.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
        return json.loads(clean_json)
    else:
        st.error(f"Erreur API: {response.text}")
        return None

def create_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Résumé"
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 40
    
    ws.merge_cells("A1:B1")
    c1 = ws.cell(row=1, column=1, value="ONCF - BORDEREAU DE LIVRAISON")
    c1.font = Font(bold=True, color="FFFFFF", size=14)
    c1.fill = PatternFill("solid", fgColor="002E5D")
    c1.alignment = Alignment(horizontal="center")
    
    fields = [
        ("N° BL", data.get("documentNumber")), 
        ("Date", data.get("date")), 
        ("Expéditeur", (data.get("sender") or {}).get("name")), 
        ("Destinataire", (data.get("recipient") or {}).get("name"))
    ]
    
    for r, (lbl, val) in enumerate(fields, 3):
        ws.cell(row=r, column=1, value=lbl).font = Font(bold=True)
        ws.cell(row=r, column=2, value=val or "—")
        
    virtual_workbook = io.BytesIO()
    wb.save(virtual_workbook)
    return virtual_workbook.getvalue()

# --- واجهة المستخدم (الفرونت إند) ---
st.markdown("<h1 style='text-align: center; color: #002E5D;'>🚆 ONCF | Scanner BL</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #F07D00; font-weight: bold;'>Extraction Intelligente par IA</p>", unsafe_allow_html=True)
st.divider()

# الأزرار ديال الكاميرا والغاليري (مدمجين مع المتصفح)
col1, col2 = st.columns(2)
with col1:
    upload_img = st.file_uploader("📁 Depuis la galerie", type=["jpg", "jpeg", "png"])
with col2:
    camera_img = st.camera_input("📸 Prendre une photo")

image_to_process = camera_img if camera_img else upload_img

if image_to_process:
    st.image(image_to_process, caption="Image prête", use_column_width=True)
    
    if st.button("⚡ Extraire les données (Excel)", use_container_width=True, type="primary"):
        with st.spinner("⏳ Analyse par IA en cours..."):
            img = Image.open(image_to_process)
            b64_img = process_image(img)
            data = extract_data(b64_img)
            
            if data:
                st.success("✅ Extraction réussie !")
                excel_data = create_excel(data)
                
                # زر التحميل المباشر للـ Excel
                st.download_button(
                    label="📥 Télécharger le fichier Excel",
                    data=excel_data,
                    file_name=f"BL_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
