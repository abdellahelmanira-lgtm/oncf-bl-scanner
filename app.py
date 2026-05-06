import streamlit as st
import json
import base64
import requests
from datetime import datetime
from PIL import Image, ImageEnhance
import io
import openpyxl
import os
import re

# --- إعدادات الصفحة ---
st.set_page_config(page_title="ONCF | Scanner BL", page_icon="🚆", layout="centered")

# --- 🔑 بلاصة API KEY ديالك ---
API_KEY = "VOTRE_CLE_GEMINI_ICI" 
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key={API_KEY}"

# --- البرومبت المترجم بالحرف من geminiService.js ---
EXTRACTION_PROMPT = """
You are an expert OCR and document analysis system specialized in ONCF (Office National des Chemins de Fer du Maroc) delivery documents called "Bordereau de Livraison" (BL).

Your task: Carefully read the provided document image and extract ALL visible information into a strictly structured JSON object.

CRITICAL RULES:
1. Return ONLY a valid JSON object — no markdown, no explanation, no code blocks.
2. If a field is not visible or not found, use null.
3. For the items table, extract EVERY row visible — do not skip any.
4. Preserve original French text exactly as written.
5. For dates, preserve the original format found in the document.
6. Numbers should be returned as numbers (not strings) where they represent quantities.
7. CRITICAL: For 'sender.name' and 'recipient.name', NEVER use the full phrase "OFFICE NATIONAL DES CHEMINS DE FER". ALWAYS extract ONLY the specific short abbreviation/code (e.g., EMC, EMIC, EMS, TMK, CLC, TMFRET) found in the header, text, or stamps.
8. STRICT SPLIT RULE: Check the "Quantité" column FIRST. If the quantity is 1 (or "01"), you MUST NEVER split the item into multiple rows, no matter how many lines the observation text takes. Only split if Quantité > 1 AND you see multiple distinct serial numbers.
9. SERIAL NUMBER (CODE) EXTRACTION: The text on these documents often uses a printed script font that resembles cursive handwriting. Do not ignore it. Look carefully in the "DETAIL DES MARCHANDISES" column, directly next to or below the main designation (e.g., next to "Pantographe"). 
10. When you find a value like "N°FS52/92" written in this script font, extract it into the "code" field, but remove the "N°" prefix (e.g., output "FS52/92" or "2459/18").
11. MULTI-LINE OBSERVATIONS: If an item has a quantity of 1, combine all lines of text in the "OBSERVATIONS" column (e.g., "Bras cassé oxydation coincement mecanique Z2M116") into a single string for that one item.

{
  "documentNumber": "string or null",
  "date": "string or null",
  "reference": "string or null",
  "sender": {
    "name": "string or null",
    "department": "string or null",
    "address": "string or null"
  },
  "recipient": {
    "name": "string or null",
    "department": "string or null",
    "address": "string or null"
  },
  "deliveryMode": "string or null",
  "subject": "string or null",
  "observations": "string or null",
  "items": [
    {
      "lineNumber": "number or string",
      "code": "string or null",
      "designation": "string or null",
      "quantity": "number or null",
      "unit": "string or null",
      "status": "string or null",
      "observations": "string or null"
    }
  ],
  "signatures": {
    "sender": "string or null",
    "recipient": "string or null"
  },
  "confidenceNote": "string describing any unclear areas or assumptions made"
}

Now extract all data from the document image:
"""

def process_image(image):
    img = image.convert('L')
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = ImageEnhance.Sharpness(img).enhance(1.5)
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=90)
    return base64.b64encode(img_byte_arr.getvalue()).decode("utf-8")

def extract_data(base64_image):
    payload = {
        "contents": [{"parts": [{"text": EXTRACTION_PROMPT}, {"inline_data": {"mime_type": "image/jpeg", "data": base64_image}}]}],
        "generationConfig": {"temperature": 0.1, "topP": 0.9}
    }
    response = requests.post(GEMINI_URL, json=payload, headers={"Content-Type": "application/json"})
    
    if response.status_code == 200:
        raw_text = response.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
        
        # Strip any accidental markdown code fences[cite: 2, 4]
        raw_text = raw_text.replace("```json", "").replace("```", "").strip()
        
        # Attempt direct parse, then fallback to regex matching JSON[cite: 2, 4]
        try:
            return json.loads(raw_text)
        except json.JSONDecodeError:
            match = re.search(r'\{[\s\S]*\}', raw_text)
            if match:
                try:
                    return json.loads(match.group(0))
                except json.JSONDecodeError:
                    st.error("Erreur de parsing JSON (Fallback échoué).")
                    return None
            st.error("Gemini n'a pas retourné un JSON valide.")
            return None
    else:
        st.error(f"Erreur API: {response.text}")
        return None

# --- اللوجيك ديالك لدمج البيانات فالملف الأصلي ---
def update_master_excel(data, template_path="BLs Etablissements ONCF 01-01-2026.xlsx"):
    if not os.path.exists(template_path):
        st.error(f"⚠️ الملف الأصلي '{template_path}' ما كاينش! تأكد أنك حطيتيه فـ GitHub.")
        return None

    wb = openpyxl.load_workbook(template_path)

    # 4. كنجبدو المُرسل والمُستقبل باش نعرفو آين ورقة غنختارو[cite: 1, 3]
    def get_nom(champ):
        if not champ: return ""
        if isinstance(champ, str): return champ
        if isinstance(champ, dict):
            # يلا كان صندوق، كنجبدو منو غير السمية[cite: 1, 3]
            return champ.get('nom') or champ.get('name') or champ.get('entreprise') or champ.get('code') or (list(champ.values())[0] if champ else "")
        return str(champ)

    sender_data = data.get('sender') or data.get('Expéditeur') or data.get('expéditeur') or data.get('expediteur')
    recipient_data = data.get('recipient') or data.get('Destinataire') or data.get('destinataire')

    exp = get_nom(sender_data)
    dest = get_nom(recipient_data)

    expediteur = str(exp).strip().upper() if exp else "INCONNU"
    destinataire = str(dest).strip().upper() if dest else "INCONNU"

    # كنجربو نقلبو على الورقة بالسهم (->) وبلا سهم (-) حيت عندك بجوج فالملف[cite: 1, 3]
    sheet_name1 = f"{expediteur}->{destinataire}"
    sheet_name2 = f"{expediteur}-{destinataire}"
    
    sheet = None
    if sheet_name1 in wb.sheetnames:
        sheet = wb[sheet_name1]
    elif sheet_name2 in wb.sheetnames:
        sheet = wb[sheet_name2]

    # يلا مالقيناش الورقة[cite: 1, 3]
    if not sheet:
        st.error(f"⚠️ مالقيناش الورقة ديال '{sheet_name1}' ولا '{sheet_name2}' فالملف الأصلي! تأكد من المُرسل والمُستقبل.")
        return None

    # 5. كنزيدو السلعة لي تسكانات فالسطر الأخير الخاوي فديك الورقة[cite: 1, 3]
    articles = data.get('items') or data.get('articles') or []
    
    # جلب التاريخ ورقم البوردرو بذكاء باش ما يزݣلهمش[cite: 1, 3]
    date_bl = data.get('date') or data.get('date_bl') or data.get('Date') or ""
    num_bl = data.get('reference') or data.get('documentNumber') or data.get('n_bl') or ""

    # أ. كنقلبو على آخر سطر عامر بصح (كنبداو من السطر 6 لي فيه العناوين)[cite: 1, 3]
    real_last_row = 6
    for i in range(1, sheet.max_row + 1):
        cell_a = sheet.cell(row=i, column=1).value
        cell_b = sheet.cell(row=i, column=2).value
        if cell_a or cell_b:
            real_last_row = i

    # ب. كنبداو نكتبو السلعة مباشرة تحت آخر سطر عامر لقيناه[cite: 1, 3]
    for item in articles:
        real_last_row += 1
        
        # 1. Kenakhdo la date mn l'article ila kant, sinon kenakhdo dyal l'BL kamel[cite: 1, 3]
        final_date = item.get('date') or date_bl
        
        # 2. Kenakhdo numBL (Reference) hwa l'awal kifma darna f l'marra li fatet[cite: 1, 3]
        final_num_bl = item.get('reference') or num_bl or item.get('documentNumber') or ""

        sheet.cell(row=real_last_row, column=1, value=final_date)
        sheet.cell(row=real_last_row, column=2, value=item.get('designation', ""))
        sheet.cell(row=real_last_row, column=3, value=item.get('code') or item.get('n_organe') or item.get('organe') or "")
        qty = item.get('quantity')
        sheet.cell(row=real_last_row, column=4, value=qty if qty is not None else 1) # Ila kant split, qté dima 1[cite: 1, 3]
        sheet.cell(row=real_last_row, column=5, value=final_num_bl)
        sheet.cell(row=real_last_row, column=6, value=item.get('observations') or item.get('observation') or "")

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
                st.json(data)
                
                updated_excel = update_master_excel(data)
                
                if updated_excel:
                    st.download_button(
                        label="📥 Télécharger le fichier de travail mis à jour",
                        data=updated_excel,
                        file_name=f"BL_ONCF_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
