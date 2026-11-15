import streamlit as st
import google.generativeai as genai
from PIL import Image
from docx import Document
import openpyxl
from pptx import Presentation
import pypdf
import requests
import io

st.set_page_config(page_title="My All-in-One AI Agent", page_icon="🤖", layout="wide")

# --- ১. সিকিউরিটি পাসওয়ার্ড ও অটো কী লোড ---
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "12345")
GEMINI_KEY = st.secrets.get("GEMINI_KEY", "")
HF_TOKEN = st.secrets.get("HF_TOKEN", "")
GAS_WEBHOOK = st.secrets.get("GAS_WEBHOOK", "")
TG_BOT_TOKEN = st.secrets.get("TG_BOT_TOKEN", "")

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 প্রাইভেট AI Agent লগইন")
    pwd = st.text_input("আপনার গোপন পাসওয়ার্ড দিন:", type="password")
    if st.button("প্রবেশ করুন"):
        if pwd == APP_PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("ভুল পাসওয়ার্ড!")
    st.stop()

# --- ২. সাইডবার কনফিগারেশন ---
with st.sidebar:
    st.header("⚙️ এজেন্ট কন্ট্রোল প্যানেল")
    gemini_api = st.text_input("Gemini API Key:", value=GEMINI_KEY, type="password")
    hf_api = st.text_input("HuggingFace Token:", value=HF_TOKEN, type="password")
    webhook_url = st.text_input("Google Script URL:", value=GAS_WEBHOOK)
    tg_token = st.text_input("Telegram Bot Token:", value=TG_BOT_TOKEN, type="password")
    
    if gemini_api:
        genai.configure(api_key=gemini_api)
        model = genai.GenerativeModel("gemini-2.5-flash")

st.title("🤖 আমার অল-ইন-ওয়ান পার্সোনাল ও অফিস AI Agent")

# অটো-ফিল ট্যাবে ডেটা সহজে নেওয়ার জন্য সেশন মেমোরি
if "last_extracted_text" not in st.session_state:
    st.session_state.last_extracted_text = ""

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📄 ফাইল ও হাতের লেখা কনভার্টার",
    "🌐 ওয়েবসাইট অটো-ফিলাপ",
    "🎨 ফ্রি আনসেন্সরড ছবি ও ভিডিও",
    "✉️ কাস্টম মেইল ও মেসেজ",
    "🔗 Gmail, Sheets ও মেসেঞ্জার কানেক্ট"
])

# ==========================================
# TAB 1: হাতের লেখা ও যেকোনো ফাইল কনভার্টার (Bangla <-> English)
# ==========================================
with tab1:
    st.subheader("১. হাতের লেখা বা যেকোনো ফাইল ➔ বাংলা/ইংরেজি Word, Excel, Slide ও Text")
    uploaded_file = st.file_uploader(
        "হাতের লেখার ছবি বা ফাইল আপলোড করুন (JPG, PNG, PDF, DOCX, XLSX, TXT)", 
        type=["jpg", "jpeg", "png", "pdf", "docx", "xlsx", "txt"]
    )
    
    col1, col2 = st.columns(2)
    with col1:
        trans_mode = st.selectbox("ভাষা ও অনুবাদ মোড:", [
            "১. ফাইলে হুবহু যা আছে তাই (Exact Same)",
            "২. বাংলা ➔ ইংরেজি (Bangla to English)",
            "৩. ইংরেজি ➔ বাংলা (English to Bangla)",
            "৪. বাংলা ও ইংরেজি পাশাপাশি (Bilingual)"
        ])
    with col2:
        out_format = st.selectbox("কোন ফাইল ডাউনলোড করবেন?", [
            "Word (.docx)", 
            "Excel (.xlsx)", 
            "Slide / PowerPoint (.pptx)",
            "Text (.txt)"
        ])

    if uploaded_file and st.button("🔄 রিড ও কনভার্ট করুন"):
        if not gemini_api:
            st.error("বাম পাশের মেনুতে আপনার Gemini API Key দিন!")
        else:
            with st.spinner("AI নিখুঁতভাবে ফাইলটি পড়ছে ও রূপান্তর করছে..."):
                ext = uploaded_file.name.split(".")[-1].lower()
                content_input = []

                if ext in ["jpg", "jpeg", "png"]:
                    img = Image.open(uploaded_file)
                    st.image(img, width=250, caption="আপলোড করা ফাইল")
                    content_input.append(img)
                elif ext == "pdf":
                    reader = pypdf.PdfReader(uploaded_file)
                    content_input.append("\n".join([p.extract_text() or "" for p in reader.pages]))
                elif ext == "docx":
                    d = Document(uploaded_file)
                    content_input.append("\n".join([p.text for p in d.paragraphs]))
                elif ext == "xlsx":
                    wb_in = openpyxl.load_workbook(uploaded_file)
                    sh = wb_in.active
                    content_input.append("\n".join([" | ".join([str(c.value or "") for c in r]) for r in sh.iter_rows()]))
                else:
                    content_input.append(uploaded_file.read().decode("utf-8", errors="ignore"))

                prompt = f"""You are an expert OCR and Bangla-English translator.
                1. Extract all text from the input with 100% accuracy.
                2. Apply translation rule: {trans_mode}.
                3. If target is Excel ({out_format}) or data is structured/tabular, separate columns with '|' on each row.
                4. Return ONLY the final converted content without any extra chat."""
                
                content_input.insert(0, prompt)
                res = model.generate_content(content_input)
                st.session_state.last_extracted_text = res.text

    if st.session_state.last_extracted_text:
        final_text = st.text_area("শনাক্ত করা লেখা (প্রয়োজনে এডিট করুন):", st.session_state.last_extracted_text, height=220)
        st.session_state.last_extracted_text = final_text

        if out_format == "Word (.docx)":
            doc = Document()
            for line in final_text.split("\n"):
                doc.add_paragraph(line)
            buf = io.BytesIO()
            doc.save(buf)
            st.download_button("📥 Word (.docx) ডাউনলোড করুন", buf.getvalue(), "converted.docx")

        elif out_format == "Excel (.xlsx)":
            wb = openpyxl.Workbook()
            ws = wb.active
            for r_idx, line in enumerate(final_text.split("\n"), 1):
                if line.strip():
                    cols = [c.strip() for c in line.split("|")] if "|" in line else [line]
                    for c_idx, val in enumerate(cols, 1):
                        ws.cell(row=r_idx, column=c_idx, value=val)
            buf = io.BytesIO()
            wb.save(buf)
            st.download_button("📥 Excel (.xlsx) ডাউনলোড করুন", buf.getvalue(), "converted.xlsx")

        elif out_format.startswith("Slide"):
            prs = Presentation()
            lines = [l.strip() for l in final_text.split("\n") if l.strip()]
            for i in range(0, len(lines), 6):
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = f"Slide {(i // 6) + 1}"
                slide.placeholders[1].text = "\n".join(lines[i:i + 6])
            buf = io.BytesIO()
            prs.save(buf)
            st.download_button("📥 Slide (.pptx) ডাউনলোড করুন", buf.getvalue(), "slides.pptx")

        elif out_format == "Text (.txt)":
            st.download_button("📥 Text (.txt) ডাউনলোড করুন", final_text.encode("utf-8"), "converted.txt")

# ==========================================
# TAB 2: ওয়েবসাইট ও অফিস সিস্টেমে অটো ফর্ম ফিলাপ
# ==========================================
with tab2:
    st.subheader("২. যেকোনো ওয়েবসাইট বা সিস্টেমে ১-ক্লিকে অটো ফর্ম ফিলাপ")
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        fill_lang = st.radio("ফর্মে কোন ভাষায় ফিলাপ হবে?", ["হুবহু যা আছে", "বাংলায় (Bangla)", "ইংরেজিতে (English)"])
        fill_method = st.selectbox("পদ্ধতি:", [
            "১-ক্লিক ব্রাউজার অটো-ফিল কোড (মোবাইল ও ল্যাপটপ)",
            "Python Playwright রোবট স্ক্রিপ্ট (বাল্ক এন্ট্রির জন্য)"
        ])
    with col_w2:
        form_img = st.file_uploader("ওয়েবসাইটের খালি ফর্মের স্ক্রিনশট দিন (ঐচ্ছিক):", type=["png", "jpg", "jpeg"])
        field_names = st.text_input("অথবা ফর্মের ঘরগুলোর নাম লিখুন (যেমন: Name, Mobile, Amount):")

    data_for_form = st.text_area("ফিলাপ করার ডেটা (১নং ট্যাব থেকে অটো চলে আসবে বা এখানে পেস্ট করুন):", value=st.session_state.last_extracted_text, height=150)

    if st.button("⚡ অটো-ফিল কোড তৈরি করুন") and gemini_api:
        with st.spinner("অটো-ফিল স্ক্রিপ্ট তৈরি হচ্ছে..."):
            p_fill = f"""Convert this data: '{data_for_form}' into '{fill_lang}' language and map it to form fields '{field_names}'.
            Generate:
            1. A smart JavaScript Console snippet that finds matching input/select/textarea fields by label, placeholder, or name and fills them automatically (triggering input/change events).
            2. A single-line Bookmarklet `javascript:(function(){{...}})();` for 1-click mobile/PC auto-fill.
            3. If Playwright is requested ({fill_method}), provide the Python script.
            Explain briefly in Bangla how to use it."""
            inputs = [p_fill, Image.open(form_img)] if form_img else [p_fill]
            st.markdown(model.generate_content(inputs).text)

# ==========================================
# TAB 3: ফ্রি আনলিমিটেড ও Uncensored ছবি ও ভিডিও
# ==========================================
with tab3:
    st.subheader("৩. ফ্রি আনলিমিটেড ছবি ও ভিডিও জেনারেটর (No Restriction)")
    engine = st.selectbox("ইঞ্জিন বেছে নিন:", [
        "Engine 1: Pollinations Flux (100% Free, Instant Image)",
        "Engine 2: HuggingFace Model (Custom Image)",
        "Engine 3: Google Colab Private GPU (100% Uncensored Image & Video)"
    ])
    
    c1, c2 = st.columns(2)
    with c1:
        m_type = st.radio("কী বানাবেন?", ["ছবি (image)", "ভিডিও (video)"])
    with c2:
        seed_val = st.number_input("Same স্টাইল রাখার Seed নম্বর:", value=42)

    m_prompt = st.text_area("আপনার বিস্তারিত Prompt লিখুন (যেকোনো ধরনের 18+/NSFW, বোল্ড বা সাধারণ প্রম্পট):")
    neg_prompt = st.text_input("Negative Prompt:", "blurry, ugly, deformed, bad anatomy, extra limbs, censored, watermark")

    if engine.startswith("Engine 1"):
        if st.button("🎨 ছবি তৈরি করুন"):
            url = f"https://image.pollinations.ai/prompt/{requests.utils.quote(m_prompt)}?seed={seed_val}&width=1024&height=1024&nologo=true"
            img_bytes = requests.get(url).content
            st.image(img_bytes)
            st.download_button("📥 ডাউনলোড করুন", img_bytes, "image.jpg")

    elif engine.startswith("Engine 2"):
        hf_model = st.text_input("HuggingFace Model ID:", "black-forest-labs/FLUX.1-schnell")
        if st.button("🎨 HF ছবি তৈরি করুন"):
            r = requests.post(f"https://api-inference.huggingface.co/models/{hf_model}", headers={"Authorization": f"Bearer {hf_api}"}, json={"inputs": m_prompt})
            if r.status_code == 200:
                st.image(r.content)
                st.download_button("📥 ডাউনলোড করুন", r.content, "hf.png")
            else:
                st.error(r.text)

    elif engine.startswith("Engine 3"):
        colab_url = st.text_input("আপনার Google Colab লিংক (.trycloudflare.com) দিন:")
        if st.button("🔥 Uncensored ছবি/ভিডিও তৈরি করুন"):
            mode_str = "video" if "video" in m_type else "image"
            with st.spinner("Colab GPU-তে তৈরি হচ্ছে..."):
                r = requests.post(f"{colab_url.rstrip('/')}/generate", json={"prompt": m_prompt, "negative_prompt": neg_prompt, "seed": int(seed_val), "media_type": mode_str}, timeout=180)
                if r.status_code == 200:
                    if mode_str == "image":
                        st.image(r.content)
                        st.download_button("📥 ছবি ডাউনলোড", r.content, "uncensored.png")
                    else:
                        st.video(r.content)
                        st.download_button("📥 ভিডিও ডাউনলোড", r.content, "uncensored.mp4")
                else:
                    st.error(r.text)

# ==========================================
# TAB 4: কাস্টম নির্দেশনা (System Prompt) অনুযায়ী মেইল ও মেসেজ
# ==========================================
with tab4:
    st.subheader("৪. স্মার্ট মেইল ও মেসেজ ড্রাফটার")
    sys_p = st.text_area("আপনার স্থায়ী নির্দেশনা (System Prompt):", "তুমি আমার অফিসিয়াল ও পার্সোনাল অ্যাসিস্ট্যান্ট। আমার দেওয়া তথ্য ও ফাইল অনুযায়ী নির্ভুল বাংলা বা ইংরেজিতে ইমেইল, রিপোর্ট বা মেসেজ লিখে দেবে।")
    att_file = st.file_uploader("রেফারেন্স ফাইল বা ছবি (ঐচ্ছিক):", type=["png", "jpg", "pdf", "txt"])
    u_task = st.text_area("কী লিখতে হবে বলুন:")

    if st.button("✍️ ড্রাফট তৈরি করুন") and gemini_api:
        inp = [f"System Prompt: {sys_p}\n\nTask: {u_task}"]
        if att_file:
            if att_file.type.startswith("image"):
                inp.append(Image.open(att_file))
            elif att_file.name.endswith(".pdf"):
                inp.append("\n".join([p.extract_text() or "" for p in pypdf.PdfReader(att_file).pages]))
            else:
                inp.append(att_file.read().decode("utf-8", errors="ignore"))
        res = model.generate_content(inp)
        st.text_area("তৈরি করা মেইল/মেসেজ:", res.text, height=220)

# ==========================================
# TAB 5: Gmail, Sheets, Calendar, Telegram ও WhatsApp কানেক্ট
# ==========================================
with tab5:
    st.subheader("৫. অ্যাপ অটোমেশন ও মেসেজিং কন্ট্রোল")
    act = st.selectbox("কোন কাজটি করতে চান?", [
        "Gmail-এ সরাসরি ইমেইল পাঠান",
        "Google Sheets-এ নতুন রো (Row) হিসেবে সেভ করুন",
        "Google Calendar-এ রিমাইন্ডার/ইভেন্ট সেট করুন",
        "Telegram Bot-এ মেসেজ/নোটিফিকেশন পাঠান",
        "WhatsApp-এ মেসেজ পাঠান (Direct Link)"
    ])
    
    to_val = st.text_input("প্রাপকের Email / Telegram Chat ID / WhatsApp নম্বর (দেশের কোডসহ, যেমন 88017...):")
    sub_val = st.text_input("সাবজেক্ট বা শিরোনাম:")
    body_val = st.text_area("মূল মেসেজ বা ডেটা (Sheets-এর জন্য কমা দিয়ে লিখুন):", value=st.session_state.last_extracted_text)

    if st.button("🚀 এক ক্লিকে পাঠান / সেভ করুন"):
        if act.startswith("Telegram"):
            if not tg_token:
                st.error("বাম পাশে Telegram Bot Token দিন!")
            else:
                r = requests.post(f"https://api.telegram.org/bot{tg_token}/sendMessage", json={"chat_id": to_val, "text": f"{sub_val}\n\n{body_val}"})
                st.success("টেলিগ্রামে মেসেজ পাঠানো হয়েছে!" if r.status_code == 200 else r.text)
        elif act.startswith("WhatsApp"):
            wa_link = f"https://wa.me/{to_val.replace('+', '').strip()}?text={requests.utils.quote(body_val)}"
            st.markdown(f"👉 **[এখানে ক্লিক করে সরাসরি WhatsApp-এ মেসেজটি পাঠান]({wa_link})**")
        else:
            if not webhook_url:
                st.error("বাম পাশে Google Script URL দিন!")
            else:
                r = requests.post(webhook_url, json={"action": act, "to": to_val, "subject": sub_val, "body": body_val})
                st.success(f"সফল! সার্ভার রেসপন্স: {r.text}")
