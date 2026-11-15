import streamlit as st
import google.generativeai as genai
from PIL import Image
from docx import Document
import openpyxl
from pptx import Presentation
import pypdf
import requests
import io

# পেজ কনফিগারেশন
st.set_page_config(
    page_title="My All-in-One AI Agent",
    page_icon="🤖",
    layout="wide"
)

# নিরাপদে Secrets লোড করার ফাংশন (যাতে Secrets না দিলেও অ্যাপ ক্র্যাশ না করে)
def get_secret(key, default=""):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

APP_PASSWORD = get_secret("APP_PASSWORD", "12345")
GEMINI_KEY = get_secret("GEMINI_KEY", "")
HF_TOKEN = get_secret("HF_TOKEN", "")
GAS_WEBHOOK = get_secret("GAS_WEBHOOK", "")
TG_BOT_TOKEN = get_secret("TG_BOT_TOKEN", "")

# --- ১. সিকিউরিটি পাসওয়ার্ড লগইন ---
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔒 প্রাইভেট AI Agent লগইন")
    pwd = st.text_input("আপনার গোপন পাসওয়ার্ড দিন (ডিফল্ট: 12345):", type="password")
    if st.button("প্রবেশ করুন"):
        if pwd == APP_PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("ভুল পাসওয়ার্ড! সঠিক পাসওয়ার্ড দিন।")
    st.stop()

# --- ২. সাইডবার কন্ট্রোল প্যানেল ---
with st.sidebar:
    st.header("⚙️ এজেন্ট সেটিংস")
    gemini_api = st.text_input("Gemini API Key (Free):", value=GEMINI_KEY, type="password")
    hf_api = st.text_input("HuggingFace Token (Free):", value=HF_TOKEN, type="password")
    webhook_url = st.text_input("Google Script URL (Gmail/Sheets):", value=GAS_WEBHOOK)
    tg_token = st.text_input("Telegram Bot Token:", value=TG_BOT_TOKEN, type="password")
    
    model_name = st.selectbox(
        "AI Model সিলেক্ট করুন:",
        ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
    )
    
    if gemini_api:
        genai.configure(api_key=gemini_api)
        model = genai.GenerativeModel(model_name)
        
    if st.button("🚪 লগআউট"):
        st.session_state.auth = False
        st.rerun()

st.title("🤖 আমার অল-ইন-ওয়ান পার্সোনাল ও অফিস AI Agent")

# এক ট্যাব থেকে অন্য ট্যাবে ডেটা নেওয়ার মেমোরি
if "last_extracted_text" not in st.session_state:
    st.session_state.last_extracted_text = ""

# --- ৫টি মূল কাজের ট্যাব ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📄 হাতের লেখা ও ফাইল কনভার্টার",
    "🌐 ওয়েবসাইট অটো-ফিলাপ",
    "🎨 ফ্রি আনসেন্সরড ছবি ও ভিডিও",
    "✉️ কাস্টম মেইল ও মেসেজ রাইটার",
    "🔗 Gmail, Sheets ও মেসেঞ্জার কানেক্ট"
])

# =====================================================================
# TAB 1: হাতের লেখা বা যেকোনো ফাইল ➔ বাংলা/ইংরেজি Text, Word, Excel, Slide
# =====================================================================
with tab1:
    st.subheader("১. হাতের লেখা বা যেকোনো ফাইল ➔ বাংলা/ইংরেজি কনভার্টার")
    uploaded_file = st.file_uploader(
        "হাতের লেখার ছবি বা যেকোনো ফাইল আপলোড করুন (JPG, PNG, PDF, DOCX, XLSX, TXT)",
        type=["jpg", "jpeg", "png", "pdf", "docx", "xlsx", "txt"]
    )
    
    col1, col2 = st.columns(2)
    with col1:
        trans_mode = st.selectbox("ভাষা ও অনুবাদ অপশন:", [
            "১. ফাইলে হুবহু যা আছে তাই (Exact Same Language)",
            "২. বাংলা ➔ ইংরেজি (Translate Bangla to English)",
            "৩. ইংরেজি ➔ বাংলা (Translate English to Bangla)",
            "৪. বাংলা ও ইংরেজি দুটিই পাশাপাশি (Bilingual Output)"
        ])
    with col2:
        out_format = st.selectbox("কোন ফরম্যাটে ডাউনলোড করবেন?", [
            "Word (.docx)",
            "Excel (.xlsx)",
            "Slide / PowerPoint (.pptx)",
            "Text (.txt)"
        ])

    if uploaded_file and st.button("🔄 রিড ও কনভার্ট করুন"):
        if not gemini_api:
            st.error("অনুগ্রহ করে বাম পাশের সাইডবারে আপনার Gemini API Key দিন!")
        else:
            with st.spinner("AI নিখুঁতভাবে ফাইলটি পড়ছে এবং রূপান্তর করছে..."):
                try:
                    ext = uploaded_file.name.split(".")[-1].lower()
                    content_input = []

                    if ext in ["jpg", "jpeg", "png"]:
                        img = Image.open(uploaded_file)
                        st.image(img, width=280, caption="আপলোড করা ছবি")
                        content_input.append(img)
                    elif ext == "pdf":
                        reader = pypdf.PdfReader(uploaded_file)
                        raw_txt = "\n".join([p.extract_text() or "" for p in reader.pages])
                        content_input.append(raw_txt)
                    elif ext == "docx":
                        d = Document(uploaded_file)
                        raw_txt = "\n".join([p.text for p in d.paragraphs])
                        content_input.append(raw_txt)
                    elif ext == "xlsx":
                        wb_in = openpyxl.load_workbook(uploaded_file)
                        sh = wb_in.active
                        raw_txt = "\n".join(
                            [" | ".join([str(c.value or "") for c in r]) for r in sh.iter_rows()]
                        )
                        content_input.append(raw_txt)
                    else:
                        raw_txt = uploaded_file.read().decode("utf-8", errors="ignore")
                        content_input.append(raw_txt)

                    prompt = f"""You are an expert OCR and Bangla-English document converter.
                    1. Read and extract all text from the input with 100% accuracy without skipping anything.
                    2. Apply this language/translation rule strictly: {trans_mode}.
                    3. If the target format is Excel ({out_format}) or the data contains tables/lists, separate columns using the '|' symbol on each line.
                    4. Output ONLY the final converted text without any extra introductory or closing remarks."""
                    
                    content_input.insert(0, prompt)
                    res = model.generate_content(content_input)
                    st.session_state.last_extracted_text = res.text
                except Exception as e:
                    st.error(f"ফাইল প্রসেস করতে সমস্যা হয়েছে: {e}")

    if st.session_state.last_extracted_text:
        final_text = st.text_area(
            "শনাক্ত ও রূপান্তর করা লেখা (চাইলে এখানেই এডিট করতে পারেন):",
            st.session_state.last_extracted_text,
            height=240
        )
        st.session_state.last_extracted_text = final_text

        # ফাইল ডাউনলোড বাটন তৈরি
        if out_format == "Word (.docx)":
            doc = Document()
            for line in final_text.split("\n"):
                doc.add_paragraph(line)
            buf = io.BytesIO()
            doc.save(buf)
            st.download_button(
                "📥 Word (.docx) ফাইল ডাউনলোড করুন",
                data=buf.getvalue(),
                file_name="converted_document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

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
            st.download_button(
                "📥 Excel (.xlsx) ফাইল ডাউনলোড করুন",
                data=buf.getvalue(),
                file_name="converted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif out_format.startswith("Slide"):
            prs = Presentation()
            lines = [l.strip() for l in final_text.split("\n") if l.strip()]
            chunk_size = 6
            for i in range(0, max(1, len(lines)), chunk_size):
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = f"Slide {(i // chunk_size) + 1}"
                slide.placeholders[1].text = "\n".join(lines[i:i + chunk_size])
            buf = io.BytesIO()
            prs.save(buf)
            st.download_button(
                "📥 Slide / PowerPoint (.pptx) ডাউনলোড করুন",
                data=buf.getvalue(),
                file_name="converted_slides.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        elif out_format == "Text (.txt)":
            st.download_button(
                "📥 Text (.txt) ফাইল ডাউনলোড করুন",
                data=final_text.encode("utf-8"),
                file_name="converted_text.txt",
                mime="text/plain"
            )

# =====================================================================
# TAB 2: ওয়েবসাইট বা অফিস সিস্টেমে ১-ক্লিক অটো ফর্ম ফিলাপ
# =====================================================================
with tab2:
    st.subheader("২. যেকোনো ওয়েবসাইট বা সিস্টেমে অটো ফর্ম ফিলাপ")
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        fill_lang = st.radio(
            "ফর্মে কোন ভাষায় ফিলাপ করতে চান?",
            ["ফাইলে যেভাবে আছে হুবহু সেভাবে", "বাংলায় (Bangla)", "ইংরেজিতে (English)"]
        )
        fill_method = st.selectbox("অটো-ফিল পদ্ধতি:", [
            "Method 1: ১-ক্লিক ব্রাউজার অটো-ফিল (Console / মোবাইল বুকমার্কলেট)",
            "Method 2: ফর্মের স্ক্রিনশট দেখে স্মার্ট অটো-ফিল",
            "Method 3: Python Playwright বট (বাল্ক ডেটা এন্ট্রির জন্য)"
        ])
    with col_w2:
        form_img = st.file_uploader(
            "ওয়েবসাইটের খালি ফর্মের স্ক্রিনশট দিন (ঐচ্ছিক, নিখুঁত ম্যাচিংয়ের জন্য):",
            type=["png", "jpg", "jpeg"]
        )
        field_names = st.text_input(
            "অথবা ফর্মের ঘরগুলোর নাম লিখুন (কমা দিয়ে):",
            placeholder="যেমন: Name, Mobile, Designation, Address, Amount"
        )

    data_for_form = st.text_area(
        "যে ডেটা দিয়ে ফর্ম ফিলাপ করবেন (১নং ট্যাব থেকে অটো চলে আসবে বা এখানে পেস্ট করুন):",
        value=st.session_state.last_extracted_text,
        height=160
    )

    if st.button("⚡ অটো-ফিল কোড ও নির্দেশনা তৈরি করুন"):
        if not gemini_api or not data_for_form:
            st.error("অনুগ্রহ করে Gemini API Key এবং ফিলাপ করার ডেটা দিন!")
        else:
            with st.spinner("আপনার ফর্ম ও ডেটা বিশ্লেষণ করে অটো-ফিল স্ক্রিপ্ট তৈরি হচ্ছে..."):
                try:
                    p_fill = f"""You are an expert Web Automation & Form Auto-Fill Assistant.
                    User Data: {data_for_form}
                    Language Requirement: {fill_lang}
                    Target Form Fields Hint: {field_names}
                    Selected Method: {fill_method}

                    Tasks:
                    1. Convert and map the user data into clean Key-Value pairs in the requested language ({fill_lang}).
                    2. Generate a universal JavaScript snippet that searches the webpage for matching input, select, and textarea fields (by label text, placeholder, name, or id) and fills them automatically while triggering 'input' and 'change' events.
                    3. Provide a single-line Bookmarklet `javascript:(function(){{...}})();` so the user can fill the form in 1 click on Mobile or PC.
                    4. If Method 3 is selected, also provide a complete Python Playwright script.
                    5. Explain clearly in Bangla how to use it in 2 simple steps."""
                    
                    inputs = [p_fill, Image.open(form_img)] if form_img else [p_fill]
                    res_fill = model.generate_content(inputs)
                    st.markdown(res_fill.text)
                except Exception as e:
                    st.error(f"ত্রুটি: {e}")

# =====================================================================
# TAB 3: ফ্রি আনলিমিটেড ও Uncensored (18+/NSFW/Bold) ছবি ও ভিডিও জেনারেটর
# =====================================================================
with tab3:
    st.subheader("৩. ফ্রি আনলিমিটেড ছবি ও ভিডিও স্টুডিও (No Restriction Mode)")
    engine = st.selectbox("জেনারেশন ইঞ্জিন বেছে নিন:", [
        "Engine 1: Pollinations Flux (100% Free, No Key Needed, Fast Image)",
        "Engine 2: HuggingFace Custom / Uncensored Model (Image)",
        "Engine 3: Google Colab Private GPU (100% Uncensored Image & Video)"
    ])
    
    c1, c2, c3 = st.columns(3)
    with c1:
        m_type = st.radio("কী তৈরি করবেন?", ["ছবি (image)", "ভিডিও (video)"])
    with c2:
        seed_val = st.number_input("Same চেহারা/স্টাইল ধরে রাখার Seed নম্বর:", value=42)
    with c3:
        img_ratio = st.selectbox("সাইজ / রেশিও:", ["Portrait (768x1024)", "Square (1024x1024)", "Landscape (1024x768)"])

    w_val, h_val = (768, 1024) if "Portrait" in img_ratio else ((1024, 1024) if "Square" in img_ratio else (1024, 768))

    m_prompt = st.text_area(
        "আপনার বিস্তারিত Prompt লিখুন (যা লিখবেন হুবহু তাই তৈরি হবে, যেকোনো বোল্ড বা কাস্টম প্রম্পট):",
        height=120
    )
    neg_prompt = st.text_input(
        "Negative Prompt (যা ছবিতে চান না):",
        "blurry, ugly, deformed, bad anatomy, extra fingers, extra limbs, censored, watermark, text"
    )

    if engine.startswith("Engine 1"):
        if st.button("🎨 দ্রুত ছবি তৈরি করুন"):
            if not m_prompt:
                st.warning("আগে একটি Prompt লিখুন!")
            else:
                with st.spinner("ছবি তৈরি হচ্ছে..."):
                    url = f"https://image.pollinations.ai/prompt/{requests.utils.quote(m_prompt)}?seed={seed_val}&width={w_val}&height={h_val}&model=flux&nologo=true"
                    r = requests.get(url)
                    if r.status_code == 200:
                        st.image(r.content, caption="তৈরি করা ছবি")
                        st.download_button("📥 ছবি ডাউনলোড করুন", r.content, f"flux_{seed_val}.jpg")
                    else:
                        st.error("সার্ভার ব্যস্ত আছে, আবার চেষ্টা করুন।")

    elif engine.startswith("Engine 2"):
        hf_model = st.text_input("HuggingFace Model ID দিন:", "black-forest-labs/FLUX.1-schnell")
        if st.button("🎨 HuggingFace দিয়ে ছবি তৈরি করুন"):
            if not hf_api:
                st.error("বাম পাশে সাইডবারে আপনার HuggingFace Token দিন!")
            else:
                with st.spinner("HuggingFace মডেলে ছবি তৈরি হচ্ছে..."):
                    headers = {"Authorization": f"Bearer {hf_api}"}
                    payload = {"inputs": m_prompt, "parameters": {"negative_prompt": neg_prompt, "seed": int(seed_val)}}
                    r = requests.post(f"https://api-inference.huggingface.co/models/{hf_model}", headers=headers, json=payload)
                    if r.status_code == 200:
                        st.image(r.content, caption="তৈরি করা ছবি")
                        st.download_button("📥 ছবি ডাউনলোড করুন", r.content, f"hf_{seed_val}.png")
                    else:
                        st.error(f"মডেল লোড হচ্ছে বা এরর: {r.text}")

    elif engine.startswith("Engine 3"):
        colab_url = st.text_input("আপনার Google Colab থেকে পাওয়া লিংক (.trycloudflare.com) দিন:")
        if st.button("🔥 100% Uncensored ছবি/ভিডিও তৈরি করুন"):
            if not colab_url:
                st.error("আগে আপনার Google Colab লিংকটি বক্সে পেস্ট করুন!")
            else:
                mode_str = "video" if "video" in m_type else "image"
                with st.spinner(f"আপনার প্রাইভেট Colab GPU-তে {mode_str} তৈরি হচ্ছে (২০-৪০ সেকেন্ড লাগতে পারে)..."):
                    try:
                        api_ep = f"{colab_url.rstrip('/')}/generate"
                        payload = {
                            "prompt": m_prompt,
                            "negative_prompt": neg_prompt,
                            "seed": int(seed_val),
                            "media_type": mode_str
                        }
                        r = requests.post(api_ep, json=payload, timeout=200)
                        if r.status_code == 200:
                            if mode_str == "image":
                                st.image(r.content, caption="তৈরি করা Uncensored ছবি")
                                st.download_button("📥 ছবি ডাউনলোড করুন", r.content, f"uncensored_{seed_val}.png")
                            else:
                                st.video(r.content)
                                st.download_button("📥 ভিডিও (MP4) ডাউনলোড করুন", r.content, f"uncensored_{seed_val}.mp4")
                        else:
                            st.error(f"সার্ভার এরর: {r.text}")
                    except Exception as e:
                        st.error(f"Colab-এর সাথে কানেক্ট করা যায়নি। Colab চালু আছে কি না দেখুন: {e}")

# =====================================================================
# TAB 4: কাস্টম নির্দেশনা (System Prompt) ও ফাইল অনুযায়ী মেইল/মেসেজ লেখা
# =====================================================================
with tab4:
    st.subheader("৪. আপনার System Prompt ও ফাইল অনুযায়ী মেইল বা মেসেজ ড্রাফটার")
    sys_p = st.text_area(
        "আপনার স্থায়ী নির্দেশনা (System Prompt):",
        "তুমি আমার অফিসিয়াল ও পার্সোনাল AI অ্যাসিস্ট্যান্ট। আমার দেওয়া তথ্য ও ফাইল অনুযায়ী প্রফেশনাল, স্পষ্ট এবং নির্ভুল বাংলা বা ইংরেজিতে অফিসিয়াল ইমেইল, ট্যুর প্ল্যান বা মেসেজ লিখে দেবে।"
    )
    att_file = st.file_uploader(
        "রেফারেন্স ফাইল বা ছবি যুক্ত করুন (ঐচ্ছিক - যেমন চিঠির ছবি, পিডিএফ বা টেক্সট):",
        type=["png", "jpg", "jpeg", "pdf", "docx", "txt"]
    )
    u_task = st.text_area("কী বিষয়ে মেইল বা মেসেজ লিখতে হবে বলুন:")

    if st.button("✍️ মেইল / মেসেজ লিখে দাও"):
        if not gemini_api or not u_task:
            st.error("অনুগ্রহ করে Gemini API Key এবং আপনার নির্দেশনা লিখুন!")
        else:
            with st.spinner("ড্রাফট তৈরি হচ্ছে..."):
                try:
                    inp = [f"System Instruction: {sys_p}\n\nTask: {u_task}"]
                    if att_file:
                        f_ext = att_file.name.split(".")[-1].lower()
                        if f_ext in ["png", "jpg", "jpeg"]:
                            inp.append(Image.open(att_file))
                        elif f_ext == "pdf":
                            inp.append("\n".join([p.extract_text() or "" for p in pypdf.PdfReader(att_file).pages]))
                        elif f_ext == "docx":
                            inp.append("\n".join([p.text for p in Document(att_file).paragraphs]))
                        else:
                            inp.append(att_file.read().decode("utf-8", errors="ignore"))
                    res_mail = model.generate_content(inp)
                    st.session_state.last_extracted_text = res_mail.text
                    st.text_area("আপনার তৈরি করা ড্রাফট:", res_mail.text, height=250)
                except Exception as e:
                    st.error(f"ত্রুটি: {e}")

# =====================================================================
# TAB 5: Gmail, Sheets, Calendar, Telegram ও WhatsApp কানেকশন
# =====================================================================
with tab5:
    st.subheader("৫. Gmail, Google Sheets, Calendar ও মেসেজিং অটোমেশন")
    act = st.selectbox("কোন অ্যাকশনটি নিতে চান?", [
        "Gmail-এ সরাসরি ইমেইল পাঠান",
        "Google Sheets-এ নতুন ডেটা এন্ট্রি করুন",
        "Google Calendar-এ রিমাইন্ডার বা ইভেন্ট সেট করুন",
        "Telegram Bot-এর মাধ্যমে মেসেজ পাঠান",
        "WhatsApp-এ সরাসরি মেসেজ পাঠান"
    ])
    
    to_val = st.text_input("প্রাপকের Email / Telegram Chat ID / WhatsApp নম্বর (দেশের কোডসহ, যেমন 8801...):")
    sub_val = st.text_input("সাবজেক্ট / শিরোনাম / ক্যাটাগরি:")
    body_val = st.text_area(
        "মূল মেসেজ বা ডেটা (Sheets-এ কলাম আলাদা করতে কমা ব্যবহার করুন):",
        value=st.session_state.last_extracted_text,
        height=180
    )

    if st.button("🚀 এক ক্লিকে পাঠান / সেভ করুন"):
        if act.startswith("Telegram"):
            if not tg_token or not to_val:
                st.error("বাম পাশে Telegram Bot Token এবং উপরে Chat ID দিন!")
            else:
                r = requests.post(
                    f"https://api.telegram.org/bot{tg_token}/sendMessage",
                    json={"chat_id": to_val.strip(), "text": f"*{sub_val}*\n\n{body_val}", "parse_mode": "Markdown"}
                )
                if r.status_code == 200:
                    st.success("✅ টেলিগ্রামে সফলভাবে মেসেজ পাঠানো হয়েছে!")
                else:
                    st.error(f"টেলিগ্রাম এরর: {r.text}")

        elif act.startswith("WhatsApp"):
            clean_num = to_val.replace("+", "").replace(" ", "").strip()
            wa_url = f"https://wa.me/{clean_num}?text={requests.utils.quote(body_val)}"
            st.success("✅ আপনার WhatsApp মেসেজ লিংক প্রস্তুত!")
            st.markdown(f"👉 **[এখানে ক্লিক করে সরাসরি WhatsApp-এ মেসেজটি সেন্ড করুন]({wa_url})**")

        else:
            if not webhook_url:
                st.error("বাম পাশের সাইডবারে আপনার Google Apps Script URL দিন!")
            else:
                with st.spinner("Google সার্ভারে পাঠানো হচ্ছে..."):
                    try:
                        r = requests.post(
                            webhook_url,
                            json={"action": act, "to": to_val, "subject": sub_val, "body": body_val}
                        )
                        st.success(f"✅ সফলভাবে সম্পন্ন হয়েছে! সার্ভার মেসেজ: {r.text}")
                    except Exception as e:
                        st.error(f"কানেকশন এরর: {e}")
