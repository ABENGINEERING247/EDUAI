import io
import json
import os
import re
import smtplib
import sqlite3
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

APP_TITLE = "EDU AI"
DB_PATH = Path("edu_ai_history.db")
st.set_page_config(page_title="EDU AI — Study & Productivity Assistant", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

# ----------------------------- Styling ---------------------------------------
st.markdown("""
<style>
:root { --navy:#102b53; --blue:#087cf0; --cyan:#12c7ed; }
.stApp { background: linear-gradient(135deg,#f5faff 0%,#edf5ff 50%,#ffffff 100%); }
header[data-testid="stHeader"] { background: rgba(255,255,255,0); }
.block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 1600px; }
.edu-header { background:linear-gradient(105deg,#0b2043,#1e477d); color:white; border-radius:18px; padding:22px 28px; margin-bottom:18px; box-shadow:0 8px 28px #102b5322; }
.edu-header h1 { margin:0; font-size:2.5rem; letter-spacing:.5px; }
.edu-header p { margin:5px 0 0; color:#d9efff; }
.ali-hero { background:linear-gradient(110deg,#e1f5ff,#f5fbff 65%,#d7efff); border:1px solid #b9e2ff; border-radius:18px; padding:24px; min-height:235px; }
.ali-hero h2 { color:#0c2d5e; font-size:2.0rem; margin:0 0 8px; }
.ali-hero .big { font-size:3.2rem; font-weight:900; color:#087cf0; line-height:1; }
.feature-card { border:1px solid #d6e7fa; border-radius:16px; padding:18px; min-height:190px; background:white; box-shadow:0 4px 15px #143a6310; }
.feature-card h3 { margin:8px 0; font-size:1.15rem; }
.small-muted { color:#58708c; font-size:.88rem; }
div[data-testid="stMetric"] { background:white; padding:12px; border:1px solid #d8e8f9; border-radius:12px; }
</style>
""", unsafe_allow_html=True)

# ----------------------------- Database --------------------------------------
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            feature TEXT NOT NULL,
            prompt TEXT NOT NULL,
            result TEXT NOT NULL
        )""")
init_db()

def save_history(feature, prompt, result):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("INSERT INTO history(created_at,feature,prompt,result) VALUES(?,?,?,?)",
                     (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), feature, prompt, result))

def load_history(limit=100):
    with sqlite3.connect(DB_PATH) as conn:
        return pd.read_sql_query("SELECT id, created_at, feature, prompt, result FROM history ORDER BY id DESC LIMIT ?", conn, params=(limit,))

# ----------------------------- AI helpers ------------------------------------
def get_api_key():
    try:
        key = st.secrets.get("OPENAI_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY", "")

def get_model():
    # Locked exclusively to gpt-6-luna
    return "gpt-6-luna"

def ask_ai(prompt, system="You are ALI, a helpful educational and productivity assistant. Be accurate, clear, and age-appropriate."):
    key = get_api_key()
    if not key or OpenAI is None:
        return demo_response(prompt)
    try:
        client = OpenAI(api_key=key)
        response = client.chat.completions.create(
            model=get_model(),
            messages=[{"role":"system","content":system},{"role":"user","content":prompt}],
            # Temperature omitted because gpt-6-luna only supports default (1) value
        )
        return response.choices[0].message.content.strip()
    except Exception as exc:
        return f"**API response unavailable.** {exc}\n\nYou can continue in Demo Mode or check your `OPENAI_API_KEY` in Streamlit secrets."

def demo_response(prompt):
    p = prompt.strip()
    return (
        "### EDU AI Demo Response\n\n"
        f"You asked: **{p[:500]}**\n\n"
        f"*(Active Model: {get_model()})*\n\n"
        "This is a local demo response, so no external AI API was called. "
        "Add `OPENAI_API_KEY` to Streamlit Secrets to enable AI-generated responses.\n\n"
        "**Suggested next steps**\n"
        "1. Define the goal and audience.\n"
        "2. Break the task into clear sections.\n"
        "3. Review the result for accuracy and adapt it to your requirements."
    )

def run_prompt(feature, prompt, system=None):
    result = ask_ai(prompt, system or f"You are ALI inside EDU AI. Help with the {feature} tool. Use a polished, practical format.")
    save_history(feature, prompt, result)
    return result

def docx_bytes(title, content):
    doc = Document()
    doc.add_heading(title, 0)
    doc.add_paragraph(f"Generated by EDU AI • {datetime.now().strftime('%d %b %Y, %H:%M')}")
    doc.add_paragraph("")
    for line in content.splitlines():
        if line.startswith("### "):
            doc.add_heading(line[4:], level=2)
        elif line.startswith("## "):
            doc.add_heading(line[3:], level=1)
        elif line.startswith("# "):
            doc.add_heading(line[2:], level=1)
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(line[2:], style="List Bullet")
        elif re.match(r"^\d+\.\s", line):
            doc.add_paragraph(re.sub(r"^\d+\.\s", "", line), style="List Number")
        elif line.strip():
            doc.add_paragraph(line.replace("**", "").replace("`", ""))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def pdf_bytes(title, content):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=0.65*inch, leftMargin=0.65*inch, topMargin=0.65*inch, bottomMargin=0.65*inch)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"]), Spacer(1, 12)]
    for line in content.splitlines():
        safe = (line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("**", "").replace("`", ""))
        if not safe.strip():
            story.append(Spacer(1, 6))
        elif line.startswith("#"):
            story.append(Paragraph(safe.lstrip("# ").strip(), styles["Heading2"]))
        elif line.startswith(("- ", "* ")):
            story.append(Paragraph("• " + safe[2:], styles["BodyText"]))
        else:
            story.append(Paragraph(safe, styles["BodyText"]))
            story.append(Spacer(1, 4))
    doc.build(story)
    return buf.getvalue()

def page_title(icon, title, subtitle=""):
    st.subheader(f"{icon}  {title}")
    if subtitle:
        st.caption(subtitle)

# ----------------------------- Header -----------------------------------------
st.markdown("""
<div class="edu-header">
  <div style="display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap">
    <div>
      <h1>🎓 EDU <span style="color:#16c7f3">AI</span></h1>
      <p><b>Your Smart Study & Productivity Assistant</b></p>
      <p style="font-size:.9rem">Assignments · Documents · PDFs · Excel · Chat · Email</p>
    </div>
    <div style="text-align:right">
      <div style="font-size:1.05rem;font-style:italic">Learn · Create · Be Productive</div>
      <div style="font-size:1.8rem;font-weight:800;color:#ffc247">with ALI 🤖</div>
      <div style="border:1px solid #16c7f3;border-radius:12px;padding:8px 14px;display:inline-block;margin-top:5px">
        <span style="color:#28e08b">●</span> Powered by GPT-6 Luna
      </div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ----------------------------- Sidebar ----------------------------------------
pages = ["Home", "Assignments Maker", "Docs Maker", "PDF Maker", "Excel Maker", "Chat with ALI", "Send Email", "History", "Settings"]
with st.sidebar:
    st.markdown("## 🎓 EDU AI")
    st.caption("Your study partner, powered by ALI")
    page = st.radio("Navigation", pages, label_visibility="collapsed")
    st.divider()
    st.markdown("### 🤖 Meet ALI")
    st.info("Small steps create big success!")
    key_active = bool(get_api_key()) and OpenAI is not None
    st.caption("Mode: " + ("OpenAI API" if key_active else "Demo"))
    st.caption("Active Model: `gpt-6-luna`")
    st.caption("Designed by Engr. Bilal Mehmood")

# ----------------------------- Home -------------------------------------------
if page == "Home":
    left, right = st.columns([2.15, 1])
    with left:
        st.markdown("""
        <div class="ali-hero">
          <div style="font-size:1rem;color:#16629b;font-weight:700">WELCOME TO</div>
          <h2>EDU <span style="color:#087cf0">AI</span></h2>
          <div style="font-size:1.04rem;max-width:680px">
          Hi! I'm <b>ALI</b>, your AI assistant. I can help you with assignments, create documents,
          make PDFs and Excel sheets, answer questions, and draft emails. Just ask me anything!
          </div>
          <div style="margin-top:14px;background:#087cf0;color:white;display:inline-block;padding:8px 16px;border-radius:24px">
          Learn smarter. Create faster. ✨
          </div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        cards = st.columns(4)
        items = [
            ("📝","Assignments Maker","Generate structured assignments.","Assignments Maker"),
            ("📄","Docs Maker","Create reports, letters and notes.","Docs Maker"),
            ("📕","PDF Maker","Create downloadable PDF files.","PDF Maker"),
            ("📊","Excel Maker","Build spreadsheets and charts.","Excel Maker"),
        ]
        for col, (icon, title, desc, target) in zip(cards, items):
            with col:
                st.markdown(f'<div class="feature-card"><div style="font-size:2rem">{icon}</div><h3>{title}</h3><div class="small-muted">{desc}</div></div>', unsafe_allow_html=True)
                if st.button("Open tool", key="home_"+target, use_container_width=True):
                    st.session_state["home_target"] = target
                    st.rerun()
        if st.session_state.get("home_target") in pages and st.session_state.get("home_target") != "Home":
            st.info(f"Choose **{st.session_state['home_target']}** from the navigation menu to continue.")
        st.write("")
        c1,c2,c3,c4 = st.columns(4)
        hist = load_history()
        c1.metric("Saved Activities", len(hist))
        c2.metric("Available Tools", "7")
        c3.metric("AI Mode", "API" if key_active else "Demo")
        c4.metric("Engine", "GPT-6 Luna")
    with right:
        st.markdown("### 💬 Chat with ALI")
        st.success("Online • Ready to help")
        st.markdown("**ALI can help you with:**")
        st.markdown("- Assignments and study notes\n- Documents and PDFs\n- Excel sheets and data\n- Explanations and brainstorming\n- Email drafts")
        quick = st.text_area("Ask ALI a quick question", placeholder="e.g. Explain artificial intelligence in simple words", height=120)
        if st.button("Ask ALI ➤", type="primary", use_container_width=True):
            if quick.strip():
                st.session_state["quick_answer"] = run_prompt("Quick Chat", quick)
            else:
                st.warning("Enter a question first.")
        if st.session_state.get("quick_answer"):
            st.markdown(st.session_state["quick_answer"])

# ----------------------------- Assignments ------------------------------------
elif page == "Assignments Maker":
    page_title("📝", "Assignments Maker", "Generate an outline or a full assignment and export it as Word or PDF.")
    with st.form("assignment_form"):
        topic = st.text_input("Assignment topic", placeholder="Artificial Intelligence in Education")
        subject = st.text_input("Subject / course", placeholder="Introduction to AI")
        level = st.selectbox("Academic level", ["School", "College", "Undergraduate", "Postgraduate", "Professional"])
        words = st.select_slider("Approximate length", options=["Short (300 words)", "Medium (700 words)", "Detailed (1200 words)"], value="Medium (700 words)")
        instructions = st.text_area("Special instructions", placeholder="Include headings, examples, conclusion, and references.")
        submitted = st.form_submit_button("✨ Generate Assignment", type="primary")
    if submitted:
        if not topic.strip():
            st.error("Please enter an assignment topic.")
        else:
            prompt = f"Create a {words} assignment for academic level {level}. Subject: {subject or 'General'}. Topic: {topic}. Requirements: include title, introduction, clear headings, examples where relevant, conclusion, and suggested references. Special instructions: {instructions or 'None'}. Do not fabricate specific citations; mark sources as suggestions to verify."
            st.session_state["assignment_result"] = run_prompt("Assignment", prompt)
    if st.session_state.get("assignment_result"):
        result = st.session_state["assignment_result"]
        st.markdown(result)
        a,b = st.columns(2)
        a.download_button("⬇️ Download Word", docx_bytes("EDU AI Assignment", result), "EDU_AI_Assignment.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        b.download_button("⬇️ Download PDF", pdf_bytes("EDU AI Assignment", result), "EDU_AI_Assignment.pdf", "application/pdf", use_container_width=True)

# ----------------------------- Documents --------------------------------------
elif page == "Docs Maker":
    page_title("📄", "Docs Maker", "Draft reports, letters, meeting notes, lesson plans, and other documents.")
    kind = st.selectbox("Document type", ["Professional report", "Formal letter", "Meeting minutes", "Study notes", "Lesson plan", "Project proposal", "Resume summary", "Custom document"])
    title = st.text_input("Document title / subject")
    details = st.text_area("Key information to include", height=170, placeholder="Provide names, goals, facts, sections, tone, and any required details.")
    if st.button("✨ Create Document", type="primary"):
        if not title.strip() or not details.strip():
            st.warning("Please enter a title and key information.")
        else:
            prompt = f"Create a polished {kind}. Title/subject: {title}. Include these details accurately: {details}. Use suitable headings, professional tone, and do not invent missing factual details."
            st.session_state["document_result"] = run_prompt("Document", prompt)
    if st.session_state.get("document_result"):
        result = st.session_state["document_result"]
        st.markdown(result)
        st.download_button("⬇️ Download DOCX", docx_bytes(title or "EDU AI Document", result), "EDU_AI_Document.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ----------------------------- PDF Maker --------------------------------------
elif page == "PDF Maker":
    page_title("📕", "PDF Maker", "Write content here or generate a draft, then download it as a PDF.")
    pdf_title = st.text_input("PDF title", "My EDU AI Document")
    pdf_content = st.text_area("Content", height=250, placeholder="Paste or write the content you want in your PDF...")
    if st.button("✨ Improve Content with ALI"):
        if pdf_content.strip():
            st.session_state["pdf_content"] = run_prompt("PDF Content", f"Improve and organize this content for a clear PDF document. Preserve facts and meaning:\n\n{pdf_content}")
        else:
            st.warning("Enter content first.")
    if "pdf_content" in st.session_state:
        pdf_content = st.session_state["pdf_content"]
        st.text_area("Edited content", value=pdf_content, height=180, key="pdf_content_edited")
        pdf_content = st.session_state["pdf_content_edited"]
    if pdf_content.strip():
        st.download_button("⬇️ Download PDF", pdf_bytes(pdf_title, pdf_content), "EDU_AI_Document.pdf", "application/pdf", type="primary")

# ----------------------------- Excel Maker ------------------------------------
elif page == "Excel Maker":
    page_title("📊", "Excel Maker", "Create a spreadsheet from a simple table or ask ALI to suggest a data structure.")
    mode = st.radio("Start with", ["Sample table", "Enter CSV data", "Ask ALI for a table"], horizontal=True)
    if mode == "Sample table":
        df = pd.DataFrame({"Student":["Ayesha","Ahmed","Fatima"],"Subject":["AI","Python","Data Science"],"Marks":[88,92,85]})
        st.dataframe(df, use_container_width=True)
    elif mode == "Enter CSV data":
        csv_text = st.text_area("Paste CSV", "Name,Course,Score\nAyesha,AI,88\nAhmed,Python,92")
        try:
            df = pd.read_csv(io.StringIO(csv_text))
            st.dataframe(df, use_container_width=True)
        except Exception as exc:
            st.error(f"Could not parse CSV: {exc}")
            df = pd.DataFrame()
    else:
        request = st.text_area("Describe your spreadsheet", placeholder="Create a student attendance sheet with date, name, roll number and status.")
        if st.button("Generate table idea", type="primary"):
            if request.strip():
                st.session_state["excel_idea"] = run_prompt("Excel Table", f"Suggest a spreadsheet table for this requirement. Return a concise CSV block with headers and 5 example rows. Requirement: {request}")
            else:
                st.warning("Describe the spreadsheet first.")
        if st.session_state.get("excel_idea"):
            st.markdown(st.session_state["excel_idea"])
            st.caption("Copy the table/CSV into the CSV input option to create an Excel file.")
    if mode != "Ask ALI for a table":
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="EDU AI Data")
        st.download_button("⬇️ Download Excel (.xlsx)", output.getvalue(), "EDU_AI_Sheet.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

# ----------------------------- Chat -------------------------------------------
elif page == "Chat with ALI":
    page_title("💬", "Chat with ALI", "Ask study questions, brainstorm ideas, or get step-by-step explanations.")
    if "chat_messages" not in st.session_state:
        st.session_state["chat_messages"] = [{"role":"assistant","content":"Hi! I'm ALI 👋 What would you like to learn or create today?"}]
    for msg in st.session_state["chat_messages"]:
        with st.chat_message(msg["role"], avatar="🤖" if msg["role"]=="assistant" else "🙂"):
            st.markdown(msg["content"])
    user_msg = st.chat_input("Type your message to ALI...")
    if user_msg:
        st.session_state["chat_messages"].append({"role":"user","content":user_msg})
        with st.chat_message("user", avatar="🙂"):
            st.markdown(user_msg)
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("ALI is thinking..."):
                answer = ask_ai(user_msg, "You are ALI, a friendly educational assistant. Explain clearly, use examples when helpful, and state uncertainty.")
            st.markdown(answer)
        st.session_state["chat_messages"].append({"role":"assistant","content":answer})
        save_history("Chat", user_msg, answer)

# ----------------------------- Email ------------------------------------------
elif page == "Send Email":
    page_title("✉️", "Send Email", "Draft an email with ALI. Actual sending requires SMTP settings in Streamlit Secrets.")
    recipient = st.text_input("Recipient email address")
    subject = st.text_input("Email subject")
    purpose = st.text_area("What should the email say?", height=120, placeholder="Explain the purpose, tone, and key points.")
    if st.button("Draft Email with ALI", type="primary"):
        if purpose.strip():
            prompt = f"Draft an email. Recipient context: {recipient or 'not specified'}. Subject: {subject or 'suggest a suitable subject'}. Requirements: {purpose}. Return subject on the first line as 'Subject: ...' followed by the email body."
            st.session_state["email_draft"] = run_prompt("Email Draft", prompt)
        else:
            st.warning("Describe what the email should say.")
    if st.session_state.get("email_draft"):
        draft = st.text_area("Review and edit email draft", value=st.session_state["email_draft"], height=220, key="email_draft_edit")
        st.caption("Sending is optional and uses SMTP credentials configured by you.")
        with st.expander("SMTP send settings"):
            st.warning("Never paste passwords into chat or source code. Configure SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD in Streamlit Secrets.")
            confirm_send = st.checkbox("I reviewed the recipient, subject, and message and want to send this email.")
        if st.button("Send Email"):
            try:
                secrets = st.secrets
                host = secrets["SMTP_HOST"]
                port = int(secrets.get("SMTP_PORT", 587))
                username = secrets["SMTP_USER"]
                password = secrets["SMTP_PASSWORD"]
                if not confirm_send:
                    st.error("Confirm the email review checkbox first.")
                elif not recipient or "@" not in recipient:
                    st.error("Enter a valid recipient email.")
                else:
                    msg = EmailMessage()
                    msg["From"] = username
                    msg["To"] = recipient
                    match = re.search(r"(?im)^Subject:\s*(.+)$", draft)
                    msg["Subject"] = subject or (match.group(1).strip() if match else "Message from EDU AI")
                    body = re.sub(r"(?im)^Subject:\s*.+\n?", "", draft, count=1).strip()
                    msg.set_content(body)
                    with smtplib.SMTP(host, port, timeout=20) as server:
                        server.starttls()
                        server.login(username, password)
                        server.send_message(msg)
                    st.success("Email sent successfully.")
                    save_history("Email Sent", recipient, f"Subject: {msg['Subject']}\n{body}")
            except Exception as exc:
                st.error(f"Email could not be sent. Check SMTP Secrets and provider settings. Details: {exc}")

# ----------------------------- History ----------------------------------------
elif page == "History":
    page_title("🕘", "History", "Your recent generated content is stored in a local SQLite database.")
    history = load_history()
    if history.empty:
        st.info("No activity yet. Create an assignment, document, or chat with ALI to see history here.")
    else:
        st.dataframe(history[["created_at","feature","prompt"]], use_container_width=True, hide_index=True)
        selected_id = st.selectbox("View saved result", history["id"].tolist(), format_func=lambda x: f"#{x} — {history.loc[history.id==x, 'feature'].iloc[0]} ({history.loc[history.id==x, 'created_at'].iloc[0]})")
        selected = history[history["id"] == selected_id].iloc[0]
        st.markdown(selected["result"])
        st.download_button("Download selected result as TXT", selected["result"], f"EDU_AI_History_{selected_id}.txt", "text/plain")
        if st.button("Clear all history", type="secondary"):
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("DELETE FROM history")
            st.success("History cleared.")
            st.rerun()

# ----------------------------- Settings ---------------------------------------
elif page == "Settings":
    page_title("⚙️", "Settings", "Configure API access and review privacy notes.")
    key_active = bool(get_api_key()) and OpenAI is not None
    st.metric("AI provider status", "Connected" if key_active else "Demo mode")
    st.write(f"**Selected model:** `gpt-6-luna` (Locked)")
    st.markdown("#### Configure Streamlit Secrets")
    st.code('OPENAI_API_KEY = "your-api-key"', language="toml")
    st.markdown("For Streamlit Community Cloud, open your app → **Settings → Secrets** and add your API key. Restart/reboot the app if needed.")
    st.markdown("#### Optional email sending")
    st.code('SMTP_HOST = "smtp.example.com"\nSMTP_PORT = 587\nSMTP_USER = "your-email@example.com"\nSMTP_PASSWORD = "your-app-password"', language="toml")
    st.warning("Do not commit API keys, SMTP passwords, or personal data to GitHub. EDU AI's SQLite history is local to the running environment and may not persist across cloud restarts.")
    st.markdown("#### About")
    st.write("EDU AI — Your Smart Study & Productivity Assistant")
    st.write("Assistant: ALI | Built with Python and Streamlit")
    st.write("Designed by Engr. Bilal Mehmood")
    st.caption("AI-generated content can contain mistakes. Verify facts, citations, calculations, and important decisions before use.")

st.divider()
st.caption("Built with ❤️ using Python & Streamlit  |  Powered by GPT-6 Luna  |  Designed by Engr. Bilal Mehmood")
