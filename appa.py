# -*- coding: utf-8 -*-
"""
医考帮学习备份工具 - Streamlit 网页版（云端适配）
所有文件操作均使用临时目录，不依赖绝对路径，安全部署到 Streamlit Cloud
"""

import json
import os
import re
import hashlib
import tempfile
import zipfile
import io
import requests
import pandas as pd
from urllib.parse import urlparse
from tqdm import tqdm
import streamlit as st

# ============================================================
# 核心函数（完全复用，路径已全部改为临时目录）
# ============================================================

def fix_encoding(text):
    if not text: return text
    try: return text.encode('latin1').decode('utf-8')
    except: return text

def _download_image(url, save_path):
    try:
        r = requests.get(url, stream=True, timeout=10)
        r.raise_for_status()
        with open(save_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    except Exception as e:
        pass

def export_exam(json_data, output_path_base, format='excel', download_images=False):
    """导出考试数据，所有文件保存在临时目录下"""
    output_dir = os.path.dirname(output_path_base)
    img_dir = os.path.join(output_dir, "images")

    rows = []
    for q in tqdm(json_data, desc="解析题目", unit="题"):
        opts = json.loads(q.get('option', '[]')) if isinstance(q.get('option'), str) else q.get('option', [])
        opt_str = "\n".join([f"{o.get('key')}: {fix_encoding(o.get('title'))}" for o in opts])
        title = fix_encoding(q.get('title', ''))
        explain = fix_encoding(q.get('explain', ''))

        if download_images:
            all_text = title + explain + opt_str
            img_urls = re.findall(r'(https?://[^\s"\'<>]+\.(?:jpg|jpeg|png|gif))', all_text, re.IGNORECASE)
            if img_urls:
                if not os.path.exists(img_dir):
                    os.makedirs(img_dir)
                for img_url in img_urls:
                    filename = os.path.basename(urlparse(img_url).path) or "image.jpg"
                    save_path = os.path.join(img_dir, filename)
                    if not os.path.exists(save_path):
                        _download_image(img_url, save_path)

        rows.append({
            "题号": q.get('number', ''),
            "题目": title,
            "选项": opt_str,
            "答案": q.get('answer', ''),
            "解析": explain
        })

    if format == 'excel':
        output_path = f"{output_path_base}.xlsx"
        pd.DataFrame(rows).to_excel(output_path, index=False)
    elif format == 'word':
        output_path = f"{output_path_base}.docx"
        _export_word(rows, output_path)
    elif format == 'pdf':
        output_path = f"{output_path_base}.pdf"
        _export_pdf(rows, output_path)

def _export_word(rows, output_path):
    from docx import Document
    from docx.shared import Pt
    doc = Document()
    doc.add_heading('医考帮题库导出', 0)
    for row in tqdm(rows, desc="生成 Word", unit="题"):
        doc.add_heading(f"第 {row['题号']} 题", level=2)
        doc.add_paragraph(f"【题目】{row['题目']}")
        doc.add_paragraph(f"【选项】\n{row['选项']}")
        p = doc.add_paragraph()
        p.add_run(f"【正确答案】{row['答案']}").bold = True
        doc.add_paragraph(f"【解析】{row['解析']}")
        doc.add_paragraph("─" * 30)
    doc.save(output_path)

def _export_pdf(rows, output_path):
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.pagesizes import A4
    import textwrap
    import os

    font_name = "Helvetica"
    font_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux 云服务器常用
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/arphic/uming.ttc",
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont('CJK', fp))
                font_name = 'CJK'
                break
            except Exception:
                continue

    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    margin = 50
    line_height = 16
    y = height - margin

    def new_page():
        nonlocal y
        c.showPage()
        c.setFont(font_name, 10)
        y = height - margin

    c.setFont(font_name, 10)

    for row in tqdm(rows, desc="生成 PDF", unit="题"):
        lines = [
            f"第 {row['题号']} 题",
            f"题目：{row['题目']}",
        ]
        for opt_line in row['选项'].split('\n'):
            lines.append(f"  {opt_line}")
        lines.append(f"正确答案：{row['答案']}")
        lines.append(f"解析：{row['解析']}")
        lines.append("")

        for line in lines:
            wrapped = textwrap.wrap(line, width=60) if line else ['']
            for wline in wrapped:
                if y < margin + line_height:
                    new_page()
                c.drawString(margin, y, wline)
                y -= line_height

    c.save()

# ============================================================
# 登录与 API 模块
# ============================================================

BASE_URL = "https://ykb-online-exam.yikaobang.com.cn/api"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded',
    'Client-Type': 'web',
}

def md5(text):
    return hashlib.md5(text.encode('utf-8')).hexdigest()

def login(mobile, password):
    url = f"{BASE_URL}/user/login"
    data = {'mobile': mobile, 'password': md5(password)}
    try:
        r = requests.post(url, data=data, headers=HEADERS, timeout=15)
        result = r.json()
        if result.get('code') == 200:
            token = result['data']['token']
            secret = result['data']['secret']
            hospitals = result['data'].get('hospital', [])
            hospital_id = hospitals[0]['id'] if hospitals else ''
            return token, secret, str(hospital_id)
        else:
            return None
    except Exception as e:
        return None

def get_exam_list(token, secret, hospital_id):
    url = f"{BASE_URL}/question/getUserExam"
    data = {'token': token, 'secret': secret, 'hospital_id': hospital_id}
    try:
        r = requests.post(url, data=data, headers=HEADERS, timeout=15)
        result = r.json()
        if result.get('code') == 200:
            return result.get('data', [])
        else:
            return []
    except Exception as e:
        return []

def download_questions(question_file_url):
    try:
        r = requests.get(question_file_url, timeout=30)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        return None

# ============================================================
# 网页界面（Streamlit）
# ============================================================

def main():
    st.set_page_config(page_title="医考帮备份工具", page_icon="📚", layout="centered")
    
    st.title("📚 医考帮个人学习备份工具")
    st.caption("基于 ykb-exporter 改造 · 网页版")
    st.divider()

    if 'auth' not in st.session_state:
        st.session_state.auth = None
        st.session_state.exams = []
        st.session_state.logged_in = False

    with st.sidebar:
        st.header("🔐 登录账号")
        mobile = st.text_input("手机号", placeholder="请输入手机号")
        password = st.text_input("密码", placeholder="请输入密码", type="password")
        
        if st.button("登录", type="primary", use_container_width=True):
            if not mobile or not password:
                st.error("请完整填写手机号和密码")
            else:
                with st.spinner("正在登录，请稍候..."):
                    auth = login(mobile, password)
                    if auth:
                        st.session_state.auth = auth
                        st.session_state.logged_in = True
                        with st.spinner("正在获取考试列表..."):
                            exams = get_exam_list(auth[0], auth[1], auth[2])
                            st.session_state.exams = exams
                        st.success("✅ 登录成功！")
                        st.rerun()
                    else:
                        st.error("❌ 登录失败，请检查账号密码")

        if st.session_state.logged_in:
            st.info(f"当前已登录：{mobile}")
            if st.button("退出登录"):
                st.session_state.auth = None
                st.session_state.exams = []
                st.session_state.logged_in = False
                st.rerun()

    if not st.session_state.logged_in:
        st.info("👈 请先在左侧侧边栏登录您的医考帮账号")
        return

    exams = st.session_state.exams
    if not exams:
        st.warning("未获取到任何考试，请确认账号权限或重新登录")
        return

    st.subheader("📋 选择要导出的考试")

    status_map = {0: "未开始", 1: "进行中", 2: "已结束"}
    exam_options = []
    for e in exams:
        status = status_map.get(e.get('status', 0), "未知")
        title = e.get('title', '未命名')
        exam_options.append(f"[{status}] {title}")

    selected_index = st.selectbox(
        "考试列表（请点击选择）",
        options=range(len(exam_options)),
        format_func=lambda x: exam_options[x]
    )

    st.subheader("📄 选择导出格式")
    formats = st.multiselect(
        "支持同时导出多种格式",
        options=['excel', 'word', 'pdf'],
        default=['excel', 'word', 'pdf'],
        format_func=lambda x: {'excel': '📊 Excel (.xlsx)', 'word': '📝 Word (.docx)', 'pdf': '📕 PDF (.pdf)'}[x]
    )

    if not formats:
        st.warning("请至少选择一种导出格式")

    if st.button("🚀 开始导出", type="primary", use_container_width=True, disabled=(not formats)):
        selected_exam = exams[selected_index]
        title = selected_exam.get('title', 'exam_export')
        question_url = selected_exam.get('question_file')

        if not question_url:
            st.error("该考试没有题目文件链接，无法导出")
            st.stop()

        status_placeholder = st.status("⏳ 正在处理，请耐心等待...", expanded=True)

        try:
            status_placeholder.write("📥 正在下载题目数据...")
            json_data = download_questions(question_url)
            if not json_data:
                st.error("题目数据下载失败，请检查网络或重试")
                status_placeholder.update(label="导出失败", state="error")
                st.stop()
            status_placeholder.write(f"✅ 下载完成，共 {len(json_data)} 道题目")

            # 使用临时目录（确保云端安全）
            with tempfile.TemporaryDirectory() as tmpdir:
                safe_title = re.sub(r'[\\/:*?"<>|]', '_', title)
                base_path = os.path.join(tmpdir, safe_title)

                status_placeholder.write("🔄 正在转换格式...")
                for fmt in formats:
                    status_placeholder.write(f"   - 生成 {fmt} 文件...")
                    export_exam(json_data, base_path, format=fmt, download_images=False)

                status_placeholder.write("📦 正在打包文件...")
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for f in os.listdir(tmpdir):
                        file_path = os.path.join(tmpdir, f)
                        if os.path.isfile(file_path) and f.endswith(('.xlsx', '.docx', '.pdf')):
                            zf.write(file_path, f)
                zip_buffer.seek(0)

                status_placeholder.update(label="✅ 导出完成！", state="complete")
                
                st.success(f"🎉 成功导出 {len(formats)} 个文件！点击下方按钮下载：")
                st.download_button(
                    label="📥 点击下载 ZIP 压缩包",
                    data=zip_buffer,
                    file_name=f"{safe_title}_export.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )

        except Exception as e:
            status_placeholder.update(label="❌ 导出出错", state="error")
            st.error(f"处理过程中发生错误：{e}")
            st.stop()

    st.divider()
    st.caption("💡 提示：所有数据仅在本地处理，不会上传至任何第三方服务器")

if __name__ == "__main__":
    main()
