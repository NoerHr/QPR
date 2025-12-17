import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import re
import google.generativeai as genai
from streamlit_mic_recorder import mic_recorder
import json

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Laporan QPR 360°", layout="wide", page_icon="🎙️")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        color: #1a1a1a;
        background-color: #f3f4f6;
    }
    .block-container { padding-top: 2rem; padding-bottom: 5rem; }

    /* CARD STYLE */
    .css-card {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 16px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 24px;
    }

    /* HEADER */
    .header-box {
        background: linear-gradient(135deg, #4c1d95 0%, #7c3aed 100%);
        color: white;
        padding: 35px;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 10px 20px -5px rgba(124, 58, 237, 0.3);
    }
    .header-box h1 { margin: 0; font-weight: 800; font-size: 2.2rem; color: white !important; }
    
    /* SCORE WIDGET */
    .score-container { text-align: center; padding: 20px; border-bottom: 2px dashed #f3f4f6; margin-bottom: 20px; }
    .score-val { font-size: 4rem; font-weight: 800; color: #6d28d9; line-height: 1; }
    .score-label { font-size: 0.85rem; font-weight: 700; color: #4b5563; text-transform: uppercase; margin-top: 10px; letter-spacing: 1.5px; }

    /* NARRATIVE BOX */
    .narrative-box {
        background-color: #fdf4ff;
        border-left: 6px solid #d946ef;
        padding: 30px;
        border-radius: 12px;
        margin-top: 20px;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    }
    .narrative-title {
        color: #86198f;
        font-weight: 800;
        text-transform: uppercase;
        margin-bottom: 15px;
        font-size: 1.1rem;
        letter-spacing: 0.5px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .narrative-text {
        color: #1f2937;
        line-height: 2; 
        text-align: justify;
        font-size: 1.05rem;
    }

    /* TABLE STYLE */
    .styled-table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 0.95rem; border-radius: 10px; overflow: hidden; }
    .styled-table th { background-color: #5b21b6; color: #ffffff; font-weight: 600; padding: 14px 16px; text-align: left; }
    .styled-table td { padding: 16px; border-bottom: 1px solid #f3f4f6; vertical-align: top; color: #374151; line-height: 1.6; word-wrap: break-word; }
    .col-catatan { width: 50%; text-align: justify; }
    .styled-table tbody tr:hover { background-color: #f9fafb; }
    
    /* AUDIO RECORDER STYLE */
    .audio-recorder-st {
        display: flex;
        justify-content: center;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# --- SIDEBAR: AI CONFIG ---
with st.sidebar:
    st.header("🧠 Pengaturan AI")
    api_key = st.text_input(
        "Google Gemini API Key:", 
        value="AIzaSyD2uwPdGO2Y8bfd64cbWiNiffJ_2Imy9kC", 
        type="password"
    )
    
    if api_key:
        try:
            genai.configure(api_key=api_key)
            st.success("✅ AI Terhubung!")
        except Exception as e:
            st.error(f"Koneksi AI Gagal: {e}")
    else:
        st.warning("Masukkan API Key.")
    
    st.markdown("---")
    st.info("Fitur Voice Note aktif. Jika transkripsi gagal, silakan ketik manual.")

# --- TITLE ---
st.markdown("""
<div class="header-box">
    <h1>QPR Report Dashboard 360°</h1>
    <p>Performance Evaluation System • Internal Organization Development</p>
</div>
""", unsafe_allow_html=True)

# --- FUNGSI AI YANG AMAN (ANTI ERROR 404) ---
def generate_content_safe(prompt, audio_bytes=None):
    # Selalu gunakan versi 1.5
    model_name = 'gemini-1.5-flash' 
    
    try:
        model = genai.GenerativeModel(model_name)
        
        # KASUS 1: ADA AUDIO
        if audio_bytes:
            response = model.generate_content([
                prompt, 
                {"mime_type": "audio/webm", "data": audio_bytes}
            ])
        # KASUS 2: HANYA TEKS
        else:
            response = model.generate_content(prompt)
            
        return response.text

    except Exception as e:
        # Jika terjadi error 404 atau lainnya, berikan pesan yang jelas
        return f"⚠️ Terjadi kesalahan pada layanan AI: {str(e)}"

# --- NAVIGATION TABS ---
tab_anggota, tab_leader = st.tabs(["👥 Evaluasi Anggota (Excel)", "🎙️ Evaluasi Ketua & Wakil (Mode Diskusi)"])

# =================================================================================================
# TAB 1: EVALUASI ANGGOTA (EXCEL)
# =================================================================================================
with tab_anggota:
    uploaded_file = st.file_uploader("Upload File Excel QPR II (.xlsx)", type=['xlsx'], key="excel_uploader")

    if uploaded_file is not None:
        try:
            df_recap_names = pd.read_excel(uploaded_file, sheet_name="Recap Point Penilaian", header=1)
            member_list = df_recap_names['Nama Anggota'].dropna().unique()

            @st.cache_data
            def process_member_data(file_input, member_name):
                sheets = ["Head of Division", "Deputy Head of Division"]
                results = {
                    "Kinerja": {"scores": [], "comments": []},
                    "Inisiatif": {"scores": [], "comments": []},
                    "Kolaborasi": {"scores": [], "comments": []},
                    "Partisipasi": {"scores": [], "comments": []},
                    "Waktu": {"scores": [], "comments": []}
                }
                for sheet in sheets:
                    try:
                        df_raw = pd.read_excel(file_input, sheet_name=sheet, header=None)
                        for keyword in results.keys():
                            matches = df_raw.index[df_raw.apply(lambda r: r.astype(str).str.contains(keyword, case=False, na=False).any(), axis=1)].tolist()
                            if not matches: continue
                            start_row = matches[0]
                            df_sliced = df_raw.iloc[start_row:]
                            head_matches = df_sliced.index[df_sliced.apply(lambda r: r.astype(str).str.contains("Nama Anggota", case=False, na=False).any(), axis=1)].tolist()
                            if not head_matches: continue
                            header_idx = head_matches[0]
                            header_row = df_raw.iloc[header_idx] 

                            total_sk_idx = -1
                            for idx, val in header_row.items():
                                if isinstance(val, str) and "Total SK" in val:
                                    total_sk_idx = idx
                                    break
                            
                            for i in range(1, 50):
                                curr = header_idx + i
                                if curr >= len(df_raw): break
                                row = df_raw.iloc[curr]
                                
                                if row.astype(str).str.contains(member_name, case=False, na=False).any():
                                    if total_sk_idx != -1:
                                        try: results[keyword]["scores"].append(float(row[total_sk_idx]))
                                        except: pass
                                    for val in row:
                                        if isinstance(val, str) and len(val) > 5:
                                            clean = val.strip()
                                            ignore = ["nan", "nil", "-", "0", member_name.lower()]
                                            if clean.lower() not in ignore:
                                                results[keyword]["comments"].append(clean)
                                    break
                    except: continue
                return results

            selected_member = st.selectbox("Pilih Anggota Tim:", member_list)
            raw_data = process_member_data(uploaded_file, selected_member)

            def calc_avg(keyword):
                scores = raw_data[keyword]["scores"]
                if not scores: return 0
                return sum(scores) / len(scores)

            s_kinerja = calc_avg("Kinerja")
            s_inisiatif = calc_avg("Inisiatif")
            s_kolab = calc_avg("Kolaborasi")
            s_partisipasi = calc_avg("Partisipasi")
            s_waktu = calc_avg("Waktu")
            final_total = (s_kinerja * 0.30) + (s_inisiatif * 0.15) + (s_kolab * 0.20) + (s_partisipasi * 0.20) + (s_waktu * 0.15)

            def get_narrative_simple(keyword, score, context="detail"):
                comments = []
                if context == "detail":
                    comments = raw_data[keyword]["comments"]
                else:
                    for k in raw_data:
                        comments.extend(raw_data[k]["comments"])
                
                points_text = "\n".join([f"- {c}" for c in comments])
                if context == "detail":
                    prompt = f"Kamu HRD. Buat 1 paragraf evaluasi bahasa Indonesia baku untuk {selected_member}, aspek {keyword}, skor {score}. Poin: {points_text}"
                else:
                    prompt = f"Kamu HRD. Buat Executive Summary kinerja {selected_member}, total skor {score}. Rangkum kekuatan & kelemahan dari poin ini: {points_text}"
                
                return generate_content_safe(prompt)

            st.markdown(f"### 📄 Laporan Anggota: {selected_member}")
            exec_sum = get_narrative_simple("All", final_total, "summary")
            st.markdown(f'<div class="narrative-box"><div class="narrative-title">📋 Executive Summary</div><div class="narrative-text">{exec_sum}</div></div>', unsafe_allow_html=True)
            
            c1, c2 = st.columns([2, 1])
            with c2:
                st.markdown('<div class="css-card">', unsafe_allow_html=True)
                st.markdown(f'<div class="score-container"><div class="score-val">{final_total:.2f}</div><div class="score-label">Total Score</div></div>', unsafe_allow_html=True)
                fig = go.Figure(data=go.Scatterpolar(r=[s_kinerja, s_inisiatif, s_kolab, s_partisipasi, s_waktu], theta=['Kinerja', 'Inisiatif', 'Kolaborasi', 'Partisipasi', 'Waktu'], fill='toself', line_color='#7c3aed'))
                fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), height=250, margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            with c1:
                st.markdown('<div class="css-card">', unsafe_allow_html=True)
                table_html = f"""
                <table class="styled-table">
                    <thead><tr><th>Komponen</th><th style='text-align:center'>Skor</th><th>Evaluasi</th></tr></thead>
                    <tbody>
                        <tr><td>Kinerja</td><td style='text-align:center; font-weight:bold'>{s_kinerja:.2f}</td><td>{get_narrative_simple("Kinerja", s_kinerja)}</td></tr>
                        <tr><td>Inisiatif</td><td style='text-align:center; font-weight:bold'>{s_inisiatif:.2f}</td><td>{get_narrative_simple("Inisiatif", s_inisiatif)}</td></tr>
                        <tr><td>Kolaborasi</td><td style='text-align:center; font-weight:bold'>{s_kolab:.2f}</td><td>{get_narrative_simple("Kolaborasi", s_kolab)}</td></tr>
                        <tr><td>Partisipasi</td><td style='text-align:center; font-weight:bold'>{s_partisipasi:.2f}</td><td>{get_narrative_simple("Partisipasi", s_partisipasi)}</td></tr>
                        <tr><td>Waktu</td><td style='text-align:center; font-weight:bold'>{s_waktu:.2f}</td><td>{get_narrative_simple("Waktu", s_waktu)}</td></tr>
                    </tbody>
                </table>
                """
                st.markdown(table_html, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Error baca Excel: {e}")
    else:
        st.info("Upload Excel untuk melihat nilai Anggota.")


# =================================================================================================
# TAB 2: LEADER EVALUATION (VOICE TO POINTS -> REPORT)
# =================================================================================================
with tab_leader:
    st.markdown("### 🎙️ Evaluasi Ketua & Wakil Divisi (Mode Diskusi)")
    st.write("Diskusikan kinerja Ketua/Wakil. Rekam suara Anda, AI akan mengubahnya menjadi **Poin Catatan**, lalu membuat laporan lengkap.")

    col_input, col_result = st.columns([1, 1.5])

    # --- STATE MANAGEMENT ---
    if 'text_discussion_val' not in st.session_state:
        st.session_state['text_discussion_val'] = ""

    with col_input:
        st.markdown('<div class="css-card">', unsafe_allow_html=True)
        st.subheader("1. Input Diskusi")
        
        target_role = st.selectbox("Siapa yang dinilai?", ["Ketua Divisi", "Wakil Ketua Divisi"])
        target_name = st.text_input("Nama Lengkap:", placeholder="Contoh: Budi Santoso")

        st.markdown("---")
        st.write("**📝 Rekam Suara (Diskusi Gabungan):**")
        
        # 1. AUDIO RECORDER
        audio_data = mic_recorder(start_prompt="🎤 Mulai Rekam", stop_prompt="⏹️ Selesai & Transkrip", key="rec_disc", format="webm")
        
        # 2. TRANSKRIPSI OTOMATIS
        if audio_data:
            with st.spinner("🤖 Mendengarkan & Meringkas poin-poin..."):
                prompt_transcribe = """
                Dengarkan audio ini. Tuliskan poin-poin penting evaluasi kinerja (Kelebihan & Kekurangan) dalam Bahasa Indonesia.
                Langsung ke poin intinya.
                """
                # Gunakan fungsi safe
                points_result = generate_content_safe(prompt_transcribe, audio_bytes=audio_data['bytes'])
                
                # Simpan ke Session State
                st.session_state['text_discussion_val'] = points_result
                st.success("✅ Selesai!")

        # 3. TEXT AREA
        text_discussion = st.text_area(
            "Catatan Poin Diskusi (Otomatis terisi dari suara atau ketik manual):", 
            value=st.session_state['text_discussion_val'],
            height=250, 
            placeholder="Jika rekaman gagal, silakan ketik hasil diskusi di sini..."
        )
        
        if text_discussion != st.session_state['text_discussion_val']:
             st.session_state['text_discussion_val'] = text_discussion

        analyze_btn = st.button("🚀 Proses Analisis Laporan", type="primary")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_result:
        if analyze_btn and api_key and target_name and text_discussion:
            with st.spinner("🤖 AI sedang memilah per kategori dan merangkai cerita akhir..."):
                try:
                    # PROMPT
                    prompt_ai = f"""
                    Kamu adalah Senior HR Evaluator. Analisis catatan diskusi tentang {target_name} ({target_role}).

                    INPUT: "{text_discussion}"

                    TUGAS:
                    1. **PER POINT**: Skor (0-100) & Narasi Detail (2-3 kalimat) untuk 5 aspek (Kinerja, Inisiatif, Kolaborasi, Partisipasi, Waktu).
                    2. **FINAL SYNTHESIS**: Gabungkan semua poin menjadi SATU NARASI PANJANG (200 kata), mengalir, profesional, ada kesimpulan tegas.

                    FORMAT OUTPUT (JSON ONLY):
                    {{
                        "scores": {{ "Kinerja": 0, "Inisiatif": 0, "Kolaborasi": 0, "Partisipasi": 0, "Waktu": 0 }},
                        "details": {{ "Kinerja": "...", "Inisiatif": "...", "Kolaborasi": "...", "Partisipasi": "...", "Waktu": "..." }},
                        "final_synthesis": "...",
                        "action_plan": "..."
                    }}
                    """

                    # Panggil AI (Pakai fungsi safe fallback)
                    json_raw = generate_content_safe(prompt_ai)
                    
                    # Bersihkan JSON
                    json_str = json_raw.replace("```json", "").replace("```", "").strip()
                    data_ai = json.loads(json_str)

                    # --- HASIL ---
                    scores = data_ai['scores']
                    final_score = (scores['Kinerja']*0.3) + (scores['Inisiatif']*0.15) + (scores['Kolaborasi']*0.2) + (scores['Partisipasi']*0.2) + (scores['Waktu']*0.15)

                    st.markdown(f"### 📊 Laporan Evaluasi: {target_name}")
                    
                    # 1. SCORE CARD
                    c_res1, c_res2 = st.columns([1, 2])
                    with c_res1:
                         st.markdown(f'<div class="score-container"><div class="score-val">{final_score:.1f}</div><div class="score-label">Final Score</div></div>', unsafe_allow_html=True)
                    with c_res2:
                        cats = list(scores.keys())
                        vals = list(scores.values())
                        fig = go.Figure(data=go.Scatterpolar(r=vals, theta=cats, fill='toself', line_color='#7c3aed'))
                        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), height=200, margin=dict(t=20, b=20, l=20, r=20))
                        st.plotly_chart(fig, use_container_width=True)

                    # 2. TABEL RINCIAN
                    st.markdown("<div class='sec-title'>1. Rincian Evaluasi Per Kategori</div>", unsafe_allow_html=True)
                    details = data_ai['details']
                    rows = ""
                    for cat, narrative in details.items():
                        score = scores[cat]
                        color = "#16a34a" if score >= 80 else "#ca8a04" if score >= 70 else "#dc2626"
                        rows += f"<tr><td><b>{cat}</b></td><td style='text-align:center; font-weight:bold; color:{color}'>{score}</td><td class='col-catatan'>{narrative}</td></tr>"

                    st.markdown(f"""
                    <table class="styled-table">
                        <thead><tr><th>Kategori</th><th style='text-align:center'>Skor</th><th>Evaluasi Detail</th></tr></thead>
                        <tbody>{rows}</tbody>
                    </table>
                    """, unsafe_allow_html=True)

                    # 3. FINAL SYNTHESIS
                    st.markdown("<div class='sec-title'>2. Kesimpulan Akhir (Sintesis)</div>", unsafe_allow_html=True)
                    st.markdown(f"""
                    <div class="narrative-box">
                        <div class="narrative-title">📋 Analisis Menyeluruh</div>
                        <div class="narrative-text">{data_ai['final_synthesis']}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.success(f"💡 **Rekomendasi Tindakan:** {data_ai['action_plan']}")

                except Exception as e:
                    st.error(f"Gagal memproses Laporan: {e}")
                    st.write("Tips: Jika error JSON, coba klik 'Proses Analisis' sekali lagi.")
        
        elif analyze_btn and not text_discussion:
            st.warning("Mohon isi catatan diskusi terlebih dahulu.")
