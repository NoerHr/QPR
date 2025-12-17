import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import re
import random

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Laporan QPR", layout="wide", page_icon="📝")

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
    .score-container {
        text-align: center;
        padding: 20px;
        border-bottom: 2px dashed #f3f4f6;
        margin-bottom: 20px;
    }
    .score-val { font-size: 4rem; font-weight: 800; color: #6d28d9; line-height: 1; }
    .score-label { font-size: 0.85rem; font-weight: 700; color: #4b5563; text-transform: uppercase; margin-top: 10px; letter-spacing: 1.5px; }

    /* EXECUTIVE SUMMARY BOX (NEW) */
    .exec-summary {
        background-color: #f5f3ff;
        border-left: 6px solid #7c3aed;
        padding: 25px;
        border-radius: 12px;
        margin-bottom: 25px;
    }
    .exec-title {
        color: #5b21b6;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 10px;
        font-size: 0.95rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .exec-text {
        color: #374151;
        line-height: 1.8;
        text-align: justify;
        font-size: 1rem;
    }

    /* TABLE STYLE */
    .styled-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: 0.9rem;
        border-radius: 10px;
        overflow: hidden;
    }
    .styled-table th {
        background-color: #5b21b6;
        color: #ffffff;
        font-weight: 600;
        text-transform: uppercase;
        font-size: 0.75rem;
        letter-spacing: 1px;
        padding: 14px 16px;
        text-align: left;
    }
    .styled-table td {
        padding: 16px;
        border-bottom: 1px solid #f3f4f6;
        vertical-align: top;
        color: #374151;
        line-height: 1.6;
        word-wrap: break-word;
        overflow-wrap: break-word;
    }
    .col-komponen { width: 25%; font-weight: 700; color: #111827; }
    .col-bobot { width: 10%; text-align: center; color: #6b7280; }
    .col-skor { width: 10%; text-align: center; font-weight: 800; color: #7c3aed; font-size: 1.05rem; }
    .col-catatan { width: 40%; text-align: justify; color: #1f2937; line-height: 1.6; }
    .col-sumber { width: 15%; color: #9ca3af; font-size: 0.8rem; font-style: italic; }
    
    .styled-table tbody tr:hover { background-color: #f9fafb; }
    .total-row td {
        background-color: #f5f3ff;
        font-weight: 800;
        color: #5b21b6;
        border-top: 2px solid #ddd6fe;
        font-size: 1rem;
    }

    /* FOOTER */
    .ttd-box { text-align: center; margin-top: 20px; }
    .ttd-role { font-size: 0.85rem; color: #6b7280; font-weight: 600; margin-bottom: 60px; text-transform: uppercase; }
    .ttd-name { font-size: 1rem; color: #111827; font-weight: 700; border-top: 2px solid #d1d5db; display: inline-block; padding-top: 10px; min-width: 180px; }
</style>
""", unsafe_allow_html=True)

# --- HEADER TITLE ---
st.markdown("""
<div class="header-box">
    <h1>QPR Report Dashboard</h1>
    <p>Performance Evaluation System • Internal Organization Development</p>
</div>
""", unsafe_allow_html=True)

# --- UPLOAD SECTION ---
uploaded_file = st.file_uploader("Upload File Excel QPR II (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # --- 1. AMBIL LIST NAMA ---
        df_recap_names = pd.read_excel(uploaded_file, sheet_name="Recap Point Penilaian", header=1)
        member_list = df_recap_names['Nama Anggota'].dropna().unique()

        # --- FUNGSI PROSES DATA ---
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

        # --- 2. SIDEBAR FILTER ---
        st.sidebar.markdown("### 🔍 Filter Anggota")
        selected_member = st.sidebar.selectbox("Pilih Nama:", member_list)
        st.sidebar.markdown("---")

        # --- 3. EKSEKUSI ---
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

        # --- FUNGSI EXECUTIVE SUMMARY (RANGKUMAN KESELURUHAN) ---
        def generate_overall_summary(score, all_data):
            # 1. Kumpulkan SEMUA komentar unik
            all_comments = []
            for key in all_data:
                for c in all_data[key]["comments"]:
                    c = re.sub(r'^[\d\-\.\)\•\"]+\s*', '', str(c)).replace('"', '').strip()
                    if c and len(c) > 5:
                        # Kapitalisasi
                        c = c[0].upper() + c[1:]
                        if c[-1] not in ['.', '!', '?']: c += "."
                        all_comments.append(c)
            
            # Hapus Duplikat
            unique_comments = list(set(all_comments))

            # 2. Tentukan Intro Berdasarkan Skor
            if score >= 90:
                intro = f"Secara keseluruhan, {selected_member} menunjukkan performa yang **sangat memuaskan** pada kuartal ini."
                tone = "positif"
            elif score >= 75:
                intro = f"Secara keseluruhan, {selected_member} menunjukkan performa yang **cukup baik**, namun terdapat beberapa area yang memerlukan perhatian."
                tone = "netral"
            else:
                intro = f"Secara keseluruhan, kinerja {selected_member} **memerlukan evaluasi mendalam** dikarenakan skor berada di bawah ekspektasi standar tim."
                tone = "kritis"

            # 3. Ambil 2-3 Poin Kunci secara Acak (Agar tidak terlalu panjang, tapi representatif)
            # Logika: Jika ada komentar 'negatif' (kata kunci: kurang, jarang, terlambat), prioritaskan itu untuk feedback
            priority_comments = []
            normal_comments = []
            
            neg_keywords = ['kurang', 'jarang', 'terlambat', 'tidak', 'kendala', 'hambatan', 'perlu']
            
            for c in unique_comments:
                if any(k in c.lower() for k in neg_keywords):
                    priority_comments.append(c)
                else:
                    normal_comments.append(c)
            
            summary_body = ""
            
            # Strategi Penyusunan Paragraf
            selected_points = []
            if priority_comments:
                selected_points.extend(priority_comments[:2]) # Ambil max 2 masalah utama
            if len(selected_points) < 3 and normal_comments:
                 selected_points.extend(normal_comments[:2]) # Ambil sisa dari komentar biasa
            
            if selected_points:
                summary_body = " Evaluator menyoroti beberapa poin penting, antara lain: " + " ".join(selected_points)
            else:
                summary_body = " Tidak ada catatan spesifik dari evaluator untuk periode ini."

            # 4. Outro
            if tone == "positif": outro = " Diharapkan prestasi ini dapat dipertahankan dan menjadi motivasi bagi anggota lainnya."
            elif tone == "netral": outro = " Diharapkan ke depannya dapat lebih proaktif dan meningkatkan konsistensi kinerja."
            else: outro = " Disarankan untuk segera melakukan sesi konseling kinerja (one-on-one) untuk merumuskan rencana perbaikan."

            return intro + summary_body + outro

        # --- FUNGSI PER DETAIL (ROW TABLE) ---
        def generate_detail_narrative(keyword):
            comments = raw_data[keyword]["comments"]
            if not comments: return "Tidak ada catatan spesifik."
            
            clean = []
            seen = set()
            for c in comments:
                c = re.sub(r'^[\d\-\.\)\•\"]+\s*', '', str(c)).replace('"', '').strip()
                if c.lower() not in seen and len(c) > 3:
                    seen.add(c.lower())
                    if c: clean.append(c[0].upper() + c[1:] + ("." if c[-1] not in ['.','!','?'] else ""))
            
            return " ".join(clean)

        def fmt_num(val):
            return f"{val:.2f}".rstrip('0').rstrip('.')

        # --- 4. LAYOUT ---
        st.markdown(f"<h2 style='margin-bottom:20px; color:#111827;'>Laporan Individu: <span style='color:#7c3aed; font-weight:800;'>{selected_member}</span></h2>", unsafe_allow_html=True)

        # --- SECTION: EXECUTIVE SUMMARY (RANGKUMAN KESELURUHAN) ---
        exec_sum_text = generate_overall_summary(final_total, raw_data)
        
        st.markdown(f"""
        <div class="exec-summary">
            <div class="exec-title">
                <span>📋 Catatan Evaluator Keseluruhan</span>
            </div>
            <div class="exec-text">
                {exec_sum_text}
            </div>
        </div>
        """, unsafe_allow_html=True)

        col_main, col_side = st.columns([1.8, 1])

        # KANAN: GRAFIK
        with col_side:
            st.markdown('<div class="css-card">', unsafe_allow_html=True)
            st.markdown(f"""
            <div class="score-container">
                <div class="score-val">{fmt_num(final_total)}</div>
                <div class="score-label">Total Performance Score</div>
            </div>
            """, unsafe_allow_html=True)
            
            cats = ['Kinerja', 'Inisiatif', 'Kolaborasi', 'Partisipasi', 'Waktu']
            vals = [s_kinerja, s_inisiatif, s_kolab, s_partisipasi, s_waktu]
            fig = go.Figure()
            fig.add_trace(go.Scatterpolar(r=vals, theta=cats, fill='toself', name=selected_member, line_color='#7c3aed', fillcolor='rgba(124, 58, 237, 0.2)'))
            fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), margin=dict(t=10, b=10, l=30, r=30), height=280, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # KIRI: TABEL DETAIL
        with col_main:
            st.markdown('<div class="css-card">', unsafe_allow_html=True)
            st.markdown("<div class='sec-title'>📝 Rincian Evaluasi per Komponen</div>", unsafe_allow_html=True)
            
            html_table = f"""
            <table class="styled-table">
                <thead>
                    <tr>
                        <th class="col-komponen">Komponen</th>
                        <th class="col-bobot">Bobot</th>
                        <th class="col-skor">Skor</th>
                        <th class="col-catatan">Catatan Evaluator</th>
                        <th class="col-sumber">Sumber</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="col-komponen">Kinerja & Deliverables</td>
                        <td class="col-bobot">30%</td>
                        <td class="col-skor">{fmt_num(s_kinerja)}</td>
                        <td class="col-catatan">{generate_detail_narrative("Kinerja")}</td>
                        <td class="col-sumber">Spreadsheet, Notion</td>
                    </tr>
                    <tr>
                        <td class="col-komponen">Inisiatif</td>
                        <td class="col-bobot">15%</td>
                        <td class="col-skor">{fmt_num(s_inisiatif)}</td>
                        <td class="col-catatan">{generate_detail_narrative("Inisiatif")}</td>
                        <td class="col-sumber">Spreadsheet, Notion</td>
                    </tr>
                    <tr>
                        <td class="col-komponen">Kolaborasi & Komunikasi</td>
                        <td class="col-bobot">20%</td>
                        <td class="col-skor">{fmt_num(s_kolab)}</td>
                        <td class="col-catatan">{generate_detail_narrative("Kolaborasi")}</td>
                        <td class="col-sumber">Spreadsheet, Notion</td>
                    </tr>
                    <tr>
                        <td class="col-komponen">Partisipasi & Kehadiran</td>
                        <td class="col-bobot">20%</td>
                        <td class="col-skor">{fmt_num(s_partisipasi)}</td>
                        <td class="col-catatan">{generate_detail_narrative("Partisipasi")}</td>
                        <td class="col-sumber">Spreadsheet, Notion</td>
                    </tr>
                    <tr>
                        <td class="col-komponen">Ketepatan Waktu</td>
                        <td class="col-bobot">15%</td>
                        <td class="col-skor">{fmt_num(s_waktu)}</td>
                        <td class="col-catatan">{generate_detail_narrative("Waktu")}</td>
                        <td class="col-sumber">Spreadsheet, Notion</td>
                    </tr>
                    <tr class="total-row">
                        <td>TOTAL AKHIR</td>
                        <td style="text-align:center;">100%</td>
                        <td class="col-skor">{fmt_num(final_total)}</td>
                        <td></td>
                        <td></td>
                    </tr>
                </tbody>
            </table>
            """
            st.markdown(html_table, unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # FOOTER
        st.markdown('<div class="css-card">', unsafe_allow_html=True)
        st.markdown("<h4 style='text-align:center; color:#4b5563; font-size:0.9rem; margin-bottom:30px; letter-spacing:1px;'>LEMBAR PENGESAHAN</h4>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1,1,1])
        with c1:
            st.markdown("<div class='ttd-box'><div class='ttd-role'>Ketua SGA</div><div class='ttd-name'>Sayyid Abdul Aziz Haidar</div></div>", unsafe_allow_html=True)
        with c3:
            st.markdown("<div class='ttd-box'><div class='ttd-role'>Ketua IOD</div><div class='ttd-name'>Ratu Bilqis</div></div>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error("Terjadi kesalahan teknis saat membaca file.")
        st.write(f"Detail Error: {e}")

else:
    st.markdown("""<div style="text-align:center; padding:60px; background:white; border-radius:16px; border:2px dashed #d1d5db; color:#6b7280; margin-top:20px;"><h3>👋 Selamat Datang</h3><p>Silakan upload file Excel <b>QPR II</b> untuk memulai.</p></div>""", unsafe_allow_html=True)