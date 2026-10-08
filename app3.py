import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ページ全体のレイアウト設定（ワイド表示）
st.set_page_config(layout="wide", page_title="銀行・CNSデータ 統合管理ツール", page_icon="🏦")
st.title("🏦 銀行・CNSデータ 統合管理ツール")

# タブ構成（全4タブ）
tab1, tab2, tab3, tab4 = st.tabs([
    "引き落としデータ作成（書き出し）",                                       # タブ1
    "入金データ変換（読み込み）",                                             # タブ2
    "CNSコンビニ収納データ作成（17列CSV）",      # タブ3
    "CNSコンビニ収納データ取込（入金データ＋入金マスタ）"          # タブ4
])


# =====================================================================
# 🛠️ 共通ヘルパー関数（Shift-JIS/CP932のバイト数基準で正確にパディング/カットする）
# =====================================================================
def sjis_ljust(string, length):
    string = str(string).strip() if pd.notnull(string) and str(string).lower() != 'nan' else ""
    encoded = string.encode('cp932', errors='ignore')
    if len(encoded) >= length:
        cut_str = encoded[:length].decode('cp932', errors='ignore')
        cut_encoded = cut_str.encode('cp932', errors='ignore')
        return cut_str + (' ' * (length - len(cut_encoded)))
    return string + (' ' * (length - len(encoded)))

def sjis_zfill(string, length):
    string = str(string).strip() if pd.notnull(string) and str(string).lower() != 'nan' else ""
    encoded = string.encode('cp932', errors='ignore')
    if len(encoded) >= length:
        cut_str = encoded[:length].decode('cp932', errors='ignore')
        cut_encoded = cut_str.encode('cp932', errors='ignore')
        return ('0' * (length - len(cut_encoded))) + cut_str
    return ('0' * (length - len(encoded))) + string

def sjis_zfill_ljust(string, length):
    string = str(string).strip() if pd.notnull(string) and str(string).lower() != 'nan' else ""
    encoded = string.encode('cp932', errors='ignore')
    if len(encoded) >= length:
        cut_str = encoded[:length].decode('cp932', errors='ignore')
        cut_encoded = cut_str.encode('cp932', errors='ignore')
        return cut_str + ('0' * (length - len(cut_encoded)))
    return string + ('0' * (length - len(encoded)))


# =====================================================================
# 📤 タブ1：引き落としデータ作成（書き出し）【千葉銀行成功事例完全一致版】
# =====================================================================
with tab1:
    st.subheader("📤 引き落とし用ファイル（全銀協形式）作成")
    st.markdown("依頼エクセルから全銀協規格フォーマットの引き落としデータを生成します。（千葉銀行の成功事例フォーマットに完全準拠）")
    target_bank = st.radio("提出先", ["千葉銀行", "京葉銀行"], key="tab1_bank_radio")
    target_date = st.date_input("振込指定日", key="tab1_target_date")
    month_day = target_date.strftime("%m%d")
    uploaded_excel = st.file_uploader("依頼エクセルをアップロードしてください（.xlsx）", type=["xlsx"], key="tab1_uploader")
    
    if uploaded_excel:
        df = pd.read_excel(uploaded_excel)
        
        required_cols = ['銀行ID', '支店ID', '普/当', '口座番号', 'ｺｳｻﾞﾒｲ', '請求額（前期未収含む）', '整理番号']
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ エラー：必須列が不足しています。必要列: {required_cols}")
        elif df['請求額（前期未収含む）'].isnull().any():
            st.warning("⚠️ 警告：金額が空欄の行があります。")
        else:
            if st.button("提出用ファイルを作成", key="tab1_exec_btn"):
                data_lines = []
                total_amount = 0
                record_count = 0
                
                for _, row in df.iterrows():
                    amt = int(row['請求額（前期未収含む）'])
                    total_amount += amt
                    record_count += 1
                    
                    bank_id_4桁 = str(row['銀行ID']).strip().zfill(4)
                    p_bank_id = "2" + bank_id_4桁
                    
                    if target_bank == "京葉銀行":
                        p_space1 = " " * 15
                        p_branch_id = sjis_zfill(str(row['支店ID']).strip(), 3)
                        p_space2 = " " * 19
                        p_acc_type = sjis_zfill(str(row['普/当']).strip(), 1)
                        p_acc_num = sjis_zfill(str(row['口座番号']).strip(), 7)
                        p_name = sjis_ljust(row['ｺｳｻﾞﾒｲ'], 30)
                        p_amt = sjis_zfill(str(amt), 10)
                        p_kbn = sjis_zfill("0", 1)
                        
                        nyukin_val = str(row.get('入金番号', '')).strip().replace("-", "")
                        if not nyukin_val or nyukin_val.lower() == 'nan': 
                            nyukin_val = "0"
                        p_nyukin = sjis_zfill(nyukin_val, 13)
                        
                        seiri_val = str(row['整理番号']).strip()
                        p_seiri = sjis_zfill(seiri_val, 7)
                        p_unused = sjis_zfill("0", 1)
                        
                        line_str = (
                            p_bank_id + p_space1 + p_branch_id + p_space2 +
                            p_acc_type + p_acc_num + p_name + p_amt +
                            p_kbn + p_nyukin + p_seiri + p_unused
                        )
                    else:
                        p_branch_id = sjis_zfill(str(row['支店ID']).strip(), 3)
                        p_acc_type = sjis_zfill(str(row['普/当']).strip(), 1)
                        p_acc_num = sjis_zfill(str(row['口座番号']).strip(), 7)
                        p_name = sjis_ljust(row['ｺｳｻﾞﾒｲ'], 30)
                        p_amt = sjis_zfill(str(amt), 10)
                        
                        kotei_val = str(row.get('固定値', '0')).strip()
                        if not kotei_val or kotei_val.lower() == 'nan': 
                            kotei_val = "0"
                        p_kbn = sjis_zfill(kotei_val, 2)
                        
                        p_bank_code_2134 = "2134"
                        
                        seiri_val = str(row['整理番号']).strip()
                        p_seiri = sjis_zfill_ljust(seiri_val, 6)
                        
                        nyukin_val = str(row.get('入金番号', '')).strip().replace("-", "")
                        if not nyukin_val or nyukin_val.lower() == 'nan': 
                            nyukin_val = "0"
                        p_nyukin = sjis_zfill_ljust(nyukin_val, 8)
                        
                        branch_code_str = sjis_zfill(str(row['支店ID']).strip(), 3)
                        
                        line_str = (
                            p_bank_id +             
                            " " * 9 +               
                            branch_code_str +       
                            sjis_ljust("ﾏｸﾊﾘ", 10) +    
                            " " * 3 +               
                            p_acc_type +            
                            p_acc_num +             
                            p_name +                
                            p_amt +                 
                            p_kbn +                 
                            p_bank_code_2134 +      
                            p_seiri +               
                            p_nyukin +              
                            " " * 4                 
                        )
                    
                    encoded_line = line_str.encode('cp932', errors='ignore')
                    if len(encoded_line) > 120:
                        encoded_line = encoded_line[:120]
                    elif len(encoded_line) < 120:
                        encoded_line = encoded_line + b' ' * (120 - len(encoded_line))
                    
                    data_lines.append(encoded_line.decode('cp932', errors='ignore'))
                
                if target_bank == "京葉銀行":
                    h_code = "1"
                    h_data_type = "911"
                    h_client_code = sjis_zfill("0000935493", 10)
                    h_client_name = sjis_ljust("ｼﾔ)ﾁﾊﾞﾆｼﾎｳｼﾞﾝｶｲ", 30)
                    h_space1 = " " * 10
                    h_date = sjis_zfill(month_day, 4)
                    h_bank_code = sjis_zfill("0522", 4)
                    h_bank_name = sjis_ljust("ｹｲﾖｳｷﾞﾝｺｳ", 10)
                    h_space_mid = " " * 1
                    h_branch_code = sjis_zfill("447", 3)
                    h_branch_name = sjis_ljust("ﾏｸﾊﾘｼﾃﾝ", 10)
                    h_space_mid2 = " " * 3
                    h_account_type = sjis_zfill("1", 1)
                    h_account_num = sjis_zfill("1273441", 7)
                    h_space2 = " " * 18
                    
                    header_str = (
                        h_code + h_data_type + h_client_code + h_client_name +
                        h_space1 + h_date + h_bank_code + h_bank_name +
                        h_space_mid + h_branch_code + h_branch_name +
                        h_space_mid2 + h_account_type + h_account_num + h_space2
                    )
                else:
                    header_str = f"19100000060922{sjis_ljust('ｼﾔ)ﾁﾊﾞﾆｼﾎｳｼﾞﾝｶｲ', 30)}{' ' * 19}{month_day}0134{sjis_ljust('ﾁﾊﾞ', 10)}{' ' * 12}002{sjis_ljust('ﾏｸﾊﾘ', 10)}{' ' * 12}11122028                 "

                enc_header = header_str.encode('cp932', errors='ignore')
                if len(enc_header) > 120:
                    enc_header = enc_header[:120]
                elif len(enc_header) < 120:
                    enc_header = enc_header + b' ' * (120 - len(enc_header))
                header = enc_header.decode('cp932', errors='ignore')

                footer_raw = f"8{str(record_count).zfill(5)}{str(total_amount).zfill(12)}{str(total_amount).zfill(12)}{'0' * 21}{str(record_count).zfill(5)}{str(total_amount).zfill(12)}{'0' * 5}"
                enc_footer = footer_raw.encode('cp932', errors='ignore')
                if len(enc_footer) > 120:
                    enc_footer = enc_footer[:120]
                elif len(enc_footer) < 120:
                    enc_footer = enc_footer + b' ' * (120 - len(enc_footer))
                footer = enc_footer.decode('cp932', errors='ignore')
                
                end_line = "9".ljust(120)
                
                final_output = [header] + data_lines + [footer] + [end_line]
                output_text = "\r\n".join(final_output)
                
                st.download_button(
                    label=f"💾 {target_bank}用データを作成", 
                    data=output_text.encode('cp932', errors='replace'), 
                    file_name="KOZKT.dat" if target_bank == "京葉銀行" else f"demand_{target_bank}_{datetime.now().strftime('%Y%m%d_%H%M')}.dat",
                    key="tab1_dl"
                )


# =====================================================================
# 📥 タブ2：入金データ変換（読み込み）【千葉=CSV / 京葉=.dat固定長（入金番号整数化対応）】
# =====================================================================
with tab2:
    st.subheader("📥 入金データ読み込み・変換")
    st.markdown("入金データを読み込み、結果区分や入金額を追加して変換します。（京葉銀行は.dat固定長 ＆ 入金番号の整数化（0頭カット）に対応）")
    
    col1, col2 = st.columns(2)
    with col1:
        bank = st.selectbox("読み込む銀行", ["千葉銀行", "京葉銀行(日本収納)"], key="tab2_bank")
    with col2:
        processing_date = st.date_input("処理日（入金日）を入力", value=datetime.now(), key="tab2_date")
        proc_date_str = processing_date.strftime("%Y/%m/%d")

    uploaded_file = st.file_uploader("入金データファイルをアップロードしてください（千葉銀行: .csv / 京葉銀行: .dat, .txt）", type=["csv", "txt", "dat"], key="tab2_uploader")
    
    result_mapping = {
        "0": "入金", "1": "残保無", "2": "口座無", "3": "預金者都合",
        "4": "振替依頼書無", "8": "委託者都合", "9": "その他"
    }

    if uploaded_file:
        try:
            try:
                raw_bytes = uploaded_file.read()
                raw = raw_bytes.decode("cp932", errors="ignore")
            except:
                raw = raw_bytes.decode("utf-8", errors="ignore")
                
            lines = raw.splitlines()
            parsed = []
            today_str = datetime.now().strftime("%Y/%m/%d")
            
            for line in lines:
                if not line.startswith("2"):  
                    continue
                
                # -------------------------------------------------------------
                # 📌 【京葉銀行】 全銀協固定長フォーマット（.dat）の処理
                # -------------------------------------------------------------
                if bank == "京葉銀行":
                    b_line = line.encode("cp932", errors="ignore")
                    if len(b_line) < 120:  
                        continue
                    try:
                        tail_str = b_line[101:120].decode("cp932", errors="ignore").strip()
                        last_digit = tail_str[-1] if len(tail_str) >= 1 else ""
                        result_text = result_mapping.get(last_digit, "その他")
                        seikyu_amt = int(b_line[80:90].decode("cp932", errors="ignore").strip())
                        nyukin_amt = seikyu_amt if last_digit == "0" else 0
                        
                        raw_nyukin_part = tail_str[:-1].strip() if len(tail_str) >= 1 else tail_str
                        nyukin_val = int(pd.to_numeric(pd.Series([raw_nyukin_part]), errors='coerce').fillna(0).iloc[0])
                        seiri_num = str(nyukin_val)
                        
                        row = {
                            "銀行ID": b_line[1:5].decode("cp932", errors="ignore").strip(), 
                            "支店ID": b_line[20:23].decode("cp932", errors="ignore").strip(),
                            "普/当": b_line[42:43].decode("cp932", errors="ignore").strip(), 
                            "口座番号": b_line[43:50].decode("cp932", errors="ignore").strip(),
                            "ｺｳｻﾞﾒｲ": b_line[50:80].decode("cp932", errors="ignore").strip(), 
                            "請求額": seikyu_amt, 
                            "入金額": nyukin_amt, 
                            "結果区分": result_text,
                            "整理番号": seiri_num, 
                            "入金番号": nyukin_val,
                            "処理日時": proc_date_str,
                            "入力日": today_str        
                        }
                        parsed.append(row)
                    except Exception:
                        continue
                
                # -------------------------------------------------------------
                # 📌 【千葉銀行】 CSV形式の戻りデータの処理
                # -------------------------------------------------------------
                else:
                    cols = [c.strip() for c in line.split(",")]
                    if len(cols) < 12:
                        continue
                    try:
                        bank_id = cols[1]
                        branch_id = cols[2]
                        acc_type = cols[3]
                        acc_num = cols[4]
                        koza_mei = cols[7]
                        seikyu_amt = int(cols[8])
                        last_digit = cols[11]
                        
                        result_text = result_mapping.get(last_digit, "その他")
                        nyukin_amt = seikyu_amt if last_digit == "0" else 0
                        
                        raw_col10 = cols[10] if len(cols) > 10 else ""
                        nyukin_val = raw_col10
                        seiri_num = raw_col10.replace("-", "")[:6]
                        
                        row = {
                            "銀行ID": bank_id, 
                            "支店ID": branch_id,
                            "普/当": acc_type, 
                            "口座番号": acc_num,
                            "ｺｳｻﾞﾒｲ": koza_mei, 
                            "請求額": seikyu_amt, 
                            "入金額": nyukin_amt, 
                            "結果区分": result_text,
                            "整理番号": seiri_num, 
                            "入金番号": nyukin_val,
                            "処理日時": proc_date_str,
                            "入力日": today_str        
                        }
                        parsed.append(row)
                    except Exception:
                        continue
            
            if parsed:
                df_parsed = pd.DataFrame(parsed)
                st.success(f"✅ {len(df_parsed)} 件のデータを正常に変換しました！")
                st.dataframe(df_parsed)
                
                csv = df_parsed.to_csv(index=False).encode('cp932', errors='ignore')
                st.download_button(
                    label="📥 処理結果をCSVとして保存",
                    data=csv,
                    file_name=f"log_{bank}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    key="tab2_dl"
                )
            else:
                st.warning("⚠️ 条件に一致するデータ行（先頭が '2' の行）が見つかりませんでした。ファイルの形式をご確認ください。")
                
        except Exception as e:
            st.error(f"❌ ファイルの処理中にエラーが発生しました: {e}")


# =====================================================================
# 🏪 タブ3：CNSコンビニ収納データコンバート（17列CSV）
# =====================================================================
with tab3:
    st.subheader("🏪 CNS コンビニ収納データ 17列CSV変換")
    st.markdown("17列構成のCSVファイルを読み込み、指定した各期限を設定してCNS指定ヘッダー名で出力します。")
    
    st.markdown("##### 1. お支払期限・バーコード取扱期限・ご請求内容の入力")
    c3_col1, c3_col2, c3_col3 = st.columns(3)
    with c3_col1:
        pay_limit_val = st.date_input("お支払期限", value=None, format="YYYY/MM/DD", key="tab3_pay_limit")
    with c3_col2:
        barcode_limit_val = st.date_input("バーコード取扱期限", value=None, format="YYYY/MM/DD", key="tab3_barcode_limit")
    with c3_col3:
        billing_content_val = st.text_input("ご請求内容（一括入力）", value="2026年度年会費", max_chars=20, key="tab3_billing_content")
        
    st.markdown("##### 2. CSVファイルの読み込み")
    uploaded_cns_csv = st.file_uploader("17列の元データCSVファイルを選択してください", type=["csv"], key="tab3_csv_uploader")
    
    target_headers = [
        "顧客コード", "払込人名", "払込人名（カナ）", "郵便番号", "電話番号", 
        "住所", "金額", "お支払期限", "バーコード取扱期限", "ご請求内容", 
        "内訳１（項目）", "内訳１（金額）", "内訳２（項目）", "内訳２（金額）", 
        "内訳３（項目）", "内訳３（金額）", "印紙有無"
    ]
    
    if uploaded_cns_csv is not None:
        try:
            try:
                df_tab3 = pd.read_csv(uploaded_cns_csv, header=None, skiprows=1)
            except UnicodeDecodeError:
                uploaded_cns_csv.seek(0)
                df_tab3 = pd.read_csv(uploaded_cns_csv, header=None, skiprows=1, encoding="cp932")
            
            actual_cols = df_tab3.shape[1]
            st.info(f"📁 読み込んだCSVの情報 - 列数: {actual_cols}列 / 行数: {df_tab3.shape[0]}行")
            
            if actual_cols != 17:
                st.warning(f"⚠️ アップロードされたCSVは17列ではありません（現在: {actual_cols}列）。自動的に17列に補正します。")
                if actual_cols < 17:
                    for i in range(actual_cols, 17):
                        df_tab3[i] = ""
                elif actual_cols > 17:
                    df_tab3 = df_tab3.iloc[:, :17]
            
            df_converted_cns = df_tab3.copy()
            df_converted_cns.columns = target_headers
            df_converted_cns["内訳１（金額）"] = df_converted_cns["金額"]
            
            if billing_content_val:
                df_converted_cns["ご請求内容"] = billing_content_val
            else:
                df_converted_cns["ご請求内容"] = ""
            
            if pay_limit_val:
                df_converted_cns["お支払期限"] = pay_limit_val.strftime("%Y%m%d")
            if barcode_limit_val:
                df_converted_cns["バーコード取扱期限"] = barcode_limit_val.strftime("%Y%m%d")
                
            st.markdown("##### 3. 変換データのプレビュー（先頭5行）")
            st.dataframe(df_converted_cns.head(5))
            
            csv_buf_sjis = io.StringIO()
            df_converted_cns.to_csv(csv_buf_sjis, index=False, encoding="cp932", lineterminator="\r\n")
            csv_bytes_sjis = csv_buf_sjis.getvalue().encode("cp932", errors="replace")
            
            csv_buf_utf8 = io.StringIO()
            df_converted_cns.to_csv(csv_buf_utf8, index=False, encoding="utf-8", lineterminator="\r\n")
            csv_bytes_utf8 = csv_buf_utf8.getvalue().encode("utf-8")
            
            st.markdown("##### 4. 変換データのダウンロード")
            c3_dl_col1, c3_dl_col2 = st.columns(2)
            
            with c3_dl_col1:
                st.download_button(
                    label="📥 CSVをダウンロード (Shift-JIS/cp932推奨)",
                    data=csv_bytes_sjis, file_name="cns_converted_sjis.csv",
                    mime="text/css", key="tab3_dl_sjis"
                )
            with c3_dl_col2:
                st.download_button(
                    label="📥 CSVをダウンロード (UTF-8)",
                    data=csv_bytes_utf8, file_name="cns_converted_utf8.csv",
                    mime="text/csv", key="tab3_dl_utf8"
                )
                
        except Exception as e:
            st.error(f"❌ コンバート処理中にエラーが発生しました: {e}")


# =====================================================================
# 🔗 タブ4：マスタ自動紐づけ（顧客コード ⇔ 整理番号）＋ 日付フォーマット変換 ＋ 不明入金切り分け
# =====================================================================
with tab4:
    st.subheader("🔗 マスタ自動紐づけ＆不明入金切り分け")
    st.markdown("入金データの「顧客コード」（10桁等）とマスタの整理番号を**数値化して桁数の違いや0埋めを自動吸収**し、正しく突合・紐づけを行います。")
    
    col_t4_1, col_t4_2 = st.columns(2)
    with col_t4_1:
        st.markdown("##### 1. 入金データ（CSV）")
        t4_nyukin_file = st.file_uploader("入金データファイルをアップロード", type=["csv"], key="tab4_nyukin_up")
    with col_t4_2:
        st.markdown("##### 2. マスタデータ（CSV / Excel）")
        t4_master_file = st.file_uploader("マスタデータファイルをアップロード", type=["csv", "xlsx"], key="tab4_master_up")

    if t4_nyukin_file and t4_master_file:
        try:
            try:
                df_t4_n = pd.read_csv(t4_nyukin_file, encoding="cp932", dtype=str)
            except UnicodeDecodeError:
                t4_nyukin_file.seek(0)
                df_t4_n = pd.read_csv(t4_nyukin_file, encoding="utf-8", dtype=str)

            if t4_master_file.name.endswith(".xlsx"):
                df_t4_m = pd.read_excel(t4_master_file, dtype=str)
            else:
                try:
                    df_t4_m = pd.read_csv(t4_master_file, encoding="cp932", dtype=str)
                except UnicodeDecodeError:
                    t4_master_file.seek(0)
                    df_t4_m = pd.read_csv(t4_master_file, encoding="utf-8", dtype=str)

            df_t4_n = df_t4_n.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
            df_t4_m = df_t4_m.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

            if "顧客コード" not in df_t4_n.columns:
                st.error("❌ 入金データ側に「顧客コード」列が見つかりません。")
            elif "整理番号" not in df_t4_m.columns or "入金番号" not in df_t4_m.columns:
                st.error("❌ マスタデータ側に「整理番号」または「入金番号」列が見つかりません。")
            else:
                st.success("✅ ファイルの読み込み・構造チェックが完了しました。")
                
                if st.button("⚡ 自動紐づけ＆日時変換・切り分けを実行", key="tab4_auto_exec"):
                    df_t4_n = df_t4_n.copy()
                    
                    df_t4_n['tmp_key'] = (
                        pd.to_numeric(df_t4_n['顧客コード'], errors='coerce')
                        .fillna(0)
                        .astype(int)
                        .astype(str)
                    )

                    df_t4_m['tmp_key'] = (
                        pd.to_numeric(df_t4_m['整理番号'], errors='coerce')
                        .fillna(0)
                        .astype(int)
                        .astype(str)
                    )

                    master_dict = dict(zip(df_t4_m['tmp_key'], df_t4_m["入金番号"]))
                    
                    is_matched = df_t4_n['tmp_key'].isin(master_dict.keys())

                    df_matched = df_t4_n[is_matched].copy()
                    matched_col = df_matched['tmp_key'].map(master_dict)
                    if "入金番号" in df_matched.columns:
                        df_matched["入金番号"] = matched_col.fillna(df_matched["入金番号"])
                    else:
                        df_matched["入金番号"] = matched_col.fillna("")
                    df_matched = df_matched.drop(columns=['tmp_key'])

                    df_unknown = df_t4_n[~is_matched].copy()
                    df_unknown["入金番号"] = ""
                    df_unknown = df_unknown.drop(columns=['tmp_key'])

                    for target_df in [df_matched, df_unknown]:
                        if "処理日時" in target_df.columns:
                            parsed_dates = pd.to_datetime(target_df["処理日時"], errors='coerce')
                            target_df["処理日時"] = parsed_dates.dt.strftime("%Y/%m/%d").fillna(target_df["処理日時"])
                        
                        for col_name in target_df.columns:
                            if "収納日付" in col_name or "収納日" in col_name or col_name == target_df.columns[8] if len(target_df.columns) > 8 else False:
                                parsed_shuno = pd.to_datetime(target_df[col_name], errors='coerce')
                                target_df[col_name] = parsed_shuno.dt.strftime("%Y/%m/%d").fillna(target_df[col_name])

                    yymmdd_str = datetime.now().strftime("%y%m%d")

                    st.markdown("---")
                    col_res1, col_res2 = st.columns(2)

                    with col_res1:
                        st.success(f"✅ 正常紐づけデータ: {len(df_matched)} 件")
                        st.dataframe(df_matched.head(5))
                        
                        buf_matched = io.BytesIO()
                        df_matched.to_csv(buf_matched, index=False, encoding="cp932", lineterminator="\r\n")
                        st.download_button(
                            label="📥 正常紐づけデータを保存 (CSV)",
                            data=buf_matched.getvalue(),
                            file_name=f"【入金完了データ{yymmdd_str}】.csv",
                            mime="text/csv",
                            key="tab4_success_dl"
                        )

                    with col_res2:
                        if len(df_unknown) > 0:
                            st.warning(f"⚠️ 不明入金（マスタなし）: {len(df_unknown)} 件")
                        else:
                            st.info(f"ℹ️ 不明入金（マスタなし）: 0 件（すべて正常に紐づきました！）")
                        
                        st.dataframe(df_unknown.head(5))
                        
                        buf_unknown = io.BytesIO()
                        df_unknown.to_csv(buf_unknown, index=False, encoding="cp932", lineterminator="\r\n")
                        st.download_button(
                            label="📥 不明入金データを保存 (入金番号ブランク・CSV)",
                            data=buf_unknown.getvalue(),
                            file_name=f"【不明入金データ{yymmdd_str}】.csv",
                            mime="text/csv",
                            key="tab4_unknown_dl"
                        )
        except Exception as e:
            st.error(f"❌ 処理中にエラーが発生しました: {e}")