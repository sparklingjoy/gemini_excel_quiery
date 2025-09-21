import streamlit as st
import pandas as pd
import requests
import json
import io
from datetime import datetime

# ページ設定
st.set_page_config(
    page_title="Excel Markdown Gemini 分析",
    page_icon="📊",
    layout="wide"
)

def call_gemini_api(prompt, api_key):
    """Gemini APIを呼び出す関数"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    
    data = {
        "contents": [{
            "parts": [{"text": prompt}]
        }]
    }
    
    try:
        response = requests.post(url, json=data, timeout=60)
        if response.status_code == 200:
            result = response.json()
            if 'candidates' in result and len(result['candidates']) > 0:
                return result['candidates'][0]['content']['parts'][0]['text']
            else:
                return "応答の解析に問題があります。"
        else:
            return f"APIエラー: {response.status_code} - {response.text}"
    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

def excel_to_markdown(uploaded_file):
    """Excelファイルを読み込み、立て積みMarkdownに変換"""
    try:
        # Excelファイルを読み込み、全シート名を取得
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        markdown_content = ""
        combined_data = []
        
        # 各シートを処理
        for sheet_name in sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # 空のシートをスキップ
            if df.empty:
                continue
            
            # CustProgカラムを追加（シート名を設定）
            df_with_custprog = df.copy()
            df_with_custprog.insert(0, 'CustProg', sheet_name)
            
            # 立て積み用のデータに追加
            combined_data.append(df_with_custprog)
            
            # Markdownセクションの作成
            markdown_content += f"## {sheet_name}\n\n"
            markdown_content += f"**データ件数**: {len(df)} 行\n\n"
            
            # テーブルをMarkdownに変換
            if not df_with_custprog.empty:
                markdown_content += df_with_custprog.to_markdown(index=False)
                markdown_content += "\n\n"
            
            markdown_content += "---\n\n"
        
        # 全シートを立て積みしたデータフレームを作成
        if combined_data:
            combined_df = pd.concat(combined_data, ignore_index=True)
            
            # 統合テーブルのMarkdownを追加
            markdown_content += "## 全シート統合データ\n\n"
            markdown_content += f"**総データ件数**: {len(combined_df)} 行\n"
            markdown_content += f"**シート数**: {len(sheet_names)} シート\n\n"
            markdown_content += combined_df.to_markdown(index=False)
            markdown_content += "\n\n"
        
        return markdown_content, combined_df, sheet_names
        
    except Exception as e:
        st.error(f"Excelファイルの処理でエラーが発生しました: {str(e)}")
        return None, None, None

def extract_result_only(gemini_response):
    """Geminiの回答から結果部分のみを抽出"""
    
    # 結果を示すキーワードリスト
    result_keywords = [
        "結果:",
        "回答:",
        "結論:",
        "答え:",
        "要約:",
        "まとめ:",
        "Result:",
        "Answer:",
        "Conclusion:",
        "Summary:"
    ]
    
    # 行で分割
    lines = gemini_response.split('\n')
    result_lines = []
    capturing = False
    
    for line in lines:
        line_stripped = line.strip()
        
        # 結果キーワードを見つけた場合
        if any(keyword in line_stripped for keyword in result_keywords):
            capturing = True
            # キーワードの後の部分を取得
            for keyword in result_keywords:
                if keyword in line_stripped:
                    after_keyword = line_stripped.split(keyword, 1)
                    if len(after_keyword) > 1 and after_keyword[1].strip():
                        result_lines.append(after_keyword[1].strip())
                    break
            continue
        
        # キャプチャ中の場合、空行や新しいセクションまで続ける
        if capturing:
            if line_stripped == "" or line_stripped.startswith("#"):
                break
            result_lines.append(line_stripped)
    
    # 結果が見つからない場合は、回答の最後の段落を使用
    if not result_lines:
        paragraphs = [p.strip() for p in gemini_response.split('\n\n') if p.strip()]
        if paragraphs:
            result_lines = [paragraphs[-1]]
    
    # 結果がまだ空の場合は、全体を返す（短縮版）
    if not result_lines:
        sentences = gemini_response.split('。')
        result_lines = [sentences[-2] + '。' if len(sentences) > 1 else gemini_response]
    
    return '\n'.join(result_lines)

def main():
    st.title("📊 Excel → Markdown → Gemini 分析アプリ")
    st.markdown("---")
    
    # サイドバーでAPI設定
    with st.sidebar:
        st.header("⚙️ 設定")
        api_key = st.text_input(
            "Gemini API Key", 
            type="password", 
            placeholder="AIza...",
            help="Google AI StudioからAPIキーを取得してください"
        )
        
        st.markdown("---")
        st.markdown("### 📋 処理フロー")
        st.markdown("""
        1. Excelファイルをアップロード
        2. 各シートに**CustProg**列を追加
        3. 全シートを立て積み
        4. Markdownファイルに変換
        5. 自然言語指示と共にGeminiに送信
        6. **結果のみ**を表示
        """)
    
    # メインコンテンツ
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📂 Excelファイル処理")
        uploaded_file = st.file_uploader(
            "Excelファイルを選択してください",
            type=['xlsx', 'xls'],
            help="複数のシートを含むExcelファイルをアップロードします"
        )
        
        if uploaded_file is not None:
            st.success(f"✅ ファイル '{uploaded_file.name}' がアップロードされました")
            
            with st.spinner("Excelファイルを処理中..."):
                # ExcelをMarkdownに変換
                markdown_content, combined_df, sheet_names = excel_to_markdown(uploaded_file)
                
                if markdown_content:
                    st.info(f"📄 処理完了: {len(sheet_names)} シート, {len(combined_df)} 行")
                    
                    # Markdownプレビュー
                    with st.expander("📝 生成されたMarkdownプレビュー"):
                        st.code(markdown_content[:2000] + "..." if len(markdown_content) > 2000 else markdown_content, language="markdown")
                    
                    # 統合データプレビュー
                    with st.expander("📊 統合データプレビュー"):
                        st.dataframe(combined_df.head(100), use_container_width=True)
                        if len(combined_df) > 100:
                            st.info(f"最初の100行を表示（全{len(combined_df)}行）")
                    
                    # セッション状態に保存
                    st.session_state.markdown_content = markdown_content
                    st.session_state.combined_df = combined_df
                    st.session_state.sheet_names = sheet_names
    
    with col2:
        st.header("💬 分析指示")
        
        if api_key and 'markdown_content' in st.session_state:
            user_instruction = st.text_area(
                "分析指示を入力してください",
                placeholder="例：各CustProgの売上合計を計算し、最も売上の高いプログラムを教えてください",
                height=120,
                help="Markdownデータに対する分析を自然言語で指示してください"
            )
            
            if st.button("🚀 Gemini で分析実行", type="primary"):
                if user_instruction:
                    with st.spinner("Geminiが分析中..."):
                        # プロンプトを構築
                        full_prompt = f"""
以下のExcelデータ（Markdown形式）を分析してください：

{st.session_state.markdown_content}

ユーザーの指示：{user_instruction}

**重要**: 必ず最後に「結果:」で始まる明確な結論を提示してください。分析過程の説明は簡潔にし、結果を重視してください。
"""
                        
                        # Gemini APIに送信
                        gemini_response = call_gemini_api(full_prompt, api_key)
                        
                        if gemini_response and not gemini_response.startswith("エラー"):
                            # 結果のみを抽出
                            result_only = extract_result_only(gemini_response)
                            
                            st.subheader("🎯 分析結果")
                            st.success(result_only)
                            
                            # 完全な回答も表示（オプション）
                            with st.expander("📝 完全な回答を表示"):
                                st.write(gemini_response)
                            
                            # セッション状態に保存
                            st.session_state.last_result = result_only
                            st.session_state.full_response = gemini_response
                        else:
                            st.error(f"❌ {gemini_response}")
                else:
                    st.error("分析指示を入力してください")
        else:
            if not api_key:
                st.warning("⚠️ サイドバーでGemini APIキーを設定してください")
            elif 'markdown_content' not in st.session_state:
                st.warning("⚠️ まずExcelファイルをアップロードしてください")
    
    # 結果表示セクション
    if 'last_result' in st.session_state:
        st.markdown("---")
        st.header("📋 最新の分析結果")
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.info(st.session_state.last_result)
        
        with col2:
            # 結果をテキストファイルとしてダウンロード
            result_text = f"""分析結果
===============

{st.session_state.last_result}

分析日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
            
            st.download_button(
                label="📄 結果をダウンロード",
                data=result_text,
                file_name=f"analysis_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )
    
    # ダウンロードセクション
    if 'markdown_content' in st.session_state:
        st.markdown("---")
        st.header("💾 ダウンロード")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Markdownファイルダウンロード
            st.download_button(
                label="📝 Markdown ファイル",
                data=st.session_state.markdown_content,
                file_name=f"excel_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                mime="text/markdown"
            )
        
        with col2:
            # 統合データCSVダウンロード
            csv_buffer = io.StringIO()
            st.session_state.combined_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
            
            st.download_button(
                label="📊 統合データ (CSV)",
                data=csv_buffer.getvalue(),
                file_name=f"combined_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        with col3:
            # 統合データExcelダウンロード
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                st.session_state.combined_df.to_excel(writer, sheet_name='Combined_Data', index=False)
            
            st.download_button(
                label="📈 統合データ (Excel)",
                data=excel_buffer.getvalue(),
                file_name=f"combined_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()