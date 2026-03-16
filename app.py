import streamlit as st
import openpyxl
import anthropic
import json
import io
from google import genai
from google.genai import types

import os

CLAUDE_API_KEY = st.secrets["CLAUDE_API_KEY"]
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "Unipo1113!")

# Excelファイルパス（ローカル or アップロードファイル）
EXCEL_FILENAME = "候補者一覧_更新版_GENOVA追加.xlsx"
LOCAL_EXCEL_PATH = r"C:\Users\takeu\Downloads\候補者一覧_更新版_GENOVA追加.xlsx"
EXCEL_PATH = LOCAL_EXCEL_PATH if os.path.exists(LOCAL_EXCEL_PATH) else EXCEL_FILENAME

st.set_page_config(page_title="求人マッチングツール", page_icon="🎯", layout="wide")

# ---- パスワード認証 ----
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔐 求人マッチングツール")
    pw = st.text_input("パスワードを入力してください", type="password")
    if st.button("ログイン"):
        if pw == APP_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    st.stop()


def extract_file_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    if name.endswith(".txt"):
        return uploaded_file.read().decode("utf-8", errors="replace")
    elif name.endswith(".pdf"):
        import pdfplumber
        text = ""
        with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    elif name.endswith(".docx"):
        from docx import Document
        doc = Document(io.BytesIO(uploaded_file.read()))
        return "\n".join([p.text for p in doc.paragraphs])
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.read()))
        text = ""
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(values_only=True):
                line = "　".join([str(c) for c in row if c is not None])
                if line.strip():
                    text += line + "\n"
        return text
    else:
        return uploaded_file.read().decode("utf-8", errors="replace")


@st.cache_data
def load_company_requirements():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    companies = {}
    for sheet_name in wb.sheetnames:
        if sheet_name == "候補者一覧":
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        info = {"company_name": "", "position": "", "must": "", "want": "", "description": ""}
        for row in rows:
            if not row[0]:
                continue
            label = str(row[0])
            value = str(row[1]) if row[1] else ""
            if "求人データ" in label:
                parts = label.split("▶▶")
                info["company_name"] = parts[-1].strip() if len(parts) > 1 else label.replace("求人データ", "").strip()
            elif "参考ポジション" in label:
                info["position"] = label
            elif "必須" in label or "Must" in label:
                info["must"] = value
            elif "歓迎" in label or "Want" in label:
                info["want"] = value
            elif "業務内容" in label or "仕事内容" in label:
                info["description"] = value
        if not info["company_name"]:
            parts = sheet_name.split("_", 1)
            info["company_name"] = parts[1] if len(parts) > 1 else sheet_name
        companies[sheet_name] = info
    return companies


@st.cache_data
def load_candidates():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["候補者一覧"]
    candidates = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if row[0]:
            candidates.append({
                "name": str(row[0] or ""),
                "job_type": str(row[2] or ""),
                "skills": str(row[3] or ""),
                "age": str(row[4] or ""),
                "job_change_count": str(row[5] or 0),
                "recent_industry": str(row[6] or ""),
                "total_experience": str(row[7] or ""),
                "summary": str(row[8] or ""),
                "comment": str(row[9] or ""),
            })
    return candidates


def build_company_list(companies: dict) -> str:
    text = ""
    for sheet_name, info in companies.items():
        text += f"""
---
【企業シート名】{sheet_name}
【企業名】{info['company_name']}
【ポジション】{info['position']}
【必須要件（Must）】{info['must']}
【歓迎要件（Want）】{info['want']}
【業務内容】{info['description']}
"""
    return text


# ---- STEP 1: Claudeで上位8社を選出 ----
def step1_rank_companies(candidate_text: str, companies: dict) -> list:
    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
    company_list = build_company_list(companies)

    prompt = f"""あなたは人材紹介会社のシニアキャリアアドバイザーです。

求職者情報と求人要件を照合し、内定確度の高い順に上位8社を選んでください。
また、求職者の転職理由・今回の転職で目指しているキャリア・将来プランを文脈から推測してください。

## 求職者情報
{candidate_text}

## 各企業の求人要件
{company_list}

## 出力形式（JSONのみ返してください）
```json
{{
  "candidate_summary": {{
    "name": "氏名（不明なら「不明」）",
    "inferred_reason": "推測した転職理由（2〜3文）",
    "inferred_career": "推測した今回の転職で目指しているキャリア（2〜3文）",
    "inferred_future": "推測した将来プラン（2〜3文）"
  }},
  "top8": [
    {{
      "rank": 1,
      "sheet_name": "シート名",
      "company_name": "企業名",
      "position": "ポジション名",
      "match_score": 85,
      "match_reason": "スキル・経験面からのマッチ理由（候補者の具体的な実績・数字を引用して3〜4文）"
    }}
  ]
}}
```
"""
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = message.content[0].text
    try:
        start = raw.find("{")
        end = raw.rfind("}") + 1
        return json.loads(raw[start:end])
    except Exception:
        return {}


# ---- STEP 2: Geminiでリアルタイム企業情報を収集 ----
def step2_search_companies(company_names: list) -> str:
    client = genai.Client(api_key=GEMINI_API_KEY)
    names_str = "、".join(company_names)

    query = f"""以下の企業について、採用担当者が求職者に訴求するために役立つ情報を各社200文字以上で調査してください。

対象企業：{names_str}

各社について以下を100文字程度で簡潔にまとめてください：
- 事業内容・強み
- 成長性・社風
- 求職者へのキャリアメリット

企業名を見出しにして、各社の情報をまとめてください。"""

    # Google検索グラウンディングで試み、失敗したら検索なしにフォールバック
    for model_name in ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-2.5-flash"]:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=query,
                config=types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())]
                ),
            )
            return response.text
        except Exception:
            pass

    # 検索なし（学習データから生成）
    for model_name in ["gemini-2.0-flash-lite", "gemini-2.0-flash", "gemini-2.5-flash"]:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=query,
            )
            return response.text
        except Exception:
            continue

    raise Exception("利用可能なGeminiモデルが見つかりませんでした")


# ---- STEP 3: Geminiで訴求文を肉付け生成 ----
def step3_enrich_pitches(candidate_text: str, step1_data: dict, gemini_info: str, companies: dict) -> dict:
    client = genai.Client(api_key=GEMINI_API_KEY)

    top8 = step1_data.get("top8", [])
    summary = step1_data.get("candidate_summary", {})

    # 上位8社の求人要件だけ抽出
    top_requirements = ""
    for item in top8:
        sheet = item.get("sheet_name", "")
        if sheet in companies:
            info = companies[sheet]
            top_requirements += f"""
---
【企業名】{item['company_name']}（rank {item['rank']}）
【ポジション】{item['position']}
【必須要件】{info['must']}
【歓迎要件】{info['want']}
【業務内容】{info['description']}
"""

    prompt = f"""あなたは人材紹介会社のシニアキャリアアドバイザーです。

以下の情報をもとに、上位8社それぞれについて求職者への訴求文を生成してください。
各訴求文は**必ず150文字以上**で、Geminiの企業リサーチ情報も積極的に活用して具体的・肉厚に書いてください。

## 求職者情報
{candidate_text}

## AIが推測した転職背景
- 転職理由：{summary.get('inferred_reason', '')}
- 今回の転職で目指しているキャリア：{summary.get('inferred_career', '')}
- 将来プラン：{summary.get('inferred_future', '')}

## 上位8社の求人要件
{top_requirements}

## Geminiが収集した企業リサーチ情報（Google検索結果）
{gemini_info}

## 出力形式（JSONのみ返してください。余分なテキストは不要です）
[
  {{
    "rank": 1,
    "sheet_name": "シート名",
    "company_name": "企業名",
    "position": "ポジション名",
    "match_score": 85,
    "match_reason": "スキル・経験面からのマッチ理由（候補者の具体的な実績・数字を引用して3〜4文）",
    "pitch_reason": "【転職理由との合致】候補者の転職理由とこの求人がどう合致するか。Geminiの企業情報も交えて150文字以上で具体的に",
    "pitch_career": "【今回の転職で目指しているキャリアとの合致】候補者が目指すキャリアに対してこの求人でどんな経験・成長が得られるか。Geminiの企業情報も交えて150文字以上で具体的に",
    "pitch_future": "【将来プランとの合致】候補者の長期ビジョンに対してこの企業・ポジションがどうつながるか。Geminiの企業情報も交えて150文字以上で具体的に",
    "pitch_message": "担当者が候補者に直接伝えるトークスクリプト。候補者の名前・企業の最新情報を入れ、感情に響く言葉で150文字以上"
  }}
]
"""
    raw = ""
    for model_name in ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-2.5-flash"]:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
            )
            raw = response.text
            break
        except Exception:
            continue

    if not raw:
        st.error("Step3: Geminiによる訴求文生成に失敗しました")
        return {}

    try:
        start = raw.find("[")
        end = raw.rfind("]") + 1
        if start == -1 or end == 0:
            st.error("Step3: JSONが見つかりませんでした")
            st.code(raw[:2000])
            return {}
        results = json.loads(raw[start:end])
        score_map = {item["sheet_name"]: item for item in top8}
        for r in results:
            sn = r.get("sheet_name", "")
            if sn in score_map and "match_score" not in r:
                r["match_score"] = score_map[sn].get("match_score", 0)
        return {"candidate_summary": summary, "results": results}
    except Exception as e:
        st.error(f"Step3 解析エラー: {e}")
        st.code(raw[:3000])
        return {}


def run_analysis(candidate_text: str, companies: dict):
    progress = st.progress(0, text="Step 1/3 : Claudeがマッチング中...")

    step1_data = step1_rank_companies(candidate_text, companies)
    if not step1_data:
        st.error("Step1の分析に失敗しました")
        return

    top8 = step1_data.get("top8", [])
    company_names = [item["company_name"] for item in top8]

    progress.progress(33, text="Step 2/3 : GeminiがGoogle検索で企業情報を収集中...")

    try:
        gemini_info = step2_search_companies(company_names)
    except Exception as e:
        st.warning(f"Gemini検索でエラーが発生しました（スキップして続行）: {e}")
        gemini_info = "企業情報の取得に失敗しました。"

    progress.progress(66, text="Step 3/3 : Geminiが訴求文を生成中...")

    final_data = step3_enrich_pitches(candidate_text, step1_data, gemini_info, companies)

    progress.progress(100, text="完了！")
    progress.empty()

    if final_data:
        show_results(final_data)
    else:
        st.error("最終出力の生成に失敗しました")


def show_results(data: dict):
    summary = data.get("candidate_summary", {})
    results = data.get("results", [])
    name = summary.get("name", "候補者")

    st.success(f"✅ 分析完了！ {name}さんへの推奨求人 上位8社")

    with st.expander("📌 AIが推測した転職背景", expanded=True):
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"**転職理由**\n\n{summary.get('inferred_reason', '')}")
        col2.markdown(f"**今回の転職で目指しているキャリア**\n\n{summary.get('inferred_career', '')}")
        col3.markdown(f"**将来プラン**\n\n{summary.get('inferred_future', '')}")

    st.markdown("---")

    for r in results:
        with st.expander(f"**#{r['rank']} {r['company_name']}**　スコア: {r.get('match_score', '-')}/100", expanded=r['rank'] <= 3):
            col_left, col_right = st.columns([3, 1])
            with col_left:
                st.write(f"**ポジション：** {r.get('position', '')}")
                st.markdown(f"**マッチ理由（スキル・経験）**\n\n{r.get('match_reason', '')}")
                st.markdown("---")
                st.markdown(f"**転職理由との合致**\n\n{r.get('pitch_reason', '')}")
                st.markdown(f"**今回の転職で目指しているキャリアとの合致**\n\n{r.get('pitch_career', '')}")
                st.markdown(f"**将来プランとの合致**\n\n{r.get('pitch_future', '')}")
                st.markdown("---")
                st.info(f"💬 **担当者トークスクリプト**\n\n{r.get('pitch_message', '')}")
            with col_right:
                st.metric("内定確度スコア", f"{r.get('match_score', '-')}/100")


# ---- UI ----
st.title("🎯 求人マッチングツール")
st.caption("候補者情報を貼り付け or ファイルアップロードするだけで、内定確度の高い求人8社＋訴求方法を自動生成します")

companies = load_company_requirements()
candidates = load_candidates()

tab1, tab2 = st.tabs(["📋 既存候補者から選ぶ", "📄 テキスト貼り付け / ファイルアップロード"])

# ---- Tab1 ----
with tab1:
    candidate_names = [f"{c['name']}（{c['job_type']}）" for c in candidates]
    selected = st.selectbox("候補者を選択", ["--- 選択してください ---"] + candidate_names)

    if selected != "--- 選択してください ---":
        idx = candidate_names.index(selected)
        c = candidates[idx]
        with st.expander("候補者情報を確認", expanded=True):
            col1, col2, col3 = st.columns(3)
            col1.metric("年齢", f"{c['age']}歳")
            col2.metric("転職回数", f"{c['job_change_count']}回")
            col3.metric("経験年数", c['total_experience'])
            st.write(f"**経験職種：** {c['job_type']}")
            st.write(f"**主要スキル：** {c['skills']}")
            st.write(f"**直近業種：** {c['recent_industry']}")
            st.write(f"**職務要約：** {c['summary']}")

        if st.button("🔍 マッチング開始", type="primary", key="btn1"):
            candidate_text = f"""氏名：{c['name']}
年齢：{c['age']}歳 / 転職回数：{c['job_change_count']}回
経験職種：{c['job_type']}
主要スキル：{c['skills']}
直近業種：{c['recent_industry']} / 合計経験年数：{c['total_experience']}
職務要約・強み：{c['summary']}
評価コメント：{c['comment']}"""
            run_analysis(candidate_text, companies)

# ---- Tab2 ----
with tab2:
    st.markdown("職務経歴書・面談メモ・候補者プロフィールなど、**何でも貼り付けかアップロード**してください。両方同時もOKです。")

    col1, col2 = st.columns([3, 2])
    with col1:
        pasted_text = st.text_area(
            "テキストを貼り付け",
            placeholder="職務経歴書、面談メモ、スカウト文面など何でもOK",
            height=300,
        )
    with col2:
        uploaded_files = st.file_uploader(
            "ファイルをアップロード（複数可）",
            type=["txt", "pdf", "docx", "xlsx"],
            accept_multiple_files=True,
        )

    if st.button("🔍 マッチング開始", type="primary", key="btn2"):
        combined_text = ""
        if pasted_text.strip():
            combined_text += "【貼り付けテキスト】\n" + pasted_text.strip() + "\n\n"
        if uploaded_files:
            for f in uploaded_files:
                file_text = extract_file_text(f)
                combined_text += f"【ファイル: {f.name}】\n" + file_text.strip() + "\n\n"

        if not combined_text.strip():
            st.warning("テキストを貼り付けるかファイルをアップロードしてください")
        else:
            run_analysis(combined_text, companies)
