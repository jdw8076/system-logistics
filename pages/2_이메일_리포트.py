import streamlit as st
import pandas as pd
import datetime
import win32com.client
import pythoncom
import matplotlib.pyplot as plt
import seaborn as sns
from collections import Counter, defaultdict
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
import numpy as np

st.set_page_config(page_title="📈 업무 효율 분석 리포트", layout="wide")
st.title("📧 Outlook 기반 업무 효율 분석 리포트")

TARGET_ACCOUNT = "jang.dw@asahikasei.co.kr"
today = datetime.date.today()
start_date = st.date_input("분석 시작 날짜", today - datetime.timedelta(days=7))
end_date = st.date_input("분석 종료 날짜", today)

@st.cache_data(show_spinner=True)
def load_emails(account_name, start_date, end_date):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    matched_account = None

    for i in range(outlook.Folders.Count):
        mailbox = outlook.Folders.Item(i + 1)
        if mailbox.Name.lower() == account_name.lower():
            matched_account = mailbox
            break

    if not matched_account:
        st.error(f"❌ 계정 '{account_name}'을(를) Outlook에서 찾을 수 없습니다.")
        return pd.DataFrame()

    try:
        inbox = matched_account.Folders["받은 편지함"]
        sent = matched_account.Folders["보낸 편지함"]
    except:
        inbox = matched_account.Folders["Inbox"]
        sent = matched_account.Folders["Sent Items"]

    messages = []
    for folder, folder_name in [(inbox, "수신"), (sent, "발신")]:
        items = folder.Items
        items.Sort("[SentOn]", True)
        restriction = f"[SentOn] >= '{start_date.strftime('%m/%d/%Y')}' AND [SentOn] <= '{end_date.strftime('%m/%d/%Y')}'"
        filtered_items = items.Restrict(restriction)

        for item in filtered_items:
            if item.Class != 43:
                continue
            messages.append({
                "타입": folder_name,
                "보낸시간": item.SentOn.replace(tzinfo=None),
                "제목": item.Subject,
                "보낸사람": item.SenderName,
                "받는사람": item.To,
                "첨부파일수": item.Attachments.Count
            })

    return pd.DataFrame(messages)

df = load_emails(TARGET_ACCOUNT, start_date, end_date)

if df.empty:
    st.warning("이 기간 동안 메일이 없습니다.")
    st.stop()

st.success(f"총 {len(df)}건의 이메일 불러오기 완료")

# 파생 컬럼
df['시간'] = df['보낸시간'].dt.hour
df['요일'] = df['보낸시간'].dt.day_name(locale='ko_KR')
df['제목'] = df['제목'].fillna("")

# 📊 기본 통계 요약
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("총 수신 메일", len(df[df['타입'] == '수신']))
    st.metric("총 발신 메일", len(df[df['타입'] == '발신']))
with col2:
    야근_기준 = 18
    night_df = df[df['보낸시간'].dt.hour >= 야근_기준]
    st.metric("야근 시간대 메일", len(night_df))
    st.metric("야근 시간대 비율 (%)", round(100 * len(night_df)/len(df), 1))
with col3:
    top_sender = df['보낸사람'].mode()[0] if '보낸사람' in df else "-"
    st.metric("가장 많이 온 발신자", top_sender)

# 🔁 평균 응답 시간 분석
st.subheader("⏱ 평균 회신 시간 분석")
response_times = []
reply_df = df[df['타입'] == '수신']
sent_df = df[df['타입'] == '발신']
sent_map = defaultdict(list)
for _, row in sent_df.iterrows():
    sent_map[row['제목'].strip().lower()].append(row['보낸시간'])
for _, row in reply_df.iterrows():
    subj = row['제목'].replace("Re:", "").strip().lower()
    if subj in sent_map:
        prev_times = sent_map[subj]
        closest = min(prev_times, key=lambda t: abs((row['보낸시간'] - t).total_seconds()))
        response_times.append((row['보낸시간'] - closest).total_seconds() / 60)

if response_times:
    avg_response = round(sum(response_times) / len(response_times), 1)
    st.metric("평균 응답 시간 (분)", avg_response)
else:
    st.info("응답 분석할 수 있는 제목 쌍이 충분하지 않습니다.")

import matplotlib.pyplot as plt
import seaborn as sns

# ✅ 일본어 지원되는 폰트로 변경
plt.rcParams['font.family'] = 'Meiryo'

st.subheader("🕒 曜日・時間帯別のメール分布")

plt.rcParams['font.family'] = 'Meiryo'  # 일본어 폰트 설정

# 요일/시간 피벗테이블
japan_days = ['月', '火', '水', '木', '金', '土', '日']
df['weekday_ja'] = df['보낸시간'].dt.dayofweek.map(lambda x: japan_days[x])
df['hour'] = df['보낸시간'].dt.hour
heatmap_data = pd.pivot_table(df, index='weekday_ja', columns='hour', values='타입', aggfunc='count').fillna(0)
heatmap_data = heatmap_data.reindex(japan_days)

fig, ax = plt.subplots(figsize=(9, 4))  # 👈 더 작고 얇게 조정
sns.heatmap(
    heatmap_data,
    cmap="BuGn",
    annot=True,
    fmt=".0f",
    linewidths=0.3,
    linecolor="white",
    cbar_kws={"label": "件数"},
    square=False,
    ax=ax
)

ax.set_title("📊 メール送受信の時間帯ヒートマップ", fontsize=13, pad=10)
ax.set_xlabel("時間帯 (時)", fontsize=11)
ax.set_ylabel("曜日", fontsize=11)
ax.tick_params(axis='x', labelrotation=0, labelsize=9)
ax.tick_params(axis='y', labelrotation=0, labelsize=10)

plt.tight_layout()  # ✅ 여백 자동 조정
st.pyplot(fig)


# 👥 주요 커뮤니케이션 대상
st.subheader("👥 커뮤니케이션 대상 상위")
통합 = pd.Series(df['보낸사람'].tolist() + df['받는사람'].tolist()).dropna()
상대_카운트 = Counter(통합)
top_n = pd.DataFrame(상대_카운트.items(), columns=['이름', '메일 수']).sort_values(by='메일 수', ascending=False).head(10)
st.dataframe(top_n)
st.bar_chart(top_n.set_index('이름'))

# 🔍 제목 키워드 분석
st.subheader("🔍 제목 키워드 분석 및 업무 클러스터")
words = [w.lower() for s in df['제목'] for w in s.split() if len(w) > 2 and not w.startswith("Re:")]
word_freq = Counter(words)
top_keywords = pd.DataFrame(word_freq.items(), columns=['키워드', '빈도']).sort_values(by='빈도', ascending=False).head(15)
st.dataframe(top_keywords)

# 📚 클러스터링
vectorizer = TfidfVectorizer(max_df=0.85, stop_words='english')
tfidf_matrix = vectorizer.fit_transform(df['제목'])
if tfidf_matrix.shape[0] >= 5:
    kmeans = KMeans(n_clusters=3, random_state=0).fit(tfidf_matrix)
    df['업무주제'] = kmeans.labels_
    st.subheader("🧠 주요 업무 주제 클러스터")
    st.dataframe(df[['제목', '업무주제']].head(10))
    st.bar_chart(df['업무주제'].value_counts())

# 📎 첨부파일 분석
st.subheader("📎 첨부파일 포함 여부")
df['첨부포함'] = df['첨부파일수'] > 0
attach_summary = df['첨부포함'].value_counts().rename(index={True: '있음', False: '없음'})
st.bar_chart(attach_summary)
st.dataframe(df[['제목', '첨부포함']].head(10))
