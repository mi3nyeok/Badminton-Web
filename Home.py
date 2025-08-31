import streamlit as st
import pandas as pd
import random
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ------------------ 페이지 설정 ------------------
st.set_page_config(page_title="🏸서천고 배드민턴부", page_icon="🏸", layout="wide")

# ------------------ Google Sheets 연결 ------------------
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# 서비스 계정 정보
gcp_info = {
    "type": "service_account",
    "project_id": "streamlit-badminton",
    "private_key_id": "05c46b2ff12bf6691ed8e05802d6a262f957937f",
    "private_key": """-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDDwjleXvs12Fgn
zOb0KhS1m6Wx9XXc4nuav3kxr+1YL1wedXUsAH5DehY1KSZF4zb7n9YzexsMQTpG
Hcccb/2pvPJEo+Iwlm20kacs77h30buThHVHHzi50QOUBo8jooaZ6QM97SyMir8L
WrO/Q+x7rFv3Zfx7dhScYVF1P7STrwwlQjZ+H4ALdK1tASBTcMlfCOGeclJuvyxb
C0BjJSk182Naa28Wx/Q+iN7Tjxokd/u6GA09ZHSjBpZwepRmy3Q46sU4BIfhsVgN
m/sg08qDIXN9WNhGgXzau8WOLjDlIqRKZAUOKqzjQTEBuYAXY3nxXjuUwh7Xg/4I
4Mx8TEo1AgMBAAECggEATlzA70yRPhg5DdGhwCraOGqylP7f7AiDl0o/nwrANXVb
6Ft3iKI36RYFrskmp4JRn793lQsaJDk2NRw1eOZBwkE/MAf0gyOsjcRiigP6MYi1
EFPNSKewYv3O82H+ybKFNDZJFHCNTUM7P7XSz2VG1KkF9Y6PV/LIRGNWottaE2Wj
efVO2/L4dYsTl5GNq9BEw0j+dculGMJMUMKpX0qwTGF14sc9FaCDu4S+/i20YeJ6
Dn2Y0RRnLew5iyQBfNly9C2Xn19c0/rEVEKsH6Rcu3CAeCmtE6FWTUy1TvK2VQk4
hn7nKNlDZkCnX+u8C1F+uwTAwQhD6+cxlyf2l+HXfwKBgQDg3MUmoRNzJmK8u04C
oEuNkxE8mnGSHuclnnjIlDc6ONpzW5gCyceyZHeo/glwQ+U1PatlEmDtUzex0p3L
Ge0J+zTn1m8AfxXlHY8RLURBq2Qy8qdUzmfDDYuHawE16RD9zctFHxzlsyzCaPjG
TGyvwguCzzPXl03+584vSa8E0wKBgQDe3b8NndhhxCBRSKCges6xV8KQEmPqHeDY
0bBhWvQL0EQXG80sntYyNY4OTLiFvfFszdwJyWkjOCQHPWcY1J6tPn2YrKiUZPLA
h4BukP5q5F0wYDR9aUYaArY2Eh5G64sDAuTRXUDXC21+S/wFfbkaa8ok5h5Wlsd3
V1oTTBGv1wKBgF8Rd9kJKOv6QqyAlE7P7qGAmD0DHxkkL53cDjzfTSo0d0dmZjJn
lSJ7D4xHOz5XtkR4OkqVJp2wHU/1N/lykKEzr/6EzpFrKZqgkKg2dzE8gTR5Gv3j
9kKlK2SpfO2KCEEGDPbAXLtQsY/QSb2s+hak97DqYYS416T55FZh1Lk5AoGBAJQD
5g2fFcTowyX4/GVY6smxpZCWBjHJIjoeGeyuwYLPEUnftwa5fzzdgLlSjwKUQjGn
D0qYr/EqKhN83moJXFFnPXcWxOe5m9fupofIWJCZlqK4YmQgdOR1hJqosd8cNLkN
RPcf3h50goXs5TaoZzV6/UhAZUnQ3i0OoM5MKTsFAoGAKbVezmwTBa2nOECoU8UL
MsIw8edthaCinW/nJLtQb769Es/X7rHrQ5nplAPgls3Y3jlNwZEzFRl7dbFpMbCp
8X+9sZdYqg8It26dkLHpvSnHCwO434qPGOPLUWgblBRYdXdrInEXuEcVVY7/i1kr
CzyjDtjgwa6aiLPMyniPHSc=
-----END PRIVATE KEY-----""",
    "client_email": "badminton@streamlit-badminton.iam.gserviceaccount.com",
    "client_id": "115918370315226724362",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/badminton@streamlit-badminton.iam.gserviceaccount.com"
}

creds = ServiceAccountCredentials.from_json_keyfile_dict(gcp_info, scope)
gc = gspread.authorize(creds)

# 시트 URL
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1G5Zaa575CwXEVjdSZYJs7BpXR0-RobwgVTZdhdN-N_A/edit?usp=sharing"
sh_cleaning = gc.open_by_url(spreadsheet_url).worksheet("뒷정리")
sh_suggestions = gc.open_by_url(spreadsheet_url).worksheet("건의사항")

# ------------------ 데이터 로드 함수 ------------------
def load_cleaning():
    try:
        data = sh_cleaning.get_all_records()
        names = [row["이름"] for row in data] if data else []
        done_check = {"목요일": [], "금요일": []}
        for day in ["목요일", "금요일"]:
            for row in data:
                done_check[day].append(bool(row.get(day, False)))
        return names, done_check
    except:
        return [], {"목요일": [], "금요일": []}

def load_suggestions():
    try:
        return sh_suggestions.get_all_records()
    except:
        return []

# ------------------ 세션 초기화 ------------------
if "cleaning" not in st.session_state or "done_check" not in st.session_state:
    names, done_check = load_cleaning()
    st.session_state.cleaning = names
    st.session_state.done_check = done_check
if "suggestions" not in st.session_state:
    st.session_state.suggestions = load_suggestions()
if "show_suggestions" not in st.session_state:
    st.session_state.show_suggestions = False
if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False

# ------------------ 사이드바: 관리자 모드 ------------------
with st.sidebar:
    st.markdown("### 🔐 관리자 모드")

    if not st.session_state.admin_mode:
        password = st.text_input("비밀번호 입력", type="password")
        if password == st.secrets["admin"]["password"]:
            st.session_state.admin_mode = True
            st.rerun()
    else:
        st.success("관리자 모드 ON")

        if st.button("🔓 관리자 모드 종료"):
            st.session_state.admin_mode = False
            st.rerun()

        if st.button("🧹 뒷정리 담당 선정"):
            try:
                df_members = pd.read_csv("명단.csv", encoding="utf-8")
                names_list = df_members["이름"].tolist()
                st.session_state.cleaning = random.sample(names_list, min(5, len(names_list)))
                st.session_state.done_check = {"목요일":[False]*5, "금요일":[False]*5}

                # 시트 초기화 후 한 번만 저장
                sh_cleaning.clear()
                sh_cleaning.append_row(["이름", "목요일", "금요일"])
                rows = [[name, False, False] for name in st.session_state.cleaning]
                sh_cleaning.append_rows(rows)
            except Exception as e:
                st.error(f"⚠️ 뒷정리 선정 오류: {e}")

        # 건의사항 보기 버튼 (토글)
        toggle_label = "📬 건의사항 닫기" if st.session_state.show_suggestions else "📬 건의사항 보기"
        if st.button(toggle_label):
            st.session_state.show_suggestions = not st.session_state.show_suggestions

# ------------------ 상단 타이틀 ------------------
st.markdown("""
<div style='display:flex; justify-content:space-between; align-items:center;'>
<h2>🏸 서천고 배드민턴부</h2>
<p style='font-size:14px; color:gray;'> 남자 기장: 최민혁 | 여자 기장: 장서윤</p>
</div>
""", unsafe_allow_html=True)

# ------------------ 기능 바로가기 버튼 ------------------
st.markdown("### ")  # 간격 조절
col1, col2 = st.columns(2)
with col1:
    st.markdown("#### 📅 목요일")
    st.page_link("pages/1_기술_코칭_매칭.py", label="🏸 기술 코칭 매칭", use_container_width=True)
with col2:
    st.markdown("#### 📅 금요일")
    st.page_link("pages/2_경기_조_편성.py", label="🎲 경기 조 편성", use_container_width=True)

# ------------------ 규칙 ------------------
st.markdown("---")
st.markdown("<h2 style='text-align:center;'>📜 배드민턴부 규칙</h2>", unsafe_allow_html=True)
rules_html = """
<div style='text-align:center; line-height:2.5;'>
1. <b>목요일</b> - 기술 코칭 / <b>금요일</b> - 경기<br>
2. <b>점심시간 전까지</b> 참가 여부 제출 필수<br>
3. <b>미제출 시</b> 페널티 부여<br>
4. <b>매주 뒷정리 5명 랜덤 선정</b> (불이행 시 페널티)<br>
5. <b>페널티 기준표</b>는 아래 참고
</div>
"""
st.markdown(rules_html, unsafe_allow_html=True)
penalty_table = pd.DataFrame({"점수": [3, 5, 7], "페널티": ["2주 뒷정리 고정", "일주일 활동 정지", "퇴출"]})
st.markdown(f"<div style='display: flex; justify-content: center;'>{penalty_table.to_html(index=False)}</div>", unsafe_allow_html=True)

# ------------------ 뒷정리 담당자 ------------------
st.markdown("---")
st.subheader("🧹 이번 주 뒷정리 담당")
if st.session_state.cleaning:
    for day in ["목요일", "금요일"]:
        st.markdown(f"### 📆 {day}")
        cols = st.columns(5)
        for i, name in enumerate(st.session_state.cleaning):
            key = f"clean_done_{day}_{i}"
            if st.session_state.admin_mode:
                with cols[i]:
                    checked = st.checkbox(name, key=key, value=st.session_state.done_check[day][i])
                    st.session_state.done_check[day][i] = checked
            else:
                color = "green" if st.session_state.done_check[day][i] else "lightgray"
                cols[i].markdown(
                    f"<div style='background-color:{color};padding:10px;border-radius:10px;text-align:center;font-size:14px'>{name}</div>",
                    unsafe_allow_html=True
                )

# ------------------ 건의사항 제출 ------------------
st.markdown("---")
st.header("📬 건의사항 제출")
is_anonymous = st.radio("익명으로 제출하시겠습니까?", ["익명", "실명"], horizontal=True, key="anon_choice")
name = st.text_input("이름을 입력하세요", key="realname_input") if is_anonymous == "실명" else ""

with st.form("suggestion_form", clear_on_submit=True):
    suggestion = st.text_area("건의사항을 입력하세요")
    submitted = st.form_submit_button("제출")
    if submitted and suggestion:
        entry = {
            "이름": name if is_anonymous == "실명" else "익명",
            "건의사항": suggestion,
            "시간": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        st.session_state.suggestions.append(entry)
        sh_suggestions.append_row([entry["이름"], entry["건의사항"], entry["시간"]])
        st.success("건의사항이 제출되었습니다!")

# ------------------ 관리자 전용 건의사항 보기 ------------------
if st.session_state.admin_mode and st.session_state.show_suggestions and st.session_state.suggestions:
    st.markdown("---")
    st.subheader("👀 제출된 건의사항")
    st.dataframe(pd.DataFrame(st.session_state.suggestions), use_container_width=True)
    최민혁
