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
    "client_x509_cert_url": "h혁
