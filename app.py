import streamlit as st
import pandas as pd
import google.auth
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import re
import plotly.express as px
import json

# Streamlit 비밀 정보 읽기
service_account_info = {
    "type": st.secrets["google"]["type"],
    "project_id": st.secrets["google"]["project_id"],
    "private_key_id": st.secrets["google"]["private_key_id"],
    "private_key": st.secrets["google"]["private_key"],
    "client_email": st.secrets["google"]["client_email"],
    "client_id": st.secrets["google"]["client_id"],
    "auth_uri": st.secrets["google"]["auth_uri"],
    "token_uri": st.secrets["google"]["token_uri"],
    "auth_provider_x509_cert_url": st.secrets["google"]["auth_provider_x509_cert_url"],
    "client_x509_cert_url": st.secrets["google"]["client_x509_cert_url"],
    "universe_domain": st.secrets["google"]["universe_domain"],
}

# Google Sheets API 설정
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)

# API 서비스 생성
service = build('sheets', 'v4', credentials=creds)

def extract_sheet_id(url):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    else:
        st.error("올바른 Google Sheets URL이 아닙니다.")
        return None

def get_data_range(spreadsheet_id, sheet_name):
    sheet = service.spreadsheets()
    try:
        response = sheet.values().get(spreadsheetId=spreadsheet_id, range=f'{sheet_name}!A1:AC1000').execute()
    except googleapiclient.errors.HttpError as e:
        st.error(f"API 요청 중 오류가 발생했습니다: {e}")
        return None
    values = response.get('values', [])
    
    if not values:
        st.error("데이터를 찾을 수 없습니다.")
        return None
    else:
        last_row = len(values)
        return f'{sheet_name}!A1:AC{last_row}'

def load_data(spreadsheet_id, sheet_name='설문지 응답 시트1'):
    sheet = service.spreadsheets()
    range_name = get_data_range(spreadsheet_id, sheet_name)
    
    if range_name:
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        if not values:
            st.error("데이터를 찾을 수 없습니다.")
            return None
        else:
            return pd.DataFrame(values[1:], columns=values[0])
    else:
        st.error("유효한 범위를 찾을 수 없습니다.")
        return None

def save_to_school_sheet(data, master_spreadsheet_id, sheet_name):
    sheet = service.spreadsheets()
    
    # 시트 목록 가져오기
    spreadsheet = sheet.get(spreadsheetId=master_spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    
    # 중복된 시트명이 존재하는지 확인
    existing_sheet = next((s for s in sheets if s['properties']['title'] == sheet_name), None)

    if existing_sheet:
        # 중복된 시트가 있으면 업데이트
        sheet_id = existing_sheet['properties']['sheetId']
        
    else:
        requests = {
            'requests': [
                {
                    'addSheet': {
                        'properties': {
                            'title': sheet_name,
                        }
                    }
                }
            ]
        }
        response = sheet.batchUpdate(spreadsheetId=master_spreadsheet_id, body=requests).execute()
        sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
        
    # 시트의 범위를 지정
    range_name = f'{sheet_name}!A1'
    
    # Google Sheets에 데이터를 추가
    body = {
        'values': data.values.tolist()
    }
    
    result = sheet.values().update(
        spreadsheetId=master_spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()

    return result

def preprocess_and_visualize(data):
    # 열 이름 지정
    data.columns = ['Timestamp', '성별', '학년', '학반', '번호'] + \
                   [f'자기관리역량{i}' for i in range(1, 7)] + \
                   [f'창의융합적사고역량{i}' for i in range(1, 7)] + \
                   [f'공감소통역량{i}' for i in range(1, 7)] + \
                   [f'공동체역량{i}' for i in range(1, 7)]
    
    # 필요한 데이터만 선택
    competency_units = {
        '자기관리역량': [f'자기관리역량{i}' for i in range(1, 7)],
        '창의융합적사고역량': [f'창의융합적사고역량{i}' for i in range(1, 7)],
        '공감소통역량': [f'공감소통역량{i}' for i in range(1, 7)],
        '공동체역량': [f'공동체역량{i}' for i in range(1, 7)]
    }
    
    # 데이터 타입을 실수형으로 변환
    competency_columns = list(sum(competency_units.values(), []))
    data[competency_columns] = data[competency_columns].astype(float)
    data[competency_columns] = data[competency_columns].fillna(0)  # NaN을 0으로 대체

    # Streamlit session state for data
    if 'data' not in st.session_state:
        st.session_state.data = data

    data = st.session_state.data

    # Add dropdown menu for selecting grade
    grades = ['전체'] + sorted(data['학년'].unique())
    selected_grade = st.selectbox("학년 선택", grades)

    # Filter data based on selected grade
    if selected_grade != '전체':
        filtered_data = data[data['학년'] == selected_grade]
    else:
        filtered_data = data

    # Calculate averages by competency
    overall_avg = {}
    for competency, columns in competency_units.items():
        overall_avg[competency] = filtered_data[columns].mean().mean()

    # DataFrame for displaying averages
    overall_avg_df = pd.DataFrame(list(overall_avg.items()), columns=['역량', '평균값'])

    # Display the DataFrame
    st.markdown("**### 학생미래역량 평균 현황**")
    st.dataframe(overall_avg_df)

    # Melt the DataFrame for plotting
    overall_avg_melted = overall_avg_df.melt(id_vars='역량', var_name='변수', value_name='평균값')

    # Plot radar chart
    fig = px.line_polar(
        overall_avg_melted,
        r='평균값',
        theta='역량',
        line_close=True,
        title=f"{selected_grade if selected_grade != '전체' else '전체 학년'} 학생 설문 조사 평균 역량"
    )
    st.plotly_chart(fig)

# Streamlit UI
custom_css = """
<style>
/* Header Styles */
.header {
    background-color: #007BFF; /* Blue background */
    color: white;
    padding: 20px; /* Increased padding for more space */
    text-align: center;
    font-size: 24px;
    font-weight: bold;
    position: fixed;
    top: 0;
    width: 100%;
    z-index: 1000; /* Ensure the header is above other content */
}

/* Footer Styles */
.footer {
    background-color: #343A40; /* Dark gray background */
    color: white;
    padding: 10px;
    text-align: center;
    font-size: 14px;
    position: fixed;
    bottom: 0;
    width: 100%;
    border-top: 1px solid #555;
}

/* General title styling */
.responsive-title {
    font-size: calc(1.5vw + 1rem); /* Responsive font size with base size */
    font-weight: bold;
    color: #333; /* Adjust text color as needed */
    text-align: center; /* Center-align the title */
    margin-bottom: 20px; /* Add space below the title */
    white-space: normal; /* Allow wrapping to avoid overflow */
    overflow: hidden; /* Hide overflowed text */
    text-overflow: ellipsis; /* Show ellipsis (...) if text overflows */
    max-width: 100%; /* Ensure it fits within the container */
}

/* Smaller screens */
@media (max-width: 768px) {
    .responsive-title {
        font-size: calc(2vw + 1rem); /* Adjust font size for smaller screens */
    }
}

/* Very small screens */
@media (max-width: 480px) {
    .responsive-title {
        font-size: calc(3vw + 1rem); /* Larger font size for very small screens */
    }
}
</style>
"""

#CSS 설정가져오기
st.markdown(custom_css, unsafe_allow_html=True)
#머릿말
st.markdown('<div class="header">대구미래학교 학생역량검사 분석 서비스</div>', unsafe_allow_html=True)
#제목
st.markdown('<div class="responsive-title">대구미래학교 학생역량검사 결과</div>', unsafe_allow_html=True)
# 사이드바에 로그인 페이지 만들기
st.sidebar.header('Login')

#학교 이름 입력
school_name = st.sidebar.text_input("학교 이름을 입력하세요")
# 사용자가 입력할 Google Sheets URL
user_url = st.sidebar.text_input("Google Sheets URL을 입력하세요")

# 데이터를 저장할 중앙 Google Sheets ID
master_spreadsheet_id = st.secrets["google"]["master_sheet_id"]  # 중앙에서 데이터를 저장할 시트 ID 입력

#데이터 처리 트리거 버튼
if st.sidebar.button("분석결과"):
    sheet_id = extract_sheet_id(user_url)
    data = load_data(sheet_id)
    if data is not None:
        # 데이터를 자동으로 저장
        save_to_school_sheet(data, master_spreadsheet_id, school_name)
        # 데이터 전처리 및 시각화
        preprocess_and_visualize(data)

#꼬리말
st.markdown('<div class="footer">&copy; 2024 대구광역시교육청 초등교육과. All rights reserved.</div>', unsafe_allow_html=True)
