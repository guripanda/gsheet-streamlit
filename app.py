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
    data[list(sum(competency_units.values(), []))] = data[list(sum(competency_units.values(), []))].astype(float)
    
    # 학년별 평균값 계산 (각 역량별 평균값)
    grade_avg = {}
    for competency, columns in competency_units.items():
        grade_avg[competency] = data.groupby('학년')[columns].mean(numeric_only=True).mean(axis=1).reset_index(name=f'{competency} 평균값')
    
    # 모든 역량별 평균값을 병합
    grade_avg_df = pd.merge(grade_avg['자기관리역량'], grade_avg['창의융합적사고역량'], on='학년')
    grade_avg_df = pd.merge(grade_avg_df, grade_avg['공감소통역량'], on='학년')
    grade_avg_df = pd.merge(grade_avg_df, grade_avg['공동체역량'], on='학년')
    
    # 데이터프레임으로 변환
    # melt로 데이터 변형
    grade_avg_melted = grade_avg_df.melt(id_vars='학년', var_name='역량', value_name='평균값')

    # 요약자료 보여주기
    st.markdown("**### 학년별 학생미래역량 현황**")
    st.dataframe(grade_avg_df)

    # 방사형 그래프 그리기
    fig = px.line_polar(
        grade_avg_melted,
        r='평균값',
        theta='역량',
        color='학년',
        line_close=True,
        title="학년별 학생 설문 조사 평균 역량"
    )
    st.plotly_chart(fig)

# Streamlit UI
st.title("대구미래학교 학생역량검사 결과 분석")

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
        if sheet_id:
        data = load_data(sheet_id)

        if data is not None:
            # 데이터를 자동으로 저장
            save_to_school_sheet(data, master_spreadsheet_id, school_name)
            
            # 데이터 전처리 및 시각화
            preprocess_and_visualize(data)
st.write("제작: 대구광역시교육청 초등교육과 | © 2024")
