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
        
    # 필요한 열들을 float 타입으로 변환, 변환되지 않는 열들은 무시
    data.iloc[:, 5:] = data.iloc[:, 5:].apply(pd.to_numeric, errors='coerce')

    # 학년 및 성별에 따른 평균값 계산
    grade_gender_avg = data.groupby(['학년', '성별']).mean(numeric_only=True).reset_index()

    # 전체 평균값 계산 (모든 학년과 성별)
    overall_avg = pd.DataFrame(data.iloc[:, 5:].mean(numeric_only=True)).T
    overall_avg['학년'] = '전체'
    overall_avg['성별'] = '전체'

    # 학년, 성별 및 전체 평균값 결합
    combined_avg = pd.concat([grade_gender_avg, overall_avg], axis=0)

    # 더 쉽게 그래프를 그릴 수 있도록 DataFrame을 melt
    melted_data = combined_avg.melt(id_vars=['학년', '성별'], 
                                    value_vars=combined_avg.columns[2:], # 평균값들만 포함
                                    var_name='역량', value_name='평균')
    
    # 깔끔한 라벨을 위해 역량 번호 제거
    melted_data['역량'] = melted_data['역량'].str.extract(r'([^\d]+)')

    # 학년 및 성별 선택 드롭다운
    grades = combined_avg['학년'].unique()
    genders = combined_avg['성별'].unique()

    # 데이터 선택
    selected_grade = st.selectbox("학년 선택", options=grades, index=0 if grades.size > 0 else None)
    selected_gender = st.selectbox("성별 선택", options=genders, index=0 if genders.size > 0 else None)

    # 필터링 쿼리 문자열 작성
    query_parts = []
    if selected_grade != '전체':
        query_parts.append(f'학년 == "{selected_grade}"')
    if selected_gender != '전체':
        query_parts.append(f'성별 == "{selected_gender}"')

    query_string = ' and '.join(query_parts) if query_parts else 'True'

    # 선택된 학년 및 성별을 기준으로 데이터 필터링
    filtered_data = melted_data.query(query_string)

    # 요약 데이터 표시
    st.markdown(f"**### {selected_grade} 학년 {selected_gender}의 역량 평균 현황**")
    st.dataframe(filtered_data.pivot(index='역량', columns='성별', values='평균'))

    # 방사형 그래프 그리기
    fig = px.line_polar(
        filtered_data,
        r='평균',
        theta='역량',
        line_close=True,
        title=f"역량 평균 - {selected_grade} 학년 {selected_gender}"
    )
    st.plotly_chart(fig)

# Streamlit UI
st.title("대구미래학교 학생역량검사 결과 분석")

# 학교 이름 입력
school_name = st.text_input("학교 이름을 입력하세요")

# 사용자가 입력할 Google Sheets URL
user_url = st.text_input("Google Sheets URL을 입력하세요")

# 데이터를 저장할 중앙 Google Sheets ID
master_spreadsheet_id = '1txggfKCwtJTpvVlWNXyUekEGdDXLo0Z-xkJTYQFi4T0'  # 중앙에서 데이터를 저장할 시트 ID 입력

if school_name and user_url:
    sheet_id = extract_sheet_id(user_url)
    
    if sheet_id:
        data = load_data(sheet_id)

        if data is not None:
            # 데이터를 자동으로 저장
            save_to_school_sheet(data, master_spreadsheet_id, school_name)
            
            # 데이터 전처리 및 시각화
            preprocess_and_visualize(data)
