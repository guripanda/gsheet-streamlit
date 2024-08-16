import streamlit as st
import pandas as pd
import google.auth
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import re
import plotly.express as px
import json

# Streamlit 비밀 정보 읽기
service_account_info = st.secrets["google"]["service_account_info"]
service_account_info = json.loads(service_account_info)

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
    response = sheet.values().get(spreadsheetId=spreadsheet_id, range=f'{sheet_name}!A1:Z1000').execute()
    values = response.get('values', [])
    
    if not values:
        st.error("데이터를 찾을 수 없습니다.")
        return None
    
    last_row = len(values)
    last_col = len(values[0])
    
    # 'A' + last_col을 열 문자로 변환
    last_col_letter = chr(65 + last_col - 1)
    
    return f'{sheet_name}!A1:{last_col_letter}{last_row}'

def load_data(spreadsheet_id, sheet_name='Sheet1'):
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

def create_or_get_sheet(1txggfKCwtJTpvVlWNXyUekEGdDXLo0Z-xkJTYQFi4T0, sheet_name):
    sheet = service.spreadsheets()
    
    # 시트 목록 가져오기
    spreadsheet = sheet.get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    
    # 시트 이름이 존재하는지 확인
    sheet_exists = any(s['properties']['title'] == sheet_name for s in sheets)
    
    if not sheet_exists:
        # 시트가 없으면 새로 생성
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
        sheet.batchUpdate(spreadsheetId=spreadsheet_id, body=requests).execute()
    
    return sheet_name

def save_to_school_sheet(data, 1txggfKCwtJTpvVlWNXyUekEGdDXLo0Z-xkJTYQFi4T0, sheet_name):
    sheet = service.spreadsheets()
    
    # 시트의 범위를 지정
    range_name = f'{sheet_name}!A1'
    
    # Google Sheets에 데이터를 추가
    body = {
        'values': data.values.tolist()
    }
    
    result = sheet.values().append(
        spreadsheetId=master_spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    ).execute()

    return result

def preprocess_and_visualize(data):
    # 데이터 전처리 및 평균값 계산
    if data.shape[1] == 6:  # 첫 열은 헤더
        data.columns = ['학교', '역량1', '역량2', '역량3', '역량4', '역량5']
        
        # 숫자 데이터만 선택하고 평균값 계산
        numerical_data = data.drop(columns=['학교']).astype(float)
        average_data = numerical_data.mean().reset_index()
        average_data.columns = ['역량', '평균값']
        
        # 방사형 그래프 그리기
        fig = px.line_polar(
            average_data,
            r='평균값',
            theta='역량',
            line_close=True,
            title="학생 설문 조사 평균 역량"
        )
        st.plotly_chart(fig)
    else:
        st.warning("데이터 형식이 올바르지 않습니다. 열 수가 맞지 않습니다.")

# Streamlit UI
st.title("학교별 데이터 수집 및 시각화")

# 학교 이름 입력
school_name = st.text_input("학교 이름을 입력하세요")

# 사용자가 입력할 Google Sheets URL
user_url = st.text_input("Google Sheets URL을 입력하세요")

# 데이터를 저장할 중앙 Google Sheets ID
master_spreadsheet_id = 'your_master_spreadsheet_id_here'  # 중앙에서 데이터를 저장할 시트 ID 입력

if school_name and user_url:
    sheet_id = extract_sheet_id(user_url)
    
    if sheet_id:
        data = load_data(sheet_id)
        
        if data is not None:
            st.dataframe(data)  # 불러온 데이터 표시
            
            # 학교 이름으로 시트를 생성하거나 가져옴
            sheet_name = create_or_get_sheet(master_spreadsheet_id, school_name)
            
            # 데이터를 자동으로 저장
            save_to_school_sheet(data, master_spreadsheet_id, sheet_name)
            
            # 데이터 전처리 및 시각화
            preprocess_and_visualize(data)
