import streamlit as st
import pandas as pd
import google.auth
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import re
import plotly.express as px
import json

# Streamlit UI 설정하기기
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
    text-align: left;
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

def preprocess_and_visualize(data, selected_grade):
    # 열 이름 지정
    data.columns = ['Timestamp', '성별', '학년', '학반', '번호'] + \
                   [f'공감소통역량{i}' for i in range(1, 7)] + \
                   [f'창의융합적사고역량{i}' for i in range(1, 7)] + \
                   [f'자기관리역량{i}' for i in range(1, 7)] + \
                   [f'공동체역량{i}' for i in range(1, 7)]
    
    # 필요한 데이터만 선택
    competency_units = {
        '공감소통역량': [f'공감소통역량{i}' for i in range(1, 7)],
        '창의융합적사고역량': [f'창의융합적사고역량{i}' for i in range(1, 7)],
        '자기관리역량': [f'자기관리역량{i}' for i in range(1, 7)],
        '공동체역량': [f'공동체역량{i}' for i in range(1, 7)]
    }

    # 텍스트에 점수 부여
    text_to_number_mapping = {
        '매우 그렇다.': 5,
        '그렇다.': 4,
        '보통이다.': 3,
        '그렇지 않다.': 2,
        '전혀 그렇지 않다.': 1
    }
    
    competency_columns = list(sum(competency_units.values(), []))
    data[competency_columns] = data[competency_columns].replace(text_to_number_mapping)
    
    # 데이터 타입을 실수형으로 변환
    data[competency_columns] = data[competency_columns].astype(float)
    data[competency_columns] = data[competency_columns].fillna(0)  # NaN을 0으로 대체

    # Filter data based on the selected grade
    if '학년' in data.columns:
        filtered_data = data if selected_grade == '전체' else data[data['학년'] == selected_grade]

        # Calculate averages by competency
        overall_avg = {competency: filtered_data[columns].mean().mean() for competency, columns in competency_units.items()}
        overall_avg_df = pd.DataFrame(list(overall_avg.items()), columns=['역량', '본교_평균'])
        overall_avg_df['본교_평균'] = overall_avg_df['본교_평균'].round(1)
        overall_avg_df['본교_10점 환산'] = (overall_avg_df['본교_평균'] * 2).round(1)
        overall_avg_df['본교_20점 환산'] = (overall_avg_df['본교_평균'] * 4).round(1)

        # 대구 평균 역량 점수 값 입력
        has_elementary = filtered_data['학년'].str.contains('초4|초5|초6').any()
        has_elementary_4 = (filtered_data['학년'].str.contains('초4') & ~filtered_data['학년'].str.contains('초5|초6')).any()
        has_elementary_5 = (filtered_data['학년'].str.contains('초5') & ~filtered_data['학년'].str.contains('초4|초6')).any()
        has_elementary_6 = (filtered_data['학년'].str.contains('초6') & ~filtered_data['학년'].str.contains('초4|초5')).any()           
        has_middle = filtered_data['학년'].str.contains('중1|중2|중3').any()
        has_middle_1 = (filtered_data['학년'].str.contains('중1') & ~filtered_data['학년'].str.contains('중2|중3')).any()
        has_middle_2 = (filtered_data['학년'].str.contains('중2') & ~filtered_data['학년'].str.contains('중1|중3')).any() 
        has_middle_3 = (filtered_data['학년'].str.contains('중3') & ~filtered_data['학년'].str.contains('중1|중2')).any() 
        
        if has_elementary and has_middle:
            # 초등과 중등이 섞여서 들어온 경우; 대구 평균 0 처리
            daegu_avg_data = {
                '공감소통역량': 0,
                '창의융합적사고역량': 0,
                '자기관리리역량': 0,
                '공동체역량': 0
            }
        elif has_middle:
            # 중학교 대구 평균
            daegu_avg_data = {
                '공감소통역량': 4.45,
                '창의융합적사고역량': 4.21,
                '자기관리역량': 4.28,
                '공동체역량': 4.35
            }
        elif has_middle and has_middle_1:
            # 중학교1 대구 평균
            daegu_avg_data = {
                '공감소통역량': 4.42,
                '창의융합적사고역량': 4.17,
                '자기관리역량': 4.25,
                '공동체역량': 4.34
            }
        elif has_middle and has_middle_2:
            # 중학교2 대구 평균
            daegu_avg_data = {
                '공감소통역량': 4.46,
                '창의융합적사고역량': 4.21,
                '자기관리역량': 4.27,
                '공동체역량': 4.33
            }
        elif has_middle and has_middle_3:
            # 중학교3 대구 평균
            daegu_avg_data = {
                '공감소통역량': 4.49,
                '창의융합적사고역량': 4.26,
                '자기관리역량': 4.31,
                '공동체역량': 4.37
            }
        elif has_elementary:
            # 초등학교 대구 평균
            daegu_avg_data = {
                '공감소통역량': 4.41,
                '창의융합적사고역량': 4.19,
                '자기관리역량': 4.28,
                '공동체역량': 4.38
            }
        elif has_elementary and has_elementary_4:
            # 초등학교4 대구 평균
            daegu_avg_data = {
                '공감소통역량': 4.38,
                '창의융합적사고역량': 4.19,
                '자기관리역량': 4.31,
                '공동체역량': 4.40
            }
        elif has_elementary and has_elementary_5:
            # 초등학교5 대구 평균
            daegu_avg_data = {
                '공감소통역량': 4.42,
                '창의융합적사고역량': 4.18,
                '자기관리역량': 4.26,
                '공동체역량': 4.38
            }
        elif has_elementary and has_elementary_6:
            # 초등학교6 대구 평균
            daegu_avg_data = {
                '공감소통역량': 4.43,
                '창의융합적사고역량': 4.21,
                '자기관리역량': 4.27,
                '공동체역량': 4.37
            }
        else:
            daegu_avg_data = {
                '공감소통역량': 0,
                '창의융합적사고역량': 0,
                '자기관리역량': 0,
                '공동체역량': 0
            }

        # 대구 평균 열 추가
        daegu_avg_df = pd.DataFrame(list(daegu_avg_data.items()), columns=['역량', '대구_평균'])
        overall_avg_df = pd.merge(overall_avg_df, daegu_avg_df, on='역량', how='left')
        
        # 합계 행 추가
        total_row = overall_avg_df[['본교_평균', '본교_10점 환산', '본교_20점 환산', '대구_평균']].mean().round(1)
        total_row['역량'] = '전체'
        overall_avg_df = pd.concat([overall_avg_df, pd.DataFrame([total_row])], ignore_index=True)

        # Display the DataFrame
        st.markdown(f"**<{selected_grade} 학년 학생 미래역량 평균 점수>**", unsafe_allow_html=True)
        st.markdown(overall_avg_df.to_html(index=False).replace('<th>', '<th style="text-align: center;">'), unsafe_allow_html=True)
        st.write(print(filtered_data.head()))
        st.write(print(filtered_data['학년'].unique()))

        # Melt the DataFrame for plotting
        overall_avg_df2=overall_avg_df[['역량', '본교_평균', '대구_평균']].iloc[:-1]
        overall_avg_melted = overall_avg_df2.melt(id_vars='역량', var_name='변수', value_name='평균값_new')

        # Plot radar chart
        fig = px.line_polar(overall_avg_melted, r='평균값_new', theta='역량',color= '변수', line_close=True,
                           color_discrete_map={'본교_평균': 'blue', '대구_평균': 'red'})
        st.write("")
        st.markdown(f"**<{selected_grade} 학년 학생 미래역량 프로파일>**", unsafe_allow_html=True)
        st.plotly_chart(fig)
    else:
        st.error("데이터에 '학년' 열이 존재하지 않습니다.")

#머릿말
st.markdown('<div class="header"></div>', unsafe_allow_html=True)
#제목
st.markdown('<div class="responsive-title">대구미래학교 학생 미래역량<br>분석 프로그램</div>', unsafe_allow_html=True)
st.write("")
st.write("")
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
            st.session_state['data'] = data
            st.session_state['initial_load'] = True
            st.session_state['selected_grade'] = '전체'
            # 데이터 전처리 및 시각화
            preprocess_and_visualize(data, '전체')

if 'data' in st.session_state and st.session_state.get('initial_load', False):
    # Add selectbox for grade selection
    selected_grade = st.sidebar.selectbox("학년 선택", ['전체'] + sorted(st.session_state['data']['학년'].unique()))
    st.session_state['selected_grade'] = selected_grade
  
    # Button to refresh the table and graph
    if st.sidebar.button("조회"):
        preprocess_and_visualize(st.session_state['data'], st.session_state['selected_grade'])
        
#꼬리말
st.markdown('<div class="footer">&copy; 2024 대구광역시교육청 초등교육과. All rights reserved.</div>', unsafe_allow_html=True)
