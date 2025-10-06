# 2025/10/01 다완테크 판넬제작 자동작도 프로그램 신규 작성
import math
import ezdxf
from ezdxf.filemanagement import readfile, new
from ezdxf.enums import TextEntityAlignment
import openpyxl
import os
import glob
import time
import sys
import io
from datetime import datetime
import json
import re
import logging
import warnings
import tkinter as tk
from tkinter import font
import requests
from gooey import Gooey, GooeyParser

# 전역 변수 초기화
if True:
    global_data = {}
    BasicXscale, BasicYscale, TargetXscale, TargetYscale, frame_scale = 0, 0, 0, 0, 0
    frameXpos = 0
    frameYpos = 0    
    thickness = 0
    selected_dimstyle = ''
    over1000dim_style = ''
    br = 0  # bending rate 신호
    saved_DimXpos = 0
    saved_DimYpos = 0
    saved_Xpos = 0
    saved_Ypos = 0
    saved_direction = "up"
    saved_text_height = 0.38
    saved_text_gap = 0.07
    saved_dim_style = ''
    pageCount = 0
    SU = 0
    exit_program = False
    program_message = ''
    text_style_name = ''
    selected_dimstyle = ''
    over1000dim_style = ''

    # 전역 변수 선언 및 초기화
    for i in range(1, 31):
        globals()[f'x{i}'] = 0
        globals()[f'y{i}'] = 0        

    # 전역 변수 선언 및 초기화
    for i in range(1, 12):
        globals()[f'P{i}_platewidth'] = 0
        globals()[f'P{i}_plateheight'] = 0
        globals()[f'P{i}_hole'] = []

# 기본 설정
if True:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')    

# 파일 경로 설정
if True:
    script_directory = os.path.dirname(os.path.abspath(__file__))
    dxf_saved_file = os.path.join(script_directory, 'dimstyle')
    
    # 엑셀 파일 경로 설정
    excel_saved_file = 'c:/python/osel/excel파일'
    xlsm_files = glob.glob(os.path.join(excel_saved_file, '*.xlsx'))
    
    # 애플리케이션 경로 설정
    application_path = script_directory
    license_file_path = os.path.join(application_path, 'data', 'hdsettings.json')

# 기본 설정
if True:
    try:
        doc = new()
        msp = doc.modelspace()
    except Exception as e:
        print(f"DXF 초기화 오류: {e}")
        doc = None
        msp = None

    try:
        if readfile is not None and os.path.exists(os.path.join(dxf_saved_file, 'style.dxf')):
            doc = readfile(os.path.join(dxf_saved_file, 'style.dxf'))
            msp = doc.modelspace()
            selected_dimstyle = 'over1000'
            over1000dim_style = 'over1000'
        else:
            print("style.dxf 파일을 찾을 수 없습니다.")
    except Exception as e:
        print(f"style.dxf 로드 오류: {e}")

    # TEXTSTYLE 정의 (한글 처리를 위한)
    text_style_name = 'H'  # 원하는 텍스트 스타일 이름
    if 'doc' in locals() and doc is not None:
        if text_style_name not in doc.styles:
            text_style = doc.styles.new(
                name=text_style_name,
                dxfattribs={
                    'font': 'Arial.ttf',  # TrueType 글꼴 파일명            
                }
            )
        else:
            text_style = doc.styles.get(text_style_name)

    # dimstyle 매핑 설정
    dimstyle_map = {
        'dim1': 'mydim1',
        'dim2': 'mydim2',
        'dim3': 'mydim3',
        'dim4': 'mydim4'
    }
    over1000dim_style_map = { 
        'dim1': 'over1000dim1', 
        'dim2': 'over1000dim2', 
        'dim3': 'over1000dim3', 
        'dim4': 'over1000dim4'
    }        
    # selected_dimstyle과 over1000dim_style 설정
    selected_dimstyle = dimstyle_map.get('dim1', 'mydim1')  # 기본값은 'mydim1'
    over1000dim_style = over1000dim_style_map.get('dim1', 'over1000dim1')  # 기본값은 'over1000dim1'

# 제작산출결과 데이터 읽기 함수
def read_manufacturing_results(sheet, start_row=3):
    """제작산출결과 시트에서 데이터를 읽어오는 함수"""
    # 제작산출결과 시트의 열과 변수 매핑
    column_mapping = {
        "A": "number",              # 번호
        "B": "unique_id",           # 고유번호
        "C": "site_name",           # 현장명
        "D": "measurement_date",    # 측정일자
        "E": "measurer",           # 측정자
        "F": "car_width",          # 카 내부 W
        "G": "car_depth",          # 카 내부 D
        "H": "car_height",         # 카 내부 H
        "I": "interior_material",  # 의장재질
        "J": "material_thickness", # 재질 두께
        "K": "panel_number",       # 패널 번호
        "L": "manufacturing_count", # 제작 대수
        "M": "panel_type",         # 패널 타입
        "N": "manufacturing_width", # 제작폭
        "O": "manufacturing_height", # 제작높이
        "P": "perforation_width",  # 타공 가로
        "Q": "perforation_length", # 타공 세로
        "R": "perforation_height", # 타공 높이(밑기준)
        "S": "distance_from_entrance" # 입구방향에서 떨어
    }

    # 결과 리스트 초기화
    manufacturing_data = []

    # 행 반복 (A열 기준으로 비어있을 때까지 반복)
    row = start_row
    while True:
        # A열 데이터 확인 (번호가 비어있으면 종료)
        cell_value = sheet[f"A{row}"].value
        if cell_value is None or cell_value == "":  # A열이 비어있으면 종료
            break

        # 현재 행 데이터를 딕셔너리로 저장
        row_data = {}
        for col, var_name in column_mapping.items():
            cell_ref = f"{col}{row}"
            row_data[var_name] = sheet[cell_ref].value  # 해당 셀 값을 딕셔너리에 저장

        # 숫자 필드 처리 (None이면 0으로 변환)
        numeric_fields = ['number', 'unique_id', 'car_width', 'car_depth', 'car_height', 
                         'material_thickness', 'panel_number', 'manufacturing_count',
                         'manufacturing_width', 'manufacturing_height', 'perforation_width',
                         'perforation_length', 'perforation_height', 'distance_from_entrance']
        
        for field in numeric_fields:
            if row_data.get(field) is None:
                row_data[field] = 0
            else:
                # 숫자로 변환 시도
                try:
                    row_data[field] = float(row_data[field])
                except (ValueError, TypeError):
                    row_data[field] = 0

        # 결과 리스트에 추가
        manufacturing_data.append(row_data)

        # 다음 행으로 이동
        row += 1

    # 제작산출결과 데이터 출력
    print("=== 제작산출결과 데이터 출력 ===")
    for i, data in enumerate(manufacturing_data, 1):
        unique_id = data.get('unique_id', 0)
        site_name = data.get('site_name', 'N/A')
        panel_number = data.get('panel_number', 0)
        manufacturing_width = data.get('manufacturing_width', 0)
        manufacturing_height = data.get('manufacturing_height', 0)
        print(f"행 {i}: 고유번호={unique_id}, 현장명={site_name}, 패널번호={panel_number}, 제작폭={manufacturing_width}, 제작높이={manufacturing_height}")
    
    return manufacturing_data

# 유틸리티 함수들
def read_excel_value(sheet, cell_ref):
    """엑셀 셀에서 값을 읽어오는 함수"""
    try:
        return sheet[cell_ref].value
    except:
        return None

def show_custom_error(message):
    """커스텀 에러 메시지 표시"""
    print(f"오류: {message}")

def log_login():
    """로그인 로그 기록"""
    global global_data, SU
    company = global_data.get("company", "다완테크")
    workplace = global_data.get("workplace", "")
    url = f"https://8440.co.kr/autopanel/savelog.php?company=다완테크&content=판넬제작_{workplace}_{SU}"
    try:
        response = requests.get(url, timeout=5)
        print(f"로그 전송: {response.status_code}")
    except:
        print("로그 전송 실패")

def parse_arguments():
    """Gooey 인수 파싱"""
    parser = GooeyParser()
    group1 = parser.add_argument_group('카판옵션')
    group1.add_argument('--opt1', action='store_true', default=True, help='기본')
    return parser.parse_args()

# 기본 도면 그리기 함수들
def line(doc, x1, y1, x2, y2, layer=None):
    """선 그리기"""
    if layer is None:
        layer = '0'
    try:
        msp.add_line((x1, y1), (x2, y2), dxfattribs={'layer': layer})
    except Exception as e:
        print(f"선 그리기 오류: {e}")

def rectangle(doc, x1, y1, dx, dy, layer=None, offset=None):
    """사각형 그리기 (4개의 점을 연결)"""
    if offset is not None:
        # 네 개의 선분으로 직사각형 그리기 offset 추가
        line(doc, x1+offset, y1+offset, dx-offset, y1+offset, layer=layer)   
        line(doc, dx-offset, y1+offset, dx-offset, dy-offset, layer=layer)  
        line(doc, dx-offset, dy-offset, x1+offset, dy-offset, layer=layer)  
        line(doc, x1+offset, dy-offset, x1+offset, y1+offset, layer=layer)  
    else:        
        # 네 개의 선분으로 직사각형 그리기
        line(doc, x1, y1, dx, y1, layer=layer)   
        line(doc, dx, y1, dx, dy, layer=layer)   
        line(doc, dx, dy, x1, dy, layer=layer)   
        line(doc, x1, dy, x1, y1, layer=layer)

def draw_Text(doc, x, y, size, text, layer=None):
    """텍스트 그리기 (한글 지원)"""
    if layer is None:
        layer = '0'
    try:
        # dawan_jamb.py와 동일한 방식으로 텍스트 스타일 설정
        text_style_name = selected_dimstyle  # 치수선 스타일을 텍스트 스타일로 사용
        
        # 텍스트 추가 및 생성된 Text 객체 가져오기
        text_entity = msp.add_text(
            text, 
            dxfattribs={
                'height': size, 
                'layer': layer,
                'style': text_style_name  # 치수선 스타일을 텍스트 스타일로 사용
            }
        )
        text_entity.set_placement((x, y), align=TextEntityAlignment.BOTTOM_LEFT)
    except Exception as e:
        print(f"텍스트 그리기 오류: {e}")

def draw_cross_mark(doc, x, y, size, layer=None):
    """X 표시 그리기 (타공 표시용)"""
    if layer is None:
        layer = '0'
    try:
        # X 표시 그리기 (대각선 두 개)
        line(doc, x - size/2, y - size/2, x + size/2, y + size/2, layer=layer)
        line(doc, x + size/2, y - size/2, x - size/2, y + size/2, layer=layer)
    except Exception as e:
        print(f"X 표시 그리기 오류: {e}")

def draw_table(doc, msp, panels_data, start_x, start_y, manufacturing_count):
    """상세 내역 테이블 그리기"""
    try:
        # 테이블 상수 정의 (5배 크기로 확대, 열 폭 15% 추가 축소, 구분열 30% 축소, 문자 크기 20% 축소)
        TABLE_ROW_HEIGHT = 250  # 50 * 5
        TABLE_COL_WIDTHS = [560, 680, 680, 680, 1360]  # 구분(30%축소), W,H,수량(15%축소), 구분(비고)(15%축소)
        TABLE_TEXT_HEIGHT = 120  # 150 * 0.8 (20% 축소)
        TABLE_TEXT_PADDING_X = 50  # 10 * 5
        TABLE_TEXT_PADDING_Y = 50  # 10 * 5
        
        current_y = start_y
        
        # 헤더 그리기
        headers = ["구분", "W", "H", "수량", "구분"]
        current_x_col = start_x
        
        for i, header in enumerate(headers):
            rectangle(doc, current_x_col, current_y, 
                     current_x_col + TABLE_COL_WIDTHS[i], current_y - TABLE_ROW_HEIGHT, layer='0')
            draw_Text(doc, current_x_col + TABLE_TEXT_PADDING_X, 
                     current_y - TABLE_ROW_HEIGHT + TABLE_TEXT_PADDING_Y, 
                     TABLE_TEXT_HEIGHT, header, layer='0')
            current_x_col += TABLE_COL_WIDTHS[i]
        
        current_y -= TABLE_ROW_HEIGHT  # 데이터 행으로 이동
        
        # 데이터 행 그리기
        for panel_data in panels_data:
            panel_number_excel = panel_data.get('panel_number', 0)
            display_panel_number = panel_number_excel - 1  # 엑셀 번호를 도면 번호로 변환
            
            w = int(panel_data.get('manufacturing_width', 0))
            h = int(panel_data.get('manufacturing_height', 0))
            quantity = int(manufacturing_count)  # 실제 제작수량 사용
            
            # 특정 패널 비고 설정
            remarks = ""
            if display_panel_number == 5:  # 패널 #5 (엑셀 #6)
                remarks = "MIRROR"
            elif display_panel_number == 8:  # 패널 #8 (엑셀 #9)
                remarks = "도면참조"
            
            row_data = [
                f"#{int(display_panel_number)}",  # 정수로 표시
                str(w),
                str(h),
                f"{quantity} EA",
                remarks
            ]
            
            current_x_col = start_x
            for i, data_item in enumerate(row_data):
                rectangle(doc, current_x_col, current_y, 
                         current_x_col + TABLE_COL_WIDTHS[i], current_y - TABLE_ROW_HEIGHT, layer='0')
                draw_Text(doc, current_x_col + TABLE_TEXT_PADDING_X, 
                         current_y - TABLE_ROW_HEIGHT + TABLE_TEXT_PADDING_Y, 
                         TABLE_TEXT_HEIGHT, data_item, layer='0')
                current_x_col += TABLE_COL_WIDTHS[i]
            
            current_y -= TABLE_ROW_HEIGHT  # 다음 행으로 이동
            
    except Exception as e:
        print(f"테이블 그리기 오류: {e}")

def draw_dimension_line(doc, x1, y1, x2, y2, distance, text, layer=None):
    """치수선 그리기 (스타일 적용)"""
    if layer is None:
        layer = selected_dimstyle
    try:
        # 치수선 그리기 (실제 ezdxf 치수선 사용)
        if abs(x2 - x1) > abs(y2 - y1):  # 수평 치수선
            dimension = msp.add_linear_dim(
                dimstyle=selected_dimstyle,
                base=(x1, y1 + 50),  # 치수선 위치 (위쪽으로)
                p1=(x1, y1),
                p2=(x2, y2),
                dxfattribs={'layer': layer}
            )
        else:  # 수직 치수선
            dimension = msp.add_linear_dim(
                dimstyle=selected_dimstyle,
                base=(x1 - 50, y1),  # 치수선 위치 (왼쪽으로)
                angle=90,
                p1=(x1, y1),
                p2=(x2, y2),
                dxfattribs={'layer': layer}
            )
    except Exception as e:
        print(f"치수선 그리기 오류: {e}")
        # 폴백: 기본 선으로 치수선 그리기
        try:
            line(doc, x1, y1, x2, y2, layer=layer)
            # 치수 텍스트 추가
            mid_x = (x1 + x2) / 2
            mid_y = (y1 + y2) / 2
            draw_Text(doc, mid_x, mid_y - 20, 20, text, layer=layer)
        except Exception as e2:
            print(f"폴백 치수선 그리기 오류: {e2}")

def insert_frame(x, y, scale, title, description, workplaceStr, sep="NOtable"):
    """도면틀 삽입"""
    try:
        # 기본 도면틀 그리기
        frame_width = 8000 * scale
        frame_height = 6000 * scale
        
        # 외곽선
        rectangle(doc, x, y, x + frame_width, y + frame_height, layer='0')
        
        # 제목
        draw_Text(doc, x + frame_width/2 - 100, y + frame_height - 100, 50, title, layer='0')
        draw_Text(doc, x + frame_width/2 - 150, y + frame_height - 150, 30, description, layer='0')
        draw_Text(doc, x + frame_width/2 - 100, y + frame_height - 180, 25, workplaceStr, layer='0')
        
        print(f"도면틀 삽입 완료: {title}")
    except Exception as e:
        print(f"도면틀 삽입 오류: {e}")

# 제작산출결과 기반 판넬제작 도면 생성 함수
def execute_panel(): 
    """제작산출결과 기반 판넬제작 도면 생성 함수"""
    global global_data, doc, msp, pageCount
    global company, drawnby, workplace, issuedate

    # ===================== (1) 기본 정보 세팅 =====================
    company = global_data.get("company", "다완테크")
    drawnby = global_data.get("drawnby", "시스템")
    workplace = global_data.get("workplace", "현장명")
    issuedate = global_data.get("issuedate", datetime.now().strftime('%Y-%m-%d'))

    # ===================== (2) 제작산출결과 데이터 가져오기 =====================
    manufacturing_data = global_data.get("manufacturing_data", [])

    # ===================== (3) 제작산출결과 기반 판넬제작 도면 작도 =====================
    t = 1.5  # 두께는 1.5로 강제
    AbsX = 0

    # 제작산출결과 데이터가 있으면 이를 기반으로 도면 생성
    if manufacturing_data:
        # 고유번호별로 그룹화
        unique_ids = {}
        for panel_data in manufacturing_data:
            unique_id = panel_data.get('unique_id', 0)
            if unique_id not in unique_ids:
                unique_ids[unique_id] = []
            unique_ids[unique_id].append(panel_data)
        
        # 각 고유번호별로 도면 생성
        for unique_id, panels in unique_ids.items():
            if not panels:
                continue
                
            # 첫 번째 패널에서 기본 정보 가져오기
            first_panel = panels[0]
            site_name = first_panel.get('site_name', '')
            car_width = first_panel.get('car_width', 0)
            car_depth = first_panel.get('car_depth', 0)
            car_height = first_panel.get('car_height', 0)
            manufacturing_height = first_panel.get('manufacturing_height', 0)
            manufacturing_count = first_panel.get('manufacturing_count', 0)
            
            rx, startYpos = 0, 0     
            pageCount += 1   
            
            print(f"제작산출결과 기반 도면 작성 중... 고유번호: {unique_id}, 현장: {site_name}")
            
            # 기본 테두리 그리기 (생략하고 패널만 그리기)
            # border_width = car_width + 200  # 카 가로 + 여백
            # border_height = car_height + 200  # 카 높이 + 여백
            # rectangle(doc, rx, startYpos, rx + border_width, startYpos + border_height, layer='0')
            
            # 패널 2번부터 10번까지 필터링 및 정렬 (9개 패널)
            panels_2_to_10 = []
            for panel_data in panels:
                panel_number = panel_data.get('panel_number', 0)
                if 2 <= panel_number <= 10:
                    panels_2_to_10.append(panel_data)
            
            # 패널 번호순으로 정렬
            panels_2_to_10.sort(key=lambda x: x.get('panel_number', 0))
            
            # 각 패널별 도면 그리기 (패널 2번부터 10번까지)
            panel_start_x = 6100  # 기존 패널들을 X좌표 6000만큼 오른쪽으로 이동
            
            # 고유번호별로 Y좌표 오프셋 적용 (두 번째 고유번호부터 5000씩 낮게)
            unique_id_index = list(unique_ids.keys()).index(unique_id)
            y_offset = unique_id_index * 5000  # 첫 번째는 0, 두 번째는 5000, 세 번째는 10000...
            panel_start_y = 100 + y_offset
            
            current_x = panel_start_x
            
            # 현장 정보 텍스트 출력 (Y좌표를 위로 올리고 행 간격 더 넓게 확대, 모든 글자 크기 20% 축소)
            draw_Text(doc, 50, panel_start_y + manufacturing_height + 1100, 120, f"현장명: {site_name}", layer='0')  # 첫 Y 시작점 100 올림
            draw_Text(doc, 50, panel_start_y + manufacturing_height + 900, 120, f"제작 대수: {int(manufacturing_count)}대", layer='0')  # 간격 200으로 확대
            
            # 카 내부 치수 정보 출력 (소수점 제거, mm 제거, 현장명과 동일한 크기)
            draw_Text(doc, 50, panel_start_y + manufacturing_height + 700, 120, f"카 내부 치수: {int(car_width)} x {int(car_depth)} x {int(car_height)}", layer='0')
            draw_Text(doc, 50, panel_start_y + manufacturing_height + 500, 120, f"제작높이: {int(manufacturing_height)}", layer='0')
            
            for panel_data in panels_2_to_10:
                panel_number = panel_data.get('panel_number', 0)
                manufacturing_width = panel_data.get('manufacturing_width', 0)
                panel_height = manufacturing_height
                
                # 엑셀의 2번부터 10번 패널을 도면에서는 1번부터 9번으로 표시
                display_panel_number = panel_number - 1
                
                print(f"패널 {panel_number} → {display_panel_number} 그리기: 너비 {manufacturing_width}mm, 높이 {panel_height}mm")
                
                # 패널 외곽선 그리기 (4개의 점을 연결하는 사각형)
                rectangle(doc, current_x, panel_start_y, current_x + manufacturing_width, panel_start_y + panel_height, layer='레이져')
                
                # 패널 번호를 패널 높이 중심에서 400 높은 위치에 표시 (레이져 레이어, 문자 크기 2배)
                panel_center_y = panel_start_y + panel_height / 2
                draw_Text(doc, current_x + manufacturing_width/2 - 10, panel_center_y + 400, 125, f"#{int(display_panel_number)}", layer='레이져')  # 62.5 * 2 = 125
                
                # 패널 폭 치수선을 패널 위에 표시
                draw_dimension_line(doc, current_x, panel_start_y + panel_height, 
                                  current_x + manufacturing_width, panel_start_y + panel_height, 
                                  manufacturing_width, f"{int(manufacturing_width)}", layer='DIM')
                
                # 타공 정보가 있으면 표시
                perforation_width = panel_data.get('perforation_width', 0)
                perforation_length = panel_data.get('perforation_length', 0)
                perforation_height = panel_data.get('perforation_height', 0)
                distance_from_entrance = panel_data.get('distance_from_entrance', 0)
                
                if perforation_width > 0 and perforation_length > 0:
                    # 타공 위치 계산 (패널 하단에서 위쪽으로)
                    hole_x = current_x + distance_from_entrance
                    hole_y = panel_start_y + panel_height - perforation_height
                    
                    # 타공 사각형 그리기
                    rectangle(doc, hole_x, hole_y, hole_x + perforation_width, hole_y + perforation_length, layer='레이져')
                    
                    # 타공 내부에 X 표시
                    cross_x = hole_x + perforation_width / 2
                    cross_y = hole_y + perforation_length / 2
                    cross_size = min(perforation_width, perforation_length) * 0.8
                    draw_cross_mark(doc, cross_x, cross_y, cross_size, layer='레이져')
                    
                    # 타공 치수선들
                    # 타공 폭 치수선
                    draw_dimension_line(doc, hole_x, hole_y - 50, 
                                      hole_x + perforation_width, hole_y - 50, 
                                      perforation_width, f"{int(perforation_width)}", layer='DIM')
                    
                    # 타공 높이 치수선
                    draw_dimension_line(doc, hole_x - 50, hole_y, 
                                      hole_x - 50, hole_y + perforation_length, 
                                      perforation_length, f"{int(perforation_length)}", layer='DIM')
                    
                    # 타공 위치 치수선들
                    # 왼쪽 여백
                    draw_dimension_line(doc, current_x, hole_y + perforation_length/2, 
                                      hole_x, hole_y + perforation_length/2, 
                                      distance_from_entrance, f"{int(distance_from_entrance)}", layer='DIM')
                    
                    # 오른쪽 여백
                    right_margin = manufacturing_width - distance_from_entrance - perforation_width
                    draw_dimension_line(doc, hole_x + perforation_width, hole_y + perforation_length/2, 
                                      current_x + manufacturing_width, hole_y + perforation_length/2, 
                                      right_margin, f"{int(right_margin)}", layer='DIM')
                    
                    # 하단 여백
                    bottom_margin = perforation_height
                    draw_dimension_line(doc, hole_x + perforation_width/2, panel_start_y + panel_height, 
                                      hole_x + perforation_width/2, hole_y + perforation_length, 
                                      bottom_margin, f"{int(bottom_margin)}", layer='DIM')
                
                # 다음 패널 위치로 이동
                current_x += manufacturing_width + 300  # 패널 간격 300mm
            
            # 전체 높이 치수선을 첫 번째 패널 왼쪽에 그리기
            if panels_2_to_10:
                first_panel = panels_2_to_10[0]
                first_panel_width = first_panel.get('manufacturing_width', 0)
                # 첫 번째 패널의 왼쪽에 수직 치수선
                draw_dimension_line(doc, panel_start_x, panel_start_y, 
                                  panel_start_x, panel_start_y + manufacturing_height, 
                                  manufacturing_height, f"{int(manufacturing_height)}", layer='DIM')
            
            # 각 현장마다 상세 내역 테이블 그리기 (현장 정보 아래에 배치)
            # 테이블을 9개 패널로 확장 (1번부터 9번까지)
            panels_for_table = panels_2_to_10[:9]  # 처음 9개 패널 사용
            
            table_start_x = 50
            table_start_y = panel_start_y + manufacturing_height + 350  # 현장 정보와 충돌하지 않도록 조정
            draw_table(doc, msp, panels_for_table, table_start_x, table_start_y, manufacturing_count)
    else:
        print("제작산출결과 데이터가 없습니다. 도면을 생성하지 않습니다.")

# 메인 함수
def main():
    global args
    global exit_program, program_message, text_style_name
    global SU
    global global_data, doc, msp

    # 현재 날짜와 시간을 가져옵니다.
    current_datetime = datetime.now()
    global_data["formatted_date"] = current_datetime.strftime('%Y-%m-%d')
    global_data["current_time"] = current_datetime.strftime("%H%M%S")

    # .xlsx 파일이 없을 경우 오류 메시지를 출력하고 실행을 중단
    if not xlsm_files:
        error_message = ".xlsx 파일이 excel파일 폴더에 없습니다. 확인바랍니다."
        show_custom_error(error_message)
        sys.exit(1)

    for file_path in xlsm_files:
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            error_message = f"엑셀 파일을 열 수 없습니다: {str(e)}"
            show_custom_error(error_message)

        # 제작산출결과 시트만 읽기
        try:
            sheet = workbook["제작산출결과"]
        except KeyError:
            error_message = "'제작산출결과' 시트를 찾을 수 없습니다."
            show_custom_error(error_message)
            sys.exit(1)

        try:
            if readfile is not None:
                doc = readfile(os.path.join(dxf_saved_file, 'style.dxf'))
                msp = doc.modelspace()
            else:
                raise AttributeError("readfile 함수를 사용할 수 없습니다.")
        except (AttributeError, FileNotFoundError) as e:
            try:
                if new is not None:
                    doc = new()
                    if readfile is not None and os.path.exists(os.path.join(dxf_saved_file, 'style.dxf')):
                        doc = readfile(os.path.join(dxf_saved_file, 'style.dxf'))
                    msp = doc.modelspace()
                else:
                    error_message = "ezdxf new 함수를 사용할 수 없습니다."
                    show_custom_error(error_message)
                    return
            except Exception as e:
                error_message = f"DXF 파일 로드 오류: {str(e)}"
                show_custom_error(error_message)
                return
        except Exception as e:
            error_message = f"DXF 파일을 읽을 수 없습니다: {str(e)}"
            show_custom_error(error_message)
            return

        # TEXTSTYLE 정의 (한글 처리를 위한)
        text_style_name = 'H'  # 원하는 텍스트 스타일 이름
        if 'doc' in locals() and doc is not None:
            if text_style_name not in doc.styles:
                text_style = doc.styles.new(
                    name=text_style_name,
                    dxfattribs={
                        'font': 'Arial.ttf',  # TrueType 글꼴 파일명            
                    }
                )
            else:
                text_style = doc.styles.get(text_style_name)

        # dimstyle 매핑 설정
        dimstyle_map = {
            'dim1': 'mydim1',
            'dim2': 'mydim2',
            'dim3': 'mydim3',
            'dim4': 'mydim4'
        }
        over1000dim_style_map = { 
            'dim1': 'over1000dim1', 
            'dim2': 'over1000dim2', 
            'dim3': 'over1000dim3', 
            'dim4': 'over1000dim4'
        }        
        # selected_dimstyle과 over1000dim_style 설정
        selected_dimstyle = dimstyle_map.get('dim1', 'mydim1')  # 기본값은 'mydim1'
        over1000dim_style = over1000dim_style_map.get('dim1', 'over1000dim1')  # 기본값은 'over1000dim1'

        variable_names = {
            "B2": "company",
            "E2": "drawnby",
            "B3": "workplace",
            "E3": "issuedate",
            "B4": "thickness_string",
            "E4": "HPI_Type",
            "B5": "usage",
            "F5": "HPI_punchWidth",
            "G5": "HPI_punchHeight",
            "I5": "HPI_holeGap",
            "N5": "HPI_punchWidth_update",
            "O5": "HPI_punchHeight_update",
            "Q5": "HPI_holeGap_update"
        }

        for cell_ref, var_name in variable_names.items():
            value = read_excel_value(sheet, cell_ref)
            global_data[var_name] = value
            globals()[var_name] = value

        # 제작산출결과 데이터 읽기
        global_data["manufacturing_data"] = read_manufacturing_results(sheet)

        # 제작산출결과 시트에서 현장명 가져오기 (C열, 마지막 행)
        site_name = None
        row = 3  # 데이터 시작 행
        while True:
            cell_value = sheet[f"C{row}"].value
            if cell_value is None or cell_value == "":
                break
            site_name = str(cell_value)  # 마지막으로 읽은 현장명 저장
            row += 1
        
        # 현장명이 없으면 기본값 사용
        if site_name is None:
            site_name = "현장명"

        thickness_string = global_data.get("thickness_string", "1.5T")
        try:
            thickness = float(re.sub("[A-Z]", "", thickness_string))
        except:
            thickness = 1.5
        global_data["WorkTitle"] = f"업체명: {global_data.get('company', '다완테크')}, 현장명: {site_name}, thickness: {thickness}"

        execute_panel()

        # 파일명에 부적합한 문자들을 제거
        invalid_chars = '<>:"/\\|?*'
        
        # 현장명에서 부적합한 문자 제거
        cleaned_workplace = re.sub(f'[{re.escape(invalid_chars)}]', '', site_name)
        
        # 현재 날짜와 시간 (년월일_시분 형태)
        current_datetime = datetime.now()
        date_time_str = current_datetime.strftime('%Y%m%d_%H%M')
        
        # 파일명 생성: 현장명_날짜시간
        cleaned_file_name = f"{cleaned_workplace}_{date_time_str}"
        
        script_directory = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_directory, "작업완료")
        
        # 작업완료 폴더가 없으면 생성
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"작업완료 폴더 생성: {output_dir}")
        
        full_file_path = os.path.join(output_dir, f"{cleaned_file_name}.dxf")
        global_data["file_name"] = full_file_path

        exit_program = False
        program_message = '''
        프로그램 실행결과입니다.
        -------------------------------------
        {0}
        -------------------------------------
        이용해 주셔서 감사합니다.
        '''
        args = parse_arguments()

        log_login()

        doc.saveas(global_data["file_name"])
        print(f" 저장 파일명: '{global_data['file_name']}' 저장 완료!")

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
