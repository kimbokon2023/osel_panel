# 2025/01/01 다완 방화쟘 자동작도 시작
# 2025/08/18 9가지 추가접수 (JD를 막판과 기둥을 결합했을때로 계산함(특이점) 발견)
import math
try:
    import ezdxf
    from ezdxf.filemanagement import readfile, new
except ImportError as e:
    print(f"ezdxf 모듈 임포트 오류: {e}")
    # 대체 임포트 시도
    try:
        import ezdxf
        readfile = getattr(ezdxf, 'readfile', None)
        new = getattr(ezdxf, 'new', None)
        if readfile is None or new is None:
            print("ezdxf 모듈에서 필요한 함수를 찾을 수 없습니다.")
    except Exception as e2:
        print(f"ezdxf 모듈 로드 실패: {e2}")
        ezdxf = None
        readfile = None
        new = None
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
    BasicXscale, BasicYscale,TargetXscale,TargetYscale, frame_scale = 0,0,0,0,0
    frameXpos = 0
    frameYpos = 0    
    thickness = 0
    selected_dimstyle = ''
    over1000dim_style = ''
    br = 0  # bending rate 연신율
    saved_DimXpos = 0
    saved_DimYpos = 0
    saved_Xpos = 0
    saved_Ypos = 0
    saved_direction = "up"
    saved_text_height = 0.38
    saved_text_gap = 0.05
    dimdistance = 0
    dim_horizontalbase = 0
    dim_verticalbase = 0
    distanceXpos = 0
    distanceYpos = 0
    start_time = 0
    secondord = None
    drawdate_str = None
    company = None
    workplace = None
    drawnby = None
    inspectedby = None    
    issuedate = None
    usage = None    
    person = None
    SU = 0    
    HPIsurang = 0
    Material = None
    Spec = None
    thickness_string = None        
    error_message = ''    
    pagex = 0
    pagey = 0
    rx, ry = 0, 0                    
    pageCount = 0

    # 전역 변수 초기화
    jambType = None
    floorDisplay = None
    material = None
    spec = None
    vcut = None
    OP = 0
    JE = 0
    JD = 0
    JD_sheet = 0
    HH = 0
    MH = 0
    HPI_height = 0
    U = 0
    C = 0
    A = 0
    grounddig = 0
    control_width = 0
    controltopbar = 0
    controlbox = 0
    poleAngle = 0
    surang = 0
    TR_width = 0
    VcutPlus = 0
    
    # HPI 관련 전역 변수 초기화
    HPIHoleWidth = 0
    HPIHoleHeight = 0
    HPIHeight = 0
    HPIholegap = 0

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

    # 경고 메시지 필터링
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")

    # 폴더 내의 모든 .xlsm 파일을 검색
    application_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    # excel_saved_file = os.path.join(application_path, 'panel_excel')
    # xlsm_files = glob.glob(os.path.join(excel_saved_file, '*.xlsm'))
    # 절대 경로를 지정
    excel_saved_file = 'c:/dawan/excel파일'
    xlsm_files = glob.glob(os.path.join(excel_saved_file, '*.xlsm'))
    license_file_path = os.path.join(application_path, 'data', 'hdsettings.json') # 하드디스크 고유번호 인식

    # DXF 파일 로드
    dxf_saved_file = 'c:/dawan/dimstyle'
    try:
        if readfile is not None:
            doc = readfile(os.path.join(dxf_saved_file, 'dawan_style.dxf'))    
            msp = doc.modelspace()
        else:
            raise AttributeError("readfile 함수를 사용할 수 없습니다.")
    except (AttributeError, FileNotFoundError) as e:
        # ezdxf 버전 호환성 문제 해결 또는 파일이 없는 경우
        try:
            if new is not None:
                doc = new()
                if readfile is not None and os.path.exists(os.path.join(dxf_saved_file, 'dawan_style.dxf')):
                    doc = readfile(os.path.join(dxf_saved_file, 'dawan_style.dxf'))
                msp = doc.modelspace()
            else:
                raise AttributeError("new 함수를 사용할 수 없습니다.")
        except Exception as e:
            print(f"DXF 파일 로드 오류: {e}")
            # 새 DXF 문서 생성
            if new is not None:
                doc = new()
                msp = doc.modelspace()
            else:
                print("ezdxf 모듈을 사용할 수 없습니다. 프로그램을 종료합니다.")
                sys.exit(1)
    except Exception as e:
        print(f"DXF 파일 로드 오류: {e}")
        # 새 DXF 문서 생성
        if new is not None:
            doc = new()
            msp = doc.modelspace()
        else:
            print("ezdxf 모듈을 사용할 수 없습니다. 프로그램을 종료합니다.")
            sys.exit(1)

    # TEXTSTYLE 정의
    text_style_name = 'H'  # 원하는 텍스트 스타일 이름
    if text_style_name not in doc.styles:
        text_style = doc.styles.new(
            name=text_style_name,
            dxfattribs={
                'font': 'Arial.ttf',  # TrueType 글꼴 파일명            
            }
        )
    else:
        text_style = doc.styles.get(text_style_name)

    # 첫 번째 .xlsm 파일에서 W1 셀 값 읽기
    if not xlsm_files:
        print(".xlsm 파일이 excel파일 폴더에 없습니다. 확인 바랍니다.")
        sys.exit(1)

    workbook = openpyxl.load_workbook(xlsm_files[0], data_only=True)
    sheet_name = '발주'
    sheet = workbook[sheet_name]
    dim_style_key = 'dim1'

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
    selected_dimstyle = dimstyle_map.get(dim_style_key, 'mydim1')  # 기본값은 'mydim1'
    over1000dim_style = over1000dim_style_map.get(dim_style_key, 'over1000dim1')  # 기본값은 'over1000dim1'

def read_excel_rows(sheet, start_row=8):
    # 열과 변수 매핑
    column_mapping = {
        "A": "jambType",
        "B": "floorDisplay",
        "C": "material",
        "D": "spec",
        "E": "vcut",
        "F": "OP",
        "G": "poleAngle",
        "H": "JE",
        "I": "JD",
        "J": "HH",
        "K": "MH",
        "L": "HPI_height",
        "M": "U",
        "N": "C",
        "O": "A",
        "P": "grounddig",
        "Q": "FireDoor" # 방화도어 여부 방화/일반 선택
        # "P": "control_width",
        # "Q": "controltopbar",
        # "R": "controlbox",
    }

    # 결과 리스트 초기화
    rows_data = []

    # 행 반복 (A열 기준으로 비어있을 때까지 반복)
    row = start_row
    while True:
        # A열 데이터 확인
        cell_value = sheet[f"A{row}"].value
        if cell_value is None:  # A열이 비어있으면 종료
            break

        # 현재 행 데이터를 딕셔너리로 저장
        row_data = {}
        for col, var_name in column_mapping.items():
            cell_ref = f"{col}{row}"
            row_data[var_name] = sheet[cell_ref].value  # 해당 셀 값을 딕셔너리에 저장

        # jambType에 따른 JD 값 조정
        if row_data.get('jambType') == '막판유' and row_data.get('JD') is not None:
            row_data['JD'] = row_data['JD'] - 10     
        # 방화도어/일반도어 구분에 따른 값을 정한다. 방화 25, 일반도어 50       
        if row_data.get('FireDoor') == '방화':
            row_data['FireDoor'] = 25
        else:
            row_data['FireDoor'] = 50

        # 결과 리스트에 추가
        rows_data.append(row_data)

        # 다음 행으로 이동
        row += 1

    # JD 값들을 화면에 출력
    print("=== JD 값 출력 ===")
    for i, data in enumerate(rows_data, 1):
        jd_value = data.get('JD', 'N/A')
        jamb_type = data.get('jambType', 'N/A')
        print(f"행 {i}: jambType={jamb_type}, 수정된 (구)JD={jd_value}")
    
    return rows_data

def setpos(a, b, c, d,e,f):
    global BasicXscale, BasicYscale,TargetXscale,TargetYscale, frame_scale, frameXpos, frameYpos
    BasicXscale, BasicYscale,TargetXscale,TargetYscale, frameXpos, frameYpos = a, b, c, d, e, f
    if(TargetXscale/BasicXscale > TargetYscale/BasicYscale ):
        frame_scale = TargetXscale/BasicXscale
    else:
        frame_scale = TargetYscale/BasicYscale
    pagex, pagey = frameXpos , frameYpos + 1123 * frame_scale
    rx, ry = pagex , pagey 
    return rx, ry , pagex , pagey, frame_scale, frameXpos, frameYpos
def log_login():
    # PHP 파일의 URL 서버에서는 아이피를 저장한다. 업체 아이피를 기록한다.
    workplace = global_data["WorkTitle"]
    url = f"https://8440.co.kr/autopanel/savelog.php?company=다완테크&content=신규쟘_{workplace}_{SU}"

    # HTTP 요청 보내기
    response = requests.get(url)
    
    # 요청이 성공했는지 확인
    if response.status_code == 200:
        # print("logged successfully.")
        print(response.json())
    else:
        print("Failed to log login time.")
        print(response.text)   
        exit(1)

def show_custom_error(message):
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨기기

    error_window = tk.Toplevel()
    error_window.title("자동작도 오류 알림")
    
    custom_font = font.Font(size=15)  # 폰트 크기 설정
    label = tk.Label(error_window, text=message, font=custom_font)
    label.pack(padx=20, pady=20)

    def close_program():
        sys.exit(1)  # 모든 프로그램 강제 종료

    close_button = tk.Button(error_window, text="확인", command=close_program)
    close_button.pack(pady=10)

    # 창을 화면 중앙에 위치시키기
    error_window.update_idletasks()
    width = error_window.winfo_width()
    height = error_window.winfo_height()
    x = (error_window.winfo_screenwidth() // 2) - (width // 2)
    y = (error_window.winfo_screenheight() // 2) - (height // 2)
    error_window.geometry(f"{width}x{height}+{x}+{y}")    
    error_window.mainloop()    
def is_number(var):
    # 변수가 숫자인지 확인하는 함수
    if isinstance(var, (int, float)):
        return True
    elif isinstance(var, str):
        try:
            float(var)  # 변환을 시도하여 숫자인지 확인
            return True
        except ValueError:
            return False
    return False
def save_file(company, workplace):
    # 현재 시간 가져오기
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")

    # 파일 이름에 사용할 수 없는 문자 정의
    invalid_chars = '<>:"/\\|?*'
    # 정규식을 사용하여 유효하지 않은 문자 제거
    cleaned_file_name = re.sub(f'[{re.escape(invalid_chars)}]', '', f"{company}_{workplace}_{thickness_string}_{drawnby}_{current_time}")

    # 결과 파일이 저장될 디렉토리
    output_directory = "c:/dawan/작업완료"

    # 디렉토리가 존재하지 않으면 생성
    os.makedirs(output_directory, exist_ok=True)

    # 결과 파일 이름
    file_name = f"{cleaned_file_name}.dxf"
    # 전체 파일 경로 생성
    full_file_path = os.path.join(output_directory, file_name)

    # 파일 경로 반환
    return full_file_path    
def read_excel_value(sheet, cell):
    value = sheet[cell].value
    if isinstance(value, str):
        try:
            float_value = float(value)  # 문자열을 float로 변환 시도
            if float_value.is_integer():  # 소수점이 없는 경우
                return int(float_value)
            else:  # 소수점이 있는 경우
                return float_value
        except ValueError:
            return value  # 변환할 수 없는 경우 원래 문자열 반환
    return value
def write_log(message):
    logging.info(message)    
def parse_arguments_settings():
    parser = GooeyParser()
    settings = parser.add_argument_group('설정')
    settings.add_argument('--config', action='store_true', default=True,  help='라이센스')
    settings.add_argument('--password', widget='PasswordField', gooey_options={'visible': True, 'show_label': True}, help='비밀번호')

    return parser.parse_args()
def parse_arguments():
    parser = GooeyParser()

    group1 = parser.add_argument_group('카판넬')
    # group1.add_argument('--opt1', action='store_true',  help='기본')
    # group1.add_argument('--opt2', action='store_true',  help='도장홀')
    # group1.add_argument('--opt3', action='store_true',  help='쪽쟘 상판끝 라운드(추영덕소장)')    
    group1.add_argument('--opt1', action='store_true', default=True, help='기본')

    settings = parser.add_argument_group('설정')
    settings.add_argument('--config', action='store_true', help='라이센스')
    settings.add_argument('--password', widget='PasswordField', gooey_options={'visible': True, 'show_label': True}, help='비밀번호')
    
    return parser.parse_args()
def display_message():
    message = program_message.format('\n'.join(sys.argv[1:])).split('\n')
    delay = 1.5 / len(message)

    for line in message:
        print(line)
        time.sleep(delay)
def load_env_settings():
# 환경설정 가져오기(하드공유 번호)    
    try:
        with open(license_file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            return data.get("DiskID")
    except FileNotFoundError:
        return None
def get_current_disk_id():
    return os.popen('wmic diskdrive get serialnumber').read().strip()
def validate_or_default(value):
# None이면 0을 리턴하는 함수    
    if value is None:
        return 0
    return value
def find_intersection(start1, end1, start2, end2):
    # 교차점 계산 (두 직선이 수직 및 수평으로 만나는 경우에 대해서만)
    if start1[0] == end1[0]:
        return (start1[0], start2[1])
    else:
        return (start2[0], start1[1])
def calculate_fillet_point(center, point, radius):
    # 필렛 접점 계산
    dx = point[0] - center[0]
    dy = point[1] - center[1]
    if abs(dx) > abs(dy):
        return (center[0] + radius * (1 if dx > 0 else -1), center[1])
    else:
        return (center[0], center[1] + radius * (1 if dy > 0 else -1))
def calculate_angle(center, point):
    # 각도 계산
    return math.degrees(math.atan2(point[1] - center[1], point[0] - center[0]))
def add_90_degree_fillet(doc, start1, end1, start2, end2, radius):
    msp = doc.modelspace()

    # 교차점 찾기
    intersection_point = find_intersection(start1, end1, start2, end2)

    # 필렛 접점 계산
    point1 = calculate_fillet_point(intersection_point, end1, radius)
    point2 = calculate_fillet_point(intersection_point, end2, radius)

    # 각도 계산
    start_angle = calculate_angle(intersection_point, point1)
    end_angle = calculate_angle(intersection_point, point2)

    # 필렛 원호 그리기
    msp.add_arc(
        center=intersection_point,
        radius=radius,
        start_angle=start_angle,
        end_angle=end_angle,
        dxfattribs={'layer': '레이져'},
    )    
    return msp
def calculate_midpoint(point1, point2):
    return ((point1[0] + point2[0]) / 2, (point1[1] + point2[1]) / 2)
def calculate_circle_center(midpoint, radius, point1, point2):
    dx = point2[0] - point1[0]
    dy = point2[1] - point1[1]
    dist = math.sqrt(dx**2 + dy**2)
    factor = math.sqrt(radius**2 - (dist / 2)**2) / dist
    return (midpoint[0] - factor * dy, midpoint[1] + factor * dx)
def add_arc_between_points(doc, point1, point2, radius):
    msp = doc.modelspace()

    # 중점 계산
    midpoint = calculate_midpoint(point1, point2)

    # 원의 중심 계산
    center = calculate_circle_center(midpoint, radius, point1, point2)

    # 각도 계산
    start_angle = calculate_angle(center, point1)
    end_angle = calculate_angle(center, point2)

    # 아크 그리기
    msp.add_arc(
        center=center,
        radius=radius,
        start_angle=start_angle,
        end_angle=end_angle,
        dxfattribs={'layer': '레이져'},
    )    
    return msp
def draw_arc(doc, x1, y1, x2, y2, radius, direction, layer='레이져'): # radius는 반지름을 넣는다. 지름이 아님
    msp = doc.modelspace()
    
    # 점들과 반지름을 이용하여 중점과 원의 중심 계산
    midpoint = calculate_midpoint((x1, y1), (x2, y2))
    center = calculate_circle_center(midpoint, radius, (x1, y1), (x2, y2))

    # 시작 각도와 끝 각도 계산
    start_angle = calculate_angle(center, (x1, y1))
    end_angle = calculate_angle(center, (x2, y2))

    # 방향에 따라 각도 조정
    if direction == 'up' or direction == 'down':
        if start_angle > end_angle:
            start_angle, end_angle = end_angle, start_angle
        if direction == 'down':
            start_angle += 180
            end_angle += 180
    elif direction == 'left' or direction == 'right':
        if start_angle > end_angle:
            start_angle, end_angle = end_angle, start_angle
        if direction == 'right':
            start_angle += 180
            end_angle += 180

    # 아크 그리기
    msp.add_arc(
        center=center,
        radius=radius,
        start_angle=start_angle,
        end_angle=end_angle,
        dxfattribs={'layer': layer},
    )
    return msp
def draw_crossmark(doc, x, y, layer='0'):
    """
    M4 육각 십자선 그림    
    """    
    line(doc, x,y-5,x,y+5,layer=layer)    
    line(doc, x-5,y,x+5,y,layer=layer)    
    return msp
def dim_leader(doc, start_x, start_y, end_x, end_y, text, text_height=30, direction=None, option=None):
    global saved_DimXpos, saved_DimYpos, saved_text_height, saved_text_gap, saved_direction, dimdistance, dim_horizontalbase, dim_verticalbase, distanceXpos, distanceYpos 
    
    msp = doc.modelspace()
    layer = '0'        
    # text_style_name = 'GHS'            
    text_style_name = selected_dimstyle
    # override 설정
    override_settings = {
        'dimasz': 15
    }

    # 텍스트 위치 조정 및 꺾이는 지점 설정
    if option is None:    
        text_offset_x = 20
        text_offset_y = 20
    else:
        text_offset_x = 0
        text_offset_y = 0        
    
    if direction == 'leftToright':
        mid_x = end_x - text_offset_x
        mid_y = end_y  # 텍스트 앞에서 꺾임
        text_position = (end_x, end_y)
    elif direction == 'rightToleft':
        mid_x = end_x + text_offset_x
        mid_y = end_y  # 텍스트 앞에서 꺾임
        text_position = (end_x - len(text) * 22, end_y)
    else:
        mid_x = (start_x + end_x) / 2
        mid_y = (start_y + end_y) / 2
        text_position = (end_x + text_offset_x, end_y + text_offset_y)

    # 지시선 추가
    leader = msp.add_leader(
        vertices=[(start_x, start_y), (mid_x, mid_y-text_height/2), (end_x, end_y-text_height/2)],  # 시작점, 중간점(문자 앞에서 꺾임), 끝점
        dxfattribs={
            'dimstyle': text_style_name,
            'layer': layer,
            'color': 3  # 녹색 (AutoCAD 색상 인덱스에서 3번은 녹색)
        },
        override=override_settings
    )

    if option is None:
        # 텍스트 추가 (선택적)
        if text:
            msp.add_mtext(text, dxfattribs={
                'insert': text_position,
                'layer': layer,
                'char_height': text_height,
                'style': text_style_name,
                'attachment_point': 1,  # 텍스트 정렬 방식 설정
                'color': 2  # 노란색 (AutoCAD 색상 인덱스에서 2번은 노란색)
            })

    return leader
def dim_linear(doc, x1, y1, x2, y2, textstr, dis, direction="up", layer='0', text_height=0.30,  text_gap=0.07):
    global saved_DimXpos, saved_DimYpos, saved_text_height, saved_text_gap, saved_direction, dimdistance, dim_horizontalbase, dim_verticalbase, distanceXpos, distanceYpos
    msp = doc.modelspace()
    layer = '0'    
    dim_style = over1000dim_style

    # 치수값 계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # 소수점이 있는지 확인
    if dimension_value % 1 != 0:
        dimdec = 1  # 소수점이 있는 경우 소수점 첫째 자리까지 표시
    else:
        dimdec = 1  # 소수점이 없는 경우 소수점 표시 없음

    # override 설정
    override_settings = {
        'dimtxt': text_height,
        'dimgap': text_gap,
        'dimscl': 1,
        'dimlfac': 1,
        # 'dimclrt': 7, 색상강제로 흰색 지정
        'dimdsep': 46,
        'dimdec': dimdec,
        # 텍스트 180도 회전
        #'dimtrot': 180        
        'dimtih': 1  # 텍스트를 항상 수평으로 표시
    }

    # 방향에 따른 치수선 추가
    if direction == "up":
        dimension = msp.add_linear_dim(
            base=(x1, y1 + dis),
            dimstyle=dim_style,
            p1=(x1, y1),
            p2=(x2, y2),
            dxfattribs={'layer': layer},
            override=override_settings
        )
    elif direction == "down":
        dimension = msp.add_linear_dim(
            dimstyle=dim_style,
            base=(x1, y1 - dis),
            p1=(x1, y1),
            p2=(x2, y2),
            dxfattribs={'layer': layer},
            override=override_settings
        )
    elif direction == "aligned":
        dimension = msp.add_aligned_dim(
            dimstyle=dim_style,
            p1=(x1, y1),
            p2=(x2, y2),
            distance=dis,            
            dxfattribs={'layer': layer},
            override=override_settings
        )
    else:
        raise ValueError("Invalid direction. Use 'up', 'down', or 'aligned'.")

    dimension.render()
    return dimension

def line(doc, x1, y1, x2, y2, layer=None):
    global saved_Xpos, saved_Ypos  # 전역 변수로 사용할 것임을 명시
    
    # 선 추가
    start_point = (x1, y1)
    end_point = (x2, y2)
    if layer:
        # 절곡선 22 layer는 ltscale을 조정한다
        if(layer=="22"):
            msp.add_line(start=start_point, end=end_point, dxfattribs={'layer': layer, 'ltscale' : 30})
        else:
            msp.add_line(start=start_point, end=end_point, dxfattribs={'layer': layer })
    else:
        msp.add_line(start=start_point, end=end_point)

    # 다음 선분의 시작점을 업데이트
    saved_Xpos = x2
    saved_Ypos = y2        
def circle_num(doc, x1, y1, x2, y2, text, option=None):
    msp = doc.modelspace()
    
    # 원 그리기
    radius = 60  # 원하는 반지름 설정
    msp.add_circle(
        center=(x2, y2),
        radius=radius,
        dxfattribs={'layer': '1'} # 녹색선
    )
    
    # 원 내부에 텍스트 추가
    msp.add_mtext(text, dxfattribs={
        'insert': (x2, y2),
        'layer': '0',
        'char_height': 40,        
        'attachment_point': 5  # 텍스트를 중앙에 배치
    })

    if option is None:
        # 원의 중심과 지시선 끝점 좌표를 사용하여 원의 둘레 상의 지시선 시작점 계산
        angle = math.atan2(y1 - y2, x1 - x2)
        start_x = x2 + radius * math.cos(angle)
        start_y = y2 + radius * math.sin(angle)

        # 지시선 그리기
        dim_leader(doc, x1, y1, start_x, start_y,  text, text_height=30, direction='up', option='nodraw')

    return msp
def lt(doc, x, y, layer=None):
    # 상대좌표로 그리는 것
    global saved_Xpos, saved_Ypos  # 전역 변수로 사용할 것임을 명시
    
    # 현재 위치를 시작점으로 설정
    start_x = saved_Xpos
    start_y = saved_Ypos

    # 끝점 좌표 계산
    end_x = start_x + x
    end_y = start_y + y

    # 모델 공간 (2D)을 가져옴
    msp = doc.modelspace()

    # 선 추가
    start_point = (start_x, start_y)
    end_point = (end_x, end_y)
    if layer:
        msp.add_line(start=start_point, end=end_point, dxfattribs={'layer': layer})
    else:
        msp.add_line(start=start_point, end=end_point)

    # 다음 선분의 시작점을 업데이트
    saved_Xpos = end_x
    saved_Ypos = end_y
def lineto(doc, x, y, layer=None):
    global saved_Xpos, saved_Ypos  # 전역 변수로 사용할 것임을 명시
    
    # 현재 위치를 시작점으로 설정
    start_x = saved_Xpos
    start_y = saved_Ypos

    # 끝점 좌표 계산
    end_x = x
    end_y = y

    # 모델 공간 (2D)을 가져옴
    msp = doc.modelspace()

    # 선 추가
    start_point = (start_x, start_y)
    end_point = (end_x, end_y)
    if layer:
        msp.add_line(start=start_point, end=end_point, dxfattribs={'layer': layer})
    else:
        msp.add_line(start=start_point, end=end_point)

    # 다음 선분의 시작점을 업데이트
    saved_Xpos = end_x
    saved_Ypos = end_y
def lineclose(doc, start_index, end_index, layer='레이져'):    

    firstX, firstY = globals()[f'X{start_index}'], globals()[f'Y{start_index}']
    prev_x, prev_y = globals()[f'X{start_index}'], globals()[f'Y{start_index}']

    # start_index+1부터 end_index까지 반복합니다.
    for i in range(start_index + 1, end_index + 1):
        # 현재 인덱스의 좌표를 가져옵니다.
        curr_x, curr_y = globals()[f'X{i}'], globals()[f'Y{i}']

        # 이전 좌표에서 현재 좌표까지 선을 그립니다.
        line(doc, prev_x, prev_y, curr_x, curr_y, layer)

        # 이전 좌표 업데이트
        prev_x, prev_y = curr_x, curr_y
        # print(f"prev_x {prev_x}" )
        # print(f"prev_y {prev_y}" )
    
    # 마지막으로 첫번째 점과 연결
    line(doc, prev_x, prev_y, firstX , firstY, layer)        
def rectangle(doc, x1, y1, dx, dy, layer=None, offset=None):
    if offset is not None:
        # 네 개의 선분으로 직사각형 그리기 offset 추가
        line(doc, x1+offset, y1+offset, dx-offset, y1+offset, layer=layer)   
        lineto(doc, dx - offset, dy - offset, layer=layer)  
        lineto(doc, x1 + offset, dy - offset, layer=layer)  
        lineto(doc, x1 + offset, y1 + offset, layer=layer)  
    else:        
        # 네 개의 선분으로 직사각형 그리기
        line(doc, x1, y1, dx, y1, layer=layer)   
        line(doc, dx, y1, dx, dy, layer=layer)   
        line(doc, dx, dy, x1, dy, layer=layer)   
        line(doc, x1, dy, x1, y1, layer=layer)   
def xrectangle(doc, x1, y1, dx, dy, layer=None):
    # 중간에 x마크로 적색선을 넣는 사각형 만들기
    line(doc, x1, y1, dx, y1, layer=layer)   
    line(doc, dx, y1, dx, dy, layer=layer)   
    line(doc, dx, dy, x1, dy, layer=layer)   
    line(doc, x1, dy, x1, y1, layer=layer)   
    line(doc, x1, y1, dx, dy, layer='3')    # 적색 센터라인
    line(doc, x1, dy, dx, y1, layer='3')  # 적색 센터라인     

def add_angular_dim_2l(drawing, base, line1, line2, location=None, text=None, text_rotation=None, dimstyle=selected_dimstyle, override=None, dxfattribs=None):
    # Create a new dimension line
    dim = drawing.dimstyle.add(
        'EZ_ANGULAR_2L',
        dxfattribs={
            'dimstyle': dimstyle,
            'dimtad': 1,  # Place text above dimension line
            'dimtih': False,  # Align text horizontally to dimension line
            'dimtoh': False,  # Align text outside horizontal
            'dimdec': 1
        }
    )
    
    # If override is provided, update the dimension style
    if override:
        dim.update(override)
    
    # Create angular dimension
    angular_dim = drawing.add_angular_dim_2l(
        base=base,
        line1=line1,
        line2=line2,
        location=location,
        text=text,
        text_rotation=text_rotation,
        dimstyle=dim.name,
        override=override,
        dxfattribs=dxfattribs,
    )
    
    # Render the dimension
    angular_dim.render()
    return angular_dim

def dim_angular(
    doc,
    x1, y1,   # 첫 번째 선분(Line1)의 시작점 좌표
    x2, y2,   # 첫 번째 선분(Line1)의 끝점 좌표
    x3, y3,   # 두 번째 선분(Line2)의 시작점 좌표
    x4, y4,   # 두 번째 선분(Line2)의 끝점 좌표
    distance=80,
    direction="left",
    dimstyle=selected_dimstyle
):
    """
    dim_angular()
    ---------------------------
    두 선분(Line1, Line2)의 좌표만 받아서 각도를 표시하는 치수(Angular Dimension)를
    ezdxf 문서(doc)에 생성하는 함수입니다. 실제로 선(Line) 엔티티를 그리지 않고,
    '가상의 선분'으로만 각도를 계산하여 각도 치수만 표시합니다.

    Parameters
    ----------
    doc : ezdxf.document.Document
        ezdxf 도큐먼트 객체. 각도 치수를 삽입할 대상 DXF 문서입니다.
    x1, y1 : float
        첫 번째 선분(Line1)의 시작점 좌표
    x2, y2 : float
        첫 번째 선분(Line1)의 끝점 좌표
    x3, y3 : float
        두 번째 선분(Line2)의 시작점 좌표
    x4, y4 : float
        두 번째 선분(Line2)의 끝점 좌표
    distance : float, default=80
        치수선(각도 표시)을 선분들의 중앙 지점에서 얼마나 떨어진 곳에 배치할지를 결정.
        값이 커질수록 치수선은 더 바깥쪽에 표시됩니다.
    direction : str, default='left'
        치수선(각도 표시)을 어느 방향으로 배치할지 결정합니다.
        - 'left'  : 왼쪽
        - 'right' : 오른쪽
        - 'up'    : 위쪽
        - 'down'  : 아래쪽
    dimstyle : str, default='mydim1'
        치수 스타일 이름. 미리 정의해둔 DimStyle 문자열을 지정할 수 있습니다.

    Returns
    -------
    ezdxf.entities.dimension.Dimension or None
        생성된 Angular Dimension 객체를 반환합니다.
        선분이 평행/0길이 등으로 각도 계산이 불가능하면 오류가 발생하거나
        None 처리 등 별도의 예외 처리가 필요할 수 있습니다.

    Notes
    -----
    - 두 선분이 완전히 평행하거나 길이가 0이면 ezdxf에서 각도 치수 생성이 불가능합니다.
      (ZeroDivisionError 또는 'Invalid colinear or parallel angle legs found' 등 발생)
    - 필요하다면, 선분 길이 0, 평행 여부를 검사하는 함수를 추가해 사용자가 예외를 처리할 수 있습니다.
    """
    msp = doc.modelspace()

    # ---------------------------------------------------------------------
    # (선분을 그리지 않고) 두 선분의 좌표만으로 각도 치수(Angular Dimension) 생성
    # ---------------------------------------------------------------------
    dimension = msp.add_angular_dim_2l(
        base=_calc_base_point(x1, y1, x2, y2, x3, y3, x4, y4, distance, direction),
        line1=((x1, y1), (x2, y2)),  # 첫 번째 선분
        line2=((x3, y3), (x4, y4)),  # 두 번째 선분
        dimstyle=dimstyle,
        override={
            'dimtxt': 0.22,    # 치수 문자 높이
            'dimgap': 0.02,    # 문자와 치수선 사이의 간격
            'dimscl': 1,       # 치수 축척
            'dimlfac': 1,      # 치수 단위 환산            
            'dimdec': 0        # 소수점 표기 자릿수 (0 = 소수점 없음)
        }
    )
    # 실제 도면에 반영
    dimension.render()
    return dimension

def _calc_base_point(x1, y1, x2, y2, x3, y3, x4, y4, distance, direction):
    """
    내부용 함수: 두 선분(Line1, Line2)의 좌표로부터
    치수선(각도 표시)의 기준점(base)을 계산.

    direction 값에 따라 base_x, base_y를 distance만큼
    '왼쪽/오른쪽/위/아래'로 이동합니다.
    """
    # 네 점(x1, y1, x2, y2, x3, y3, x4, y4)의 평균점(중심점) 계산
    base_x = (x1 + x2 + x3 + x4) / 4.0
    base_y = (y1 + y2 + y3 + y4) / 4.0

    if direction == "left":
        base_x -= distance
    elif direction == "right":
        base_x += distance
    elif direction == "up":
        base_y += distance
    elif direction == "down":
        base_y -= distance
    else:
        # default = left
        base_x -= distance

    return (base_x, base_y)

def is_zero_length(x1, y1, x2, y2):
    return (x1 == x2) and (y1 == y2)

def is_parallel(x1, y1, x2, y2, x3, y3, x4, y4):
    v1x = x2 - x1
    v1y = y2 - y1
    v2x = x4 - x3
    v2y = y4 - y3
    # 외적이 0이면 평행(또는 선분 길이가 0인 경우도 여기에 해당)
    return (v1x * v2y - v1y * v2x) == 0

def dim_diameter(doc, center, diameter, angle, dimstyle=selected_dimstyle, override=None):
    msp = doc.modelspace()
    
    # 기본 지름 치수선 추가
    dimension = msp.add_diameter_dim(
        center=center,  # 원의 중심점
        radius=diameter/2,  # 반지름
        angle=angle,    # 치수선 각도
        dimstyle=dimstyle,  # 치수 스타일
        override={"dimtoh": 1}    # 추가 스타일 설정 (옵션) 지시선이 한번 꺾여서 글자각도가 표준형으로 나오는 옵션
    )
    
    # 치수선의 기하학적 형태 생성
    dimension.render()    
def dim_string(doc, x1, y1, x2, y2, dis,  textstr,  text_height=0.30, text_gap=0.05, direction="up"):
    global saved_DimXpos, saved_DimYpos, saved_text_height, saved_text_gap, saved_direction, dimdistance, dim_horizontalbase, dim_verticalbase
    msp = doc.modelspace()
    dim_style = selected_dimstyle
    layer = selected_dimstyle

    # 치수값 계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # 소수점이 있는지 확인
    if dimension_value % 1 != 0:
        dimdec = 1  # 소수점이 있는 경우 소수점 첫째 자리까지 표시
    else:
        dimdec = 0  # 소수점이 없는 경우 소수점 표시 없음

    # override 설정
    override_settings = {
        'dimtxt': text_height,
        'dimgap': text_gap,
        'dimscl': 1,
        'dimlfac': 1,
        'dimclrt': 7,        
        'dimdec': dimdec,
        'dimtix': 1,  # 치수선 내부에 텍스트를 표시 (필요에 따라)
        'dimtad': 1 # 치수선 상단에 텍스트를 표시 (필요에 따라)               
    }

    # 방향에 따른 치수선 추가
    if direction == "up":
        dimension = msp.add_linear_dim(
            base=(x1, y1 + dis),
            dimstyle=dim_style,
            p1=(x1, y1),
            p2=(x2, y2),
            dxfattribs={'layer': layer},
            text = textstr,
            override=override_settings
        )
    elif direction == "down":
        dimension = msp.add_linear_dim(
            dimstyle=dim_style,
            base=(x1, y1 - dis),
            p1=(x1, y1),
            p2=(x2, y2),
            dxfattribs={'layer': layer},
            text = textstr,            
            override=override_settings
        )
    elif direction == "left":        
        dimension =  msp.add_linear_dim(
            dimstyle=dim_style,
            base=(x1 - dis, y1),
            angle=90,
            p1=(x1, y1),
            p2=(x2, y2),
            dxfattribs={'layer': layer},
            text = textstr,  
            override=override_settings
        )     
    elif direction == "right":        
        dimension =  msp.add_linear_dim(
            dimstyle=dim_style,
            base=(x1 + dis, y1),
            p1=(x1, y1),
            p2=(x2, y2),
            dxfattribs={'layer': layer},
            text = textstr,  
            angle=270,            
            override=override_settings
        )
    else:
        raise ValueError("Invalid direction. Use 'up', 'down', or 'aligned'.")

    dimension.render()
    return dimension
def d(doc, x1, y1, x2, y2, dis, text_height=0.30, text_gap=0.05, direction="up", option=None, starbottomtion=None, text=None, dim_style=None):
    global saved_DimXpos, saved_DimYpos, saved_text_height, saved_text_gap, saved_direction, dimdistance, dim_horizontalbase, dim_verticalbase, distanceXpos, distanceYpos

    # text_height=0.80 # 강제로 크게 해봄
    # Option 처리
    if option == 'reverse':   
        x1, x2 = x2, x1
        y1, y2 = y2, y1    
        saved_DimXpos, saved_DimYpos = x1, y1
    else:
        saved_DimXpos, saved_DimYpos = x2, y2

    dimdistance = dis
    saved_text_height = text_height
    saved_text_gap = text_gap
    saved_direction = direction

    # 연속선 구현을 위한 구문
    if starbottomtion is None:        
        if direction == "left":             
            distance = min(x1, x2) - dis            
            dim_horizontalbase = distance
        elif direction == "right":   
            distance = max(x1, x2) + dis                       
            dim_horizontalbase = distance 
        elif direction == "up":   
            distance = max(y1, y2) + dis            
            dim_verticalbase = distance
        elif direction == "down":     
            distance = min(y1, y2) - dis                                  
            dim_horizontalbase = distance
    else:
        if direction == "left":             
            distance = distanceXpos
        elif direction == "right":                        
            distance = distanceXpos
        elif direction == "up":            
            distance = distanceYpos
        elif direction == "down":     
            distance = distanceYpos

    msp = doc.modelspace()
    layer = selected_dimstyle
    if dim_style is None:        
        dim_style = selected_dimstyle
    else:        
        dim_style = selected_dimstyle

    # 치수값 계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # 소수점 확인
    dimdec = 1 if dimension_value % 1 != 0 else 0

    # override 설정
    if text is not None:                
        override_settings = {      
            'dimtxt': text_height,  
            'dimgap': text_gap if text_gap is not None else 0.05,
            'dimscl': 1,
            'dimlfac': 1,
            'dimclrt': 7,        
            'dimdec': dimdec,
            'dimtix': 1, 
            'dimtad': 1,
            'dimtmove': 0,  # 텍스트 이동 옵션
            'dimpost': f"{text} " # 치수선 앞에 텍스트 표시
        }
    else:
        override_settings = {      
            'dimtxt': text_height,  
            'dimgap': text_gap if text_gap is not None else 0.05,
            'dimscl': 1,
            'dimlfac': 1,
            'dimclrt': 7,        
            'dimdec': dimdec,
            'dimtix': 1, 
            'dimtad': 1,
            'dimtmove': 2,  # 텍스트 이동 옵션
        }

    base_point = ((x1 + x2) / 2, distance)    
    add_dim_args = {
        'dimstyle': dim_style,
        'base': base_point,
        'p1': (x1, y1),
        'p2': (x2, y2),
        'override': override_settings
    }

    if direction in ["up", "down"]:
        add_dim_args['base'] = ((x1 + x2) / 2, distance)
    elif direction in ["left", "right"]:
        add_dim_args['base'] = (distance, (y1 + y2) / 2)
        add_dim_args['angle'] = 90

    if starbottomtion is None:
        if direction == "up":
            distanceYpos = max(y1, y2) + dis
        elif direction == "down":
            distanceYpos = min(y1, y2) - dis
        elif direction == "left":
            distanceXpos = min(x1, x2) - dis
        elif direction == "right":
            distanceXpos = max(x1, x2) + dis

    if text is not None:
        add_dim_args['text'] = text  # 추가 텍스트만 설정

    if option != 'noprint':
        return msp.add_linear_dim(**add_dim_args)
    else:
        return

def dc(doc, x, y, distance=None, option=None, text=None, dim_style=None) :    
    global saved_DimXpos, saved_DimYpos, saved_text_height, saved_text_gap, saved_direction, dimdistance, dim_horizontalbase, dim_verticalbase, distanceXpos, distanceYpos
    x1 = saved_DimXpos
    y1 = saved_DimYpos
    x2 = x
    y2 = y        
    if distance is not None :
        dimdistance = distance    
    # reverse 옵션 처리
    if option == 'reverse':
        x1, x2 = x2, x1
        y1, y2 = y2, y1             
    if dim_style is None:        
        dim_style = selected_dimstyle
    else:        
        dim_style = dim_style               
    d(doc, x1, y1, x2, y2, dimdistance, text_height=saved_text_height, text_gap=saved_text_gap, direction=saved_direction, starbottomtion='continue', text=text,  dim_style = dim_style, option=option )
def dim(doc, x1, y1, x2, y2, dis, text_height=0.30, text_gap=0.05, direction="up", option=None, starbottomtion=None, text=None):
    global saved_DimXpos, saved_DimYpos, saved_text_height, saved_text_gap, saved_direction, dimdistance, dim_horizontalbase, dim_verticalbase

    # Option 처리
    if option == 'reverse':
        saved_DimXpos, saved_DimYpos = x1, y1
    else:
        saved_DimXpos, saved_DimYpos = x2, y2

    dimdistance = dis
    saved_text_height = text_height
    saved_text_gap = text_gap
    saved_direction = direction

    # 연속선 구현을 위한 구문
    if starbottomtion is None:
        if direction == "left":             
            dim_horizontalbase = dis - (x1 - x2)
        elif direction == "right":                        
            dim_horizontalbase = dis 
        elif direction == "up":            
            dim_verticalbase = dis
        elif direction == "down":                        
            dim_verticalbase = dis   

    # flip을 선언하면 치수선의 시작과 끝을 바꾼다. 저장된 좌표는 지장없다.
    # 치수선의 시작점 끝점에 따라 치수선이 나오는 것을 만들기 위함이다. 
    # 연속치수선때는 좌표가 바뀌면 안되기때문에 고려한 부분이다.

    msp = doc.modelspace()
    dim_style = selected_dimstyle
    layer = selected_dimstyle

    # 치수값 계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # 소수점 확인
    dimdec = 1 if dimension_value % 1 != 0 else 0

    # override 설정
    override_settings = {      
        'dimtxt': text_height,  
        'dimgap': text_gap if text_gap is not None else 0.05,  # 여기에서 dimgap에 기본값 설정
        'dimscl': 1,
        'dimlfac': 1,
        'dimclrt': 7,        
        'dimdec': dimdec,
        'dimtix': 1, 
        'dimtad': 1  
    }

    # 방향에 따른 치수선 추가
    if direction == "up":
        add_dim_args = {
            'base': (x1, y1 + dis),
            'dimstyle': dim_style,
            'p1': (x1, y1),
            'p2': (x2, y2),
            'override': override_settings
        }
        if text is not None :
            add_dim_args['text'] = text
            msp.add_linear_dim(**add_dim_args).render()
            return
        else:
            return msp.add_linear_dim(**add_dim_args)

    elif direction == "down":
        add_dim_args = {
            'dimstyle': dim_style,
            'base': (x1, y1 - dis),
            'p1': (x1, y1),
            'p2': (x2, y2),
            'override': override_settings
        }
        if text is not None :
            add_dim_args['text'] = text
            msp.add_linear_dim(**add_dim_args).render()
            return
        else:
            return msp.add_linear_dim(**add_dim_args)

    elif direction == "left":
        if option == 'reverse':   
            add_dim_args = {
                'dimstyle': dim_style,
                'base': (x2 - dis, y2),
                'angle': 90,
                'p1': (x1, y1),
                'p2': (x2, y2),
                'override': override_settings
            }
        else:
            add_dim_args = {
                'dimstyle': dim_style,
                'base': (x1 - dis, y1),
                'angle': 90,
                'p1': (x2, y2),
                'p2': (x1, y1),
                'override': override_settings
            }        
        if text is not None :
            add_dim_args['text'] = text
            msp.add_linear_dim(**add_dim_args).render()
            return
        else:
            return msp.add_linear_dim(**add_dim_args)

    elif direction == "right":
        add_dim_args = {
            'dimstyle': dim_style,
            'base': (x1 + dis, y1),
            'angle': 90,
            'p1': (x1, y1),
            'p2': (x2, y2),
            'override': override_settings
        }
        if text is not None :
            add_dim_args['text'] = text
            msp.add_linear_dim(**add_dim_args).render()
            return
        else:
            return msp.add_linear_dim(**add_dim_args)

    elif direction == "aligned":
        add_dim_args = {
            'dimstyle': dim_style,
            'points': [(x1, y1), (x2, y2)],
            'base': (x1 + dis, y1),
            'dxfattribs': {'layer': layer},
            'override': override_settings
        }
        if text is not None :
            add_dim_args['text'] = text
        return msp.add_multi_point_linear_dim(**add_dim_args)

    else:
        raise ValueError("Invalid direction. Use 'up', 'down', 'left', 'right', or 'aligned'.")
def dimcontinue(doc, x, y, distance=None, option=None) :    
    global saved_DimXpos
    global saved_DimYpos
    global saved_text_height
    global saved_text_gap
    global saved_direction
    global dimdistance
    global dim_horizontalbase
    global dim_verticalbase

    x1 = saved_DimXpos
    y1 = saved_DimYpos
    x2 = x
    y2 = y        

# 방향에 대한 정의를 하고 연속적으로 차이를 감지해서 처리하는 것이다.
    if saved_direction=="left" :
        dimdistance = dim_horizontalbase
        # 재계산해야 함.        
        dim_horizontalbase = dimdistance - (x1 - x2)        
    if saved_direction=="right" :
        dimdistance = dim_horizontalbase 
        # 재계산해야 함.
        dim_horizontalbase = dimdistance - (x2 - x1)
    if saved_direction=="up" :
        dimdistance = dim_verticalbase
        # 재계산해야 함.
        dim_verticalbase = dimdistance - (y2 - y1)
    if saved_direction=="down" :
        dimdistance = dim_verticalbase 
        # 재계산해야 함.
        dim_verticalbase = dimdistance - (y1 - y2)

    if distance is not None :
        dimdistance = distance    

    # reverse 옵션 처리
    if option == 'reverse':
        x1, x2 = x2, x1
        y1, y2 = y2, y1                

    dim(doc, x1, y1, x2, y2, dimdistance, text_height=saved_text_height, text_gap=saved_text_gap, direction=saved_direction, starbottomtion='continue')
def dimto(doc, x2, y2, dis, text_height=0.20, text_gap=None, option=None):    
    global saved_DimXpos
    global saved_DimYpos
    global saved_text_height
    global saved_text_gap
    global saved_direction

    # 오류 수정: text_gap이 None이 아닐 때만 saved_text_gap을 갱신해야 함
    if text_gap is not None:
        saved_text_gap = text_gap
    else:
        text_gap = saved_text_gap

    # 오류 수정: text_height가 None이 아닐 때만 saved_text_height을 갱신해야 함
    if text_height is not None:
        saved_text_height = text_height
    else:
        text_height = saved_text_height

    dim(doc, saved_DimXpos, saved_DimYpos, x2, y2, dis, text_height=text_height, text_gap=text_gap, direction=saved_direction, option=option)
def create_vertical_dim(doc, x1, y1, x2, y2, dis, angle, layer=None, text_height=0.30, text_gap=0.05):
    msp = doc.modelspace()
    dim_style = layer  # 치수 스타일 이름
    points = [(x1, y1), (x2, y2)]

    if angle==None :
        angle = 270

    # 치수값 계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # 소수점이 있는지 확인
    if dimension_value % 1 != 0:
        dimdec = 1  # 소수점이 있는 경우 소수점 첫째 자리까지 표시
    else:
        dimdec = 1  # 소수점이 없는 경우 소수점 표시 없음    

    return msp.add_multi_point_linear_dim(
        base=(x1 + dis if angle == 270 else x1 - dis , y1),  #40은 보정
        points = points,
        angle = angle,
        dimstyle = dim_style,
        discard = True,
        dxfattribs = {'layer': layer},
        # 치수 문자 위치 조정 (0: 치수선 위, 1: 치수선 옆) 'dimtmove': 1 
        # override={'dimtxt': text_height, 'dimscl': 1, 'dimlfac': 1, 'dimclrt': 7, 'dimdec': dimdec, 'dimdsep': 46, 'dimtmove': 3  }
        override = {'dimtxt': text_height, 'dimgap': text_gap, 'dimscl': 1, 'dimlfac': 1, 'dimclrt': 7, 'dimdec': dimdec, 'dimtmove': 3  }
    )
def dim_vertical_right(doc, x1, y1, x2, y2, dis, layer=selected_dimstyle, text_height=0.30,  text_gap=0.07, angle=None):
    return create_vertical_dim(doc, x1, y1, x2, y2, dis, angle, layer, text_height, text_gap)
def dim_vertical_left(doc, x1, y1, x2, y2, dis, layer=selected_dimstyle, text_height=0.30,  text_gap=0.07):
    return create_vertical_dim(doc, x1, y1, x2, y2, dis, 90, layer, text_height, text_gap)
def create_vertical_dim_string(doc, x1, y1, x2, y2, dis, angle, textstr, text_height=0.30, text_gap=0.07):    
    msp = doc.modelspace()
    dim_style = selected_dimstyle
    layer = selected_dimstyle
    points = [(x1, y1), (x2, y2)]

    # 치수값 계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # 소수점이 있는지 확인
    if dimension_value % 1 != 0:
        dimdec = 1  # 소수점이 있는 경우 소수점 첫째 자리까지 표시
    else:
        dimdec = 1  # 소수점이 없는 경우 소수점 표시 없음    

    return msp.add_multi_point_linear_dim(
        base=(x1 + dis if angle == 270 else x1 - dis , y1),  #40은 보정
        points=points,
        angle=angle,
        dimstyle=dim_style,
        discard=True,
        dxfattribs={'layer': layer},
        # 치수 문자 위치 조정 (0: 치수선 위, 1: 치수선 옆) 'dimtmove': 1 
        override={'dimtxt': text_height, 'dimgap': text_gap, 'dimscl': 1, 'dimlfac': 1, 'dimclrt': 7, 'dimdec': dimdec,  'dimtmove': 3 , 'text' : textstr      }
    )
def dim_vertical_right_string(doc, x1, y1, x2, y2, dis, textstr, text_height=0.30,  text_gap=0.07):
    return create_vertical_dim_string(doc, x1, y1, x2, y2, dis, 270, textstr, text_height, text_gap)
def dim_vertical_left_string(doc, x1, y1, x2, y2, dis, textstr, text_height=0.30,  text_gap=0.07):
    return create_vertical_dim_string(doc, x1, y1, x2, y2, dis, 90, textstr, text_height, text_gap)
def draw_Text_direction(doc, x, y, size, text, layer=None, rotation = 90):
    # 모델 공간 (2D)을 가져옴
    msp = doc.modelspace()

    # text_style_name ='GHS'
    text_style_name =selected_dimstyle
    # MText 객체 생성
    mtext = msp.add_mtext(
        text,  # 텍스트 내용
        dxfattribs={
            'layer': layer,  # 레이어 지정
            'style': text_style_name,  # 텍스트 스타일 지정
            'char_height': size,  # 문자 높이 (크기) 지정
        }
    )

    # MText 위치와 회전 설정
    mtext.set_location(insert=(x, y), attachment_point=1, rotation=rotation)

    return mtext

def draw_Text(doc, x, y, size, text, layer=None):
    # 모델 공간 (2D)을 가져옴
    msp = doc.modelspace()

    # layer가 None이면 기본값 설정
    if layer is None:
        layer = "mydim"  # 기본 레이어 이름 설정 (필요 시 변경 가능)

    # text_style_name을 설정
    text_style_name = selected_dimstyle

    # 텍스트 추가 및 생성된 Text 객체 가져오기
    text_entity = msp.add_text(
        text,  # 텍스트 내용
        dxfattribs={
            'layer': layer,  # 레이어 지정
            'style': text_style_name,  # 텍스트 스타일 지정
            'height': size,  # 텍스트 높이 (크기) 지정
        }
    )

    # Text 객체의 위치 설정
    text_entity.set_placement((x, y), align=TextEntityAlignment.BOTTOM_LEFT)

    # MIDDLE_LEFT로 놓으면 다른 도면에 붙일때 문자 위치가 달라지는 경우가 있다. 주의요함 (일해이엔지 개발시 발견함)
def draw_circle(doc, center_x, center_y, radius, layer='0', color='7'):
    """
    원을 그리는 함수 radius는 지름으로 넣도록 수정
    :param doc: ezdxf 문서 객체
    :param center_x: 원의 중심 x 좌표
    :param center_y: 원의 중심 y 좌표
    :param radius: 원의 반지름이 아닌 지름입력
    : radius=radius/2 적용
    :param layer: 원을 추가할 레이어 이름 (기본값은 '0')
    지름으로 수정 
    """
    msp = doc.modelspace()
    circle = msp.add_circle(center=(center_x, center_y), radius=radius/2, dxfattribs={'layer': layer, 'color' : color})
    return circle
def circle_cross(doc, center_x, center_y, radius, layer='0', color='7'):
    """
    DXF 문서에 원을 그리는 함수
    :param doc: ezdxf 문서 객체
    :param center_x: 원의 중심 x 좌표
    :param center_y: 원의 중심 y 좌표
    :param radius: 원의 반지름
    :param layer: 원을 추가할 레이어 이름 (기본값은 '0')
    지름으로 수정 /2 적용
    """
    draw_circle(doc, center_x, center_y, radius, layer=layer, color=color)
    # 적색 십자선 그려주기    
    line(doc, center_x - radius/2 - 5, center_y, center_x +  radius/2 + 5, center_y, layer="CEN" )
    line(doc, center_x , center_y - radius/2 - 5, center_x , center_y +  radius/2 + 5, layer="CEN" )
    return circle_cross        
def cross10(doc, center_x, center_y):
# 10미리 십자선 레이져 만들기   
    line(doc, center_x - 5, center_y, center_x +  5, center_y, layer="레이져" )
    line(doc, center_x, center_y - 5, center_x , center_y + 5, layer="레이져" )
    return cross10
def cross(doc, center_x, center_y, length, layer='레이져'):
    line(doc, center_x - length, center_y, center_x +  length, center_y, layer=layer )
    line(doc, center_x, center_y - length, center_x , center_y + length, layer=layer )
    return cross
def crossslot(doc, center_x, center_y, direction=None):
# 10미리 십자선 레이져 만들기   
    line(doc, center_x - 10, center_y, center_x +  10, center_y, layer="CL" )
    line(doc, center_x, center_y - 10, center_x , center_y + 10, layer="CL" )
    if direction== 'vertical':
        insert_block(doc,center_x, center_y , "8x16_vertical_draw", layer='0')
    else:
        insert_block(doc,center_x, center_y , "8x16_horizontal_draw", layer='0')
    return cross10
def m14(doc, center_x, center_y,layer='0', color='4'):
    radius = 14
    draw_circle(doc, center_x, center_y, 14 , layer=layer, color=color)
    draw_circle(doc, center_x, center_y, 8 , layer=layer, color=color)
    draw_circle(doc, center_x, center_y, 4.9 , layer=layer, color=color)
    # 적색 십자선 그려주기    
    line(doc, center_x - radius/2 - 5, center_y, center_x +  radius/2 + 5, center_y, layer="CEN" )
    line(doc, center_x , center_y - radius/2 - 5, center_x , center_y +  radius/2 + 5, layer="CEN" )
    return 
def extract_abs(a, b):
    return abs(a - b)   
def insert_block(doc, x, y, block_name, layer='레이져'):
    # 도면틀 삽입    
    scale = 1
    insert_point = (x, y, scale)

    # 블록 삽입하는 방법           
    doc.modelspace().add_blockref(block_name, insert_point, dxfattribs={
        'xscale': scale,
        'yscale': scale,
        'rotation': 0,
        'layer': layer
    })
def insert_frame(x, y, scale, title, description, workplaceStr, sep="NOtable"):
    # issuedate 오류 원인 설명:
    # issuedate가 None이거나, 문자열이지만 형식이 "%Y-%m-%d %H:%M:%S"가 아닐 경우 datetime.strptime에서 에러가 발생합니다.
    # 또한, issuedate가 아예 값이 없거나, 타입이 예상과 다를 때도 문제가 생깁니다.
    # 안전하게 처리하려면 None 체크와 문자열 포맷 예외처리가 필요합니다.

    global issuedate

    formatted_date = ""
    if issuedate is None:
        formatted_date = ""
    elif isinstance(issuedate, str):
        # 문자열이지만 포맷이 다를 수 있으므로 예외처리
        try:
            date_object = datetime.strptime(issuedate, "%Y-%m-%d %H:%M:%S")
            formatted_date = date_object.strftime("%y.%m.%d")
        except Exception:
            # 다른 포맷 시도 또는 그냥 원본 사용
            try:
                date_object = datetime.strptime(issuedate, "%Y-%m-%d")
                formatted_date = date_object.strftime("%y.%m.%d")
            except Exception:
                formatted_date = issuedate
    elif isinstance(issuedate, datetime):
        formatted_date = issuedate.strftime("%y.%m.%d")
    else:
        formatted_date = str(issuedate)

    sep = "NOtable"
    # 도면틀 삽입
    if sep == "basic":
        block_name = "drawings_frame"
    # 2열 삽입 도면 ASSY도면
    if sep == "column2":
        block_name = "drawings_frame_column2"
    # 7열 삽입 도면 ASSY도면
    if sep == "column7":
        block_name = "drawings_frame_column7"
    # 1열 삽입 도면 ASSY도면
    if sep == "NOtable":
        block_name = "drawings_frame_NOtable"

    insert_point = (x, y, scale)

    # 블록 삽입하는 방법
    msp.add_blockref(block_name, insert_point, dxfattribs={
        'xscale': scale,
        'yscale': scale,
        'rotation': 0
    })

    draw_Text(doc, x + (3545 + 900) * scale, y + 420 * scale, 50 * scale, str(description), '0')
    draw_Text(doc, x + (3545 + 900 + 100) * scale, y + 630 * scale, 60 * scale, f"{title}", '0')
    draw_Text(doc, x + (3545 + 900) * scale, y + 850 * scale, 100 * scale, f"현장명 : {workplaceStr}", '0')

def envsettings():
    # 하드디스크 고유번호를 가져오는 코드 (시스템에 따라 다를 수 있음)
    disk_id = os.popen('wmic diskdrive get serialnumber').read().strip()
    data = {"DiskID": disk_id}
    with open(license_file_path, 'w', encoding='utf-8') as file:
        json.dump(data, file)        
def adjust_coordinates(coord1, coord2, coord3):
    # Function to adjust coordinates with rounding and adding truncated decimals    
    adjusted_coord1 = round(coord1)
    decimal_part = coord1 - adjusted_coord1
    adjusted_coord2 = round(coord2 + decimal_part)
    adjusted_coord3 = round(coord3 + (coord2 - adjusted_coord2))

    return adjusted_coord1, adjusted_coord2, adjusted_coord3
def calculate_holeArray(startnum, interval, limit, length):
    # 결과를 저장할 리스트 초기화
    hole_array = []

    # 현재 숫자를 startnum으로 설정
    current_num = startnum

    # current_num이 limit과 length를 넘지 않을 때까지 반복
    while current_num <= limit and current_num <= length:
        # 리스트에 현재 숫자 추가
        hole_array.append(current_num)
        # 다음 숫자를 interval만큼 증가
        current_num += interval
    hole_array.append(length-85)
    return hole_array
def calculate_splitholeArray(startnum, interval, limit, length):    
    hole_array = []

    # 현재 숫자를 startnum으로 설정
    current_num = startnum

    # current_num이 limit과 length를 넘지 않을 때까지 반복
    while current_num <= limit and current_num <= length:
        # 리스트에 현재 숫자 추가
        hole_array.append(current_num)
        # 다음 숫자를 interval만큼 증가
        current_num += interval
    hole_array.append(length)
    return hole_array

def calSplitHole(start, interval, limit):
    """
    start: 시작값
    interval: 간격
    limit: 최대 범위 (예: MH + HH + grounddig)
    """
    hole_array = []
    current_num = start
    
    while current_num <= limit:
        hole_array.append(current_num)
        current_num += interval
    
    return hole_array

def draw_arc_slot(doc, center, radius, start_angle, end_angle, layer):
    """
    Draws an arc in the DXF document.
    
    Parameters:
    doc (ezdxf.document): The DXF document to draw on.
    center (tuple): The (x, y) coordinates of the arc's center.
    radius (float): The radius of the arc.
    start_angle (float): The starting angle of the arc in degrees.
    end_angle (float): The ending angle of the arc in degrees.
    layer (str): The layer to draw the arc on.
    """
    msp = doc.modelspace()
    msp.add_arc(
        center=center,
        radius=radius,
        start_angle=start_angle,
        end_angle=end_angle,
        dxfattribs={'layer': layer}
    )

def draw_slot(doc, x, y, size, direction="가로", option=None, layer='0'):
    """
    Draws a slot (장공) with specified parameters in a DXF document.
    
    Parameters:
    doc (ezdxf.document): The DXF document to draw on.
    x (float): The x-coordinate of the slot's center.
    y (float): The y-coordinate of the slot's center.
    size (str): The size of the slot in the format 'WxH' (e.g., '8x16').
    direction (str): The direction of the slot, either '가로' (horizontal) or '세로' (vertical). Default is '가로'.
    option (str): Optional feature for the slot. If 'cross', draws center lines extending beyond the slot. Default is None.
    layer (str): The layer to draw the slot on. Default is '0'.
    """    
    msp = doc.modelspace()

    # size를 분리하여 폭과 높이 계산
    width, height = map(int, size.lower().split('x'))

    if direction == "가로":
        slot_length = height - width
        slot_width = width
    else:  # 세로
        slot_length = width
        slot_width = height - width

    radius = slot_width / 2

    # 중심점을 기준으로 시작점과 끝점 계산
    if direction == "가로":
        start_point = (x - slot_length / 2, y)
        end_point = (x + slot_length / 2, y)
    else:  # 세로
        start_point = (x, y - slot_length / 2)
        end_point = (x, y + slot_length / 2)

    # 직선 부분 그리기
    if direction == "가로":
        msp.add_line((start_point[0], start_point[1] + radius), (end_point[0], end_point[1] + radius), dxfattribs={'layer': layer})
        msp.add_line((start_point[0], start_point[1] - radius), (end_point[0], end_point[1] - radius), dxfattribs={'layer': layer})
    else:  # 세로
        msp.add_line((start_point[0] + radius, start_point[1]), (end_point[0] + radius, end_point[1]), dxfattribs={'layer': layer})
        msp.add_line((start_point[0] - radius, start_point[1]), (end_point[0] - radius, end_point[1]), dxfattribs={'layer': layer})

    # 양 끝의 반원 그리기
    if direction == "가로":
        draw_arc_slot(doc, start_point, radius, 90, 270, layer)  # 반원 방향 수정
        draw_arc_slot(doc, end_point, radius, 270, 90, layer)  # 반원 방향 수정
    else:  # 세로
        draw_arc_slot(doc, start_point, radius, 180, 360, layer)  # 반원 방향 수정
        draw_arc_slot(doc, end_point, radius, 0, 180, layer)  # 반원 방향 수정

    # 옵션이 "cross"인 경우 중심선을 추가로 그리기
    if option == "cross":
        if direction == "가로":
            msp.add_line((x - slot_length / 2 - 8, y), (x + slot_length / 2 + 8, y), dxfattribs={'layer': 'CL'})
            msp.add_line((x, y - slot_width / 2 - 4), (x, y + slot_width / 2 + 4), dxfattribs={'layer': 'CL'})
        else:  # 세로
            msp.add_line((x - slot_width / 2 - 4, y), (x + slot_width / 2 + 4, y), dxfattribs={'layer': 'CL'})
            msp.add_line((x, y - slot_length / 2 - 8), (x, y + slot_length / 2 + 8), dxfattribs={'layer': 'CL'})

    return msp

def read_panel_data():
    panel_data = []
    for i in range(1, 12):
        panel = {
            "material": globals()[f"P{i}_material"],
            "width": globals()[f"P{i}_width"],
            "widthReal": globals()[f"P{i}_widthReal"],
            "holegap": globals()[f"P{i}_holegap"],
            "holes": [globals()[f"P{i}_hole{j}"] for j in range(1, 6)],
            "COPType": globals()[f"P{i}_COP"]            
        }
        panel_data.append(panel)    
    return panel_data
def generate_drawing(row_data):
    """
    도면 생성 로직을 처리하는 함수
    """
    global jambType, floorDisplay, material, spec, vcut, OP, JE, JD, HH, MH
    global HPI_height, U, C, A, grounddig, control_width, controltopbar, controlbox, poleAngle

    # 행 데이터를 전역 변수로 설정
    jambType = row_data.get("jambType", 0)
    floorDisplay = row_data.get("floorDisplay", 0)
    material = row_data.get("material", 0)
    spec = row_data.get("spec", 0)
    vcut = row_data.get("vcut", 0)
    OP = row_data.get("OP", 0)
    JE = row_data.get("JE", 0)
    JD = row_data.get("JD", 0)
    HH = row_data.get("HH", 0)
    MH = row_data.get("MH", 0)
    HPI_height = row_data.get("HPI_height", 0)
    U = row_data.get("U", 0)
    C = row_data.get("C", 0)
    A = row_data.get("A", 0)
    grounddig = row_data.get("grounddig", 0)
    control_width = row_data.get("control_width", 0)
    controltopbar = row_data.get("controltopbar", 0)
    controlbox = row_data.get("controlbox", 0)
    poleAngle = round(row_data.get("poleAngle", 0), 1)  # 소수점 둘째 자리까지 제한

    # 도면 생성 작업
    # print(f"Generating drawing for jambType: {jambType}, material: {material}, spec: {spec}")
    # 여기에 실제 도면 생성 코드를 추가

def calculate_base(height, angle_degrees, option=None):
    """
    주어진 높이와 끼인각(도)을 이용해 밑변을 계산합니다.

    # # 예제 사용법
    # height = 200
    # angle_degrees = 10
    # base = calculate_base(height, angle_degrees)
    # print(f"높이가 {height}이고 끼인각이 {angle_degrees}도일 때 밑변은 {base}입니다.")    

    :param height: 높이 (세로 변)
    :param angle_degrees: 끼인각 (도)
    :return: 밑변 길이
    """
    if angle_degrees <= 0 or angle_degrees >= 90:
        raise ValueError("각도는 0도 초과 90도 미만이어야 합니다.")
    
    # 각도를 라디안으로 변환
    angle_radians = math.radians(angle_degrees)
    # 밑변 계산: 높이와 탄젠트의 곱
    base = height * math.tan(angle_radians)
    if option=="height" :
        base = height * math.cos(angle_radians)
    return round(base, 2)

def set_point(a, index, x_value, y_value):
    """
    주어진 딕셔너리에 x, y 값을 설정하는 함수
    """
    a[f'x{index}'], a[f'y{index}'] = x_value, y_value    
def calcuteHoleArray(totalLength, startLength, holeNumber):
    """
    totalLength : 전체 길이
    startLength : 좌우 여유 길이(시작·끝 지점에서 떨어진 거리)
    holeNumber  : 전체적으로 뚫을 홀(구멍)의 개수

    예) totalLength=250, startLength=20, holeNumber=5
       -> 5개의 홀 위치가 [20, 72.5, 125.0, 177.5, 230.0]으로 계산
    """
    # 혹시 holeNumber가 1 이하라면 처리 불가능하니 예외 처리
    if holeNumber < 2:
        raise ValueError("holeNumber는 최소 2 이상이어야 합니다.")

    # 결과 리스트
    holes = []

    # "양쪽 가장자리"를 이미 포함해야 하므로,
    # 가운데에 들어갈 홀 수는 (holeNumber - 2),
    # 하지만 간격을 구할 땐 (holeNumber - 1)로 나눈다.
    #
    # --예시--
    #  holeNumber=5라면, 총 5개의 홀을 찍을 것이고,
    #  간격은 4구간(=5-1)으로 나눈다.
    #
    # 양쪽 startLength를 제외한 순수 중간 영역
    middle_length = totalLength - 2 * startLength

    # 구간 간격
    interval = middle_length / (holeNumber - 1)

    # 0번 홀부터 holeNumber-1번 홀까지 계산
    for i in range(holeNumber):
        pos = startLength + interval * i
        holes.append(pos)

    # 간격 보정 로직 추가
    # 모든 간격을 소수점 첫째자리까지 반올림해서 비교
    if len(holes) >= 3:
        # 실제 간격 리스트 생성
        intervals = [round(holes[i+1] - holes[i], 1) for i in range(len(holes)-1)]
        # 가장 많이 나오는 간격을 기준으로 삼음
        from collections import Counter
        cnt = Counter(intervals)
        most_common_gap, _ = cnt.most_common(1)[0]
        # 첫번째 간격이 다른 경우 보정
        if abs(intervals[0] - most_common_gap) >= 0.1:
            # 첫번째 위치를 보정
            diff = intervals[0] - most_common_gap
            # holes[0]은 startLength, holes[1]만 보정
            holes[1] = holes[0] + most_common_gap
            # 이후 값들도 간격을 맞춰서 재계산
            for i in range(2, len(holes)):
                holes[i] = holes[i-1] + most_common_gap
            # 마지막 값이 endLength를 넘으면 마지막만 endLength로 맞춤
            endLength = totalLength - startLength
            if abs(holes[-1] - endLength) > 0.1:
                holes[-1] = endLength

    return holes

def simulate_hole_positions_from_bottom(totalLength, target_first_distance, hole_count, poleAngle):
    """
    손이 들어가는 구조를 위한 홀 위치 시뮬레이션
    
    - 첫 번째 홀: 수직거리 52mm
    - 홀 간격: 사선거리 64mm 전후 (각도 고려)
    """
    holes = []
    
    # 기존 방식으로 기본 홀 위치 계산 (조건 순서 수정)
    if hole_count >= 5:
        holes = calcuteHoleArray(totalLength, 48.5, 5)
    elif hole_count >= 4:
        holes = calcuteHoleArray(totalLength, 48.5, 4)
    else:
        holes = calcuteHoleArray(totalLength, 48.5, 3)
    
    # 안전한 접근: 기존 홀 위치를 거의 그대로 유지하면서 미세 조정만
    # 문제: offset이 음수가 되어 홀들이 아래로 내려감
    
    # 현재 홀들이 정상 위치에 있는지 확인 후 미세 조정만 적용
    # 41.6 → 52로 증가시키려면 홀을 아래쪽으로 이동해야 함
    # (홀이 아래로 가면 상단까지 거리가 증가)
    
    # 안전하게 5mm만 아래쪽으로 이동 (41.6 + 약 5 ≈ 47 정도)
    safe_adjustment = -0.6  # 음수 = 아래쪽 이동 = 상단 거리 증가
    
    # 모든 홀을 동일하게 이동 (간격 유지)
    holes = [hole + safe_adjustment for hole in holes]
    
    return holes

def ds(doc, x1, y1, x2, y2, dis, text_height=0.30, text_gap=0.05,
       direction="up", option=None, starbottomtion=None, text=None,
       dim_style=None):
    """
    문자열 + 치수선 함께 구현하는 예제 함수.

    Parameters
    ----------
    doc : ezdxf.document.Document
        ezdxf 도큐먼트 객체(치수를 추가할 DXF 문서).
    x1, y1 : float
        첫 번째 점 P1 (x, y)
    x2, y2 : float
        두 번째 점 P2 (x, y)
    dis : float
        치수선(문자 표시 위치)을 기준점에서 얼마나 떨어뜨릴지 결정하는 거리.
    text_height : float, default=0.30
        치수 문자 높이.
    text_gap : float, default=0.05
        치수 문자와 치수선 사이 간격.
    direction : {"up", "down", "left", "right"}
        치수선을 어느 방향(수평/수직)으로 배치할지 결정.
        - "up"    : P1, P2를 잇는 선분 위쪽에 수평 치수선
        - "down"  : P1, P2를 잇는 선분 아래쪽에 수평 치수선
        - "left"  : P1, P2를 잇는 선분 왼쪽에  수직 치수선
        - "right" : P1, P2를 잇는 선분 오른쪽에 수직 치수선
    option : str or None
        'reverse'이면 P1, P2를 바꾸어(뒤집어서) 치수 표시.
    text : str or None
        치수 값에 앞서 표시할 문구. 예: "상판 높이"
    dim_style : str or None
        사용할 치수 스타일. None이면 기본 "over1000dim1" 사용.

    Returns
    -------
    ezdxf.entities.dimension.Dimension 
        생성된 치수 엔티티 객체를 반환.
    """

    # 옵션이 'reverse'인 경우, 두 점을 서로 바꿔서 치수를 반대로 표시
    if option == 'reverse':
        x1, x2 = x2, x1
        y1, y2 = y2, y1

    # 방향에 따라 angle(0 or 90)과 base(치수선 기준점) 계산
    if direction == "up":
        angle = 0  # 수평 치수선
        base_x = (x1 + x2) / 2
        base_y = max(y1, y2) + dis
    elif direction == "down":
        angle = 0  # 수평 치수선
        base_x = (x1 + x2) / 2
        base_y = min(y1, y2) - dis
    elif direction == "left":
        angle = 90  # 수직 치수선
        base_x = min(x1, x2) - dis
        base_y = (y1 + y2) / 2
    elif direction == "right":
        angle = 90  # 수직 치수선
        base_x = max(x1, x2) + dis
        base_y = (y1 + y2) / 2
    else:
        # 혹은 기본값을 "up"으로 처리
        angle = 0
        base_x = (x1 + x2) / 2
        base_y = max(y1, y2) + dis

    # Model Space 획득
    msp = doc.modelspace()

    # 치수선 생성
    dim = msp.add_linear_dim(
        base=(base_x, base_y),  # 치수선 기준점
        p1=(x1, y1),            # 첫 점
        p2=(x2, y2),            # 둘째 점
        angle=angle,            # 0°=가로, 90°=세로
        dimstyle=dim_style or "over1000dim1",
        override={
            # 문자열 + 치수값 : 예) "상판 높이 200.0"
            'dimpost': f"{text} <>" if text else "<>",
            'dimtxt': text_height,  # 텍스트 높이
            'dimdec': 1,            # 소수점 자리수 (1자리)
            'dimgap': text_gap,     # 텍스트와 치수선 사이 간격
        }
    )

    # 도면에 반영
    dim.render()
    return dim
def calculate_jb(JE, JD_plus_10):
    """
    JE (float) : 밑변 (예: 60)
    JD_plus_10 (float) : 높이 JD + 10 (예: 300)

    Returns
    -------
    (jb_pythagoras, jb_trig)
      jb_pythagoras : 피타고라스 정리를 이용해 계산한 빗변
      jb_trig       : 삼각함수를 이용해 계산한 빗변
    """

    # 1) 피타고라스 정리로 빗변 JB 구하기
    jb_pythagoras = math.sqrt(JE**2 + JD_plus_10**2)

    # 2) 삼각함수를 이용해 빗변 JB 구하기
    #    tanθ = (JD+10) / JE  ->  θ = arctan((JD+10)/JE)
    #    JB = (JE / cosθ) 또는 (JD+10) / sinθ
    theta = math.atan(JD_plus_10 / JE)  # 라디안(radian) 값
    jb_trig = JE / math.cos(theta)

    # return jb_pythagoras, jb_trig
    return jb_pythagoras
def aggregate_rows(rows_data):
    """
    rows_data: 각 행의 정보를 담은 리스트
               (예: global_data["rows_data"] 형태)

    반환값:
      동일한 (jambType, material, spec, vcut, OP, JE, JD, HH, MH, HPI_height,
              U, C, A, grounddig, poleAngle, FireDoor)
      에 대해서 floorDisplay를 콤마로 합쳐 'floorDisplay'로 두고,
      surang에 합쳐진 개수를 저장한 리스트를 반환.
    """
    global SU, HPIsurang, surang
    aggregated = {}

    SU = 0
    HPIsurang = 0
    for row_data in rows_data:
        jambType    = row_data["jambType"]
        material    = row_data["material"]
        spec        = row_data["spec"]
        vcut        = row_data["vcut"]
        OP          = row_data["OP"]
        JE          = row_data["JE"]
        JD          = row_data["JD"]
        HH          = row_data["HH"]
        MH          = row_data["MH"]
        HPI_height  = row_data["HPI_height"]
        U           = row_data["U"]
        C           = row_data["C"]
        A           = row_data["A"]
        grounddig   = row_data["grounddig"]
        # cwidth      = row_data["control_width"]
        # ctopbar     = row_data["controltopbar"]
        # cbox        = row_data["controlbox"]
        pangle      = row_data["poleAngle"]
        FireDoor    = row_data.get("FireDoor", 25)  # 기본값 25 (방화도어)

        floorDisplay = row_data["floorDisplay"]

        # floorDisplay 제외한 모든 속성으로 key 구성 (FireDoor 포함)
        key = (
            jambType, material, spec, vcut, OP, JE, JD, HH, MH,
            HPI_height, U, C, A, grounddig, pangle, FireDoor
        )

        if key not in aggregated:
            aggregated[key] = {
                "floorDisplays": [],
                "surang": 0
            }

        aggregated[key]["floorDisplays"].append(floorDisplay)
        aggregated[key]["surang"] += 1
        if jambType == '막판유' :
            HPIsurang += 1
        SU += 1

    # 이제 aggregated -> 최종 리스트 생성
    final_list = []
    for key, val in aggregated.items():
        (
            jambType, material, spec, vcut, OP, JE, JD, HH, MH,
            HPI_height, U, C, A, grounddig, pangle, FireDoor
        ) = key

        # floorDisplay를 콤마로 연결
        floorDisplays_str = ",".join(val["floorDisplays"])
        surang_value = val["surang"]

        final_list.append({
            "jambType": jambType,
            "floorDisplay": floorDisplays_str,
            "material": material,
            "spec": spec,
            "vcut": vcut,
            "OP": OP,
            "JE": JE,
            "JD": JD,
            "HH": HH,
            "MH": MH,
            "HPI_height": HPI_height,
            "U": U,
            "C": C,
            "A": A,
            "grounddig": grounddig,
            # "control_width": cwidth,
            # "controltopbar": ctopbar,
            # "controlbox": cbox,
            "poleAngle": pangle,
            "FireDoor": FireDoor,
            "surang": surang_value
        })
    return final_list
def draw_hatshape(doc, basex, basey, angle, bottomLength, topLength, height, layer="0", dim=None):
    """
    모자보강(Hat Shape) 도면을 그리는 예시 함수.
    7번 점이 기본적으로 (0,0)에 위치하도록 좌표를 정의한 뒤,
    angle만큼 시계방향 회전 후, 최종적으로 (basex, basey)에 놓습니다.

    Parameters
    ----------
    doc : ezdxf.document.Document
        ezdxf 도큐먼트 객체
    basex, basey : float
        7번 점을 최종적으로 위치시킬 기준 좌표
    angle : float
        시계방향 회전 각도(도 단위). 예: 10
    bottomLength : float
        예: 1→2, 6→7 같은 하단 길이
    topLength : float
        11→10 구간 길이
    height : float
        1→11 구간 높이
    layer : str
        도면에 사용할 레이어 이름
    """

    # -----------------------------
    # 1) 좌표 정의(각도=0, 7번=(0,0))
    # -----------------------------
    # 이 예시에선 “7번이 원점”이고,
    # 나머지 점들(1~6,8~13)을 적당히 배치.
    # 실제 도면 상황에 맞추어 조정 가능합니다.

    points = {}

    t = 1.6

    # 7번 점을 원점(0,0)
    points[6] = (0.0, 0.0)

    # 8번은 7→8이 bottomLength라면, x+ 방향에 배치
    points[7] = (0.0, t)

    # 9번은 모자 내부로 꺾이는 부분 - 대략 (someX, someY)
    points[8] = (-bottomLength + t, t)
    
    # 여기서는 임시 좌표
    points[9] = (-bottomLength + t, height )
    points[10] = (-bottomLength - topLength + t, height )
    points[11] = (-bottomLength - topLength + t, t )
    points[12] = (-bottomLength*2 - topLength + t*2, t )
    points[1] = (-bottomLength*2 - topLength + t*2, 0 )
    points[2] = (-bottomLength - topLength + t*2, 0 )
    points[3] = (-bottomLength - topLength + t*2, height - t )
    points[4] = (-bottomLength , height - t )
    points[5] = (-bottomLength , 0 )
    
    # -----------------------------
    # 2) 회전 변환 (시계방향 angle)
    #    7번점 (지금은 (0,0)) 기준으로 모든 점을 회전
    # -----------------------------
    # 시계방향 angle도 = “반시계 -angle”와 동일
    # 회전 공식 (반시계 θ):
    #   x' = x*cosθ - y*sinθ
    #   y' = x*sinθ + y*cosθ
    # 여기서는 θ = -angle (degree) → rad = -angle * π/180
    theta = math.radians(-angle)
    cos_t = math.cos(theta)
    sin_t = math.sin(theta)

    def rotate_clockwise_about_7(pt):
        # pt는 (x, y)
        x, y = pt
        # 이미 7번점은 (0,0)이므로, 
        # 굳이 ‘원점 이동 → 회전 → 원점 복귀’ 과정에서
        # 원점 이동이 필요없다.
        x_new = x * cos_t - y * sin_t
        y_new = x * sin_t + y * cos_t
        return (x_new, y_new)

    # 모든 점을 회전
    for i in points:
        points[i] = rotate_clockwise_about_7(points[i])

    # -----------------------------
    # 3) 평행이동 (7번을 (basex, basey)에 둠)
    # -----------------------------
    # 회전 후에도 7번은 (0,0) 상태 → 이를 (basex, basey)로 이동
    px7, py7 = points[6]
    shift_x = basex - px7
    shift_y = basey - py7

    for i in points:
        x0, y0 = points[i]
        points[i] = (x0 + shift_x, y0 + shift_y)

    # -----------------------------
    # 4) 선 그리기: 1~13을 원하는 순서대로 연결
    #    (line 함수는 이미 제공되었다고 가정)
    # -----------------------------
    def draw_line(idx1, idx2):
        x1, y1 = points[idx1]
        x2, y2 = points[idx2]
        line(doc, x1, y1, x2, y2, layer=layer)

    # 실제 모양에 맞게 연결 (예: 1→2, 2→6, 6→7, 7→8, 8→9, 9→10, 10→11, 11→12, 12→13, 13→1, ...)
    draw_line(1, 2)
    draw_line(2, 3)
    draw_line(3, 4)
    draw_line(4, 5)
    draw_line(5, 6)
    draw_line(6, 7)
    draw_line(7, 8)
    draw_line(8, 9)
    draw_line(9, 10)
    draw_line(10, 11)
    draw_line(11, 12)
    draw_line(12, 1)

    return points

def load_excel(row_data):
    # row_data 각 행에서 꺼내 쓸 전역 변수들
    # global jambType, floorDisplay, material, spec, vcut, OP, JE, JD, HH, MH
    # global HPI_height, U, C, A, grounddig, control_width, controltopbar, controlbox, poleAngle, surang

    jambType    = row_data["jambType"]
    floorDisplay= row_data["floorDisplay"]  # 예: "1F,2F"
    material    = row_data["material"]
    spec        = row_data["spec"]
    vcut        = row_data["vcut"]
    OP          = row_data["OP"]
    JE          = row_data["JE"]
    JD          = row_data["JD"]
    HH          = row_data["HH"]
    MH          = row_data["MH"]
    HPI_height  = row_data["HPI_height"]
    U           = row_data["U"]
    C           = row_data["C"]
    A           = row_data["A"]
    grounddig   = row_data["grounddig"]
    # control_width   = row_data["control_width"]
    # controltopbar   = row_data["controltopbar"]
    # controlbox  = row_data["controlbox"] 
    poleAngle   = row_data["poleAngle"]
    surang      = row_data["surang"]  # 중복된 행의 개수
    FireDoor    = row_data.get("FireDoor", 25)  # 기본값 25 (방화도어)

    return jambType, floorDisplay, material, spec, vcut, OP, poleAngle, JE, JD, HH, MH, HPI_height, U, C, A, grounddig, surang, FireDoor 

##########################################################################################################################################
# 와이드 작도 (막판유, 막판무) 
##########################################################################################################################################
def execute_wide(): 
    global global_data, doc, msp, SU, pageCount
    global company, drawnby, workplace, issuedate, thickness_string, HPI_Type, usage
    global HPI_punchWidth, HPI_punchHeight, HPI_holeGap
    global HPI_punchWidth_update, HPI_punchHeight_update, HPI_holeGap_update

    # row_data 각 행에서 꺼내 쓸 전역 변수들
    global jambType, floorDisplay, material, spec, vcut, OP, JE, JD, HH, MH
    global HPI_height, U, C, A, grounddig, FireDoor, control_width, controltopbar, controlbox, poleAngle, surang

    # ===================== (1) 전역 변수 기초정보 세팅 =====================
    company          = global_data["company"]
    drawnby          = global_data["drawnby"]
    workplace        = global_data["workplace"]
    issuedate        = global_data["issuedate"]
    thickness_string = global_data["thickness_string"]
    HPI_Type         = global_data["HPI_Type"]
    usage            = global_data["usage"]
    HPI_punchWidth   = global_data["HPI_punchWidth"]
    HPI_punchHeight  = global_data["HPI_punchHeight"]
    HPI_holeGap      = global_data["HPI_holeGap"]
    HPI_punchWidth_update   = global_data["HPI_punchWidth_update"]
    HPI_punchHeight_update  = global_data["HPI_punchHeight_update"]
    HPI_holeGap_update      = global_data["HPI_holeGap_update"]

    # ===================== (2) rows_data 가져오기 =====================
    rows_data = global_data["rows_data"]

    # ===================== (3) 중복 행 병합 (floorDisplay 콤마 연결, surang 계산) =====================
    merged_rows = aggregate_rows(rows_data)

    # ===================== (4) 병합된 rows를 활용해 도면 작도(또는 그 외 처리) =====================
    t = 1.5  # 두께는 1.5로 강제
    AbsX = 0

    # 막판유, 막판무 도면 그리기 시작
    for index, row_data in enumerate(merged_rows, start=1):        
        jambType, floorDisplay, material, spec, vcut, OP, poleAngle, JE, JD, HH, MH, HPI_height, U, C, A, grounddig, surang, FireDoor = load_excel(row_data)

        rx, startYpos = AbsX + index*10000,3000     
        pageCount += 1   
        # 첫페이지에 헤더표시
        if index == 1 :
            ################################################################
            # 헤더는 전체수량 적용
            # 좌표 초기화
            kk = {f'x{i}': 0 for i in range(1, 31)}
            kk.update({f'y{i}': 0 for i in range(1, 31)})

            # 테두리 설정
            rectangle(doc, 0, startYpos- 1500 + 800, 1000 + OP*1.2 + 300, startYpos- 3200, layer='0' )

            bottom_length = 40
            top_length = 25
            height = 36
            # 좌표 설정
            set_point(kk, 1, 1000, startYpos- 1500)
            set_point(kk, 2, kk['x1'] + (OP-170)   , kk['y1']   )
            set_point(kk, 3, kk['x2']   , kk['y2'] + bottom_length - 2 )
            set_point(kk, 4, kk['x3']   , kk['y3'] + 5  )
            set_point(kk, 5, kk['x4'] + 50  , kk['y4']   )
            set_point(kk, 6, kk['x5']   , kk['y5'] + 27   )
            set_point(kk, 7, kk['x6']   , kk['y6'] + top_length- 2  )
            set_point(kk, 8, kk['x7'] - OP + 70  , kk['y7']   )
            set_point(kk, 9, kk['x8']   , kk['y8'] - top_length + 2  )
            set_point(kk, 10, kk['x9']  , kk['y9'] - 27  )
            set_point(kk, 11, kk['x10'] + 50 , kk['y10']  )
            set_point(kk, 12, kk['x11'] , kk['y11'] - 5 )		

            prev_x, prev_y = kk['x1'], kk['y1']  # 첫 번째 점으로 초기화
            lastNum = 12
            for i in range(1, lastNum + 1):
                cukk_x, cukk_y = kk[f'x{i}'], kk[f'y{i}']
                line(doc, prev_x, prev_y, cukk_x, cukk_y, layer="레이져")
                prev_x, prev_y = cukk_x, cukk_y
            line(doc, prev_x, prev_y, kk['x1'], kk['y1'], layer="레이져")          

            # 절곡선 2개소
            bx1, by1, bx2,by2 = kk['x12'], kk['y12'] , kk['x3'], kk['y3'] 
            line(doc, bx1, by1, bx2,by2 , layer="절곡선")        
            bx3, by3, bx4,by4 = kk['x9'], kk['y9'] , kk['x6'], kk['y6']
            line(doc, bx3, by3, bx4,by4 , layer="절곡선")

            kk['x13'] , kk['y13'] = kk['x8'] + 15, kk['y8'] - 8
            kk['x14'] , kk['y14'] = kk['x7'] - 15, kk['y7'] - 8

            draw_circle(doc,  kk['x13'] , kk['y13'] , 11 , layer='레이져') # 11 파이
            draw_circle(doc,  kk['x14'] , kk['y14'] , 11 , layer='레이져')
            draw_circle(doc,  (kk['x1'] +  kk['x2'])/2 , kk['y1'] + 3 , 3 , layer='레이져')  # 3파이
            dim_leader(doc, (kk['x1'] +  kk['x2'])/2 , kk['y1'] + 3 , (kk['x1'] +  kk['x2'])/2 + 100 , kk['y1'] + 3 + 50, "%%c3", direction="leftToright")               
            
            # 2.3T PLATE 치수선
            d(doc, kk['x1'], kk['y1'], kk['x2'], kk['y2'], 100, direction="down", dim_style=over1000dim_style)   
            d(doc, kk['x7'], kk['y7'], kk['x2'], kk['y2'], 120, direction="right", dim_style=over1000dim_style)   
            d(doc, kk['x9'], kk['y9'], kk['x8'], kk['y8'], 160, direction="left", dim_style=over1000dim_style)   
            d(doc, kk['x12'], kk['y12'], kk['x9'], kk['y9'], 250, direction="left", dim_style=over1000dim_style)   
            d(doc, kk['x12'], kk['y12'], kk['x1'], kk['y1'], 350, direction="left", dim_style=over1000dim_style)   

            # 상부
            d(doc, kk['x13'], kk['y13'], kk['x8'], kk['y8'], 80, direction="up", dim_style=over1000dim_style)   
            string =f"OP-100 = "
            ds(doc, kk['x13'], kk['y13'], kk['x14'], kk['y14'], 80 + 8, direction="up", text= string, dim_style=over1000dim_style)   
            d(doc, kk['x14'], kk['y14'], kk['x7'], kk['y7'], 80, direction="up", dim_style=over1000dim_style)   
            string =f"OP-100+30 = "
            ds(doc, kk['x8'], kk['y8'], kk['x7'], kk['y7'], 200, direction="up", text= string, dim_style=over1000dim_style)   

            ##################### 헤더 좌측 단면도 ################################
            topLength = 25
            height = 36       
            bottomLength = 40       
            tt = 2.3 

            aa = {f'x{i}': 0 for i in range(1, 31)}
            aa.update({f'y{i}': 0 for i in range(1, 31)})        
            # 좌표 설정
            set_point(aa, 1, kk['x1'] - 700, kk['y9'] )
            set_point(aa, 2, aa['x1'] + topLength , aa['y1'] )
            set_point(aa, 3, aa['x2'] , aa['y2'] - height )
            set_point(aa, 4, aa['x3'] - bottomLength , aa['y3']  )
            set_point(aa, 5, aa['x4'] , aa['y4'] + tt  )
            set_point(aa, 6, aa['x5'] + bottomLength - tt,  aa['y5'] )		
            set_point(aa, 7, aa['x6'] ,  aa['y6'] + height - tt*2 )		
            set_point(aa, 8, aa['x7'] - topLength + tt,  aa['y7'] )		
            # 타공자리
            set_point(aa, 9, aa['x1']  + 10,  aa['y1'] + 10 )		
            set_point(aa, 10, aa['x1']  + 10,  aa['y1'] - 10 )		

            prev_x, prev_y = aa['x1'], aa['y1']  # 첫 번째 점으로 초기화
            lastNum = 8
            for i in range(1, lastNum + 1):
                cuaa_x, cuaa_y = aa[f'x{i}'], aa[f'y{i}']
                line(doc, prev_x, prev_y, cuaa_x, cuaa_y, layer="구성선")
                prev_x, prev_y = cuaa_x, cuaa_y
            line(doc, prev_x, prev_y, aa['x1'], aa['y1'], layer="구성선")  
            
            line(doc,  aa['x9'], aa['y9'], aa['x10'], aa['y10'], layer="CL")  

            d(doc, aa['x1'] , aa['y1'], aa['x2'] , aa['y2'], 250, text_height=0.20, direction="up", dim_style=over1000dim_style)                
            d(doc, aa['x9'] , aa['y9'], aa['x1'] , aa['y1'], 100, text_height=0.20, direction="up", dim_style=over1000dim_style)

            d(doc, aa['x2'] , aa['y2'], aa['x3'] , aa['y3'], 130, text_height=0.20, direction="right", dim_style=over1000dim_style)
            d(doc, aa['x3'] , aa['y3'], aa['x4'] , aa['y4'], 100, text_height=0.20, direction="down", dim_style=over1000dim_style)
            
            string = f"{SU} EA" 
            draw_Text(doc, (kk['x1'] + kk['x2'])/2 - len(string)*120/2 , kk['y1'] - 300, 120, text=string, layer='레이져')
            string = f"2.3T PLATE" 
            draw_Text(doc, (kk['x1'] + kk['x2'])/2 - len(string)*90/2 , kk['y1'] + 600, 90, text=string, layer='레이져')
            string = f"{OP} 홀헤더 보강" 
            draw_Text(doc, (kk['x1'] + kk['x2'])/2 - len(string)*90/2 , kk['y1'] + 450, 90, text=string, layer='레이져')

        ################################################################ 헤더 좌측 단면도 끝 ################################################################
        ################################################################ HPI B/K 모델별 그려주기 시작 ################################################################

            def HPIbracket(doc, model_type, x_position, y_position, OP_value, HPIsurang_value, over1000dim_style_param):
                """
                HPI B/K 모델별 그려주기 함수
                
                Parameters:
                doc: ezdxf 문서 객체
                model_type: 모델 타입 (예: 'LGS(250X70)')
                x_position: X 좌표 위치
                y_position: Y 좌표 위치 
                OP_value: OP 값
                HPIsurang_value: HPI 수량
                over1000dim_style_param: 치수 스타일
                """
                # 좌표 초기화
                oo = {f'x{i}': 0.0 for i in range(1, 31)}
                oo.update({f'y{i}': 0.0 for i in range(1, 31)})

                # 50x50 중간 타공 4.5파이 형태
                if model_type == 'LGS(250X70)' or model_type == 'ANT(254X70)':
                    # 좌표 설정            
                    set_point(oo, 1, x_position + OP_value/2 - 200, y_position - 2200)
                    set_point(oo, 2, oo['x1'] + 50 , oo['y1']   )
                    set_point(oo, 3, oo['x2']   , oo['y2'] - 50 )
                    set_point(oo, 4, oo['x3'] - 50  , oo['y3']  )		

                    prev_x, prev_y = oo['x1'], oo['y1']  # 첫 번째 점으로 초기화
                    lastNum = 4
                    for i in range(1, lastNum + 1):
                        cuoo_x, cuoo_y = oo[f'x{i}'], oo[f'y{i}']
                        line(doc, prev_x, prev_y, cuoo_x, cuoo_y, layer="레이져")
                        prev_x, prev_y = cuoo_x, cuoo_y
                    line(doc, prev_x, prev_y, oo['x1'], oo['y1'], layer="레이져")          

                    oo['x5'] , oo['y5'] = oo['x1'] + 25, oo['y1'] - 25

                    draw_circle(doc,  oo['x5'] , oo['y5'] , 4.5 , layer='레이져') # 4.5 파이
                    dim_leader(doc, oo['x5'] , oo['y5'],oo['x5'] + 50, oo['y5'] + 50,  "%%c4.5", direction="leftToright")               

                    d(doc, oo['x1'], oo['y1'], oo['x2'], oo['y2'], 60, direction="up", dim_style=over1000dim_style_param)   
                    d(doc, oo['x5'], oo['y5'], oo['x3'], oo['y3'], 120, direction="right", dim_style=over1000dim_style_param)   
                    d(doc, oo['x1'], oo['y1'], oo['x4'], oo['y4'], 120, direction="left", dim_style=over1000dim_style_param)     
                    d(doc, oo['x5'], oo['y5'], oo['x4'], oo['y4'], 150, direction="down", dim_style=over1000dim_style_param)   

                    string = f"HPI B/K" 
                    draw_Text(doc, (oo['x1'] + oo['x2'])/2 - len(string)*45/2 , oo['y1'] + 150, 90, text=string, layer='레이져')
                    string = f"EGI 2.3T {HPIsurang_value * 2}EA" 
                    draw_Text(doc, (oo['x1'] + oo['x2'])/2 - len(string)*45/2 , oo['y1'] - 400, 90, text=string, layer='레이져')

                # 60x34 갈지모형 중간 타공 5파이 팝너트 형태 3파이, 7파이 타공 (일해이엔지 EMCC반매립형과 유사함)
                if model_type == '영진(505X128)':
                    # 좌표 설정                               
                    set_point(oo, 1, x_position + OP_value/2 - 200, y_position - 2400)
                    set_point(oo, 2, oo['x1'] + 58.5 , oo['y1']   )
                    set_point(oo, 3, oo['x2'] + 31 , oo['y1']   )
                    set_point(oo, 4, oo['x3'] + 13.5 , oo['y1']   )
                    set_point(oo, 5, oo['x4'] , oo['y1'] - 128 )
                    set_point(oo, 6, oo['x5'] - 13.5  , oo['y5']  )		
                    set_point(oo, 7, oo['x6'] - 31  , oo['y5']  )		
                    set_point(oo, 8, oo['x7'] - 58.5  , oo['y5']  )		

                    prev_x, prev_y = oo['x1'], oo['y1']  # 첫 번째 점으로 초기화
                    lastNum = 8
                    for i in range(1, lastNum + 1):
                        cuoo_x, cuoo_y = oo[f'x{i}'], oo[f'y{i}']
                        line(doc, prev_x, prev_y, cuoo_x, cuoo_y, layer="레이져")
                        prev_x, prev_y = cuoo_x, cuoo_y
                    line(doc, prev_x, prev_y, oo['x1'], oo['y1'], layer="레이져")          

                    oo['x9'] , oo['y9'] = oo['x1'] + 3, oo['y1'] - 64
                    oo['x10'] , oo['y10'] = oo['x1'] + 74, oo['y1'] - 64

                    # 절곡선 그리기
                    line(doc, oo['x2'], oo['y2'], oo['x7'], oo['y7'], layer="구성선")          
                    line(doc, oo['x3'], oo['y3'], oo['x6'], oo['y6'], layer="구성선")                              

                    draw_circle(doc,  oo['x9'] , oo['y9'] , 3 , layer='레이져') # 3 파이
                    draw_circle(doc,  oo['x10'] , oo['y10'] , 7 , layer='레이져') # 7 파이
                    dim_leader(doc, oo['x9'] , oo['y9'],oo['x10'] - 250, oo['y10'] + 100,  "%%c3", direction="rightToleft")               
                    dim_leader(doc, oo['x10'] , oo['y10'],oo['x10'] + 100, oo['y9'] + 100,  "%%c7", direction="leftToright")               
                    dim_leader(doc, oo['x10'] , oo['y10'],oo['x10'] + 100, oo['y9'] - 100,  "M5 POP NUT", direction="leftToright")               

                    d(doc, oo['x1'], oo['y1'], oo['x2'], oo['y2'], 60, direction="up", dim_style=over1000dim_style_param)   
                    d(doc, oo['x3'], oo['y3'], oo['x2'], oo['y2'], 60, direction="up", dim_style=over1000dim_style_param)   
                    d(doc, oo['x3'], oo['y3'], oo['x4'], oo['y4'], 120, direction="up", dim_style=over1000dim_style_param)   
                    d(doc, oo['x1'], oo['y1'], oo['x8'], oo['y8'], 150, direction="left", dim_style=over1000dim_style_param)   
                    d(doc, oo['x8'], oo['y8'], oo['x5'], oo['y5'], 150, direction="down", dim_style=over1000dim_style_param)     

                    string = f"HPI B/K" 
                    draw_Text(doc, (oo['x1'] + oo['x2'])/2 - len(string)*45/2 , oo['y1'] + 250, 90, text=string, layer='레이져')
                    string = f"EGI 2.3T {HPIsurang_value * 2}EA" 
                    draw_Text(doc, (oo['x1'] + oo['x2'])/2 - len(string)*45/2 , oo['y1'] - 600, 90, text=string, layer='레이져')

                    # 좌표 초기화
                    # 갈지자 혀앹 단면도
                    previousX = oo['x5']
                    previousY = oo['y5']
                    oo = {f'x{i}': 0.0 for i in range(1, 31)}
                    oo.update({f'y{i}': 0.0 for i in range(1, 31)})                    
                    
                    ribT = 2.3
                    set_point(oo, 1, previousX + 500 , previousY   )
                    set_point(oo, 2, oo['x1'] + 60 , oo['y1']   )
                    set_point(oo, 3, oo['x2']  , oo['y2'] + 17  )
                    set_point(oo, 4, oo['x3'] ,  oo['y3'] + 17 - ribT )
                    set_point(oo, 5, oo['x4'] + 15 - ribT , oo['y4']  )
                    set_point(oo, 6, oo['x5'] , oo['y5'] + ribT  )		
                    set_point(oo, 7, oo['x6'] - 15  , oo['y6']  )		
                    set_point(oo, 8, oo['x7']  , oo['y7'] - 17  )		
                    set_point(oo, 9, oo['x8']  , oo['y8'] - 17 + ribT )		
                    set_point(oo, 10, oo['x9']-60 + ribT  , oo['y9']  )		

                    prev_x, prev_y = oo['x1'], oo['y1']  # 첫 번째 점으로 초기화
                    lastNum = 10
                    for i in range(1, lastNum + 1):
                        cuoo_x, cuoo_y = oo[f'x{i}'], oo[f'y{i}']
                        line(doc, prev_x, prev_y, cuoo_x, cuoo_y, layer="0")
                        prev_x, prev_y = cuoo_x, cuoo_y
                    line(doc, prev_x, prev_y, oo['x1'], oo['y1'], layer="0")          

                    # 팝너트 선 그리기
                    line(doc, oo['x8']-20, oo['y8'], oo['x3']+ 20, oo['y3'], layer="구성선")          

                    # 파이 그리기                    

                    # 치수선
                    d(doc, oo['x2'], oo['y2'], oo['x1'], oo['y1'], 150, direction="down", dim_style=over1000dim_style_param)   
                    d(doc, oo['x1'], oo['y1'], oo['x7'], oo['y7'], 150, direction="left", dim_style=over1000dim_style_param)   
                    d(doc, oo['x8'], oo['y8'], oo['x7'], oo['y7'], 50, direction="left", dim_style=over1000dim_style_param)   
                    d(doc, oo['x7'], oo['y7'], oo['x6'], oo['y6'], 150, direction="up", dim_style=over1000dim_style_param)       
        
                # 60x34 갈지모형 중간 타공 5파이 팝너트 형태 3파이, 7파이 타공 (일해이엔지 EMCC반매립형과 유사함)
                if model_type == '영진(412X83)':               
                    # 좌표 설정            
                    set_point(oo, 1, x_position + OP_value/2 - 200, y_position - 2400)
                    set_point(oo, 2, oo['x1'] + 58 , oo['y1']   )
                    set_point(oo, 3, oo['x2'] + 30 , oo['y1']   )
                    set_point(oo, 4, oo['x3'] + 13 , oo['y1']   )
                    set_point(oo, 5, oo['x4'] , oo['y1'] - 83 )
                    set_point(oo, 6, oo['x5'] - 13  , oo['y5']  )		
                    set_point(oo, 7, oo['x6'] - 30  , oo['y5']  )		
                    set_point(oo, 8, oo['x7'] - 58  , oo['y5']  )		

                    prev_x, prev_y = oo['x1'], oo['y1']  # 첫 번째 점으로 초기화
                    lastNum = 8
                    for i in range(1, lastNum + 1):
                        cuoo_x, cuoo_y = oo[f'x{i}'], oo[f'y{i}']
                        line(doc, prev_x, prev_y, cuoo_x, cuoo_y, layer="레이져")
                        prev_x, prev_y = cuoo_x, cuoo_y
                    line(doc, prev_x, prev_y, oo['x1'], oo['y1'], layer="레이져")          

                    oo['x9'] , oo['y9'] = oo['x1'] + 3, oo['y1'] - 41.5
                    oo['x10'] , oo['y10'] = oo['x1'] + 73, oo['y1'] - 41.5

                    # 절곡선 그리기
                    line(doc, oo['x2'], oo['y2'], oo['x7'], oo['y7'], layer="구성선")          
                    line(doc, oo['x3'], oo['y3'], oo['x6'], oo['y6'], layer="구성선")                              

                    draw_circle(doc,  oo['x9'] , oo['y9'] , 3 , layer='레이져') # 3 파이
                    draw_circle(doc,  oo['x10'] , oo['y10'] , 7 , layer='레이져') # 7 파이
                    dim_leader(doc, oo['x9'] , oo['y9'],oo['x10'] - 250, oo['y10'] + 100,  "%%c3", direction="rightToleft")               
                    dim_leader(doc, oo['x10'] , oo['y10'],oo['x10'] + 100, oo['y9'] + 100,  "%%c7", direction="leftToright")               
                    dim_leader(doc, oo['x10'] , oo['y10'],oo['x10'] + 100, oo['y9'] - 100,  "M5 POP NUT", direction="leftToright")               

                    d(doc, oo['x1'], oo['y1'], oo['x2'], oo['y2'], 60, direction="up", dim_style=over1000dim_style_param)   
                    d(doc, oo['x3'], oo['y3'], oo['x2'], oo['y2'], 60, direction="up", dim_style=over1000dim_style_param)   
                    d(doc, oo['x3'], oo['y3'], oo['x4'], oo['y4'], 120, direction="up", dim_style=over1000dim_style_param)   
                    d(doc, oo['x1'], oo['y1'], oo['x8'], oo['y8'], 150, direction="left", dim_style=over1000dim_style_param)   
                    d(doc, oo['x8'], oo['y8'], oo['x5'], oo['y5'], 150, direction="down", dim_style=over1000dim_style_param)     

                    string = f"HPI B/K" 
                    draw_Text(doc, (oo['x1'] + oo['x2'])/2 - len(string)*45/2 , oo['y1'] + 250, 90, text=string, layer='레이져')
                    string = f"EGI 2.3T {HPIsurang_value * 2}EA" 
                    draw_Text(doc, (oo['x1'] + oo['x2'])/2 - len(string)*45/2 , oo['y1'] - 600, 90, text=string, layer='레이져')

                    # 좌표 초기화
                    # 'z' 자 브라켓 단면도
                    previousX = oo['x5']
                    previousY = oo['y5']
                    oo = {f'x{i}': 0 for i in range(1, 31)}
                    oo.update({f'y{i}': 0 for i in range(1, 31)})                    
                    
                    ribT = 2.3
                    set_point(oo, 1, previousX + 500 , previousY   )
                    set_point(oo, 2, oo['x1'] + 60 , oo['y1']   )
                    set_point(oo, 3, oo['x2']  , oo['y2'] + 17  )
                    set_point(oo, 4, oo['x3'] ,  oo['y3'] + 17 - ribT )
                    set_point(oo, 5, oo['x4'] + 15 - ribT , oo['y4']  )
                    set_point(oo, 6, oo['x5'] , oo['y5'] + ribT  )		
                    set_point(oo, 7, oo['x6'] - 15  , oo['y6']  )		
                    set_point(oo, 8, oo['x7']  , oo['y7'] - 17  )		
                    set_point(oo, 9, oo['x8']  , oo['y8'] - 17 + ribT )		
                    set_point(oo, 10, oo['x9']-60 + ribT  , oo['y9']  )		

                    prev_x, prev_y = oo['x1'], oo['y1']  # 첫 번째 점으로 초기화
                    lastNum = 10
                    for i in range(1, lastNum + 1):
                        cuoo_x, cuoo_y = oo[f'x{i}'], oo[f'y{i}']
                        line(doc, prev_x, prev_y, cuoo_x, cuoo_y, layer="0")
                        prev_x, prev_y = cuoo_x, cuoo_y
                    line(doc, prev_x, prev_y, oo['x1'], oo['y1'], layer="0")          

                    # 팝너트 선 그리기
                    line(doc, oo['x8']-20, oo['y8'], oo['x3']+ 20, oo['y3'], layer="구성선")          

                    # 파이 그리기                    

                    # 치수선
                    d(doc, oo['x2'], oo['y2'], oo['x1'], oo['y1'], 150, direction="down", dim_style=over1000dim_style_param)   
                    d(doc, oo['x1'], oo['y1'], oo['x7'], oo['y7'], 150, direction="left", dim_style=over1000dim_style_param)   
                    d(doc, oo['x8'], oo['y8'], oo['x7'], oo['y7'], 50, direction="left", dim_style=over1000dim_style_param)   
                    d(doc, oo['x7'], oo['y7'], oo['x6'], oo['y6'], 150, direction="up", dim_style=over1000dim_style_param)       

                # 'ㄴ'자 타공4개의 브라켓 
                if model_type == 'MTK(302X62)':               
                    # 좌표 설정            
                    set_point(oo, 1, x_position + OP_value/2 - 200, y_position - 2400)
                    set_point(oo, 2, oo['x1'] + 100 , oo['y1']   )
                    set_point(oo, 3, oo['x2']  , oo['y2'] + 38  )
                    set_point(oo, 4, oo['x3']  , oo['y3'] + 48  )
                    set_point(oo, 5, oo['x4'] - 100, oo['y4'] )
                    set_point(oo, 6, oo['x5']  , oo['y5'] - 48 )			

                    prev_x, prev_y = oo['x1'], oo['y1']  # 첫 번째 점으로 초기화
                    lastNum = 6
                    for i in range(1, lastNum + 1):
                        cuoo_x, cuoo_y = oo[f'x{i}'], oo[f'y{i}']
                        line(doc, prev_x, prev_y, cuoo_x, cuoo_y, layer="레이져")
                        prev_x, prev_y = cuoo_x, cuoo_y
                    line(doc, prev_x, prev_y, oo['x1'], oo['y1'], layer="레이져")          

                    # 절곡선 그리기
                    line(doc, oo['x6'], oo['y6'], oo['x3'], oo['y3'], layer="구성선")                              

                    # 단공 및 슬롯 그리기   좌표설정
                    set_point(oo, 7, oo['x1'] + 19, oo['y6'])
                    set_point(oo, 8, oo['x2'] - 19, oo['y3'])
                    set_point(oo, 9, oo['x1'] + 27.5, oo['y1'] + 20)
                    set_point(oo, 10, oo['x2'] - 27.5, oo['y1'] + 20)

                    # 단공 슬롯 그리기
                    draw_circle(doc,  oo['x7'] , oo['y7'] , 3 , layer='레이져') # 3 파이
                    draw_circle(doc,  oo['x8'] , oo['y8'] , 3 , layer='레이져') # 3 파이
                    draw_slot(doc, oo['x9'] , oo['y9'], "7x14", direction="세로", option="cross", layer='레이져')                    
                    draw_slot(doc, oo['x10'] , oo['y10'], "7x14", direction="세로", option="cross", layer='레이져')                    
                    
                    dim_leader(doc, oo['x8'] , oo['y8'],oo['x8'] + 150, oo['y8'] + 100,  "%%c3", direction="leftToright")                
                    dim_leader(doc, oo['x10'] , oo['y10'],oo['x10'] + 150, oo['y9'] - 100,  "slot %%c7X14", direction="leftToright")               

                    d(doc, oo['x9'], oo['y9'], oo['x10'], oo['y10'], 150, direction="down", dim_style=over1000dim_style_param)   
                    d(doc, oo['x1'], oo['y1'], oo['x2'], oo['y2'], 200, direction="down", dim_style=over1000dim_style_param)   
                    d(doc, oo['x2'], oo['y2'], oo['x4'], oo['y4'], 100, direction="right", dim_style=over1000dim_style_param)   
                    d(doc, oo['x6'], oo['y6'], oo['x1'], oo['y1'], 80, direction="left", dim_style=over1000dim_style_param)   
                    d(doc, oo['x6'], oo['y6'], oo['x5'], oo['y5'], 80, direction="left", dim_style=over1000dim_style_param)   
                    # 단공 간격 치수선
                    d(doc, oo['x7'], oo['y7'], oo['x8'], oo['y8'], 120, direction="up", dim_style=over1000dim_style_param)   

                    string = f"HPI B/K" 
                    draw_Text(doc, (oo['x1'] + oo['x2'])/2 - len(string)*45/2 , oo['y1'] + 250, 90, text=string, layer='레이져')
                    string = f"EGI 2.3T {HPIsurang_value * 2}EA" 
                    draw_Text(doc, (oo['x1'] + oo['x2'])/2 - len(string)*45/2 , oo['y1'] - 600, 90, text=string, layer='레이져')

                    # 좌표 초기화
                    # 'ㄴ' 자 브라켓 단면도
                    previousX = oo['x1']
                    previousY = oo['y1']
                    oo = {f'x{i}': 0 for i in range(1, 31)}
                    oo.update({f'y{i}': 0 for i in range(1, 31)})                    
                    
                    ribT = 2.3
                    set_point(oo, 1, previousX - 400 , previousY   )
                    set_point(oo, 2, oo['x1'] + 40 , oo['y1']   )
                    set_point(oo, 3, oo['x2']  , oo['y2'] + 50  )
                    set_point(oo, 4, oo['x3'] - ribT,  oo['y3']  )
                    set_point(oo, 5, oo['x4'] , oo['y4'] - 50 + ribT )
                    set_point(oo, 6, oo['x5']  - 40 + ribT, oo['y5']  )			

                    prev_x, prev_y = oo['x1'], oo['y1']  # 첫 번째 점으로 초기화
                    lastNum = 6
                    for i in range(1, lastNum + 1):
                        cuoo_x, cuoo_y = oo[f'x{i}'], oo[f'y{i}']
                        line(doc, prev_x, prev_y, cuoo_x, cuoo_y, layer="0")
                        prev_x, prev_y = cuoo_x, cuoo_y
                    line(doc, prev_x, prev_y, oo['x1'], oo['y1'], layer="0")          

                    # 치수선
                    d(doc, oo['x2'], oo['y2'], oo['x1'], oo['y1'], 150, direction="down", dim_style=over1000dim_style_param)   
                    d(doc, oo['x2'], oo['y2'], oo['x3'], oo['y3'], 150, direction="right", dim_style=over1000dim_style_param)   
                    
                    # HPI B/K 모델별 그려주기, 도면의 헤더 도면 아래 위치함
            # HPIbracket 함수 호출
            HPIbracket(doc, HPI_Type, 1000, startYpos, OP, HPIsurang, over1000dim_style)

        ################################################################ HPI B/K 모델별 그려주기 끝 ################################################################
        ################################################################ 쟘기둥 단면도 시작 ################################################################
        SW = FireDoor # 방화도어 25, 일반도어 50
        SBW = 10 # 출입구쪽 기둥 뒷날개
        # 좌표 초기화
        a = {f'x{i}': 0 for i in range(1, 31)}
        a.update({f'y{i}': 0 for i in range(1, 31)})

        # 막판무, 막판유 차이는 JD+10, JD+5 차이 설계 상판 평면도 좌측 기둥 형상
        # 다완테크는 막판유는 JD가 기둥을 평면으로 봤을때 길이(실제 상판은 - 10 적용, JD + 5를 적용해야 맞다.
        # 상판이 기둥보다 돌출값을 addSpace로 설정
        # 기둥을 그리는 관점에서 치수 조정함 (막판유는 막판에서 10mm 돌출, 판무는 기둥에서 5mm 돌출(상판이 더 큼))        
        if jambType=='막판유':
            addSpace = 10
        elif jambType == '막판무':
            addSpace = 5
        # 막판유 상판 조립도 그리기
        ############################################################## 막판유 상판 조립도 그리기 시작 ##############################################################
        if jambType=='막판유':
            # 막판유 상판 조립도 좌측의 기둥 단면도 표시 좌표 설정
            # Side의 JE는 JD값으로 새로 계산함            
            # SideJE = JE
            SideJE = (JD + 10 ) * math.tan(math.radians(poleAngle))
            set_point(a, 1, rx, startYpos)
            set_point(a, 2, a['x1'] - t, a['y1'])
            set_point(a, 3, a['x2'], a['y2'] - A)
            set_point(a, 4, a['x3'] + C, a['y3'])
            set_point(a, 5, a['x4'] + SideJE, a['y4'] + JD + addSpace)  # 기둥이 JD보다  들어간 치수 (막판유, 막판무 차이 있음 주의)
            set_point(a, 6, a['x5'] - SW, a['y5'])
            set_point(a, 7, a['x6'], a['y6'] - addSpace)
            set_point(a, 8, a['x7'] + t, a['y7'])
            set_point(a, 9, a['x8'], a['y8'] + addSpace - t)
            set_point(a, 10, a['x9'] + SW - t * 2, a['y9'])
            set_point(a, 11, a['x10'] - SideJE , a['y10'] - JD - addSpace + t * 2)
            set_point(a, 12, a['x11'] - C + t * 2, a['y11'])

            prev_x, prev_y = a['x1'], a['y1']  # 첫 번째 점으로 초기화
            lastNum = 12
            for i in range(1, lastNum + 1):
                curr_x, curr_y = a[f'x{i}'], a[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="0")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, a['x1'], a['y1'], layer="0")

            # 상판 평면도 상판 가상 형상 좌표 설정
            b = {f'x{i}': 0 for i in range(1, 31)}
            b.update({f'y{i}': 0 for i in range(1, 31)})

            # 높이와 각도로 밑면 계산
            height = JD - t
            base = calculate_base(height, poleAngle) # JD에 대한 밑면을 계산하니, JE를 적용하지말고 base를 적용해야 한다.

            # print(f"각도: {poleAngle}, 치수: {base:.1f}")

            # 막판유 상판 조립도 상판 본판의 좌표 설정
            set_point(b, 1, a['x5'], a['y5'])
            set_point(b, 2, b['x1'] + OP, b['y1'])
            set_point(b, 3, b['x2'] + base, b['y2'] - JD + t)
            set_point(b, 4, b['x3'] - OP - base * 2, b['y3'])
            
            prev_x, prev_y = b['x1'], b['y1']  # 첫 번째 점으로 초기화
            lastNum = 4
            for i in range(1, lastNum + 1):
                curr_x, curr_y = b[f'x{i}'], b[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="0")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, b['x1'], b['y1'], layer="0")   
            rectangle(doc,b['x3'], b['y3'], b['x4'], b['y4'] - t, layer='0')     

            ############################################################## 막판유 상판 조립도 그리기 끝 ##############################################################
            ############################################################## 막판무 막판유 상판 조립도  ##############################################################
            # 막판유 좌기둥 좌표 초기화
            c = {f'x{i}': 0 for i in range(1, 31)}
            c.update({f'y{i}': 0 for i in range(1, 31)})
             
            set_point(c, 1, a['x5'] + OP, a['y5'])
            set_point(c, 2, c['x1'] + SW, c['y1'])
            set_point(c, 3, c['x2'], c['y2'] - addSpace)
            set_point(c, 4, c['x3'] - t, c['y3'])
            set_point(c, 5, c['x4'], c['y4'] + addSpace - t)
            set_point(c, 6, c['x5'] - SW + t * 2, c['y5'])
            set_point(c, 7, c['x6'] + SideJE, c['y6'] - JD - addSpace + t * 2)
            set_point(c, 8, c['x7'] + C - t * 2, c['y7'])
            set_point(c, 9, c['x8'], c['y8'] + A - t)
            set_point(c, 10, c['x9'] + t, c['y9'])
            set_point(c, 11, c['x10'], c['y10'] - A)
            set_point(c, 12, c['x11'] - C, c['y11'])

            prev_x, prev_y = c['x1'], c['y1']  # 첫 번째 점으로 초기화
            lastNum = 12
            for i in range(1, lastNum + 1):
                curr_x, curr_y = c[f'x{i}'], c[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="0")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, c['x1'], c['y1'], layer="0")              

            # 상판 조립 치수선 표시
            # 하부 표시
            d(doc, b['x3'], b['y3'], b['x4'], b['y4'], 100, direction="down", dim_style=over1000dim_style)        
            d(doc, a['x3'], a['y3'], c['x11'], c['y11'], 200, direction="down", dim_style=over1000dim_style)        
            # 상부 표시
            d(doc, b['x1'], b['y1'], b['x2'], b['y2'], 80, direction="up", dim_style=over1000dim_style)        
            # 우측 표시
            d(doc, c['x2'], c['y2'], c['x11'], c['y11'], 150, direction="right", dim_style=over1000dim_style)        
            d(doc, (b['x1'] + b['x2'])/2, b['y1'], (b['x1'] + b['x2'])/2 , b['y3'] - t , 100, direction="right", dim_style=over1000dim_style)        

            string = f"{floorDisplay} 상판" 
            draw_Text(doc, (b['x1'] + b['x2'])/2 - len(string)*80/2 , b['y3'] - 400, 80, text=string)

            # 좌표 초기화
            aa = {f'x{i}': 0 for i in range(1, 31)}
            aa.update({f'y{i}': 0 for i in range(1, 31)})

            ########################################## 막판유 상판 전개도(가공품) 좌표 설정 시작 ##########################################
            # 막판유 상판 전개도(가공품) 좌표 설정
            set_point(aa, 1, (b['x1'] + b['x2'])/2 - OP/2 - JE + 10 , a['y2'] -  850)
            set_point(aa, 2, aa['x1'] + OP + JE*2 - 10*2, aa['y1'])
            set_point(aa, 3, aa['x2'] + 10 , aa['y2'] - 25 + t)
            set_point(aa, 4, aa['x3'] , aa['y3'] - MH + t*2)
            set_point(aa, 5, aa['x4'] - JE , aa['y4'] - JD + t*2 )
            set_point(aa, 6, aa['x5'] , aa['y5'] - 15 + t)
            set_point(aa, 7, aa['x6'] - OP, aa['y6'] )
            set_point(aa, 8, aa['x7'] , aa['y7'] + 15 - t)
            set_point(aa, 9, aa['x8'] - JE , aa['y8'] + JD - t*2)
            set_point(aa, 10, aa['x9'] , aa['y9'] + MH - t*2)

            prev_x, prev_y = aa['x1'], aa['y1']  # 첫 번째 점으로 초기화
            lastNum = 10
            for i in range(1, lastNum + 1):
                curr_x, curr_y = aa[f'x{i}'], aa[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="레이져")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, aa['x1'], aa['y1'], layer="레이져")     

            # 발주처에 '영진'인 경우는 105파이 JD부위 센터에 타공 후 치수선 그려주기
            if '영진' in company:
                # 장공 위치 좌표설정
                centerX, centerY = (aa['x4'] + aa['x9'])/2, (aa['y4'] + aa['y5'])/2
                draw_circle(doc, centerX, centerY, 105, layer="레이져")
                dim_leader(doc, centerX, centerY, centerX, centerY + 105, "%%c105", direction="leftToright")                
                # 치수선 우측에 2개
                d(doc, (aa['x9'] + aa['x4'])/2, aa['y4'], centerX, centerY , 280, direction="right", dim_style=over1000dim_style)                
                d(doc, (aa['x6'] + aa['x7'])/2, aa['y5'], centerX, centerY , 280, direction="right", dim_style=over1000dim_style)                


            # 상판의 중심위치 좌표 설정
            centerX =  (aa['x6'] +  aa['x7'])/2
            # 상판 중심 3파이 타공 
            draw_circle(doc, centerX , aa['y7'] + 3, 3 , layer='레이져', color='7')
            # 절곡선
            line(doc, aa['x10'], aa['y10'], aa['x3'], aa['y3'], layer="절곡선")      
            line(doc, aa['x9'], aa['y9'], aa['x4'], aa['y4'], layer="절곡선")      
            line(doc, aa['x8'], aa['y8'], aa['x5'], aa['y5'], layer="절곡선")      

            # 치수선
            d(doc, aa['x1'] , aa['y1'], aa['x7'] , aa['y7'], 150, direction="left", dim_style=over1000dim_style)        
            d(doc, aa['x7'] , aa['y7'], aa['x9'] , aa['y9'], 100, direction="down", dim_style=over1000dim_style)        
            d(doc, aa['x6'] , aa['y6'], aa['x4'] , aa['y4'], 100, direction="down", dim_style=over1000dim_style)        
            d(doc, aa['x7'] , aa['y7'], aa['x6'] , aa['y6'], 100, direction="down", dim_style=over1000dim_style)        
            # 우측치수선
            d(doc, aa['x5'] , aa['y5'], aa['x6'] , aa['y6'], extract_abs(aa['x4'],aa['x5']) + 300, direction="right", dim_style=over1000dim_style)        
            d(doc, aa['x5'] , aa['y5'], aa['x4'] , aa['y4'], extract_abs(aa['x4'],aa['x5']) + 300, direction="right", dim_style=over1000dim_style)        
            d(doc, aa['x4'] , aa['y4'], aa['x3'] , aa['y3'], extract_abs(aa['x4'],aa['x5']) + 300, direction="right", dim_style=over1000dim_style)        
            d(doc, aa['x3'] , aa['y3'], aa['x2'] , aa['y2'], 400, direction="right", dim_style=over1000dim_style)        
            # 상부 치수선
            d(doc, aa['x1'] , aa['y1'], aa['x10'] , aa['y10'], 50, direction="up", dim_style=over1000dim_style)        
            d(doc, aa['x2'] , aa['y2'], aa['x3'] , aa['y3'], 50, direction="up", dim_style=over1000dim_style)        
            d(doc, aa['x10'] , aa['y10'], aa['x3'] , aa['y3'], 230, direction="up", dim_style=over1000dim_style)        

            # 시작점 JE OP/4 적용 마지막 
            lastNum = 5
            GapSum = 0
            for i in range(1, lastNum + 1):
                # 8x16 장공
                Gap = OP/4
                insert_block(doc, aa['x10'] + JE + GapSum, aa['y10'] + 11.5, "8x16_vertical", layer="레이져")
                if i==1 :
                    d(doc, aa['x10'] , aa['y10'], aa['x10'] + JE + GapSum , aa['y10'] + 11.5 , 120, direction="up", dim_style=over1000dim_style, option='reverse')        
                else:
                    dc(doc, aa['x10'] + JE + GapSum , aa['y10'] + 11.5)
                    lastX, lastY = aa['x10'] + JE + GapSum , aa['y10'] + 11.5
                GapSum = GapSum + Gap            

            dc(doc,  aa['x3'] , aa['y3'])
            d(doc, lastX, lastY , aa['x2'] , aa['y2'],  200, direction="right", dim_style=over1000dim_style )
            d(doc, lastX, lastY , aa['x3'] , aa['y3'],  250, direction="right", dim_style=over1000dim_style )


######################################################## HPI 전개도 홀타공 시작 ###################################################################################            
            # HPI 타공홀 HPI홀타공 그리기(일반적 타공만 있어서 함수화가 필요함)
            if HPI_Type not in ('', None):
                if HPI_punchWidth_update in ('', None):
                    HPI_punchWidth_update = HPI_punchWidth

                if HPI_punchHeight_update in ('', None):
                    HPI_punchHeight_update = HPI_punchHeight

                if HPI_holeGap_update in ('', None):
                    HPI_holeGap_update = HPI_holeGap

                centerX, centerY = aa['x7'] + OP/2, aa['y9'] + HPI_height + HPI_punchHeight_update/2 - t
                # 좌표 초기화
                bb = {f'x{i}': 0 for i in range(1, 31)}
                bb.update({f'y{i}': 0 for i in range(1, 31)})

                # 좌표 설정
                set_point(bb, 1, centerX-HPI_punchWidth_update/2, centerY )
                set_point(bb, 2, centerX+HPI_punchWidth_update/2, centerY )
                set_point(bb, 3, centerX+HPI_punchWidth_update/2, centerY-HPI_punchHeight_update )
                set_point(bb, 4, centerX-HPI_punchWidth_update/2, centerY-HPI_punchHeight_update )
                prev_x, prev_y = bb['x1'], bb['y1']  # 첫 번째 점으로 초기화
                lastNum = 4
                for i in range(1, lastNum + 1):
                    curr_x, curr_y = bb[f'x{i}'], bb[f'y{i}']
                    line(doc, prev_x, prev_y, curr_x, curr_y, layer="레이져")
                    prev_x, prev_y = curr_x, curr_y
                line(doc, prev_x, prev_y, bb['x1'], bb['y1'], layer="레이져")

                # 치수선
                d(doc,  bb['x2'], centerY-HPI_punchHeight_update/2 , bb['x2'],  aa['y9'] , 100, direction="right", dim_style=over1000dim_style )
                d(doc,  bb['x1'],  bb['y1'],  bb['x2'],  bb['y2'] , 40, direction="up", dim_style=over1000dim_style )
                d(doc,  bb['x1'],  bb['y1'],  bb['x4'],  bb['y4'] , 60, direction="left", dim_style=over1000dim_style )

                # HPI 주변홀 타공이나 치수선 추가로 모델별로 그려주기                
                # 특정 노출형 HPI 처리
                if "성광(310X75)" in HPI_Type:
                    X1, X2 = centerX - HPI_holeGap_update / 2, centerX + HPI_holeGap_update / 2
                    Y1 = centerY - HPI_punchHeight_update / 2 
                    Y2 = Y1        
                    X3 = X1 + 22.5
                    X4 = X2 - 22.5             

                    draw_circle(doc, X1, Y1, 7, layer="레이져")
                    draw_circle(doc, X2, Y2, 7, layer="레이져")
                    draw_circle(doc, X3, Y2, 7, layer="레이져")
                    draw_circle(doc, X4, Y2, 7, layer="레이져")
                    d(doc, X3 , Y1, X4 , Y2,  170, direction="down")   
                    d(doc, X1 , Y1, X2 , Y2,  270, direction="down")                    
                    dim_leader(doc,  X2 , Y1  , X2 + 100, Y1 + 50, "4-%%c7", direction="leftToright")                                

                if "구룡산전(170X61)" in HPI_Type:
                    X1, X2 = centerX - HPI_holeGap_update / 2, centerX + HPI_holeGap_update / 2
                    Y1 = centerY - HPI_punchHeight_update / 2 
                    Y2 = Y1        
                    draw_circle(doc, X1, Y1, 7, layer="레이져")
                    draw_circle(doc, X2, Y2, 7, layer="레이져")
                    d(doc, X1 , Y1, X2 , Y2,  270, direction="down")                    
                    dim_leader(doc,  X2 , Y1  , X2 + 100, Y1 + 50, "2-%%c7", direction="leftToright")                                

                if "승강기블루(290X85)" in HPI_Type or "지앤디(290X85)" in HPI_Type or "KB산업(290X85)" in HPI_Type:
                    X1, X2 = centerX - HPI_holeGap_update / 2, centerX + HPI_holeGap_update / 2
                    Y1 = centerY - HPI_punchHeight_update / 2 
                    Y2 = Y1        
                    draw_circle(doc, X1, Y1, 7, layer="레이져")
                    draw_circle(doc, X2, Y2, 7, layer="레이져")
                    d(doc, X1 , Y1, X2 , Y2,  150, direction="up")                    
                    dim_leader(doc,  X2 , Y1  , X2 + 100, Y1 + 50, "2-%%c7", direction="leftToright")                                

                if "서명(170X50)" in HPI_Type :
                    X1, X2 = centerX - HPI_holeGap_update / 2, centerX + HPI_holeGap_update / 2
                    Y1 = centerY - HPI_punchHeight_update / 2 
                    Y2 = Y1        
                    draw_slot(doc, X1 , Y1, "7x14", direction="가로", option="cross", layer='레이져')
                    draw_slot(doc, X2 , Y2, "7x14", direction="세로", option="cross", layer='레이져')                    
                    d(doc, X1 , Y1, X2 , Y2,  150, direction="up")                    
                    dim_leader(doc,  X2 , Y1  , X2 + 100, Y1 + 50, "2-7x14", direction="leftToright")                                        
            
######################################################## HPI 전개도 홀타공 끝 ###################################################################################            
            # 상판 각도표시 4개소 aa변수 사용
            # 상판 좌측 하단
            dim_angular(doc,   aa['x9'], aa['y9'], aa['x8'] , aa['y8'],  aa['x8'] , aa['y8'],  aa['x8'] + 200, aa['y8'],  20, direction="up" )   # up,down, left, right각도를 표시할때는 기존치수선 연장선에 해줘야 제대로 값을 얻는다        
            # 상판 우측 하단
            dim_angular(doc,   aa['x5'], aa['y5'], aa['x5'] - 100 , aa['y5'],  aa['x5'] , aa['y5'],  aa['x4'] , aa['y4'],  20, direction="up" )   # up,down, left, right각도를 표시할때는 기존치수선 연장선에 해줘야 제대로 값을 얻는다        
            # 상판 JD 우측 상단
            dim_angular(doc,   aa['x5'], aa['y5'], aa['x4']  , aa['y4'],  aa['x4'] , aa['y4'],  aa['x4']-100 , aa['y4'],  70, direction="left" )   # up,down, left, right각도를 표시할때는 기존치수선 연장선에 해줘야 제대로 값을 얻는다        
            # 상판 우측 하단
            dim_angular(doc,   aa['x9'] + 100, aa['y9'], aa['x9']  , aa['y9'],  aa['x9'] , aa['y9'],  aa['x8'] , aa['y8'],  70, direction="right" )   # up,down, left, right각도를 표시할때는 기존치수선 연장선에 해줘야 제대로 값을 얻는다        
            
        # 막판무 상판 조립도 그리기            
        if jambType=='막판무':
            ##############################################################
            # 막판무 상판 조립도 (상판 + 기둥형상)
            ##############################################################
            # 높이와 각도로 밑면 계산 막판무 밑면 높이는  다완의 JD가 5mm 크다. 이 각도 보정해야 함. JD가 기준이다.
            height = JD  + 5
            base = calculate_base(height, poleAngle) # JD에 대한 밑면을 계산하니, JE를 적용하지말고 base를 적용해야 한다.            
            set_point(a, 1, rx, startYpos )
            set_point(a, 2, a['x1'] - t, a['y1'])
            set_point(a, 3, a['x2'], a['y2'] + A)
            set_point(a, 4, a['x3'] + C, a['y3'])
            set_point(a, 5, a['x4'] + base, a['y4'] - JD ) 
            set_point(a, 6, a['x5'] - SW, a['y5'])
            set_point(a, 7, a['x6'], a['y6'] + SBW)
            set_point(a, 8, a['x7'] + t, a['y7'])
            set_point(a, 9, a['x8'], a['y8'] - SBW + t)
            set_point(a, 10, a['x9'] + SW - t * 2, a['y9'])
            set_point(a, 11, a['x10'] - base, a['y10'] + JD - t * 2 )
            set_point(a, 12, a['x11'] - C + t * 2, a['y11'])

            prev_x, prev_y = a['x1'], a['y1']  # 첫 번째 점으로 초기화
            lastNum = 12
            for i in range(1, lastNum + 1):
                curr_x, curr_y = a[f'x{i}'], a[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="0")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, a['x1'], a['y1'], layer="0")

            # 상판 평면도 상판 형상
            b = {f'x{i}': 0 for i in range(1, 31)}
            b.update({f'y{i}': 0 for i in range(1, 31)})           

            # 높이와 각도로 밑면 계산            
            base = calculate_base(height, poleAngle) # JD에 대한 밑면을 계산하니, JE를 적용하지말고 base를 적용해야 한다.

            height = JD  # 상판의 각도
            baseTop = calculate_base(height, poleAngle) # JD에 대한 밑면을 계산하니, JE를 적용하지말고 base를 적용해야 한다.

            # print(f"각도: {poleAngle}, 치수: {base:.1f}")

            # 좌표 설정
            # 막판무 상판 좌표 설정
            set_point(b, 1, a['x5'], a['y5'])
            set_point(b, 2, b['x1'] + OP, b['y1'])
            set_point(b, 3, b['x2'] + baseTop, b['y2'] + JD + 5)
            set_point(b, 4, b['x3'] - OP - baseTop * 2, b['y3'])
            
            # 좌표 초기화
            c = {f'x{i}': 0 for i in range(1, 31)}
            c.update({f'y{i}': 0 for i in range(1, 31)})

            # 막판무 조립도 우기둥 좌표 설정
            set_point(c, 1, a['x5'] + OP, a['y5'])
            set_point(c, 2, c['x1'] + SW, c['y1'])
            set_point(c, 3, c['x2'], c['y2'] + SBW)
            set_point(c, 4, c['x3'] - t, c['y3'])
            set_point(c, 5, c['x4'], c['y4'] - SBW + t)
            set_point(c, 6, c['x5'] - SW + t * 2, c['y5'])
            set_point(c, 7, c['x6'] + base, c['y6'] + JD - t * 2)
            set_point(c, 8, c['x7'] + C - t * 2, c['y7'])
            set_point(c, 9, c['x8'], c['y8'] - A + t)
            set_point(c, 10, c['x9'] + t, c['y9'])
            set_point(c, 11, c['x10'], c['y10'] + A)
            set_point(c, 12, c['x11'] - C, c['y11'])

            prev_x, prev_y = c['x1'], c['y1']  # 첫 번째 점으로 초기화
            lastNum = 12
            for i in range(1, lastNum + 1):
                curr_x, curr_y = c[f'x{i}'], c[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="0")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, c['x1'], c['y1'], layer="0")              

            # 상부 ㄱ자 연결 좌우
            line(doc, a['x3'], a['y3']+5 ,c['x11'], c['y11'] + 5 , layer="0")               
            line(doc, a['x3'], a['y3'], a['x3'], a['y3'] + 5 , layer="0")               
            line(doc, c['x11'], c['y11'] , c['x11'], c['y11'] + 5 , layer="0")     
            # OP 연결          
            line(doc, a['x5'], a['y5'] , c['x1'], c['y1'] , layer="0")    


            # 막판무 상판 조립시 ㄱ자 보강연결부분 그려주기
            xGap1 = 16.9 # 첫번째 장공X 위치
            xGap2 = 30 # 두번째 보강테두리 위치
            line(doc, a['x4']-xGap1 , a['y4']  , a['x5'] - xGap1, a['y5'] , layer="구성선") 
            line(doc, a['x4']-xGap2 , a['y4']  , a['x5'] - xGap2, a['y5'] , layer="구성선") 

            # JD 330 - 312 = 18 줄어듬
            # JD 방향 장공 - 시뮬레이션 함수로 정확한 52mm 위치 계산
            if(JD >= 600) :
                hole_count = 4
            elif (JD >= 900) :
                hole_count = 5
            else:
                hole_count = 3
                
            # 새로운 시뮬레이션 함수 사용: 첫 번째 홀이 정확히 52mm 위치에 오도록 계산
            # 상부떨어지는 값을 계산한다.
            # JD_upperGap 계산 공식:
            # JD + 5가 100일 때 1.5, 50일 때 0.75, 400일 때 1.5가 되도록 한다.
            # 즉, (JD + 5) / 100 * 1.5
            JD_upperGap = 0
            if JD >= 300:
                JD_upperGap = (JD - 300) * 0.015

            JDholes = simulate_hole_positions_from_bottom(
                totalLength=JD - 4.5 - 1.1 + JD_upperGap,
                target_first_distance=52.0,
                hole_count=hole_count,
                poleAngle=poleAngle
            )

            # print(f" JDholes: {JDholes}")
            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(JDholes):
                # 변화에 따른 x위치 삼각함수 적용                
                value = calculate_base(hole + 11.5, poleAngle) 
                Heightchange = calculate_base(hole + 11.5, poleAngle, option="height") 
                Heightchange = round(Heightchange, 2)
                xpos, ypos = a['x5'] - xGap1 - value  , a['y5'] + Heightchange
                insert_block(doc, xpos, ypos, "8x20_horizontal", layer="0")     
                if i == 0 :
                    dim_linear(doc, xpos, ypos,  a['x5'] , a['y5'], "", 100 + i*60,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 안쪽 치수선
                    firstx, firsty = xpos, ypos
                else:
                    dim_linear(doc, lastx, lasty , xpos, ypos, "", dis=100,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 안쪽 치수선
                lastx, lasty = xpos, ypos                 
                

            # 상판 조립 평면도 치수선
            # 하부 표시
            d(doc, a['x5'], a['y5'], c['x1'], c['y1'], 100, direction="down", dim_style=over1000dim_style)        
            d(doc, a['x6'], a['y6'], a['x5'], c['y5'], 200, direction="down", dim_style=over1000dim_style)        
            # 좌측
            # d(doc, a['x7'], a['y7'], a['x6'], a['y6'], 150, direction="left", dim_style=over1000dim_style)        
            d(doc, a['x2'], a['y2'], a['x3'], a['y3'], 300, direction="left", dim_style=over1000dim_style)        
            # 우측 표시
            d(doc, c['x11'], c['y11'], c['x2'], c['y2'], 220, direction="right", dim_style=over1000dim_style)        
            d(doc, a['x3'] + OP/2, a['y3'] + 5 , a['x3'] + OP/2 , a['y5'], 100, direction="right", dim_style=over1000dim_style)        
            # 상부 표시
            d(doc, a['x3'], a['y3'], c['x11'], c['y11'], 150, direction="up", dim_style=over1000dim_style)    

            # 추가 치수선     
            d(doc, c['x1'], c['y1'], c['x11'], c['y11'], 200, direction="left", dim_style=over1000dim_style)        
            line(doc, a['x2'] , lasty , c['x11'] , lasty , layer="구성선")             
            d(doc,  a['x3'], a['y3'], lastx, lasty ,300, direction="right", dim_style=over1000dim_style)        
            d(doc,  a['x3'], a['y3'] + 5, lastx, lasty , 400, direction="right", dim_style=over1000dim_style)        

            string = f"{floorDisplay} 상판" 
            draw_Text(doc, (b['x1'] + b['x2'])/2 - len(string)*80/2 , b['y3'] - 700, 80, text=string)

            # 좌표 초기화
            aa = {f'x{i}': 0 for i in range(1, 31)}
            aa.update({f'y{i}': 0 for i in range(1, 31)})

            ############################################################################
            # 막판무 상판 전개도 좌표 설정
            ############################################################################
            # 막판무 상판은 JD + 5 적용 막판무 Normal JD 값 계산
            Normal_JD = JD + 5
            set_point(aa, 1, a['x3'] , a['y2'] -  1850)  # 단면도의 크기에 맞춤
            set_point(aa, 2, c['x11'] , aa['y1'])
            set_point(aa, 3, aa['x2'] , aa['y2'] + 15 - t)
            set_point(aa, 4, aa['x3'] , aa['y3'] + Normal_JD - t*2)
            set_point(aa, 5, aa['x4'] , aa['y4'] + C - t*2 )
            set_point(aa, 6, aa['x5'] , aa['y5'] + U - t)
            set_point(aa, 7,  a['x3'] , aa['y6'] )
            set_point(aa, 8, aa['x7'] , aa['y7'] - U + t)
            set_point(aa, 9, aa['x8'] , aa['y8'] - C + t*2)
            set_point(aa, 10, aa['x9'] , aa['y9'] - Normal_JD + t*2)

            prev_x, prev_y = aa['x1'], aa['y1']  # 첫 번째 점으로 초기화
            lastNum = 10
            for i in range(1, lastNum + 1):
                curr_x, curr_y = aa[f'x{i}'], aa[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="레이져")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, aa['x1'], aa['y1'], layer="레이져")        

            # 절곡선
            line(doc, aa['x10'], aa['y10'], aa['x3'], aa['y3'], layer="절곡선")      
            line(doc, aa['x9'], aa['y9'], aa['x4'], aa['y4'], layer="절곡선")      
            line(doc, aa['x8'], aa['y8'], aa['x5'], aa['y5'], layer="절곡선")      

            # 상판의 중심위치 좌표 설정
            centerX =  (aa['x1'] +  aa['x2'])/2
            # 상판 중심 3파이 타공 
            draw_circle(doc, centerX , aa['y1'] + 3, 3 , layer='레이져', color='7')

            # 하부 치수선
            d(doc, aa['x1'] , aa['y1'], aa['x2'] , aa['y2'], 120, direction="down", dim_style=over1000dim_style)        
            # 좌측 치수선
            d(doc, aa['x1'] , aa['y1'], aa['x7'] , aa['y7'], 300, direction="left", dim_style=over1000dim_style)        

            # 우측치수선
            d(doc, aa['x5'] , aa['y5'], aa['x4'] , aa['y4'], 100, direction="right", dim_style=over1000dim_style)        
            d(doc, aa['x5'] , aa['y5'], aa['x6'] , aa['y6'], 200, direction="right", dim_style=over1000dim_style)        
            d(doc, aa['x4'] , aa['y4'], aa['x3'] , aa['y3'], 200, direction="right", dim_style=over1000dim_style)        
            d(doc, aa['x3'] , aa['y3'], aa['x2'] , aa['y2'], 200, direction="right", dim_style=over1000dim_style)        
            d(doc, aa['x6'] , aa['y6'], aa['x2'] , aa['y2'], 300, direction="right", dim_style=over1000dim_style)        

            # print(f" JDholes: {JDholes}")
            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(JDholes):
                # 변화에 따른 x위치 삼각함수 적용                
                value = calculate_base(hole + 11.5, poleAngle) 
                Heightchange = calculate_base(hole + 11.5, poleAngle, option="height") 
                Heightchange = round(Heightchange, 2)
                xpos, ypos = centerX - OP/2 - xGap1 - value  , aa['y3'] + Heightchange - t
                rightxpos = centerX + OP/2 + xGap1 + value 
                insert_block(doc, xpos, ypos, "8x20_horizontal_laser")     
                insert_block(doc, rightxpos, ypos, "8x20_horizontal_laser")     
                if i == 0 :
                    dim_linear(doc, xpos, ypos,   centerX - OP/2  , aa['y3'] - t , "", 100 + i*60,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 안쪽 치수선
                    # dim_linear(doc, rightxpos, ypos,   centerX + OP/2  , aa['y3'] - t , "", 100 + i*60,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 안쪽 치수선
                    firstx, firsty = xpos, ypos
                else:
                    dim_linear(doc, lastx, lasty , xpos, ypos, "", dis=100,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 안쪽 치수선
                    dim_linear(doc, lastrightx, lasty , rightxpos, ypos, "", dis=100,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 안쪽 치수선
                lastx, lasty = xpos, ypos                
                lastrightx = rightxpos

            # 장공과 U과 간격
            d(doc, aa['x9'] , aa['y9'], lastx, lasty , 180, direction="left", dim_style=over1000dim_style)  

        ######################################################################################
        # 막판유 상판 우측에 위치한 단면도 그리기 10개 좌표
        # ######################################################################################
        # 좌표 초기화
        ee = {f'x{i}': 0 for i in range(1, 31)}
        ee.update({f'y{i}': 0 for i in range(1, 31)})

        UW = 15 # 상판 뒷날개
        U_innerWing = 25 # 상판 상부날개 높이
        if jambType=='막판유':            
            # 좌표 설정        
            set_point(ee, 1, aa['x4'] + 1000, aa['y4'])
            set_point(ee, 2, ee['x1'] + JD  , ee['y1']   )
            set_point(ee, 3, ee['x2']   , ee['y2'] + UW   )
            set_point(ee, 4, ee['x3'] - t  , ee['y3']   )
            set_point(ee, 5, ee['x4']   , ee['y4'] - UW + t  )
            set_point(ee, 6, ee['x5'] -JD + t*2  , ee['y5']   )
            set_point(ee, 7, ee['x6']   , ee['y6'] + MH - t*2   )
            set_point(ee, 8, ee['x7']  + U_innerWing - t , ee['y7']   )
            set_point(ee, 9, ee['x8']   , ee['y8'] + t  )
            set_point(ee, 10, ee['x9'] -U_innerWing , ee['y9']   )

            prev_x, prev_y = ee['x1'], ee['y1']  # 첫 번째 점으로 초기화
            lastNum = 10
            for i in range(1, lastNum + 1):
                curr_x, curr_y = ee[f'x{i}'], ee[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="0")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, ee['x1'], ee['y1'], layer="0")            

            # 상판 쫄대 연결홀 위치 표기
            tmpX1,tmpY1,tmpX2,tmpY2 = ee['x10'] + 13, ee['y10'] -5, ee['x10'] + 13, ee['y10'] + 5
            line(doc,tmpX1,tmpY1,tmpX2,tmpY2 , layer="CL")    
            d(doc, ee['x10'] , ee['y10'], ee['x9'] , ee['y9'], 100, direction="up", dim_style=over1000dim_style)   
            d(doc, tmpX2, tmpY2, ee['x10'] , ee['y10'], 200, direction="up", dim_style=over1000dim_style)   
            d(doc, ee['x10'] , ee['y10'], ee['x1'] , ee['y1'], 150, direction="left", dim_style=over1000dim_style)   
            d(doc, ee['x1'] , ee['y1'], ee['x2'] , ee['y2'], 200, direction="down", dim_style=over1000dim_style)   
            d(doc, ee['x2'] , ee['y2'], ee['x3'] , ee['y3'], 150, direction="right", dim_style=over1000dim_style)   

            ######################################################################################
            # 와이드쟘 상판 수량표시 기둥 주기
            ######################################################################################        
            string = f"{surang} EA" 
            draw_Text(doc, (b['x1'] + b['x2'])/2 - len(string)*80/2 , aa['y6'] - 350, 120, text=string, layer='레이져')
            draw_Text(doc, (b['x1'] + b['x2'])/2 - len(string)*80/2 , aa['y6'] - 1200, 120, text=string, layer='레이져')
            draw_Text(doc, (b['x1'] + b['x2'])/2 - len(string)*80/2 - 4200 , aa['y6'] - (1400+JD), 120, text=string, layer='레이져')           # 좌기둥 수량
            draw_Text(doc, (b['x1'] + b['x2'])/2 - len(string)*80/2 - 4200 + 1900, aa['y6'] - (1400+JD), 120, text=string, layer='레이져')     # 우기둥 수량

            ######################################################################################
            # 와이드쟘 상판 쫄대 표시
            ######################################################################################        
            string = f"{floorDisplay} 상판 쫄대" 
            draw_Text(doc, (b['x1'] + b['x2'])/2 - len(string)*80/2 , aa['y6'] - 600, 80, text=string, layer='레이져')

            # 좌표 초기화
            ff = {f'x{i}': 0 for i in range(1, 31)}
            ff.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(ff, 1, centerX - OP/2 - SideJE - C , aa['y6'] - 900) # 2는 10미리 각도 돌출에 따른 보정값 추후 삼각함수로 계산해야 함
            set_point(ff, 2, ff['x1'] + OP + SideJE*2 + C*2  , ff['y1']   )
            set_point(ff, 3, ff['x2'] , ff['y2'] - 94   )
            set_point(ff, 4, ff['x3'] + (OP + SideJE*2 + C*2 ) * -1  , ff['y3']   )

            prev_x, prev_y = ff['x1'], ff['y1']  # 첫 번째 점으로 초기화
            lastNum = 4
            for i in range(1, lastNum + 1):
                curr_x, curr_y = ff[f'x{i}'], ff[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="레이져")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, ff['x1'], ff['y1'], layer="레이져")    
            # 절곡선 2개소
            bx1, by1, bx2,by2 = ff['x4'], ff['y4'] + 28.5, ff['x3'], ff['y3'] + 28.5
            line(doc, bx1, by1, bx2,by2 , layer="절곡선")
            d(doc, bx1, by1, ff['x4'] , ff['y4'], 50, direction="left", dim_style=over1000dim_style)  
            bx3, by3, bx4,by4 = ff['x4'], ff['y4'] + 55.5, ff['x3'], ff['y3'] + 55.5
            line(doc, bx3, by3, bx4,by4 , layer="절곡선")
            # 좌측 치수선
            d(doc, bx3, by3,  bx1, by1, 300, direction="left", dim_style=over1000dim_style)  
            d(doc, bx3, by3,  ff['x1'], ff['y1'] , 200, direction="left", dim_style=over1000dim_style)  
            d(doc, ff['x4'], ff['y4'] ,  ff['x1'], ff['y1'] , 400, direction="left", dim_style=over1000dim_style)  

            # 상부치수선
            d(doc, ff['x1'], ff['y1'] ,  ff['x2'], ff['y2'] , 210, direction="up", dim_style=over1000dim_style)  

            # 쫄대 장공 가로형 8x16 넣기
            # 시작점 SideJE OP/4 적용 마지막 
            lastNum = 5
            GapSum = 0
            for i in range(1, lastNum + 1):
                # 8x16 장공
                Gap = OP/4
                leftstartX = ff['x1'] + SideJE + C + GapSum 
                ypos =  ff['y1'] - 12
                insert_block(doc, leftstartX ,ypos , "8x16_horizontal", layer="레이져")
                if i==1 :
                    d(doc, ff['x1'] , ff['y1'], leftstartX , ypos , 120, direction="up", dim_style=over1000dim_style, option='reverse')        
                else:
                    dc(doc, leftstartX  , ypos )
                    lastX, lastY = leftstartX , ypos
                GapSum = GapSum + Gap            

            dc(doc,  ff['x2'] , ff['y2'])
            d(doc, lastX, lastY , ff['x2'] , ff['y2'],  100, direction="right", dim_style=over1000dim_style )        

            ######################################################################################
            # 와이드쟘 상판 쫄대 우측 단면도 그리기
            ######################################################################################
            # 좌표 초기화
            gg = {f'x{i}': 0 for i in range(1, 31)}
            gg.update({f'y{i}': 0 for i in range(1, 31)})

            UW = 15 # 상판 뒷날개
            # 좌표 설정        
            set_point(gg, 1, ff['x2'] + 400, ff['y2'])
            set_point(gg, 2, gg['x1'] , gg['y1'] - 30  )
            set_point(gg, 3, gg['x2'] + 30, gg['y2']   )
            set_point(gg, 4, gg['x3'] , gg['y3'] + t  )
            set_point(gg, 5, gg['x4'] - 30 + t , gg['y4']   )
            set_point(gg, 6, gg['x5'] , gg['y5'] + 30 - t*2   )
            set_point(gg, 7, gg['x6'] + 40 - t, gg['y6']   )
            set_point(gg, 8, gg['x7'] , gg['y7'] + t   )        

            prev_x, prev_y = gg['x1'], gg['y1']  # 첫 번째 점으로 초기화
            lastNum = 8
            for i in range(1, lastNum + 1):
                curr_x, curr_y = gg[f'x{i}'], gg[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="0")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, gg['x1'], gg['y1'], layer="0")            

            # 상판 쫄대 연결홀 위치 표기
            tmpX1,tmpY1,tmpX2,tmpY2 = gg['x8'] - 12, gg['y8'] -5, gg['x8'] - 12, gg['y8'] + 5
            line(doc,tmpX1,tmpY1,tmpX2,tmpY2 , layer="CL")    
            d(doc, gg['x1'] , gg['y1'], gg['x8'] , gg['y8'], 100, direction="up", dim_style=over1000dim_style)
            d(doc, gg['x1'] , gg['y1'], gg['x2'] , gg['y2'], 150, direction="left", dim_style=over1000dim_style)
            d(doc, gg['x2'] , gg['y2'], gg['x3'] , gg['y3'], 150, direction="down", dim_style=over1000dim_style)            
        if jambType=='막판무':            
            # 막판무는 U값이 30임
            U = 30
            # 막판무 상판은 JD + 5 적용 막판무 Normal JD 값 계산
            Normal_JD = JD + 5
            # 좌표 설정        
            set_point(ee, 1, aa['x4'] + 800, aa['y4'])
            set_point(ee, 2, ee['x1'] , ee['y1'] - Normal_JD  )
            set_point(ee, 3, ee['x2'] + UW , ee['y2']  )
            set_point(ee, 4, ee['x3'] , ee['y3'] + t  )
            set_point(ee, 5, ee['x4'] - UW + t  , ee['y4']   )
            set_point(ee, 6, ee['x5'] , ee['y5'] +Normal_JD - t*2    )
            set_point(ee, 7, ee['x6'] + C - t*2  , ee['y6']   )
            set_point(ee, 8, ee['x7'] , ee['y7'] - U + t  )
            set_point(ee, 9, ee['x8']  + t , ee['y8']   )
            set_point(ee, 10, ee['x9'] , ee['y9'] + U   )

            prev_x, prev_y = ee['x1'], ee['y1']  # 첫 번째 점으로 초기화
            lastNum = 10
            for i in range(1, lastNum + 1):
                curr_x, curr_y = ee[f'x{i}'], ee[f'y{i}']
                line(doc, prev_x, prev_y, curr_x, curr_y, layer="0")
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, ee['x1'], ee['y1'], layer="0")            

            # 상판 쫄대 연결홀 위치 표기            
            d(doc, ee['x10'] , ee['y10'], ee['x9'] , ee['y9'], 100, direction="right", dim_style=over1000dim_style)               
            d(doc, ee['x10'] , ee['y10'], ee['x1'] , ee['y1'], 150, direction="up", dim_style=over1000dim_style)   
            d(doc, ee['x1'] , ee['y1'], ee['x2'] , ee['y2'], 150, direction="left", dim_style=over1000dim_style)   
            d(doc, ee['x2'] , ee['y2'], ee['x3'] , ee['y3'], 150, direction="down", dim_style=over1000dim_style)   

            ######################################################################################
            # 와이드쟘(막판무) 상판 수량표시 기둥 주기
            ######################################################################################        
            string = f"{surang} EA" 
            draw_Text(doc, centerX - len(string)*80/2 , aa['y2'] - 450, 120, text=string, layer='레이져')            
            draw_Text(doc, centerX - len(string)*80/2 - 4200 , aa['y2'] - 1050, 120, text=string, layer='레이져')           # 좌기둥 수량
            draw_Text(doc, centerX - len(string)*80/2 - 4200 + 1900, aa['y2'] - 1050, 120, text=string, layer='레이져')     # 우기둥 수량	

        ########################################################################################
        # 막판유 좌기둥 그리기
        ########################################################################################
        if jambType=='막판유':
            ######################################################################################
            # 와이드쟘 좌기둥 작도  
            # 2025/08/18 좌기둥 하부 따임 높이 15 추가 (기존은 치수선만 존재했음)
            # 막판과 기둥 조립홀 위치 수정(횡보강 홀위치 일부 수정) 34.5 기본, 20mm 간격 유지 등
            ######################################################################################        
            rx, startYpos = AbsX + index*10000 - 4000 , 3000
            SW = FireDoor # 방화도어 25, 일반도어 50
            # 좌표 초기화
            rr = {f'x{i}': 0 for i in range(1, 31)}
            rr.update({f'y{i}': 0 for i in range(1, 31)})

            # JB 값 계산 공식
            JB = round(calculate_jb(JE, JD + 10),0)
            # SBW 사이드 뒷날개
            SBW = 10

            # print(f"피타고라스 정리로 구한 JB  : {jb_pythagoras:.2f}")  

            # 일반도어/방화도어에 따라 U형태 따임위치 지정 변수설정
            if FireDoor == 25:
                Utrim_gap = 6
            else:
                Utrim_gap = 6 + 25/2

            # 좌표 설정
            set_point(rr, 1, rx, a['y1']-100)
            set_point(rr, 2, rr['x1']   , rr['y1'] - MH - HH )
            set_point(rr, 3, rr['x2']   , rr['y2'] - grounddig )
            set_point(rr, 4, rr['x3'] + A -t  , rr['y3'] )
            set_point(rr, 5, rr['x4'] + C -t*2  , rr['y4']   )
            set_point(rr, 6, rr['x5'] + JB - t*2  , rr['y5']   )
            set_point(rr, 7, rr['x6'] + Utrim_gap , rr['y6']  )
            set_point(rr, 8, rr['x7']  , rr['y7'] + 9  )
            set_point(rr, 9, rr['x8']  + 10 , rr['y8']   )
            set_point(rr, 10, rr['x9']  , rr['y9'] - 9  )
            set_point(rr, 11, rr['x10'] + Utrim_gap, rr['y10']  )
            set_point(rr, 12, rr['x11'] , rr['y11'] + 15  ) # 기본적으로 15미리 높이 따임
            set_point(rr, 13, rr['x12'] + SBW - t , rr['y12']   )
            set_point(rr, 14, rr['x13'] , rr['y13'] + HH + (grounddig - 15)  )
            set_point(rr, 15, rr['x14'] , rr['y14'] + 3  )
            set_point(rr, 16, rr['x15'] - SBW + t, rr['y15']   )
            set_point(rr, 17, rr['x16'] - SW + t*2, rr['y16']   )
            set_point(rr, 18, rr['x17'] - 15 + t, rr['y17']   )
            set_point(rr, 19, rr['x18'] , rr['y18'] + 27  )
            set_point(rr, 20, rr['x19'] - (JB - t*2 - 38.5 - 13.5), rr['y19']   )
            set_point(rr, 21, rr['x20'] , rr['y20'] + MH - 30 )
            set_point(rr, 22, rr['x21'] - 38.5 , rr['y21']  )		
            set_point(rr, 23, rr['x22'] - C + t*2, rr['y22']  )		

            prev_x, prev_y = rr['x1'], rr['y1']  # 첫 번째 점으로 초기화
            lastNum = 23
            for i in range(1, lastNum + 1):
                curr_x, curr_y = rr[f'x{i}'], rr[f'y{i}']
                if i == 9:
                    draw_arc(doc,  rr['x8'], rr['y8'] , rr['x9'], rr['y9'] , 5, direction='up')  # 반지름 5R 반원 그리기            
                else:
                    line(doc, prev_x, prev_y, curr_x, curr_y, layer="레이져")        
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, rr['x1'], rr['y1'], layer="레이져")   

            # 띠보강 용접홀 3개 좌표 초기화
            weldingHole = {f'x{i}': 0 for i in range(1, 4)}
            weldingHole.update({f'y{i}': 0 for i in range(1, 4)})

            # 띠보강 용접홀 3개 좌표 설정
            set_point(weldingHole, 1, rr['x3'] + 5 , rr['y3']  + 170 )
            set_point(weldingHole, 2, rr['x3'] + 5 , rr['y3']  + 1070 )
            set_point(weldingHole, 3, rr['x3'] + 5 , rr['y3']  + 1970 )

            draw_circle(doc, weldingHole['x1'], weldingHole['y1'], 3 , layer='레이져')
            draw_circle(doc, weldingHole['x2'], weldingHole['y2'], 3 , layer='레이져')
            draw_circle(doc, weldingHole['x3'], weldingHole['y3'], 3 , layer='레이져')

            # 절곡선 4개소
            bx1, by1, bc1, bd1 = rr['x4'], rr['y4'], rr['x23'], rr['y23'] 
            bx2, by2, bc2, bd2 = rr['x5'], rr['y5'], rr['x22'], rr['y22'] 
            bx3, by3, bc3, bd3 = rr['x6'], rr['y6'], rr['x17'], rr['y17'] 
            bx4, by4, bc4, bd4 = rr['x11'], rr['y11'], rr['x16'], rr['y16'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")        
            line(doc, bx2, by2, bc2, bd2, layer="절곡선")        
            line(doc, bx3, by3, bc3, bd3, layer="절곡선")        
            line(doc, bx4, by4, bc4, bd4, layer="절곡선")        

            # HH 자리 CL선 표기
            bx5, by5, bc5, bd5 = rr['x1'] - 50, rr['y14'] , rr['x14'] + 15, rr['y14'] 
            line(doc, bx5, by5, bc5, bd5, layer="CL")                

            # 기둥 하부 치수선
            d(doc, rr['x4'] , rr['y4'], rr['x3'] , rr['y3'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x4'] , rr['y4'], rr['x5'] , rr['y5'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x6'] , rr['y6'], rr['x5'] , rr['y5'], 80, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x6'] , rr['y6'], rr['x11'] , rr['y11'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x11'] , rr['y11'], rr['x13'] , rr['y13'], 80, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x3'] , rr['y3'], rr['x13'] , rr['y13'], 300, direction="down", dim_style=over1000dim_style)        

            # 기둥 우측 치수선
            d(doc, rr['x13'] , rr['y13'], rr['x11'] , rr['y11'],  200, direction="right", dim_style=over1000dim_style)                
            d(doc, rr['x15'] , rr['y15'], rr['x11'] , rr['y11'],  300, direction="right", dim_style=over1000dim_style)        
            d(doc, rr['x14'] , rr['y14'], rr['x15'] , rr['y15'],  150, direction="right", dim_style=over1000dim_style)        
            d(doc, rr['x14'] , rr['y14'], rr['x19'] , rr['y19'],  220, direction="right", dim_style=over1000dim_style)    # 27mm    
            d(doc, rr['x19'] , rr['y19'], rr['x21'] , rr['y21'],  400, direction="right", dim_style=over1000dim_style)        

            # 치수선 상부
            d(doc, rr['x22'] , rr['y22'], rr['x1'] , rr['y1'],  200, direction="up", dim_style=over1000dim_style)        
            d(doc, rr['x22'] , rr['y22'], rr['x21'] , rr['y21'],  200, direction="up", dim_style=over1000dim_style)        
            d(doc, rr['x1'] , rr['y1'], rr['x21'] , rr['y21'],  300, direction="up", dim_style=over1000dim_style)        

            # 치수선 + 문자 구현 예제임 
            ds(doc,  rr['x1'] , by5, rr['x1'] , rr['y1'], dis = 340,  direction="left", text="상판H")
            ds(doc,  rr['x1'] , rr['y1'] , rr['x1'] , rr['y1'] + U, dis = 400,  direction="left", text="쫄대")
            ds(doc,  rr['x2'] , rr['y2'] , rr['x1'] , rr['y1'] , dis = 400,  direction="left", text="마감과 쫄대를 뺀치수")
            ds(doc,  rr['x3'] , rr['y3'] , rr['x1'] , rr['y1'] , dis = 480,  direction="left", text="절단 치수")
            ds(doc,  rr['x2'] , rr['y2'] , rr['x1'] , rr['y1'] + U , dis = 560,  direction="left", text="마감 빠지고 쫄대 포함 높이")
            ds(doc,  rr['x3'] , rr['y3'] , rr['x1'] , rr['y1'] + U , dis = 640,  direction="left", text="쫄대, 마감 포함 높이")

            #A값이 40보다 큰 경우는 MRL 제어반 연결홀 타공하기
            if A>40 :
                holeArray = calSplitHole(125, 300, MH + HH + grounddig)  # 시작점, 인터벌(간격), 제한 limit, length 길이
                # print(f" holeArray: {holeArray}")
                for i, hole in enumerate(holeArray):            
                    xpos, ypos = rr['x1'] + 10 , rr['y3'] + hole
                    draw_circle(doc,  xpos, ypos, 9 , layer='레이져', color='7')
                    if i == 0 :
                        d(doc, rr['x12'] , rr['y12'], xpos, ypos,  30, direction="left", dim_style=over1000dim_style, option='reverse')       
                    else:
                        dc(doc,  xpos, ypos )
                    lastXX, lastYY =xpos, ypos 

                # 마지막 치수선 연결
                d(doc, lastXX, lastYY, xpos, rr['y1'] ,  30 , direction="left", dim_style=over1000dim_style)    

            # MH방향 장공
            if(MH >= 350) :
                result = calcuteHoleArray(MH, 20, 4)
            else:
                result = calcuteHoleArray(MH, 20, 3)

            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(result):            
                xpos, ypos = rr['x21']-15 , rr['y21'] - hole
                insert_block(doc, xpos, ypos, "8x20_vertical_laser")     
                if i == 0 :
                    d(doc, rr['x21'] , rr['y21'], xpos, ypos,  200, direction="right", dim_style=over1000dim_style, option='reverse')       
                else:
                    d(doc, lastXX, lastYY, xpos, ypos,  130 + i*60, direction="left", dim_style=over1000dim_style)     
                lastXX, lastYY =xpos, ypos 

            # 마지막 치수선 연결
            d(doc, lastXX, lastYY, xpos, rr['y21'] - MH,  130 + (i+1)*60, direction="left", dim_style=over1000dim_style)             
            d(doc, xpos , rr['y1'] - 20, rr['x21'] , rr['y21'],  60, direction="up", dim_style=over1000dim_style)        

            # JD 방향 장공
            if JD >= 900:
                JDholes = calcuteHoleArray(JD - 39.5, 20, 5)
            elif JD >= 600: 
                JDholes = calcuteHoleArray(JD - 39.5, 20, 4)
            else:
                JDholes = calcuteHoleArray(JD - 39.5, 20, 3)

            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(JDholes):
                xpos, ypos = rr['x19'] - hole , rr['y19'] - 13 
                insert_block(doc, xpos, ypos, "8x20_horizontal_laser")     
                if i == 0 :
                    d(doc, rr['x19'] , rr['y19'], xpos, ypos,  160, direction="down", dim_style=over1000dim_style, option='reverse')    
                    firstx, firsty = xpos, ypos
                else:
                    d(doc, lastx, lasty, xpos, ypos,  160 + i*55, direction="down", dim_style=over1000dim_style)     
                lastx, lasty = xpos, ypos                 

            # 마지막 치수선 연결
            d(doc, lastXX, lastYY , xpos, ypos, 350, direction="down", dim_style=over1000dim_style, option='reverse')               
            d(doc,  rr['x14'] , rr['y14'],   firstx, firsty,50, direction="left", dim_style=over1000dim_style)    #높이 표시            

            ######################################################################################
            # 와이드쟘 우기둥 작도  
            ######################################################################################        
            rx, startYpos = AbsX + index*10000 - 2000 + JD + 80 , 3000        
            # 좌표 초기화
            rr = {f'x{i}': 0 for i in range(1, 31)}
            rr.update({f'y{i}': 0 for i in range(1, 31)})
            # 좌표 설정
            set_point(rr, 1, rx, a['y1']-100)
            set_point(rr, 2, rr['x1']   , rr['y1'] - MH - HH )
            set_point(rr, 3, rr['x2']   , rr['y2'] - grounddig )
            set_point(rr, 4, rr['x3'] - A + t  , rr['y3'] )
            set_point(rr, 5, rr['x4'] - C +t*2  , rr['y4']   )
            set_point(rr, 6, rr['x5'] - JB + t*2  , rr['y5']   )
            set_point(rr, 7, rr['x6'] - Utrim_gap , rr['y6']  )
            set_point(rr, 8, rr['x7']  , rr['y7'] + 9  )
            set_point(rr, 9, rr['x8']  - 10 , rr['y8']   )
            set_point(rr, 10, rr['x9']  , rr['y9'] - 9  )
            set_point(rr, 11, rr['x10'] - Utrim_gap, rr['y10']  )
            set_point(rr, 12, rr['x11'] , rr['y11'] + 15  ) # 기본적으로 15미리 높이 따임
            set_point(rr, 13, rr['x12'] - SBW + t, rr['y12']   )
            set_point(rr, 14, rr['x13'] , rr['y13'] + HH + (grounddig - 15)  )
            set_point(rr, 15, rr['x14'] , rr['y14'] + 3  )
            set_point(rr, 16, rr['x15'] + SBW - t, rr['y15']   )
            set_point(rr, 17, rr['x16'] + SW - t*2, rr['y16']   )
            set_point(rr, 18, rr['x17'] + 15 - t, rr['y17']   )
            set_point(rr, 19, rr['x18'] , rr['y18'] + 27  )
            set_point(rr, 20, rr['x19'] + (JB - t*2 - 38.5 - 13.5), rr['y19']   )
            set_point(rr, 21, rr['x20'] , rr['y20'] + MH - 30 )
            set_point(rr, 22, rr['x21'] + 38.5 , rr['y21']  )		
            set_point(rr, 23, rr['x22'] + C - t*2, rr['y22']  )		

            prev_x, prev_y = rr['x1'], rr['y1']  # 첫 번째 점으로 초기화
            lastNum = 23
            for i in range(1, lastNum + 1):
                curr_x, curr_y = rr[f'x{i}'], rr[f'y{i}']
                if i == 9:
                    draw_arc(doc,  rr['x8'], rr['y8'] , rr['x9'], rr['y9'] , 5, direction='up')  # 반지름 5R 반원 그리기            
                else:
                    line(doc, prev_x, prev_y, curr_x, curr_y, layer="레이져")        
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, rr['x1'], rr['y1'], layer="레이져")   

            # 절곡선 4개소
            bx1, by1, bc1, bd1 = rr['x4'], rr['y4'], rr['x23'], rr['y23'] 
            bx2, by2, bc2, bd2 = rr['x5'], rr['y5'], rr['x22'], rr['y22'] 
            bx3, by3, bc3, bd3 = rr['x6'], rr['y6'], rr['x17'], rr['y17'] 
            bx4, by4, bc4, bd4 = rr['x11'], rr['y11'], rr['x16'], rr['y16'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")        
            line(doc, bx2, by2, bc2, bd2, layer="절곡선")        
            line(doc, bx3, by3, bc3, bd3, layer="절곡선")        
            line(doc, bx4, by4, bc4, bd4, layer="절곡선")        

            # HH 자리 CL선 표기
            bx5, by5, bc5, bd5 = rr['x1'] - 50, rr['y14'] , rr['x14'] + 15, rr['y14'] 
            line(doc, bx5, by5, bc5, bd5, layer="CL")                

            # 기둥 하부 치수선
            d(doc, rr['x4'] , rr['y4'], rr['x3'] , rr['y3'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x4'] , rr['y4'], rr['x5'] , rr['y5'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x6'] , rr['y6'], rr['x5'] , rr['y5'], 80, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x6'] , rr['y6'], rr['x11'] , rr['y11'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x11'] , rr['y11'], rr['x13'] , rr['y13'], 80, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x3'] , rr['y3'], rr['x13'] , rr['y13'], 300, direction="down", dim_style=over1000dim_style)        

            # 기둥 우측 치수선
            d(doc, rr['x13'] , rr['y13'], rr['x11'] , rr['y11'],  200, direction="left", dim_style=over1000dim_style)                
            d(doc, rr['x15'] , rr['y15'], rr['x11'] , rr['y11'],  300, direction="left", dim_style=over1000dim_style)        
            d(doc, rr['x14'] , rr['y14'], rr['x15'] , rr['y15'],  150, direction="left", dim_style=over1000dim_style)        
            d(doc, rr['x14'] , rr['y14'], rr['x19'] , rr['y19'],  220, direction="left", dim_style=over1000dim_style)    # 27mm    
            d(doc, rr['x19'] , rr['y19'], rr['x21'] , rr['y21'],  400, direction="left", dim_style=over1000dim_style)        

            # 치수선 상부
            d(doc, rr['x22'] , rr['y22'], rr['x1'] , rr['y1'],  200, direction="up", dim_style=over1000dim_style)        
            d(doc, rr['x22'] , rr['y22'], rr['x21'] , rr['y21'],  200, direction="up", dim_style=over1000dim_style)        
            d(doc, rr['x1'] , rr['y1'], rr['x21'] , rr['y21'],  300, direction="up", dim_style=over1000dim_style)        

            # 치수선 + 문자 구현 예제임 
            ds(doc,  rr['x1'] , by5, rr['x1'] , rr['y1'], dis = 400,  direction="right", text="상판H")
            ds(doc,  rr['x1'] , rr['y1'] , rr['x1'] , rr['y1'] + U, dis = 500,  direction="right", text="쫄대")
            ds(doc,  rr['x2'] , rr['y2'] , rr['x1'] , rr['y1'] , dis = 420,  direction="right", text="마감과 쫄대를 뺀치수")
            ds(doc,  rr['x3'] , rr['y3'] , rr['x1'] , rr['y1'] , dis = 500,  direction="right", text="절단 치수")
            ds(doc,  rr['x2'] , rr['y2'] , rr['x1'] , rr['y1'] + U , dis = 580,  direction="right", text="마감 빠지고 쫄대 포함 높이")
            ds(doc,  rr['x3'] , rr['y3'] , rr['x1'] , rr['y1'] + U , dis = 660,  direction="right", text="쫄대, 마감 포함 높이")

            # 띠보강 용접홀 3개 좌표 초기화
            weldingHole = {f'x{i}': 0 for i in range(1, 4)}
            weldingHole.update({f'y{i}': 0 for i in range(1, 4)})

            # 띠보강 용접홀 3개 좌표 설정
            set_point(weldingHole, 1, rr['x3'] - 5 , rr['y3']  + 170 )
            set_point(weldingHole, 2, rr['x3'] - 5 , rr['y3']  + 1070 )
            set_point(weldingHole, 3, rr['x3'] - 5 , rr['y3']  + 1970 )

            draw_circle(doc, weldingHole['x1'], weldingHole['y1'], 3 , layer='레이져')
            draw_circle(doc, weldingHole['x2'], weldingHole['y2'], 3 , layer='레이져')
            draw_circle(doc, weldingHole['x3'], weldingHole['y3'], 3 , layer='레이져')


            #A값이 40보다 큰 경우는 MRL 제어반 연결홀 타공하기
            if A>40 :            
                holeArray = calSplitHole(125, 300, MH + HH + grounddig)  # 시작점, 인터벌(간격), 제한 limit, length 길이
                # print(f" holeArray: {holeArray}")
                for i, hole in enumerate(holeArray):            
                    xpos, ypos = rr['x1'] - 10 , rr['y3'] + hole
                    draw_circle(doc,  xpos, ypos, 9 , layer='레이져', color='7')
                    if i == 0 :
                        d(doc, rr['x12'] , rr['y12'], xpos, ypos,  80, direction="right", dim_style=over1000dim_style, option='reverse')       
                    else:
                        dc(doc,  xpos, ypos )
                    lastXX, lastYY =xpos, ypos 

                # 마지막 치수선 연결
                d(doc, lastXX, lastYY, xpos, rr['y1'] ,  80 , direction="right", dim_style=over1000dim_style)    


            # MH방향 장공
            if(MH >= 350) :
                result = calcuteHoleArray(MH, 20, 4)
            else:
                result = calcuteHoleArray(MH, 20, 3)
            # print(result)
            sum = len(result)
            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(result):            
                xpos, ypos = rr['x21']+15 , rr['y21'] - hole
                insert_block(doc, xpos, ypos, "8x20_vertical_laser")     
                if i == 0 :
                    d(doc, rr['x21'] , rr['y21'], xpos, ypos,  200, direction="left", dim_style=over1000dim_style, option='reverse')       
                else:
                    d(doc, lastXX, lastYY , xpos, ypos,  160 + i*60, direction="right", dim_style=over1000dim_style)
                lastXX, lastYY =xpos, ypos 

            # 마지막 치수선 연결
            d(doc, lastXX, lastYY, xpos, rr['y21'] - MH,  160 + (i+1)*60, direction="right", dim_style=over1000dim_style)             
            d(doc, xpos , rr['y1'] - 20, rr['x21'] , rr['y21'],  60, direction="up", dim_style=over1000dim_style)        

            # JD 방향 장공
            if JD >= 900:
                JDholes = calcuteHoleArray(JD - 39.5, 20, 5)
            elif JD >= 600: 
                JDholes = calcuteHoleArray(JD - 39.5, 20, 4)
            else:
                JDholes = calcuteHoleArray(JD - 39.5, 20, 3)

            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(JDholes):
                xpos, ypos = rr['x19'] + hole , rr['y19'] - 13 
                insert_block(doc, xpos, ypos, "8x20_horizontal_laser")     
                if i == 0 :
                    d(doc, rr['x19'] , rr['y19'], xpos, ypos,  30, direction="up", dim_style=over1000dim_style, option='reverse')    
                    firstx, firsty = xpos, ypos
                else:
                    d(doc, lastx, lasty, xpos, ypos,  30 + i*60, direction="down", dim_style=over1000dim_style)     
                lastx, lasty = xpos, ypos                 

            # 마지막 치수선 연결
            d(doc, xpos, ypos, lastXX, lastYY , 300, direction="down", dim_style=over1000dim_style, option='reverse')               
            d(doc, rr['x14'] , rr['y14'], firstx, firsty,  50, direction="right", dim_style=over1000dim_style)    #높이 표시         

            ######################################################################################
            # 와이드쟘 기둥 주기
            ######################################################################################                
            draw_Text(doc, rr['x15'] , rr['y1'] +  450, 120, text="우측기둥")
            draw_Text(doc, rr['x15'] - 2000 , rr['y1'] +  450, 120, text="좌측기둥")
            string =f"{floorDisplay} 기둥 JD={JD+10}" # JD를 전체 기둥과 결합한 형태로 다완테크는 표현합니다. 주의요함
            draw_Text(doc, rr['x15'] - 1300 , rr['y1'] +  720, 120, text=string)

            ######################################################################################
            # 와이드쟘 기둥 단면도 (좌기둥 우기둥 중간에 위치) 띠보강 들어감
            ######################################################################################                
            tmpx, tmpy = AbsX + index*10000 - 3000 , 3300            

            SWC = FireDoor # 오른쪽 하단 방화도어 25, 일반도어 50
            SW = 10  # 오른쪽 그 위
            JD_plus_10 = JD + 10

            angle_deg = 90 - poleAngle
            angle_rad = math.radians(angle_deg)

            base = JD_plus_10 / math.tan(angle_rad)   # 밑면 구하기      

            # ㄷ자 형상 만든기
            # 좌표 초기화
            jj = {f'x{i}': 0 for i in range(1, 31)}
            jj.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(jj, 1, tmpx, tmpy )
            set_point(jj, 2, jj['x1'] - A , jj['y1']   )
            set_point(jj, 3, jj['x2']   , jj['y2'] - C )
            set_point(jj, 4, jj['x3'] + JD_plus_10  , jj['y3'] - base )
            set_point(jj, 5, jj['x4']   , jj['y4'] + SWC  )
            set_point(jj, 6, jj['x5'] - SW  , jj['y5']   )
            set_point(jj, 7, jj['x6']   , jj['y6'] -t  )
            set_point(jj, 8, jj['x7'] + SW - t  , jj['y7']   )
            set_point(jj, 9, jj['x8']   , jj['y8'] - SWC + t*2  )
            set_point(jj, 10, jj['x9'] - JD_plus_10 + t*2 , jj['y9'] + base  )
            set_point(jj, 11, jj['x10'] , jj['y10'] + C - t*2  )
            set_point(jj, 12, jj['x11'] + A - t , jj['y11']  )		

            prev_x, prev_y = jj['x1'], jj['y1']  # 첫 번째 점으로 초기화
            lastNum = 12
            for i in range(1, lastNum + 1):
                cujj_x, cujj_y = jj[f'x{i}'], jj[f'y{i}']
                line(doc, prev_x, prev_y, cujj_x, cujj_y, layer="0")
                prev_x, prev_y = cujj_x, cujj_y
            line(doc, prev_x, prev_y, jj['x1'], jj['y1'], layer="0")  	

            d(doc, jj['x1'], jj['y1'], jj['x2'], jj['y2'] , 100, direction="up", dim_style=over1000dim_style)    #높이 표시              
            d(doc, jj['x2'], jj['y2'], jj['x3'], jj['y3'] , 100, direction="left", dim_style=over1000dim_style)    #높이 표시              
            d(doc, jj['x5'], jj['y5'], jj['x4'], jj['y4'] , 120, direction="right", dim_style=over1000dim_style)    #높이 표시       
            d(doc, jj['x4'], jj['y4'], jj['x3'], jj['y3'],  250, direction="down", dim_style=over1000dim_style)    # JD 표시       
            d(doc, jj['x6'], jj['y6'], jj['x5'], jj['y5'],  120, direction="up", dim_style=over1000dim_style)    # 뒷날개       
            dim_linear(doc, jj['x9'], jj['y9'], jj['x10'], jj['y10'], "", dis=50,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 안쪽 치수선
            dim_linear(doc, jj['x4'], jj['y4'], jj['x3'], jj['y3'], "", dis=150,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 바깥쪽 치수선
            # 각도표시
            dim_angular(doc,   jj['x10'], jj['y10'], jj['x9'] , jj['y9'],  jj['x9'] , jj['y9'],  jj['x8'] , jj['y8'] + 50,  50, direction="up" )   # up,down, left, right각도를 표시할때는 기존치수선 연장선에 해줘야 제대로 값을 얻는다        

            #  조인트 형상 만든기 판때기 (약간 절곡 부분) 띠보강 들어감
            tmpx, tmpy =  jj['x1'], jj['y1'] - t
            # 좌표 초기화
            pp = {f'x{i}': 0 for i in range(1, 31)}
            pp.update({f'y{i}': 0 for i in range(1, 31)})

            h = JD_plus_10-78 + (25 - A)
            theta_deg = poleAngle
            theta_rad = math.radians(theta_deg)
            height = math.floor(h * math.cos(theta_rad)*10)/10 + 5
            base = h / math.tan(angle_rad)   # 밑면 구하기              

            # 막판유 띠보강 단면도 좌표 설정
            set_point(pp, 1, tmpx - 5, tmpy )            
            set_point(pp, 2, pp['x1'] + height, pp['y1'] - base - 3 )  # 3.4 공차를 제거함
            set_point(pp, 3, pp['x2'] - 0.2  , pp['y2'] - t  )            
            set_point(pp, 4, tmpx - 5 - 0.2 , pp['y1'] - t  )

            prev_x, prev_y = pp['x1'], pp['y1']  # 첫 번째 점으로 초기화
            lastNum = 4
            for i in range(1, lastNum + 1):
                cupp_x, cupp_y = pp[f'x{i}'], pp[f'y{i}']
                line(doc, prev_x, prev_y, cupp_x, cupp_y, layer="구성선")
                prev_x, prev_y = cupp_x, cupp_y
            line(doc, prev_x, prev_y, pp['x1'], pp['y1'], layer="구성선")  	

            dim_linear(doc, pp['x1'], pp['y1'], pp['x2'], pp['y2'], "", dis=180,  direction="aligned" ,text_height=0.30, text_gap=0.07) 

            # 모자보강 그리기
            basex, basey = jj['x4'] - 21, jj['y4'] + 21 / math.tan(angle_rad) + t + 0.5
            angle = poleAngle                 
            topLength = 33     #모자보강 상단 폭
            bottomLength = 25  #모자보강 높이
            height = 25        #모자보강 밑면

            # 도면 그리기
            pts = draw_hatshape(doc, basex, basey, angle, bottomLength, topLength, height, layer="구성선")

        ########################################################################################
        # 막판무 좌기둥 그리기
        ########################################################################################
        if jambType=='막판무':            
            rx, startYpos = AbsX + index*10000 - 4000 , 3000
            SW = FireDoor # 방화도어 25, 일반도어 50
            # 좌표 초기화
            rr = {f'x{i}': 0 for i in range(1, 31)}
            rr.update({f'y{i}': 0 for i in range(1, 31)})

            # JB 값 계산 공식 막판무는 JD가 5mm 크다. 이 각도 보정해야 함. JD가 기준이다.
            JB = round(calculate_jb(JE, JD ),0)
            # SBW 사이드 뒷날개
            SBW = 10

            # print(f"피타고라스 정리로 구한 JB  : {jb_pythagoras:.2f}")  

            # 일반도어/방화도어에 따라 U형태 따임위치 지정 변수설정
            if FireDoor == 25:
                Utrim_gap = 6
            else:
                Utrim_gap = 6 + 25/2            

            # 좌표 설정
            set_point(rr, 1, rx, a['y1']-100)
            set_point(rr, 2, rr['x1']   , rr['y1'] - HH )
            set_point(rr, 3, rr['x2']   , rr['y2'] - grounddig )
            set_point(rr, 4, rr['x3'] + A -t  , rr['y3'] )
            set_point(rr, 5, rr['x4'] + C -t*2  , rr['y4']   )
            set_point(rr, 6, rr['x5'] + JB - t*2  , rr['y5']   )
            set_point(rr, 7, rr['x6'] + Utrim_gap , rr['y6']  )
            set_point(rr, 8, rr['x7']  , rr['y7'] + 9  )
            set_point(rr, 9, rr['x8']  + 10 , rr['y8']   )
            set_point(rr, 10, rr['x9']  , rr['y9'] - 9  )
            set_point(rr, 11, rr['x10'] + Utrim_gap, rr['y10']  )
            set_point(rr, 12, rr['x11'] , rr['y11'] + 15  ) # 기본적으로 15미리 높이 따임
            set_point(rr, 13, rr['x12'] + SBW - t , rr['y12'] )
            set_point(rr, 14, rr['x13'] , rr['y13'] + HH + (grounddig - 15)  )
            set_point(rr, 15, rr['x14'] - SBW + t , rr['y14']  )
            set_point(rr, 16, rr['x15'] - SW + t*2 , rr['y15']   )
            set_point(rr, 17, rr['x16'] - (JB - t*2), rr['y16']   )
            set_point(rr, 18, rr['x17'] - C + t, rr['y17']   )

            prev_x, prev_y = rr['x1'], rr['y1']  # 첫 번째 점으로 초기화
            lastNum = 18
            for i in range(1, lastNum + 1):
                curr_x, curr_y = rr[f'x{i}'], rr[f'y{i}']
                if i == 9:
                    draw_arc(doc,  rr['x8'], rr['y8'] , rr['x9'], rr['y9'] , 5, direction='up')  # 반지름 5R 반원 그리기            
                else:
                    line(doc, prev_x, prev_y, curr_x, curr_y, layer="레이져")        
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, rr['x1'], rr['y1'], layer="레이져")   

            # 절곡선 4개소
            bx1, by1, bc1, bd1 = rr['x4'], rr['y4'], rr['x18'], rr['y18'] 
            bx2, by2, bc2, bd2 = rr['x5'], rr['y5'], rr['x17'], rr['y17'] 
            bx3, by3, bc3, bd3 = rr['x6'], rr['y6'], rr['x16'], rr['y16'] 
            bx4, by4, bc4, bd4 = rr['x11'], rr['y11'], rr['x15'], rr['y15'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")        
            line(doc, bx2, by2, bc2, bd2, layer="절곡선")        
            line(doc, bx3, by3, bc3, bd3, layer="절곡선")        
            line(doc, bx4, by4, bc4, bd4, layer="절곡선")                    

            # 기둥 하부 치수선
            d(doc, rr['x4'] , rr['y4'], rr['x3'] , rr['y3'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x4'] , rr['y4'], rr['x5'] , rr['y5'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x6'] , rr['y6'], rr['x5'] , rr['y5'], 80, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x6'] , rr['y6'], rr['x11'] , rr['y11'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x11'] , rr['y11'], rr['x13'] , rr['y13'], 80, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x3'] , rr['y3'], rr['x13'] , rr['y13'], 300, direction="down", dim_style=over1000dim_style)        

            # 기둥 우측 치수선
            d(doc, rr['x13'] , rr['y13'], rr['x11'] , rr['y11'],  200, direction="right", dim_style=over1000dim_style)                
            d(doc, rr['x15'] , rr['y15'], rr['x11'] , rr['y11'],  300, direction="right", dim_style=over1000dim_style)        

            # 치수선 + 문자 구현 예제임             
            ds(doc,  rr['x1'] , rr['y1'] , rr['x1'] , rr['y1'] + U, dis = 200,  direction="left", text="쫄대")
            d(doc, rr['x1'] , rr['y1'], rr['x2'] , rr['y2'], 200, direction="left", dim_style=over1000dim_style)        
            d(doc, rr['x1'] , rr['y1'] + U, rr['x3'] , rr['y3'], 300, direction="left", dim_style=over1000dim_style)        

            # 띠보강 용접홀 3개 좌표 초기화
            weldingHole = {f'x{i}': 0 for i in range(1, 4)}
            weldingHole.update({f'y{i}': 0 for i in range(1, 4)})

            # 띠보강 용접홀 3개 좌표 설정
            set_point(weldingHole, 1, rr['x3'] + 5 , rr['y3']  + 170 )
            set_point(weldingHole, 2, rr['x3'] + 5 , rr['y3']  + 1070 )
            set_point(weldingHole, 3, rr['x3'] + 5 , rr['y3']  + 1970 )

            draw_circle(doc, weldingHole['x1'], weldingHole['y1'], 3 , layer='레이져')
            draw_circle(doc, weldingHole['x2'], weldingHole['y2'], 3 , layer='레이져')
            draw_circle(doc, weldingHole['x3'], weldingHole['y3'], 3 , layer='레이져')            
            
            #A값이 40보다 큰 경우는 MRL 제어반 연결홀 타공하기
            if A>40 :
                holeArray = calSplitHole(125, 300, HH + grounddig)  # 시작점, 인터벌(간격), 제한 limit, length 길이
                # print(f" holeArray: {holeArray}")
                for i, hole in enumerate(holeArray):            
                    xpos, ypos = rr['x1'] + 10 , rr['y3'] + hole
                    draw_circle(doc,  xpos, ypos, 9 , layer='레이져', color='7')
                    if i == 0 :
                        d(doc, rr['x3'] , rr['y3'], xpos, ypos,  130, direction="left", dim_style=over1000dim_style, option='reverse')       
                    else:
                        dc(doc,  xpos, ypos )
                    lastXX, lastYY =xpos, ypos 

                # 마지막 치수선 연결
                d(doc, lastXX, lastYY, xpos, rr['y1'] ,  130 , direction="left", dim_style=over1000dim_style)    

            ######################################################################################
            # 막판무 우기둥 작도  
            ######################################################################################        
            rx, startYpos = AbsX + index*10000 - 2000 + JD + 80 , 3000        
            # 좌표 초기화
            rr = {f'x{i}': 0 for i in range(1, 31)}
            rr.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(rr, 1, rx, a['y1']-100)
            set_point(rr, 2, rr['x1']   , rr['y1'] - HH )
            set_point(rr, 3, rr['x2']   , rr['y2'] - grounddig )
            set_point(rr, 4, rr['x3'] - A + t  , rr['y3'] )
            set_point(rr, 5, rr['x4'] - C +t*2  , rr['y4']   )
            set_point(rr, 6, rr['x5'] - JB + t*2  , rr['y5']   )
            set_point(rr, 7, rr['x6'] - Utrim_gap , rr['y6']  )
            set_point(rr, 8, rr['x7']  , rr['y7'] + 9  )
            set_point(rr, 9, rr['x8']  - 10 , rr['y8']   )
            set_point(rr, 10, rr['x9']  , rr['y9'] - 9  )
            set_point(rr, 11, rr['x10'] - Utrim_gap, rr['y10']  )
            set_point(rr, 12, rr['x11'] , rr['y11'] + 15  )
            set_point(rr, 13, rr['x12'] - SBW + t, rr['y12']  )
            set_point(rr, 14, rr['x13'] , rr['y13'] + HH + (grounddig - 15)  )
            set_point(rr, 15, rr['x14'] + SBW - t , rr['y14']  )
            set_point(rr, 16, rr['x15'] + SW  - t*2 , rr['y15']   )
            set_point(rr, 17, rr['x16'] + (JB - t*2),  rr['y16']   )
            set_point(rr, 18, rr['x17'] + C - t , rr['y17']   )

            prev_x, prev_y = rr['x1'], rr['y1']  # 첫 번째 점으로 초기화
            lastNum = 18
            for i in range(1, lastNum + 1):
                curr_x, curr_y = rr[f'x{i}'], rr[f'y{i}']
                if i == 9:
                    draw_arc(doc,  rr['x8'], rr['y8'] , rr['x9'], rr['y9'] , 5, direction='up')  # 반지름 5R 반원 그리기            
                else:
                    line(doc, prev_x, prev_y, curr_x, curr_y, layer="레이져")        
                prev_x, prev_y = curr_x, curr_y
            line(doc, prev_x, prev_y, rr['x1'], rr['y1'], layer="레이져")   

            # 절곡선 4개소
            bx1, by1, bc1, bd1 = rr['x4'], rr['y4'], rr['x18'], rr['y18'] 
            bx2, by2, bc2, bd2 = rr['x5'], rr['y5'], rr['x17'], rr['y17'] 
            bx3, by3, bc3, bd3 = rr['x6'], rr['y6'], rr['x16'], rr['y16'] 
            bx4, by4, bc4, bd4 = rr['x11'], rr['y11'], rr['x15'], rr['y15'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")        
            line(doc, bx2, by2, bc2, bd2, layer="절곡선")        
            line(doc, bx3, by3, bc3, bd3, layer="절곡선")        
            line(doc, bx4, by4, bc4, bd4, layer="절곡선")         

            # 기둥 하부 치수선
            d(doc, rr['x4'] , rr['y4'], rr['x3'] , rr['y3'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x4'] , rr['y4'], rr['x5'] , rr['y5'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x6'] , rr['y6'], rr['x5'] , rr['y5'], 80, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x6'] , rr['y6'], rr['x11'] , rr['y11'], 150, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x11'] , rr['y11'], rr['x13'] , rr['y13'], 80, direction="down", dim_style=over1000dim_style)        
            d(doc, rr['x3'] , rr['y3'], rr['x13'] , rr['y13'], 300, direction="down", dim_style=over1000dim_style)        

            # 기둥 우측 치수선
            d(doc, rr['x13'] , rr['y13'], rr['x11'] , rr['y11'],  150, direction="left", dim_style=over1000dim_style)                
            d(doc, rr['x15'] , rr['y15'], rr['x11'] , rr['y11'],  250, direction="left", dim_style=over1000dim_style)        

            # 치수선 + 문자 구현 예제임             
            ds(doc,  rr['x1'] , rr['y1'] , rr['x1'] , rr['y1'] + U, dis = 250,  direction="right", text="쫄대")
            d(doc, rr['x1'] , rr['y1'], rr['x2'] , rr['y2'], 250, direction="right", dim_style=over1000dim_style)        
            d(doc, rr['x1'] , rr['y1'] + U, rr['x3'] , rr['y3'], 350, direction="right", dim_style=over1000dim_style)     

            # 띠보강 용접홀 3개 좌표 초기화
            weldingHole = {f'x{i}': 0 for i in range(1, 4)}
            weldingHole.update({f'y{i}': 0 for i in range(1, 4)})

            # 띠보강 용접홀 3개 좌표 설정
            set_point(weldingHole, 1, rr['x3'] - 5 , rr['y3']  + 170 )
            set_point(weldingHole, 2, rr['x3'] - 5 , rr['y3']  + 1070 )
            set_point(weldingHole, 3, rr['x3'] - 5 , rr['y3']  + 1970 )

            draw_circle(doc, weldingHole['x1'], weldingHole['y1'], 3 , layer='레이져')
            draw_circle(doc, weldingHole['x2'], weldingHole['y2'], 3 , layer='레이져')
            draw_circle(doc, weldingHole['x3'], weldingHole['y3'], 3 , layer='레이져')               

            if A>40 :            
                holeArray = calSplitHole(125, 300,  HH + grounddig)  # 시작점, 인터벌(간격), 제한 limit, length 길이
                # print(f" holeArray: {holeArray}")
                for i, hole in enumerate(holeArray):            
                    xpos, ypos = rr['x1'] - 10 , rr['y3'] + hole
                    draw_circle(doc,  xpos, ypos, 9 , layer='레이져', color='7')
                    if i == 0 :
                        d(doc, rr['x3'] , rr['y3'], xpos, ypos,  180, direction="right", dim_style=over1000dim_style, option='reverse')       
                    else:
                        dc(doc,  xpos, ypos )
                    lastXX, lastYY =xpos, ypos 

                # 마지막 치수선 연결
                d(doc, lastXX, lastYY, xpos, rr['y1'] ,  180 , direction="right", dim_style=over1000dim_style)    
            
            ######################################################################################
            # 막판무 기둥 주기
            ######################################################################################                
            draw_Text(doc, rr['x15'] , rr['y1'] +  450, 120, text="우측기둥")
            draw_Text(doc, rr['x15'] - 2000 , rr['y1'] +  450, 120, text="좌측기둥")
            string =f"{floorDisplay} 기둥 JD={JD}" # JD를 전체 기둥과 결합한 형태로 다완테크는 표현합니다. 주의요함
            draw_Text(doc, rr['x15'] - 1300 , rr['y1'] +  720, 120, text=string)

            ######################################################################################
            # 와이드쟘 기둥 단면도 (좌기둥 우기둥 중간에 위치)
            ######################################################################################                
            tmpx, tmpy = AbsX + index*10000 - 3000 , 3300            

            SWC = FireDoor # 오른쪽 하단 방화도어 25, 일반도어 50
            SW = 10  # 오른쪽 그 위
            JB = JD  # 막판무는 JD 그대로 적용

            angle_deg = 90 - poleAngle
            angle_rad = math.radians(angle_deg)

            base = JB / math.tan(angle_rad)   # 밑면 구하기      

            # ㄷ자 형상 만든기
            # 좌표 초기화
            jj = {f'x{i}': 0 for i in range(1, 31)}
            jj.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(jj, 1, tmpx, tmpy )
            set_point(jj, 2, jj['x1'] - A , jj['y1']   )
            set_point(jj, 3, jj['x2']   , jj['y2'] - C )
            set_point(jj, 4, jj['x3'] + JB  , jj['y3'] - base )
            set_point(jj, 5, jj['x4']   , jj['y4'] + SWC  )
            set_point(jj, 6, jj['x5'] - SW  , jj['y5']   )
            set_point(jj, 7, jj['x6']   , jj['y6'] -t  )
            set_point(jj, 8, jj['x7'] + SW - t  , jj['y7']   )
            set_point(jj, 9, jj['x8']   , jj['y8'] - SWC + t*2  )
            set_point(jj, 10, jj['x9'] - JB + t*2 , jj['y9'] + base  )
            set_point(jj, 11, jj['x10'] , jj['y10'] + C - t*2  )
            set_point(jj, 12, jj['x11'] + A - t , jj['y11']  )		

            prev_x, prev_y = jj['x1'], jj['y1']  # 첫 번째 점으로 초기화
            lastNum = 12
            for i in range(1, lastNum + 1):
                cujj_x, cujj_y = jj[f'x{i}'], jj[f'y{i}']
                line(doc, prev_x, prev_y, cujj_x, cujj_y, layer="0")
                prev_x, prev_y = cujj_x, cujj_y
            line(doc, prev_x, prev_y, jj['x1'], jj['y1'], layer="0")  	

            d(doc, jj['x1'], jj['y1'], jj['x2'], jj['y2'] , 100, direction="up", dim_style=over1000dim_style)    #높이 표시              
            d(doc, jj['x2'], jj['y2'], jj['x3'], jj['y3'] , 100, direction="left", dim_style=over1000dim_style)    #높이 표시              
            d(doc, jj['x5'], jj['y5'], jj['x4'], jj['y4'] , 120, direction="right", dim_style=over1000dim_style)    #높이 표시       
            d(doc, jj['x4'], jj['y4'], jj['x3'], jj['y3'],  250, direction="down", dim_style=over1000dim_style)    # JD 표시       
            d(doc, jj['x6'], jj['y6'], jj['x5'], jj['y5'],  120, direction="up", dim_style=over1000dim_style)    # 뒷날개       
            dim_linear(doc, jj['x9'], jj['y9'], jj['x10'], jj['y10'], "", dis=50,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 안쪽 치수선
            dim_linear(doc, jj['x4'], jj['y4'], jj['x3'], jj['y3'], "", dis=150,  direction="aligned" ,text_height=0.30, text_gap=0.07)   # 바깥쪽 치수선
            # 각도표시
            dim_angular(doc,   jj['x10'], jj['y10'], jj['x9'] , jj['y9'],  jj['x9'] , jj['y9'],  jj['x8'] , jj['y8'] + 50,  50, direction="up" )   # up,down, left, right각도를 표시할때는 기존치수선 연장선에 해줘야 제대로 값을 얻는다        

            #  조인트 형상 만든기 판때기 (약간 절곡 부분) 띠보강 들어감

            tmpx, tmpy =  jj['x1'], jj['y1'] - t
            # 좌표 초기화
            pp = {f'x{i}': 0 for i in range(1, 31)}
            pp.update({f'y{i}': 0 for i in range(1, 31)})

            h = JB-78 + (45 - A)
            theta_deg = poleAngle
            theta_rad = math.radians(theta_deg)
            height = math.floor(h * math.cos(theta_rad)*10)/10 - 15.2
            base = (h - (45 - A)) / math.tan(angle_rad)   # 밑면 구하기              

            # 막판무 띠보강 단면도 좌표 설정
            # 띠보강은 A값에 따라 다르게 그려야 한다. 기본 45가 기준인데, 최상층이 아니면 작아지는 부분을 고려해야 함 띠보강이 더 길어진다.
            set_point(pp, 1, tmpx - 5, tmpy )            
            set_point(pp, 2, pp['x1'] + height, pp['y1'] - base - 3 )  # 6.4 공차를 제거함
            set_point(pp, 3, pp['x2'] - 0.2  , pp['y2'] - t  )            
            set_point(pp, 4, tmpx - 5 - 0.2 , pp['y1'] - t  )

            prev_x, prev_y = pp['x1'], pp['y1']  # 첫 번째 점으로 초기화
            lastNum = 4
            for i in range(1, lastNum + 1):
                cupp_x, cupp_y = pp[f'x{i}'], pp[f'y{i}']
                line(doc, prev_x, prev_y, cupp_x, cupp_y, layer="구성선")
                prev_x, prev_y = cupp_x, cupp_y
            line(doc, prev_x, prev_y, pp['x1'], pp['y1'], layer="구성선")  	

            dim_linear(doc, pp['x1'], pp['y1'], pp['x2'], pp['y2'], "", dis=180,  direction="aligned" ,text_height=0.30, text_gap=0.07) 

            # 모자보강 그리기
            basex, basey = jj['x4'] - 21, jj['y4'] + 21 / math.tan(angle_rad) + t + 0.5
            angle = poleAngle                 
            topLength = 33     #모자보강 상단 폭
            bottomLength = 25  #모자보강 높이
            height = 25        #모자보강 밑면

            # 도면 그리기
            pts = draw_hatshape(doc, basex, basey, angle, bottomLength, topLength, height, layer="구성선")


        BasicXscale = 6851
        BasicYscale = 4870
        TargetXscale = 7200 
        TargetYscale = 5800 
        if (TargetXscale / BasicXscale > TargetYscale / BasicYscale):
            frame_scale = TargetXscale / BasicXscale
        else:
            frame_scale = TargetYscale / BasicYscale        
        frameXpos = rx
        frameYpos = -1500

        # 막판무는 JD 그대로 적용 (기둥기준으로 직선거리를 JD로 명기함 다완테크)
        JD_sheet = JD 

        insert_frame(frameXpos-3500 , frameYpos  , frame_scale, "JAMB ASSY", f"JD:{JD_sheet}, OP:{OP}", f"{workplace}","basic" )
  
    #####################################################################################
    # 와이드쟘 부속자재 그려주기 merged_rows만 순회
    #####################################################################################
    for index, row_data in enumerate(merged_rows, start=1):
        jambType, floorDisplay, material, spec, vcut, OP, poleAngle, JE, JD, HH, MH, HPI_height, U, C, A, grounddig, surang, FireDoor = load_excel(row_data)

        rx, startYpos = (pageCount+1) * 10000 + index*2000 - 5000, 3000      

        if index == 1: 
            # 모자보강 단면도 
            basex, basey = rx - 2200 , startYpos
            tt = 1.6
            angle = 0
            topLength = 33     #모자보강 상단 폭
            bottomLength = 25  #모자보강 높이
            height = 25        #모자보강 밑면            
            # 좌표 초기화
            yy = {f'x{i}': 0 for i in range(1, 31)}
            yy.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(yy, 1, basex, basey )
            set_point(yy, 2, yy['x1'] + bottomLength  , yy['y1']   )
            set_point(yy, 3, yy['x2']   , yy['y2'] - height + tt  )
            set_point(yy, 4, yy['x3'] + topLength - tt*2  , yy['y3']   )
            set_point(yy, 5, yy['x4']   , yy['y4'] + height - t  )
            set_point(yy, 6, yy['x5'] + bottomLength  , yy['y5']   )
            set_point(yy, 7, yy['x6']   , yy['y6'] - t  )
            set_point(yy, 8, yy['x7'] -bottomLength + t  , yy['y7']   )
            set_point(yy, 9, yy['x8']  , yy['y8'] - height +t  )
            set_point(yy, 10, yy['x9'] - topLength , yy['y9']   )
            set_point(yy, 11, yy['x10'] , yy['y10'] + height -t )
            set_point(yy, 12, yy['x11'] - bottomLength + t , yy['y11']  )

            prev_x, prev_y = yy['x1'], yy['y1']  # 첫 번째 점으로 초기화
            lastNum = 12
            for i in range(1, lastNum + 1):
                cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
                line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="0")
                prev_x, prev_y = cuyy_x, cuyy_y
            line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="0")    

            d(doc, yy['x1'] , yy['y1'], yy['x2'] , yy['y2'], 120,  direction="up", dim_style=over1000dim_style)          
            d(doc, yy['x10'] , yy['y10'], yy['x9'] , yy['y9'], 120,  direction="down", dim_style=over1000dim_style)          
            d(doc, yy['x9'] , yy['y9'], yy['x6'] , yy['y6'],  150,  direction="right", dim_style=over1000dim_style)      

            SideHatribLength = HH - 115 
            basex, basey = rx - 2200 , startYpos - 320 -  SideHatribLength              
            # 좌표 초기화
            yy = {f'x{i}': 0 for i in range(1, 31)}
            yy.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(yy, 1, basex, basey )
            set_point(yy, 2, yy['x1'] + 23.5  , yy['y1']   )
            set_point(yy, 3, yy['x2'] + 22  , yy['y2']   )
            set_point(yy, 4, yy['x3'] + 15  , yy['y3']   )
            set_point(yy, 5, yy['x4'] + 15  , yy['y4']   )
            set_point(yy, 6, yy['x5'] + 22  , yy['y5']   )
            set_point(yy, 7, yy['x6'] + 23.5 , yy['y6']   )
            set_point(yy, 8, yy['x7']   , yy['y7'] + SideHatribLength )
            set_point(yy, 9, yy['x8']  - 23.5 , yy['y8']   )
            set_point(yy, 10, yy['x9']  - 22 , yy['y9']   )
            set_point(yy, 11, yy['x10']  - 15, yy['y10']  )
            set_point(yy, 12, yy['x11']  - 15, yy['y11']  )
            set_point(yy, 13, yy['x12']  -22 , yy['y12']   )
            set_point(yy, 14, yy['x13']  - 23.5 , yy['y13']   )

            prev_x, prev_y = yy['x1'], yy['y1']  # 첫 번째 점으로 초기화
            lastNum = 14
            for i in range(1, lastNum + 1):
                cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
                line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="레이져")
                prev_x, prev_y = cuyy_x, cuyy_y
            line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="레이져")    

            # 절곡선 4개소
            bx1, by1, bc1, bd1 = yy['x13'], yy['y13'], yy['x2'], yy['y2'] 
            bx2, by2, bc2, bd2 = yy['x12'], yy['y12'], yy['x3'], yy['y3'] 
            bx3, by3, bc3, bd3 = yy['x5'], yy['y5'], yy['x10'], yy['y10'] 
            bx4, by4, bc4, bd4 = yy['x6'], yy['y6'], yy['x9'], yy['y9'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")        
            line(doc, bx2, by2, bc2, bd2, layer="절곡선")        
            line(doc, bx3, by3, bc3, bd3, layer="절곡선")        
            line(doc, bx4, by4, bc4, bd4, layer="절곡선")               

            d(doc, yy['x1'] , yy['y1'], yy['x7'] , yy['y7'], 120,  direction="down", dim_style=over1000dim_style)          
            d(doc, yy['x9'] , yy['y9'], yy['x8'] , yy['y8'], 80,  direction="up", dim_style=over1000dim_style)          
            d(doc, yy['x13'] , yy['y13'], yy['x14'] , yy['y14'], 80,  direction="up", dim_style=over1000dim_style)          
            d(doc, yy['x14'] , yy['y14'], yy['x1'] , yy['y1'],  150,  direction="left", dim_style=over1000dim_style)        

            string = f"1.6T EGI" 
            draw_Text(doc, (yy['x1'] + yy['x7'])/2 - len(string)*50/2 , yy['y14'] + 800, 50, text=string, layer='레이져')            
            string = f"기둥 모자 보강" 
            draw_Text(doc, (yy['x1'] + yy['x7'])/2 - len(string)*60/2 , yy['y14'] + 600, 60, text=string, layer='레이져')        
            string = f"{SU*2} EA" 
            draw_Text(doc, (yy['x1'] + yy['x7'])/2 - len(string)*90/2 , yy['y1'] - 300, 90, text=string, layer='레이져')

           
        ######################################################################################
        # 세로형 상판+기둥 조립용  MH (270) - 5 = 265 막판 타입 상판 보강 1
        ######################################################################################

        if jambType == '막판유':
            # 좌표 초기화
            Length = MH-5

            yy = {f'x{i}': 0 for i in range(1, 31)}
            yy.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(yy, 1, rx, startYpos )
            set_point(yy, 2, yy['x1'] + Length  , yy['y1'] )
            set_point(yy, 3, yy['x2'] , yy['y2'] + 48 )
            set_point(yy, 4, yy['x3'] , yy['y3'] + 26 )
            set_point(yy, 5, yy['x4'] - Length  , yy['y4'] )
            set_point(yy, 6, yy['x5'] , yy['y5'] - 26 )		

            prev_x, prev_y = yy['x1'], yy['y1']  # 첫 번째 점으로 초기화
            lastNum = 6
            for i in range(1, lastNum + 1):
                cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
                line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="레이져")
                prev_x, prev_y = cuyy_x, cuyy_y
            line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="레이져")  

            d(doc, yy['x1'] , yy['y1'], yy['x2'] , yy['y2'], 100, text_height=0.20, direction="down", dim_style=over1000dim_style)
            d(doc, yy['x6'] , yy['y6'], yy['x1'] , yy['y1'], 80, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x6'] , yy['y6'], yy['x5'] , yy['y5'], 80, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x1'] , yy['y1'], yy['x5'] , yy['y5'], 150, text_height=0.20, direction="left", dim_style=over1000dim_style)

            # MH방향 장공
            if(MH >= 350) :
                result = calcuteHoleArray(MH-5, 17.5, 4)
            else:
                result = calcuteHoleArray(MH-5, 17.5, 3)
            # print(result)
            sum = len(result)
            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(result):            
                xpos, ypos = yy['x5'] + hole , yy['y5'] - 14.5
                insert_block(doc, xpos, ypos, "8x20_vertical_laser")     
                if i == 0 :
                    d(doc, yy['x5'] , yy['y5'], xpos, ypos,  140, text_height=0.20, direction="up", dim_style=over1000dim_style, option='reverse')       
                else:
                    d(doc, lastx, lasty, xpos, ypos,  140 + i*80 , text_height=0.20, direction="up", dim_style=over1000dim_style, option='reverse')       
                lastx, lasty =xpos, ypos 

            # 마지막 치수선 연결
            dc(doc, yy['x4'] , yy['y4'] )
            d(doc, lastx, lasty , yy['x4'] , yy['y4'],  200, text_height=0.20,direction="right", dim_style=over1000dim_style)         # 장공 치수선

            # 절곡선 
            bx1, by1, bc1, bd1 = yy['x6'], yy['y6'], yy['x3'], yy['y3'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")                   

            string = f"2.3T PLATE" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*50/2 , yy['y1'] + 860, 50, text=string, layer='레이져')
            string = f"{floorDisplay}" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] + 680, 90, text=string, layer='레이져')        
            string = f"막판 타입 상판 보강 1" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*60/2 , yy['y1'] + 550, 60, text=string, layer='레이져')        
            string = f"{surang*2} EA" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] - 300, 90, text=string, layer='레이져')

            ###########################################################################
            # 기둥과 상판 조립 MH 좌측 단면도 'ㄱ'자 형상 
            ###########################################################################
            h = 28
            theta_deg = poleAngle
            theta_rad = math.radians(theta_deg)
            height = math.floor(h * math.cos(theta_rad)*10)/10
            base = h / math.tan(angle_rad)   # 밑면 구하기   

            aa = {f'x{i}': 0 for i in range(1, 31)}
            aa.update({f'y{i}': 0 for i in range(1, 31)})        
            # 좌표 설정
            set_point(aa, 1, yy['x6'] - 350, yy['y6'] )
            set_point(aa, 2, aa['x1'] + height  , aa['y1']  + base)
            set_point(aa, 3, aa['x2'] , aa['y2'] - 50 )
            set_point(aa, 4, aa['x3'] - 2.3 , aa['y3']  )
            set_point(aa, 5, aa['x4'] , aa['y4'] + 50 - 2.3 + 0.2 )
            set_point(aa, 6, aa['x5'] - height + 2.3, aa['y5'] - base )		

            prev_x, prev_y = aa['x1'], aa['y1']  # 첫 번째 점으로 초기화
            lastNum = 6
            for i in range(1, lastNum + 1):
                cuaa_x, cuaa_y = aa[f'x{i}'], aa[f'y{i}']
                line(doc, prev_x, prev_y, cuaa_x, cuaa_y, layer="0")
                prev_x, prev_y = cuaa_x, cuaa_y
            line(doc, prev_x, prev_y, aa['x1'], aa['y1'], layer="0")  
            
            d(doc, aa['x2'] , aa['y2'], aa['x3'] , aa['y3'], 100, text_height=0.20, direction="right", dim_style=over1000dim_style)
            dim_linear(doc, aa['x1'], aa['y1'], aa['x2'], aa['y2'],  "", dis=150,  direction="aligned" ,text_height=0.30, text_gap=0.07) 
            dim_angular(doc,   aa['x4'], aa['y4'], aa['x5'] , aa['y5'],  aa['x5'] , aa['y5'],  aa['x6'] , aa['y6'] ,  100, direction="left" )   # up,down, left, right각도를 표시할때는 기존치수선 연장선에 해줘야 제대로 값을 얻는다        

            
            ######################################################################################
            # 막판 타입 상판 보강 2   JD(245) - 42 = 203
            ######################################################################################
            # 좌표 초기화
            Length = JD-42

            yy = {f'x{i}': 0 for i in range(1, 31)}
            yy.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(yy, 1, rx, startYpos - 1500 )
            set_point(yy, 2, yy['x1'] + Length  , yy['y1'] )
            set_point(yy, 3, yy['x2'] , yy['y2'] + 48 )
            set_point(yy, 4, yy['x3'] , yy['y3'] + 26 )
            set_point(yy, 5, yy['x4'] - Length  , yy['y4'] )
            set_point(yy, 6, yy['x5'] , yy['y5'] - 26 )		

            prev_x, prev_y = yy['x1'], yy['y1']  # 첫 번째 점으로 초기화
            lastNum = 6
            for i in range(1, lastNum + 1):
                cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
                line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="레이져")
                prev_x, prev_y = cuyy_x, cuyy_y
            line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="레이져")  

            d(doc, yy['x1'] , yy['y1'], yy['x2'] , yy['y2'], 100, text_height=0.20, direction="down", dim_style=over1000dim_style)
            d(doc, yy['x6'] , yy['y6'], yy['x1'] , yy['y1'], 80, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x6'] , yy['y6'], yy['x5'] , yy['y5'], 80, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x1'] , yy['y1'], yy['x5'] , yy['y5'], 150, text_height=0.20, direction="left", dim_style=over1000dim_style)

            # JD 방향 장공
            if(JD >= 600) :
                JDholes = calcuteHoleArray(Length , 19, 4)
            elif (JD >= 900) :
                JDholes = calcuteHoleArray(Length , 19,  5)
            else:
                JDholes = calcuteHoleArray(Length , 19,  3)

            # print(result)
            sum = len(result)
            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(JDholes):            
                xpos, ypos = yy['x5'] + hole , yy['y5'] - 13
                insert_block(doc, xpos, ypos, "8x20_vertical_laser")     
                if i == 0 :
                    d(doc, yy['x5'] , yy['y5'], xpos, ypos,  140, text_height=0.15, direction="up", dim_style=over1000dim_style, option='reverse')       
                else:
                    d(doc, lastx, lasty, xpos, ypos,  140 + i*80 , text_height=0.20, direction="up", dim_style=over1000dim_style, option='reverse')       
                lastx, lasty =xpos, ypos 

            # 마지막 치수선 연결
            dc(doc, yy['x4'] , yy['y4'] )
            d(doc, lastx, lasty , yy['x4'] , yy['y4'],  180, text_height=0.20,direction="right", dim_style=over1000dim_style)         # 장공 치수선
            d(doc, lastx, lasty , yy['x3'] , yy['y3'],  260, text_height=0.20,direction="right", dim_style=over1000dim_style)         # 장공 치수선

            # 절곡선 
            bx1, by1, bc1, bd1 = yy['x6'], yy['y6'], yy['x3'], yy['y3'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")                   

            string = f"2.3T PLATE" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*50/2 , yy['y1'] + 860, 50, text=string, layer='레이져')
            string = f"{floorDisplay}" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] + 680, 90, text=string, layer='레이져')        
            string = f"막판 타입 상판 보강 2" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*60/2 , yy['y1'] + 550, 60, text=string, layer='레이져')        
            string = f"{surang*2} EA" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] - 300, 90, text=string, layer='레이져')

            ###########################################################################
            # 기둥과 상판 조립 MH 좌측 단면도 'ㄱ'자 형상 각도 직각
            ###########################################################################
            h = 28        

            aa = {f'x{i}': 0 for i in range(1, 31)}
            aa.update({f'y{i}': 0 for i in range(1, 31)})        
            # 좌표 설정
            set_point(aa, 1, yy['x6'] - 350, yy['y6'] )
            set_point(aa, 2, aa['x1'] + h  , aa['y1'] )
            set_point(aa, 3, aa['x2'] , aa['y2'] - 50 )
            set_point(aa, 4, aa['x3'] - 2.3 , aa['y3']  )
            set_point(aa, 5, aa['x4'] , aa['y4'] + 50 - 2.3  )
            set_point(aa, 6, aa['x5'] - h + 2.3, aa['y5'] )		

            set_point(aa, 7, aa['x1'] + 13, aa['y1'] + 10 )		
            set_point(aa, 8, aa['x1'] + 13, aa['y1'] - 10 )		

            prev_x, prev_y = aa['x1'], aa['y1']  # 첫 번째 점으로 초기화
            lastNum = 6
            for i in range(1, lastNum + 1):
                cuaa_x, cuaa_y = aa[f'x{i}'], aa[f'y{i}']
                line(doc, prev_x, prev_y, cuaa_x, cuaa_y, layer="0")
                prev_x, prev_y = cuaa_x, cuaa_y
            line(doc, prev_x, prev_y, aa['x1'], aa['y1'], layer="0")  

            line(doc, aa['x7'], aa['y7'],  aa['x8'], aa['y8'], layer="CL")  
            
            d(doc, aa['x1'] , aa['y1'], aa['x2'] , aa['y2'], 200, text_height=0.20, direction="up", dim_style=over1000dim_style)        
            d(doc, aa['x7'] , aa['y7'], aa['x1'] , aa['y1'],  100, text_height=0.20, direction="up", dim_style=over1000dim_style)        
            d(doc, aa['x2'] , aa['y2'], aa['x3'] , aa['y3'], 100, text_height=0.20, direction="right", dim_style=over1000dim_style)
        
            ######################################################################################
            # 기둥 조인트 보강  예시 JD(245) - 50 = 195
            ######################################################################################
            # 좌표 초기화
            Length = JD - 50

            yy = {f'x{i}': 0 for i in range(1, 31)}
            yy.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(yy, 1, rx, startYpos - 1500*2 )
            set_point(yy, 2, yy['x1'] + Length - 23*2  , yy['y1'] )
            set_point(yy, 3, yy['x2'] , yy['y2'] + 23 )
            set_point(yy, 4, yy['x3'] + 23, yy['y3']  )
            set_point(yy, 5, yy['x4'] , yy['y4'] + 48)
            set_point(yy, 6, yy['x5'] - Length , yy['y5']  )		
            set_point(yy, 7, yy['x6'] , yy['y6'] - 48 )		
            set_point(yy, 8, yy['x7'] + 23, yy['y7']  )		

            prev_x, prev_y = yy['x1'], yy['y1']  # 첫 번째 점으로 초기화
            lastNum = 8
            for i in range(1, lastNum + 1):
                cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
                line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="레이져")
                prev_x, prev_y = cuyy_x, cuyy_y
            line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="레이져")  

            d(doc, yy['x7'] , yy['y7'], yy['x4'] , yy['y4'], 150, text_height=0.20, direction="down", dim_style=over1000dim_style)
            d(doc, yy['x7'] , yy['y7'], yy['x6'] , yy['y6'], 80, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x7'] , yy['y7'], yy['x1'] , yy['y1'], 80, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x6'] , yy['y6'], yy['x1'] , yy['y1'], 160, text_height=0.20, direction="left", dim_style=over1000dim_style)

            # JD 방향 장공
            if(JD >= 600) :
                JDholes = calcuteHoleArray(Length , 15, 4)
            elif (JD >= 900) :
                JDholes = calcuteHoleArray(Length , 15,  5)
            else:
                JDholes = calcuteHoleArray(Length , 15,  3)

            # print(result)
            sum = len(result)
            # 각 변수에 대해 값이 0보다 크고 limit_point에 없으면 rightholes에 추가
            for i, hole in enumerate(JDholes):            
                xpos, ypos = yy['x6'] + hole , yy['y6'] - 13
                insert_block(doc, xpos, ypos, "8x20_horizontal_laser")     
                if i == 0 :
                    d(doc, yy['x6'] , yy['y6'], xpos, ypos,  140, text_height=0.15, direction="up", dim_style=over1000dim_style, option='reverse')       
                else:
                    d(doc, lastx, lasty, xpos, ypos,  140 + i*80 , text_height=0.20, direction="up", dim_style=over1000dim_style, option='reverse')       
                lastx, lasty =xpos, ypos 

            # 마지막 치수선 연결
            dc(doc, yy['x5'] , yy['y5'] )
            d(doc, lastx, lasty , yy['x5'] , yy['y5'],  180, text_height=0.20,direction="right", dim_style=over1000dim_style)         # 장공 치수선
            d(doc, lastx, lasty , yy['x4'] , yy['y4'],  260, text_height=0.20,direction="right", dim_style=over1000dim_style)         # 장공 치수선

            # 절곡선 
            bx1, by1, bc1, bd1 = yy['x8'], yy['y8'], yy['x3'], yy['y3'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")                   

            string = f"2.3T PLATE" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*50/2 , yy['y1'] + 860, 50, text=string, layer='레이져')
            string = f"{floorDisplay}" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] + 680, 90, text=string, layer='레이져')        
            string = f"기둥 조인트 보강" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*60/2 , yy['y1'] + 550, 60, text=string, layer='레이져')        
            string = f"{surang*2} EA" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] - 300, 90, text=string, layer='레이져')

            ###########################################################################
            # 기둥과 조인트 보강 'ㄴ'자 형상 각도 직각 단면도
            ###########################################################################
            h = 25
            w = 50       
            tt = 2.3 

            aa = {f'x{i}': 0 for i in range(1, 31)}
            aa.update({f'y{i}': 0 for i in range(1, 31)})        
            # 좌표 설정
            set_point(aa, 1, yy['x6'] - 450, yy['y6'] )
            set_point(aa, 2, aa['x1'] + h  , aa['y1'] )
            set_point(aa, 3, aa['x2'] , aa['y2'] + w )
            set_point(aa, 4, aa['x3'] - tt , aa['y3']  )
            set_point(aa, 5, aa['x4'] , aa['y4'] - w + tt  )
            set_point(aa, 6, aa['x5'] - h + tt,  aa['y5'] )		

            set_point(aa, 7, aa['x3'] - 10, aa['y3'] - 13 )		
            set_point(aa, 8, aa['x3'] + 10, aa['y3'] - 13 )		        

            prev_x, prev_y = aa['x1'], aa['y1']  # 첫 번째 점으로 초기화
            lastNum = 6
            for i in range(1, lastNum + 1):
                cuaa_x, cuaa_y = aa[f'x{i}'], aa[f'y{i}']
                line(doc, prev_x, prev_y, cuaa_x, cuaa_y, layer="0")
                prev_x, prev_y = cuaa_x, cuaa_y
            line(doc, prev_x, prev_y, aa['x1'], aa['y1'], layer="0")  

            line(doc, aa['x7'], aa['y7'],  aa['x8'], aa['y8'], layer="CL")  
            
            d(doc, aa['x1'] , aa['y1'], aa['x2'] , aa['y2'], 100, text_height=0.20, direction="down", dim_style=over1000dim_style)                
            d(doc, aa['x2'] , aa['y2'], aa['x3'] , aa['y3'], 100, text_height=0.20, direction="right", dim_style=over1000dim_style)
            
            
        ######################################################################################
        # 막판유 기둥 하부 보강   JD - 80   * 2 적용
        ######################################################################################
        # 좌표 초기화
        Length = JD - 80

        yy = {f'x{i}': 0 for i in range(1, 31)}
        yy.update({f'y{i}': 0 for i in range(1, 31)})

        # 좌표 설정
        set_point(yy, 1, rx, startYpos - 1500*3 )
        set_point(yy, 2, yy['x1'] + Length , yy['y1'] )
        set_point(yy, 3, yy['x2'] , yy['y2'] + 18 )
        set_point(yy, 4, yy['x3'] , yy['y3'] + 48 )
        set_point(yy, 5, yy['x4'] - Length , yy['y4'] )
        set_point(yy, 6, yy['x5'] , yy['y5'] - 48 )		

        prev_x, prev_y = yy['x1'], yy['y1']  # 첫 번째 점으로 초기화
        lastNum = 6
        for i in range(1, lastNum + 1):
            cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
            line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="레이져")
            prev_x, prev_y = cuyy_x, cuyy_y
        line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="레이져")  

        d(doc, yy['x1'] , yy['y1'], yy['x2'] , yy['y2'], 150,  direction="down", dim_style=over1000dim_style)
        d(doc, yy['x6'] , yy['y6'], yy['x5'] , yy['y5'], 80, direction="left", dim_style=over1000dim_style)
        d(doc, yy['x6'] , yy['y6'], yy['x1'] , yy['y1'], 80,  direction="left", dim_style=over1000dim_style)
        d(doc, yy['x2'] , yy['y2'], yy['x4'] , yy['y4'], 160,  direction="right", dim_style=over1000dim_style)

        # 절곡선 
        bx1, by1, bc1, bd1 = yy['x6'], yy['y6'], yy['x3'], yy['y3'] 
        line(doc, bx1, by1, bc1, bd1, layer="절곡선")                   

        string = f"2.3T PLATE" 
        draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*50/2 , yy['y1'] + 660, 50, text=string, layer='레이져')
        string = f"{floorDisplay}" 
        draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] + 480, 90, text=string, layer='레이져')        
        string = f"기둥 하부 보강"  
        draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*60/2 , yy['y1'] + 350, 60, text=string, layer='레이져')        
        string = f"{surang*2} EA" 
        draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] - 300, 90, text=string, layer='레이져')

        ###########################################################################
        # 기둥과 상판 조립 MH 좌측 단면도 'ㄴ'자 형상 각도 직각
        ###########################################################################
        h = 20
        w = 50       
        tt = 2.3 

        aa = {f'x{i}': 0 for i in range(1, 31)}
        aa.update({f'y{i}': 0 for i in range(1, 31)})        
        # 좌표 설정
        set_point(aa, 1, yy['x6'] - 350, yy['y6'] )
        set_point(aa, 2, aa['x1'] + h  , aa['y1'] )
        set_point(aa, 3, aa['x2'] , aa['y2'] + w )
        set_point(aa, 4, aa['x3'] - tt , aa['y3']  )
        set_point(aa, 5, aa['x4'] , aa['y4'] - w + tt  )
        set_point(aa, 6, aa['x5'] - h + tt,  aa['y5'] )		

        prev_x, prev_y = aa['x1'], aa['y1']  # 첫 번째 점으로 초기화
        lastNum = 6
        for i in range(1, lastNum + 1):
            cuaa_x, cuaa_y = aa[f'x{i}'], aa[f'y{i}']
            line(doc, prev_x, prev_y, cuaa_x, cuaa_y, layer="0")
            prev_x, prev_y = cuaa_x, cuaa_y
        line(doc, prev_x, prev_y, aa['x1'], aa['y1'], layer="0")  
        
        d(doc, aa['x1'] , aa['y1'], aa['x2'] , aa['y2'], 130, text_height=0.20, direction="down", dim_style=over1000dim_style)                
        d(doc, aa['x2'] , aa['y2'], aa['x3'] , aa['y3'], 100, text_height=0.20, direction="right", dim_style=over1000dim_style)
               
        
        ######################################################################################
        # 기둥 띄보강  전개도 JD - 68 길이 : JD 245인 경우 177
        # 펴진 띠보강으로 수정요청 25/08/18
        # b = JD - 5 - 19.8 - 56.4 # 막판무 공식
        # xangle = 10  # 도 단위

        # # 빗변 x 계산 (밑변과 각도 이용)
        # x = b / math.cos(math.radians(xangle))

        # # 정수 반올림
        # x_rounded = round(x)

        # print("밑변 b =", b)
        # print("빗변 x =", x_rounded)        
        ######################################################################################
        # 좌표 초기화
        b = 0
        if jambType == '막판유':
            b = JD + 10 - 20 - 56.1 + (25 - A) 
        elif jambType =='막판무':
            b = JD - 40 - 56 + (45 - A) # 막판무 공식 

        # 빗변 x 계산 (밑변과 각도 이용)
        try:
            # poleAngle이 딕셔너리인 경우 0을 사용하고, 아니면 float로 변환
            angle_value = 0 if isinstance(poleAngle, dict) else float(poleAngle or 0)
            x = float(b) / math.cos(math.radians(angle_value))
        except (TypeError, ValueError, ZeroDivisionError):
            x = 0
        # 소수점이면 무조건 올림
        if isinstance(x, float):
            Length = int(x) if x == int(x) else int(x) + 1
        else:
            Length = 0

        yy = {f'x{i}': 0 for i in range(1, 31)}
        yy.update({f'y{i}': 0 for i in range(1, 31)})

        # 좌표 설정 (4개 점, 각도 없이 직사각형 형태)
        set_point(yy, 1, rx, startYpos - 1500*4 )
        set_point(yy, 2, yy['x1'] + Length, yy['y1'])
        set_point(yy, 3, yy['x2'], yy['y2'] + 60)
        set_point(yy, 4, yy['x1'], yy['y1'] + 60)

        prev_x, prev_y = yy['x1'], yy['y1']  # 첫 번째 점으로 초기화
        lastNum = 4
        for i in range(1, lastNum + 1):
            cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
            line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="레이져")
            prev_x, prev_y = cuyy_x, cuyy_y
        line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="레이져")  

        # 치수선 표시 (각도 없이 단순 사각형)
        d(doc, yy['x1'], yy['y1'], yy['x2'], yy['y2'], 130, direction="down", dim_style=over1000dim_style)
        d(doc, yy['x2'], yy['y2'], yy['x3'], yy['y3'], 100, direction="right", dim_style=over1000dim_style)
        d(doc, yy['x3'], yy['y3'], yy['x4'], yy['y4'], 80, direction="up", dim_style=over1000dim_style)
        d(doc, yy['x4'], yy['y4'], yy['x1'], yy['y1'], 80, direction="left", dim_style=over1000dim_style)

        string = f"2.3T PLATE" 
        draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*50/2 , yy['y1'] + 860, 50, text=string, layer='레이져')
        string = f"{floorDisplay}" 
        draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] + 720, 90, text=string, layer='레이져')        
        string = f"기둥 띄보강"  
        draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*60/2 , yy['y1'] + 590, 60, text=string, layer='레이져')        
        string = f"{surang*6} EA" 
        draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] - 300, 90, text=string, layer='레이져')

        ###########################################################################
        # 띄보강 단면도: 4개의 점으로 평판으로 수정
        ###########################################################################
        tmpx, tmpy = yy['x1'], yy['y1'] + 400
        # 좌표 초기화
        pp = {f'x{i}': 0 for i in range(1, 31)}
        pp.update({f'y{i}': 0 for i in range(1, 31)})

        # 4개의 점으로 평판 사각형 설정
        set_point(pp, 1, tmpx, tmpy)
        set_point(pp, 2, pp['x1'] + Length, pp['y1'])
        set_point(pp, 3, pp['x2'], pp['y2'] - 9.4)
        set_point(pp, 4, pp['x1'], pp['y1'] - 9.4)

        prev_x, prev_y = pp['x1'], pp['y1']  # 첫 번째 점으로 초기화
        lastNum = 4
        for i in range(1, lastNum + 1):
            cupp_x, cupp_y = pp[f'x{i}'], pp[f'y{i}']
            line(doc, prev_x, prev_y, cupp_x, cupp_y, layer="0")
            prev_x, prev_y = cupp_x, cupp_y
        line(doc, prev_x, prev_y, pp['x1'], pp['y1'], layer="0")
        d(doc, pp['x1'], pp['y1'], pp['x2'], pp['y2'], 100, text_height=0.30, direction="up", dim_style=over1000dim_style)


        if jambType == '막판무' :
            ######################################################################################
            # 기둥 조인트 보강  JD - 18  막판유와 상하 반전 모양임 단면도 역시 반전
            ######################################################################################
            # 좌표 초기화
            Length = JD - 18
            yy = {f'x{i}': 0 for i in range(1, 31)}
            yy.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 설정
            set_point(yy, 1, rx, startYpos - 1500*2 )
            set_point(yy, 2, yy['x1'] + Length  , yy['y1'] )
            set_point(yy, 3, yy['x2'] , yy['y2'] + 43 )
            set_point(yy, 4, yy['x3'] - 18, yy['y3']  )
            set_point(yy, 5, yy['x4'] , yy['y4'] + 5 )
            set_point(yy, 6, yy['x5'] , yy['y5'] +26 )		
            set_point(yy, 7, yy['x6'] - Length + 18*2, yy['y6']  )		
            set_point(yy, 8, yy['x7'] , yy['y7'] - 26 )		
            set_point(yy, 9, yy['x8'] , yy['y8'] - 5)		
            set_point(yy, 10, yy['x9'] - 18 , yy['y9'] )		

            prev_x, prev_y = yy['x1'], yy['y1']  # 첫 번째 점으로 초기화
            lastNum = 10
            for i in range(1, lastNum + 1):
                cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
                line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="레이져")
                prev_x, prev_y = cuyy_x, cuyy_y
            line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="레이져")  

            d(doc, yy['x1'] , yy['y1'], yy['x2'] , yy['y2'], 150, text_height=0.20, direction="down", dim_style=over1000dim_style)
            d(doc, yy['x3'] , yy['y3'], yy['x2'] , yy['y2'], 220, text_height=0.20, direction="right", dim_style=over1000dim_style)
            d(doc, yy['x3'] , yy['y3'], yy['x6'] , yy['y6'], 220, text_height=0.20, direction="right", dim_style=over1000dim_style)

            d(doc, yy['x8'] , yy['y8'], yy['x7'] , yy['y7'], 120, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x8'] , yy['y8'], yy['x1'] , yy['y1'], 120, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x7'] , yy['y7'], yy['x1'] , yy['y1'], 220, text_height=0.20, direction="left", dim_style=over1000dim_style)

            # JD 방향 장공
            if(JD >= 600) :
                JDholes = calcuteHoleArray(Length , 42, 4)
            elif (JD >= 900) :
                JDholes = calcuteHoleArray(Length , 42,  5)
            else:
                JDholes = calcuteHoleArray(Length , 42,  3)

            for i, hole in enumerate(JDholes):            
                xpos, ypos = yy['x1'] + hole , yy['y6'] - 13
                insert_block(doc, xpos, ypos, "8x20_horizontal_laser")     
                if i == 0 :
                    d(doc, yy['x7'] , yy['y7'], xpos, ypos,  100, text_height=0.15, direction="up", dim_style=over1000dim_style, option='reverse')       
                else:
                    d(doc, lastx, lasty, xpos, ypos,  100 + i*80 , text_height=0.20, direction="up", dim_style=over1000dim_style, option='reverse')       
                lastx, lasty =xpos, ypos 

            # 마지막 치수선 연결
            dc(doc, yy['x6'] , yy['y6'] )
            d(doc, lastx, lasty , yy['x6'] , yy['y6'],  150, text_height=0.20,direction="right", dim_style=over1000dim_style)         # 장공 치수선

            # 절곡선 
            bx1, by1, bc1, bd1 = yy['x8'], yy['y8'], yy['x5'], yy['y5'] 
            line(doc, bx1, by1, bc1, bd1, layer="절곡선")                   

            string = f"2.3T PLATE" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*50/2 , yy['y1'] + 860, 50, text=string, layer='레이져')
            string = f"{floorDisplay}" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] + 680, 90, text=string, layer='레이져')        
            string = f"기둥 조인트 보강" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*60/2 , yy['y1'] + 550, 60, text=string, layer='레이져')        
            string = f"{surang*2} EA" 
            draw_Text(doc, (yy['x1'] + yy['x2'])/2 - len(string)*90/2 , yy['y1'] - 300, 90, text=string, layer='레이져')

            ###########################################################################
            # 기둥과 조인트 보강 'ㄴ'자 형상 각도 직각 단면도  28 H는 막판은 25
            ###########################################################################
            h = 28
            w = 50       
            tt = 2.3 

            aa = {f'x{i}': 0 for i in range(1, 31)}
            aa.update({f'y{i}': 0 for i in range(1, 31)})        
            # 좌표 설정
            set_point(aa, 1, yy['x7'] - 550, yy['y7'] )
            set_point(aa, 2, aa['x1'] + h  , aa['y1'] )
            set_point(aa, 3, aa['x2'] , aa['y2'] - w )
            set_point(aa, 4, aa['x3'] - tt , aa['y3']  )
            set_point(aa, 5, aa['x4'] , aa['y4'] + w - tt  )
            set_point(aa, 6, aa['x5'] - h + tt,  aa['y5'] )		

            set_point(aa, 7, aa['x1'] + 13, aa['y1'] - 10 )		
            set_point(aa, 8, aa['x1'] + 13, aa['y1'] + 10 )		        

            prev_x, prev_y = aa['x1'], aa['y1']  # 첫 번째 점으로 초기화
            lastNum = 6
            for i in range(1, lastNum + 1):
                cuaa_x, cuaa_y = aa[f'x{i}'], aa[f'y{i}']
                line(doc, prev_x, prev_y, cuaa_x, cuaa_y, layer="0")
                prev_x, prev_y = cuaa_x, cuaa_y
            line(doc, prev_x, prev_y, aa['x1'], aa['y1'], layer="0")  

            line(doc, aa['x7'], aa['y7'],  aa['x8'], aa['y8'], layer="CL")  
            
            d(doc, aa['x1'] , aa['y1'], aa['x2'] , aa['y2'], 200, text_height=0.20, direction="up", dim_style=over1000dim_style)            
            d(doc, aa['x7'] , aa['y7'], aa['x1'] , aa['y1'],  100, text_height=0.20, direction="up", dim_style=over1000dim_style)            
            d(doc, aa['x2'] , aa['y2'], aa['x3'] , aa['y3'], 100, text_height=0.20, direction="right", dim_style=over1000dim_style)

############################################################################################################################################################
@Gooey(encoding='utf-8', program_name='다완테크 신규쟘 자동작도', tabbed_groups=True, navigation='Tabbed', show_success_modal=False,  default_size=(1200, 600))

def main():
    global args  #수정할때는 global 선언이 필요하다. 단순히 읽기만 할때는 필요없다.        
    global exit_program, program_message, text_style_name    
    global SU
    # 전역 데이터 딕셔너리 초기화    
    global global_data, doc, msp

    # 현재 날짜와 시간을 가져옵니다.
    current_datetime = datetime.now()
    global_data["formatted_date"] = current_datetime.strftime('%Y-%m-%d')
    global_data["current_time"] = current_datetime.strftime("%H%M%S")

    # .xlsm 파일이 없을 경우 오류 메시지를 출력하고 실행을 중단
    if not xlsm_files:
        error_message = ".xlsm 파일이 excel파일 폴더에 없습니다. 확인바랍니다."
        show_custom_error(error_message)
        sys.exit(1)

    # 찾은 .xlsm 파일 목록 순회
    for file_path in xlsm_files:
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            error_message = f"엑셀 파일을 열 수 없습니다: {str(e)}"
            show_custom_error(error_message)

        sheet_name = "발주"  # 원하는 시트명
        sheet = workbook[sheet_name]

        try:
            if readfile is not None:
                doc = readfile(os.path.join(dxf_saved_file, 'dawan_style.dxf'))
                msp = doc.modelspace()
            else:
                raise AttributeError("readfile 함수를 사용할 수 없습니다.")
        except (AttributeError, FileNotFoundError) as e:
            # ezdxf 버전 호환성 문제 해결 또는 파일이 없는 경우
            try:
                if new is not None:
                    doc = new()
                    if readfile is not None and os.path.exists(os.path.join(dxf_saved_file, 'dawan_style.dxf')):
                        doc = readfile(os.path.join(dxf_saved_file, 'dawan_style.dxf'))
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

        # 엑셀 셀 값을 global_data와 전역 변수에 저장
        for cell_ref, var_name in variable_names.items():
            value = read_excel_value(sheet, cell_ref)
            global_data[var_name] = value  # global_data에 저장
            globals()[var_name] = value  # 전역 변수로도 저장

        global_data["rows_data"] =read_excel_rows(sheet)

        # # 데이터 출력 (global_data에 저장된 데이터 확인)
        # for key, value in global_data.items():
        #     if key == "rows_data":  # 행 데이터는 따로 출력
        #         print(f"{key}:")
        #         for index, row in enumerate(value, start=1):
        #             print(f"  Row {index}: {row}")
        #     else:
        #         print(f"{key}: {value}")

        # 초기 변수 할당
        thickness = float(re.sub("[A-Z]", "", global_data["thickness_string"]))
        global_data["WorkTitle"] = f"업체명: {global_data['company']}, 현장명: {global_data['workplace']}, thickness: {thickness}"

        # 와이드쟘 2CO 신규쟘
        execute_wide()

        # 파일 이름 생성
        invalid_chars = '<>:"/\\|?*'
        cleaned_file_name = re.sub(
            f'[{re.escape(invalid_chars)}]', '',
            f"{global_data['company']}_{global_data['workplace']}_{global_data['thickness_string']}_{global_data['current_time']}"
        )
        script_directory = os.path.dirname(os.path.abspath(__file__))
        full_file_path = os.path.join(script_directory, f"c:/dawan/작업완료/{cleaned_file_name}.dxf")
        global_data["file_name"] = full_file_path

        # gooey 호출하는 부분 (화면에 창을 보이게 하기)
        exit_program = False
        program_message = \
            '''
        프로그램 실행결과입니다.
        -------------------------------------
        {0}
        -------------------------------------
        이용해 주셔서 감사합니다.
        '''    
        args = parse_arguments()            

        # 서버에 로그인 기록을 남김
        log_login()

        # 도면 저장
        doc.saveas(global_data["file_name"])
        print(f" 저장 파일명: '{global_data['file_name']}' 저장 완료!")

if __name__ == '__main__':
    main()
