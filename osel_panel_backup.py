# 2025/10/01 ?�성?�엘 ?�넬 ?�동?�도 ?�로그램 ?�작 ?�작
import math
try:
    import ezdxf
    from ezdxf.filemanagement import readfile, new
except ImportError as e:
    print(f"ezdxf 모듈 ?�포???�류: {e}")
    # ?��??�포???�도
    try:
        import ezdxf
        readfile = getattr(ezdxf, 'readfile', None)
        new = getattr(ezdxf, 'new', None)
        if readfile is None or new is None:
            print("ezdxf 모듈?�서 ?�요???�수�?찾을 ???�습?�다.")
    except Exception as e2:
        print(f"ezdxf 모듈 로드 ?�패: {e2}")
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

# ?�역 변??초기??if True:
    global_data = {}
    BasicXscale, BasicYscale,TargetXscale,TargetYscale, frame_scale = 0,0,0,0,0
    frameXpos = 0
    frameYpos = 0    
    thickness = 0
    selected_dimstyle = ''
    over1000dim_style = ''
    br = 0  # bending rate ?�신??    saved_DimXpos = 0
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

    # ?�역 변??초기??    jambType = None
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
    
    # HPI 관???�역 변??초기??    HPIHoleWidth = 0
    HPIHoleHeight = 0
    HPIHeight = 0
    HPIholegap = 0

    # ?�역 변???�언 �?초기??    for i in range(1, 31):
        globals()[f'x{i}'] = 0
        globals()[f'y{i}'] = 0        

    # ?�역 변???�언 �?초기??    for i in range(1, 12):
        globals()[f'P{i}_platewidth'] = 0
        globals()[f'P{i}_plateheight'] = 0
        globals()[f'P{i}_hole'] = []

# 기본 ?�정
if True:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')    

    # 경고 메시지 ?�터�?    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")

    # ?�더 ?�의 모든 .xlsm ?�일??검??    application_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    # excel_saved_file = os.path.join(application_path, 'panel_excel')
    # xlsm_files = glob.glob(os.path.join(excel_saved_file, '*.xlsm'))
    # ?��? 경로�?지??    excel_saved_file = 'c:/python/osel/excel?�일'
    xlsm_files = glob.glob(os.path.join(excel_saved_file, '*.xlsx'))
    license_file_path = os.path.join(application_path, 'data', 'hdsettings.json') # ?�드?�스??고유번호 ?�식

    # DXF ?�일 로드
    dxf_saved_file = 'c:/python/osel/dimstyle'
    try:
        if readfile is not None:
            doc = readfile(os.path.join(dxf_saved_file, 'style.dxf'))    
            msp = doc.modelspace()
        else:
            raise AttributeError("readfile ?�수�??�용?????�습?�다.")
    except (AttributeError, FileNotFoundError) as e:
        # ezdxf 버전 ?�환??문제 ?�결 ?�는 ?�일???�는 경우
        try:
            if new is not None:
                doc = new()
                if readfile is not None and os.path.exists(os.path.join(dxf_saved_file, 'style.dxf')):
                    doc = readfile(os.path.join(dxf_saved_file, 'style.dxf'))
                msp = doc.modelspace()
            else:
                raise AttributeError("new ?�수�??�용?????�습?�다.")
        except Exception as e:
            print(f"DXF ?�일 로드 ?�류: {e}")
            # ??DXF 문서 ?�성
            if new is not None:
                doc = new()
                msp = doc.modelspace()
            else:
                print("ezdxf 모듈???�용?????�습?�다. ?�로그램??종료?�니??")
                sys.exit(1)
    except Exception as e:
        print(f"DXF ?�일 로드 ?�류: {e}")
        # ??DXF 문서 ?�성
        if new is not None:
            doc = new()
            msp = doc.modelspace()
        else:
            print("ezdxf 모듈???�용?????�습?�다. ?�로그램??종료?�니??")
            sys.exit(1)

    # TEXTSTYLE ?�의
    text_style_name = 'H'  # ?�하???�스???��????�름
    if text_style_name not in doc.styles:
        text_style = doc.styles.new(
            name=text_style_name,
            dxfattribs={
                'font': 'Arial.ttf',  # TrueType 글�??�일�?           
            }
        )
    else:
        text_style = doc.styles.get(text_style_name)

    # �?번째 .xlsm ?�일?�서 W1 ?� �??�기
    if not xlsm_files:
        print(".xlsm ?�일??excel?�일 ?�더???�습?�다. ?�인 바랍?�다.")
        sys.exit(1)

    workbook = openpyxl.load_workbook(xlsm_files[0], data_only=True)
    sheet_name = '발주'
    sheet = workbook[sheet_name]
    dim_style_key = 'dim1'

    # dimstyle 매핑 ?�정
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
    # selected_dimstyle�?over1000dim_style ?�정
    selected_dimstyle = dimstyle_map.get(dim_style_key, 'mydim1')  # 기본값�? 'mydim1'
    over1000dim_style = over1000dim_style_map.get(dim_style_key, 'over1000dim1')  # 기본값�? 'over1000dim1'

def read_manufacturing_results(sheet, start_row=3):
    # ?�작?�출결과 ?�트???�과 변??매핑
    column_mapping = {
        "A": "number",              # 번호
        "B": "unique_id",           # 고유번호
        "C": "site_name",           # ?�장�?        "D": "measurement_date",    # 측정?�자
        "E": "measurer",           # 측정??        "F": "car_width",          # �??��? W
        "G": "car_depth",          # �??��? D
        "H": "car_height",         # �??��? H
        "I": "interior_material",  # ?�장?�질
        "J": "material_thickness", # ?�질 ?�께
        "K": "panel_number",       # ?�널 번호
        "L": "manufacturing_count", # ?�작 ?�??        "M": "panel_type",         # ?�널 ?�??        "N": "manufacturing_width", # ?�작??        "O": "manufacturing_height", # ?�작?�이
        "P": "perforation_width",  # ?��?가�?        "Q": "perforation_length", # ?��??�로
        "R": "perforation_height", # ?��??�이(밑기준)
        "S": "distance_from_entrance" # ?�구방향?�서 ?�어
    }

    # 결과 리스??초기??    manufacturing_data = []

    # ??반복 (A??기�??�로 비어?�을 ?�까지 반복)
    row = start_row
    while True:
        # A???�이???�인 (번호가 비어?�으�?종료)
        cell_value = sheet[f"A{row}"].value
        if cell_value is None or cell_value == "":  # A?�이 비어?�으�?종료
            break

        # ?�재 ???�이?��? ?�셔?�리�??�??        row_data = {}
        for col, var_name in column_mapping.items():
            cell_ref = f"{col}{row}"
            row_data[var_name] = sheet[cell_ref].value  # ?�당 ?� 값을 ?�셔?�리???�??
        # ?�자 ?�드 처리 (None?�면 0?�로 변??
        numeric_fields = ['number', 'unique_id', 'car_width', 'car_depth', 'car_height', 
                         'material_thickness', 'panel_number', 'manufacturing_count',
                         'manufacturing_width', 'manufacturing_height', 'perforation_width',
                         'perforation_length', 'perforation_height', 'distance_from_entrance']
        
        for field in numeric_fields:
            if row_data.get(field) is None:
                row_data[field] = 0
            else:
                # ?�자�?변???�도
                try:
                    row_data[field] = float(row_data[field])
                except (ValueError, TypeError):
                    row_data[field] = 0

        # 결과 리스?�에 추�?
        manufacturing_data.append(row_data)

        # ?�음 ?�으�??�동
        row += 1

    # ?�작?�출결과 ?�이??출력
    print("=== ?�작?�출결과 ?�이??출력 ===")
    for i, data in enumerate(manufacturing_data, 1):
        unique_id = data.get('unique_id', 0)
        site_name = data.get('site_name', 'N/A')
        panel_number = data.get('panel_number', 0)
        manufacturing_width = data.get('manufacturing_width', 0)
        manufacturing_height = data.get('manufacturing_height', 0)
        print(f"??{i}: 고유번호={unique_id}, ?�장�?{site_name}, ?�널번호={panel_number}, ?�작??{manufacturing_width}, ?�작?�이={manufacturing_height}")
    
    return manufacturing_data

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
    # PHP ?�일??URL ?�버?�서???�이?��? ?�?�한?? ?�체 ?�이?��? 기록?�다.
    workplace = global_data["WorkTitle"]
    url = f"https://8440.co.kr/autopanel/savelog.php?company=?�완?�크&content=?�넬?�작_{workplace}_{SU}"

    # HTTP ?�청 보내�?    response = requests.get(url)
    
    # ?�청???�공?�는지 ?�인
    if response.status_code == 200:
        # print("logged successfully.")
        print(response.json())
    else:
        print("Failed to log login time.")
        print(response.text)   
        exit(1)

def show_custom_error(message):
    root = tk.Tk()
    root.withdraw()  # 메인 ?�도???�기�?
    error_window = tk.Toplevel()
    error_window.title("?�동?�도 ?�류 ?�림")
    
    custom_font = font.Font(size=15)  # ?�트 ?�기 ?�정
    label = tk.Label(error_window, text=message, font=custom_font)
    label.pack(padx=20, pady=20)

    def close_program():
        sys.exit(1)  # 모든 ?�로그램 강제 종료

    close_button = tk.Button(error_window, text="?�인", command=close_program)
    close_button.pack(pady=10)

    # 창을 ?�면 중앙???�치?�키�?    error_window.update_idletasks()
    width = error_window.winfo_width()
    height = error_window.winfo_height()
    x = (error_window.winfo_screenwidth() // 2) - (width // 2)
    y = (error_window.winfo_screenheight() // 2) - (height // 2)
    error_window.geometry(f"{width}x{height}+{x}+{y}")    
    error_window.mainloop()    
def is_number(var):
    # 변?��? ?�자?��? ?�인?�는 ?�수
    if isinstance(var, (int, float)):
        return True
    elif isinstance(var, str):
        try:
            float(var)  # 변?�을 ?�도?�여 ?�자?��? ?�인
            return True
        except ValueError:
            return False
    return False
def save_file(company, workplace):
    # ?�재 ?�간 가?�오�?    current_time = datetime.now().strftime("%Y%m%d%H%M%S")

    # ?�일 ?�름???�용?????�는 문자 ?�의
    invalid_chars = '<>:"/\\|?*'
    # ?�규?�을 ?�용?�여 ?�효?��? ?��? 문자 ?�거
    cleaned_file_name = re.sub(f'[{re.escape(invalid_chars)}]', '', f"{company}_{workplace}_{thickness_string}_{drawnby}_{current_time}")

    # 결과 ?�일???�?�될 ?�렉?�리
    output_directory = "c:/python/osel/?�업?�료"

    # ?�렉?�리가 존재?��? ?�으�??�성
    os.makedirs(output_directory, exist_ok=True)

    # 결과 ?�일 ?�름
    file_name = f"{cleaned_file_name}.dxf"
    # ?�체 ?�일 경로 ?�성
    full_file_path = os.path.join(output_directory, file_name)

    # ?�일 경로 반환
    return full_file_path    
def read_excel_value(sheet, cell):
    value = sheet[cell].value
    if isinstance(value, str):
        try:
            float_value = float(value)  # 문자?�을 float�?변???�도
            if float_value.is_integer():  # ?�수?�이 ?�는 경우
                return int(float_value)
            else:  # ?�수?�이 ?�는 경우
                return float_value
        except ValueError:
            return value  # 변?�할 ???�는 경우 ?�래 문자??반환
    return value
def write_log(message):
    logging.info(message)    
def parse_arguments_settings():
    parser = GooeyParser()
    settings = parser.add_argument_group('?�정')
    settings.add_argument('--config', action='store_true', default=True,  help='?�이?�스')
    settings.add_argument('--password', widget='PasswordField', gooey_options={'visible': True, 'show_label': True}, help='비�?번호')

    return parser.parse_args()
def parse_arguments():
    parser = GooeyParser()

    group1 = parser.add_argument_group('카판옵션')
    # group1.add_argument('--opt1', action='store_true',  help='기본')
    # group1.add_argument('--opt2', action='store_true',  help='?�장?�')
    # group1.add_argument('--opt3', action='store_true',  help='쪽쟘 ?�판???�운??추영?�소??')    
    group1.add_argument('--opt1', action='store_true', default=True, help='기본')

    settings = parser.add_argument_group('?�정')
    settings.add_argument('--config', action='store_true', help='?�이?�스')
    settings.add_argument('--password', widget='PasswordField', gooey_options={'visible': True, 'show_label': True}, help='비�?번호')
    
    return parser.parse_args()
def display_message():
    message = program_message.format('\n'.join(sys.argv[1:])).split('\n')
    delay = 1.5 / len(message)

    for line in message:
        print(line)
        time.sleep(delay)
def load_env_settings():
# ?�경?�정 가?�오�??�드공유 번호)    
    try:
        with open(license_file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            return data.get("DiskID")
    except FileNotFoundError:
        return None
def get_current_disk_id():
    return os.popen('wmic diskdrive get serialnumber').read().strip()
def validate_or_default(value):
# None?�면 0??리턴?�는 ?�수    
    if value is None:
        return 0
    return value
def find_intersection(start1, end1, start2, end2):
    # 교차??계산 (??직선???�직 �??�평?�로 만나??경우???�?�서�?
    if start1[0] == end1[0]:
        return (start1[0], start2[1])
    else:
        return (start2[0], start1[1])
def calculate_fillet_point(center, point, radius):
    # ?�렛 ?�점 계산
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

    # 교차??찾기
    intersection_point = find_intersection(start1, end1, start2, end2)

    # ?�렛 ?�점 계산
    point1 = calculate_fillet_point(intersection_point, end1, radius)
    point2 = calculate_fillet_point(intersection_point, end2, radius)

    # 각도 계산
    start_angle = calculate_angle(intersection_point, point1)
    end_angle = calculate_angle(intersection_point, point2)

    # ?�렛 ?�호 그리�?    msp.add_arc(
        center=intersection_point,
        radius=radius,
        start_angle=start_angle,
        end_angle=end_angle,
        dxfattribs={'layer': '?�이??},
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

    # ?�의 중심 계산
    center = calculate_circle_center(midpoint, radius, point1, point2)

    # 각도 계산
    start_angle = calculate_angle(center, point1)
    end_angle = calculate_angle(center, point2)

    # ?�크 그리�?    msp.add_arc(
        center=center,
        radius=radius,
        start_angle=start_angle,
        end_angle=end_angle,
        dxfattribs={'layer': '?�이??},
    )    
    return msp
def draw_arc(doc, x1, y1, x2, y2, radius, direction, layer='?�이??): # radius??반�?름을 ?�는?? 지름이 ?�님
    msp = doc.modelspace()
    
    # ?�들�?반�?름을 ?�용?�여 중점�??�의 중심 계산
    midpoint = calculate_midpoint((x1, y1), (x2, y2))
    center = calculate_circle_center(midpoint, radius, (x1, y1), (x2, y2))

    # ?�작 각도?� ??각도 계산
    start_angle = calculate_angle(center, (x1, y1))
    end_angle = calculate_angle(center, (x2, y2))

    # 방향???�라 각도 조정
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

    # ?�크 그리�?    msp.add_arc(
        center=center,
        radius=radius,
        start_angle=start_angle,
        end_angle=end_angle,
        dxfattribs={'layer': layer},
    )
    return msp
def draw_crossmark(doc, x, y, layer='0'):
    """
    M4 ?�각 ??��??그림    
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
    # override ?�정
    override_settings = {
        'dimasz': 15
    }

    # ?�스???�치 조정 �?꺾이??지???�정
    if option is None:    
        text_offset_x = 20
        text_offset_y = 20
    else:
        text_offset_x = 0
        text_offset_y = 0        
    
    if direction == 'leftToright':
        mid_x = end_x - text_offset_x
        mid_y = end_y  # ?�스???�에??꺾임
        text_position = (end_x, end_y)
    elif direction == 'rightToleft':
        mid_x = end_x + text_offset_x
        mid_y = end_y  # ?�스???�에??꺾임
        text_position = (end_x - len(text) * 22, end_y)
    else:
        mid_x = (start_x + end_x) / 2
        mid_y = (start_y + end_y) / 2
        text_position = (end_x + text_offset_x, end_y + text_offset_y)

    # 지?�선 추�?
    leader = msp.add_leader(
        vertices=[(start_x, start_y), (mid_x, mid_y-text_height/2), (end_x, end_y-text_height/2)],  # ?�작?? 중간??문자 ?�에??꺾임), ?�점
        dxfattribs={
            'dimstyle': text_style_name,
            'layer': layer,
            'color': 3  # ?�색 (AutoCAD ?�상 ?�덱?�에??3번�? ?�색)
        },
        override=override_settings
    )

    if option is None:
        # ?�스??추�? (?�택??
        if text:
            msp.add_mtext(text, dxfattribs={
                'insert': text_position,
                'layer': layer,
                'char_height': text_height,
                'style': text_style_name,
                'attachment_point': 1,  # ?�스???�렬 방식 ?�정
                'color': 2  # ?��???(AutoCAD ?�상 ?�덱?�에??2번�? ?��???
            })

    return leader
def dim_linear(doc, x1, y1, x2, y2, textstr, dis, direction="up", layer='0', text_height=0.30,  text_gap=0.07):
    global saved_DimXpos, saved_DimYpos, saved_text_height, saved_text_gap, saved_direction, dimdistance, dim_horizontalbase, dim_verticalbase, distanceXpos, distanceYpos
    msp = doc.modelspace()
    layer = '0'    
    dim_style = over1000dim_style

    # 치수�?계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # ?�수?�이 ?�는지 ?�인
    if dimension_value % 1 != 0:
        dimdec = 1  # ?�수?�이 ?�는 경우 ?�수??첫째 ?�리까�? ?�시
    else:
        dimdec = 1  # ?�수?�이 ?�는 경우 ?�수???�시 ?�음

    # override ?�정
    override_settings = {
        'dimtxt': text_height,
        'dimgap': text_gap,
        'dimscl': 1,
        'dimlfac': 1,
        # 'dimclrt': 7, ?�상강제�??�색 지??        'dimdsep': 46,
        'dimdec': dimdec,
        # ?�스??180???�전
        #'dimtrot': 180        
        'dimtih': 1  # ?�스?��? ??�� ?�평?�로 ?�시
    }

    # 방향???�른 치수??추�?
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
    global saved_Xpos, saved_Ypos  # ?�역 변?�로 ?�용??것임??명시
    
    # ??추�?
    start_point = (x1, y1)
    end_point = (x2, y2)
    if layer:
        # ?�곡??22 layer??ltscale??조정?�다
        if(layer=="22"):
            msp.add_line(start=start_point, end=end_point, dxfattribs={'layer': layer, 'ltscale' : 30})
        else:
            msp.add_line(start=start_point, end=end_point, dxfattribs={'layer': layer })
    else:
        msp.add_line(start=start_point, end=end_point)

    # ?�음 ?�분???�작?�을 ?�데?�트
    saved_Xpos = x2
    saved_Ypos = y2        
def circle_num(doc, x1, y1, x2, y2, text, option=None):
    msp = doc.modelspace()
    
    # ??그리�?    radius = 60  # ?�하??반�?�??�정
    msp.add_circle(
        center=(x2, y2),
        radius=radius,
        dxfattribs={'layer': '1'} # ?�색??    )
    
    # ???��????�스??추�?
    msp.add_mtext(text, dxfattribs={
        'insert': (x2, y2),
        'layer': '0',
        'char_height': 40,        
        'attachment_point': 5  # ?�스?��? 중앙??배치
    })

    if option is None:
        # ?�의 중심�?지?�선 ?�점 좌표�??�용?�여 ?�의 ?�레 ?�의 지?�선 ?�작??계산
        angle = math.atan2(y1 - y2, x1 - x2)
        start_x = x2 + radius * math.cos(angle)
        start_y = y2 + radius * math.sin(angle)

        # 지?�선 그리�?        dim_leader(doc, x1, y1, start_x, start_y,  text, text_height=30, direction='up', option='nodraw')

    return msp
def lt(doc, x, y, layer=None):
    # ?��?좌표�?그리??�?    global saved_Xpos, saved_Ypos  # ?�역 변?�로 ?�용??것임??명시
    
    # ?�재 ?�치�??�작?�으�??�정
    start_x = saved_Xpos
    start_y = saved_Ypos

    # ?�점 좌표 계산
    end_x = start_x + x
    end_y = start_y + y

    # 모델 공간 (2D)??가?�옴
    msp = doc.modelspace()

    # ??추�?
    start_point = (start_x, start_y)
    end_point = (end_x, end_y)
    if layer:
        msp.add_line(start=start_point, end=end_point, dxfattribs={'layer': layer})
    else:
        msp.add_line(start=start_point, end=end_point)

    # ?�음 ?�분???�작?�을 ?�데?�트
    saved_Xpos = end_x
    saved_Ypos = end_y
def lineto(doc, x, y, layer=None):
    global saved_Xpos, saved_Ypos  # ?�역 변?�로 ?�용??것임??명시
    
    # ?�재 ?�치�??�작?�으�??�정
    start_x = saved_Xpos
    start_y = saved_Ypos

    # ?�점 좌표 계산
    end_x = x
    end_y = y

    # 모델 공간 (2D)??가?�옴
    msp = doc.modelspace()

    # ??추�?
    start_point = (start_x, start_y)
    end_point = (end_x, end_y)
    if layer:
        msp.add_line(start=start_point, end=end_point, dxfattribs={'layer': layer})
    else:
        msp.add_line(start=start_point, end=end_point)

    # ?�음 ?�분???�작?�을 ?�데?�트
    saved_Xpos = end_x
    saved_Ypos = end_y
def lineclose(doc, start_index, end_index, layer='?�이??):    

    firstX, firstY = globals()[f'X{start_index}'], globals()[f'Y{start_index}']
    prev_x, prev_y = globals()[f'X{start_index}'], globals()[f'Y{start_index}']

    # start_index+1부??end_index까�? 반복?�니??
    for i in range(start_index + 1, end_index + 1):
        # ?�재 ?�덱?�의 좌표�?가?�옵?�다.
        curr_x, curr_y = globals()[f'X{i}'], globals()[f'Y{i}']

        # ?�전 좌표?�서 ?�재 좌표까�? ?�을 그립?�다.
        line(doc, prev_x, prev_y, curr_x, curr_y, layer)

        # ?�전 좌표 ?�데?�트
        prev_x, prev_y = curr_x, curr_y
        # print(f"prev_x {prev_x}" )
        # print(f"prev_y {prev_y}" )
    
    # 마�?막으�?첫번�??�과 ?�결
    line(doc, prev_x, prev_y, firstX , firstY, layer)        
def rectangle(doc, x1, y1, dx, dy, layer=None, offset=None):
    if offset is not None:
        # ??개의 ?�분?�로 직사각형 그리�?offset 추�?
        line(doc, x1+offset, y1+offset, dx-offset, y1+offset, layer=layer)   
        lineto(doc, dx - offset, dy - offset, layer=layer)  
        lineto(doc, x1 + offset, dy - offset, layer=layer)  
        lineto(doc, x1 + offset, y1 + offset, layer=layer)  
    else:        
        # ??개의 ?�분?�로 직사각형 그리�?        line(doc, x1, y1, dx, y1, layer=layer)   
        line(doc, dx, y1, dx, dy, layer=layer)   
        line(doc, dx, dy, x1, dy, layer=layer)   
        line(doc, x1, dy, x1, y1, layer=layer)   
def xrectangle(doc, x1, y1, dx, dy, layer=None):
    # 중간??x마크�??�색?�을 ?�는 ?�각??만들�?    line(doc, x1, y1, dx, y1, layer=layer)   
    line(doc, dx, y1, dx, dy, layer=layer)   
    line(doc, dx, dy, x1, dy, layer=layer)   
    line(doc, x1, dy, x1, y1, layer=layer)   
    line(doc, x1, y1, dx, dy, layer='3')    # ?�색 ?�터?�인
    line(doc, x1, dy, dx, y1, layer='3')  # ?�색 ?�터?�인     

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
    x1, y1,   # �?번째 ?�분(Line1)???�작??좌표
    x2, y2,   # �?번째 ?�분(Line1)???�점 좌표
    x3, y3,   # ??번째 ?�분(Line2)???�작??좌표
    x4, y4,   # ??번째 ?�분(Line2)???�점 좌표
    distance=80,
    direction="left",
    dimstyle=selected_dimstyle
):
    """
    dim_angular()
    ---------------------------
    ???�분(Line1, Line2)??좌표�?받아??각도�??�시?�는 치수(Angular Dimension)�?    ezdxf 문서(doc)???�성?�는 ?�수?�니?? ?�제�???Line) ?�티?��? 그리지 ?�고,
    '가?�의 ?�분'?�로�?각도�?계산?�여 각도 치수�??�시?�니??

    Parameters
    ----------
    doc : ezdxf.document.Document
        ezdxf ?�큐먼트 객체. 각도 치수�??�입???�??DXF 문서?�니??
    x1, y1 : float
        �?번째 ?�분(Line1)???�작??좌표
    x2, y2 : float
        �?번째 ?�분(Line1)???�점 좌표
    x3, y3 : float
        ??번째 ?�분(Line2)???�작??좌표
    x4, y4 : float
        ??번째 ?�분(Line2)???�점 좌표
    distance : float, default=80
        치수??각도 ?�시)???�분?�의 중앙 지?�에???�마???�어�?곳에 배치?��?�?결정.
        값이 커질?�록 치수?��? ??바깥쪽에 ?�시?�니??
    direction : str, default='left'
        치수??각도 ?�시)???�느 방향?�로 배치?��? 결정?�니??
        - 'left'  : ?�쪽
        - 'right' : ?�른�?        - 'up'    : ?�쪽
        - 'down'  : ?�래�?    dimstyle : str, default='mydim1'
        치수 ?��????�름. 미리 ?�의?�둔 DimStyle 문자?�을 지?�할 ???�습?�다.

    Returns
    -------
    ezdxf.entities.dimension.Dimension or None
        ?�성??Angular Dimension 객체�?반환?�니??
        ?�분???�행/0길이 ?�으�?각도 계산??불�??�하�??�류가 발생?�거??        None 처리 ??별도???�외 처리가 ?�요?????�습?�다.

    Notes
    -----
    - ???�분???�전???�행?�거??길이가 0?�면 ezdxf?�서 각도 치수 ?�성??불�??�합?�다.
      (ZeroDivisionError ?�는 'Invalid colinear or parallel angle legs found' ??발생)
    - ?�요?�다�? ?�분 길이 0, ?�행 ?��?�?검?�하???�수�?추�????�용?��? ?�외�?처리?????�습?�다.
    """
    msp = doc.modelspace()

    # ---------------------------------------------------------------------
    # (?�분??그리지 ?�고) ???�분??좌표만으�?각도 치수(Angular Dimension) ?�성
    # ---------------------------------------------------------------------
    dimension = msp.add_angular_dim_2l(
        base=_calc_base_point(x1, y1, x2, y2, x3, y3, x4, y4, distance, direction),
        line1=((x1, y1), (x2, y2)),  # �?번째 ?�분
        line2=((x3, y3), (x4, y4)),  # ??번째 ?�분
        dimstyle=dimstyle,
        override={
            'dimtxt': 0.22,    # 치수 문자 ?�이
            'dimgap': 0.02,    # 문자?� 치수???�이??간격
            'dimscl': 1,       # 치수 축척
            'dimlfac': 1,      # 치수 ?�위 ?�산            
            'dimdec': 0        # ?�수???�기 ?�릿??(0 = ?�수???�음)
        }
    )
    # ?�제 ?�면??반영
    dimension.render()
    return dimension

def _calc_base_point(x1, y1, x2, y2, x3, y3, x4, y4, distance, direction):
    """
    ?��????�수: ???�분(Line1, Line2)??좌표로�???    치수??각도 ?�시)??기�???base)??계산.

    direction 값에 ?�라 base_x, base_y�?distance만큼
    '?�쪽/?�른�????�래'�??�동?�니??
    """
    # ????x1, y1, x2, y2, x3, y3, x4, y4)???�균??중심?? 계산
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
    # ?�적??0?�면 ?�행(?�는 ?�분 길이가 0??경우???�기???�당)
    return (v1x * v2y - v1y * v2x) == 0

def dim_diameter(doc, center, diameter, angle, dimstyle=selected_dimstyle, override=None):
    msp = doc.modelspace()
    
    # 기본 지�?치수??추�?
    dimension = msp.add_diameter_dim(
        center=center,  # ?�의 중심??        radius=diameter/2,  # 반�?�?        angle=angle,    # 치수??각도
        dimstyle=dimstyle,  # 치수 ?��???        override={"dimtoh": 1}    # 추�? ?��????�정 (?�션) 지?�선???�번 꺾여??글?�각?��? ?��??�으�??�오???�션
    )
    
    # 치수?�의 기하?�적 ?�태 ?�성
    dimension.render()    
def dim_string(doc, x1, y1, x2, y2, dis,  textstr,  text_height=0.30, text_gap=0.05, direction="up"):
    global saved_DimXpos, saved_DimYpos, saved_text_height, saved_text_gap, saved_direction, dimdistance, dim_horizontalbase, dim_verticalbase
    msp = doc.modelspace()
    dim_style = selected_dimstyle
    layer = selected_dimstyle

    # 치수�?계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # ?�수?�이 ?�는지 ?�인
    if dimension_value % 1 != 0:
        dimdec = 1  # ?�수?�이 ?�는 경우 ?�수??첫째 ?�리까�? ?�시
    else:
        dimdec = 0  # ?�수?�이 ?�는 경우 ?�수???�시 ?�음

    # override ?�정
    override_settings = {
        'dimtxt': text_height,
        'dimgap': text_gap,
        'dimscl': 1,
        'dimlfac': 1,
        'dimclrt': 7,        
        'dimdec': dimdec,
        'dimtix': 1,  # 치수???��????�스?��? ?�시 (?�요???�라)
        'dimtad': 1 # 치수???�단???�스?��? ?�시 (?�요???�라)               
    }

    # 방향???�른 치수??추�?
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

    # text_height=0.80 # 강제�??�게 ?�봄
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

    # ?�속??구현???�한 구문
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

    # 치수�?계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # ?�수???�인
    dimdec = 1 if dimension_value % 1 != 0 else 0

    # override ?�정
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
            'dimtmove': 0,  # ?�스???�동 ?�션
            'dimpost': f"{text} " # 치수???�에 ?�스???�시
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
            'dimtmove': 2,  # ?�스???�동 ?�션
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
        add_dim_args['text'] = text  # 추�? ?�스?�만 ?�정

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
    # reverse ?�션 처리
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

    # ?�속??구현???�한 구문
    if starbottomtion is None:
        if direction == "left":             
            dim_horizontalbase = dis - (x1 - x2)
        elif direction == "right":                        
            dim_horizontalbase = dis 
        elif direction == "up":            
            dim_verticalbase = dis
        elif direction == "down":                        
            dim_verticalbase = dis   

    # flip???�언?�면 치수?�의 ?�작�??�을 바꾼?? ?�?�된 좌표??지?�없??
    # 치수?�의 ?�작???�점???�라 치수?�이 ?�오??것을 만들�??�함?�다. 
    # ?�속치수?�때??좌표가 바뀌면 ?�되기때문에 고려??부분이??

    msp = doc.modelspace()
    dim_style = selected_dimstyle
    layer = selected_dimstyle

    # 치수�?계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # ?�수???�인
    dimdec = 1 if dimension_value % 1 != 0 else 0

    # override ?�정
    override_settings = {      
        'dimtxt': text_height,  
        'dimgap': text_gap if text_gap is not None else 0.05,  # ?�기?�서 dimgap??기본�??�정
        'dimscl': 1,
        'dimlfac': 1,
        'dimclrt': 7,        
        'dimdec': dimdec,
        'dimtix': 1, 
        'dimtad': 1  
    }

    # 방향???�른 치수??추�?
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

# 방향???�???�의�??�고 ?�속?�으�?차이�?감�??�서 처리?�는 것이??
    if saved_direction=="left" :
        dimdistance = dim_horizontalbase
        # ?�계?�해????        
        dim_horizontalbase = dimdistance - (x1 - x2)        
    if saved_direction=="right" :
        dimdistance = dim_horizontalbase 
        # ?�계?�해????
        dim_horizontalbase = dimdistance - (x2 - x1)
    if saved_direction=="up" :
        dimdistance = dim_verticalbase
        # ?�계?�해????
        dim_verticalbase = dimdistance - (y2 - y1)
    if saved_direction=="down" :
        dimdistance = dim_verticalbase 
        # ?�계?�해????
        dim_verticalbase = dimdistance - (y1 - y2)

    if distance is not None :
        dimdistance = distance    

    # reverse ?�션 처리
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

    # ?�류 ?�정: text_gap??None???�닐 ?�만 saved_text_gap??갱신?�야 ??    if text_gap is not None:
        saved_text_gap = text_gap
    else:
        text_gap = saved_text_gap

    # ?�류 ?�정: text_height가 None???�닐 ?�만 saved_text_height??갱신?�야 ??    if text_height is not None:
        saved_text_height = text_height
    else:
        text_height = saved_text_height

    dim(doc, saved_DimXpos, saved_DimYpos, x2, y2, dis, text_height=text_height, text_gap=text_gap, direction=saved_direction, option=option)
def create_vertical_dim(doc, x1, y1, x2, y2, dis, angle, layer=None, text_height=0.30, text_gap=0.05):
    msp = doc.modelspace()
    dim_style = layer  # 치수 ?��????�름
    points = [(x1, y1), (x2, y2)]

    if angle==None :
        angle = 270

    # 치수�?계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # ?�수?�이 ?�는지 ?�인
    if dimension_value % 1 != 0:
        dimdec = 1  # ?�수?�이 ?�는 경우 ?�수??첫째 ?�리까�? ?�시
    else:
        dimdec = 1  # ?�수?�이 ?�는 경우 ?�수???�시 ?�음    

    return msp.add_multi_point_linear_dim(
        base=(x1 + dis if angle == 270 else x1 - dis , y1),  #40?� 보정
        points = points,
        angle = angle,
        dimstyle = dim_style,
        discard = True,
        dxfattribs = {'layer': layer},
        # 치수 문자 ?�치 조정 (0: 치수???? 1: 치수???? 'dimtmove': 1 
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

    # 치수�?계산
    dimension_value = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)

    # ?�수?�이 ?�는지 ?�인
    if dimension_value % 1 != 0:
        dimdec = 1  # ?�수?�이 ?�는 경우 ?�수??첫째 ?�리까�? ?�시
    else:
        dimdec = 1  # ?�수?�이 ?�는 경우 ?�수???�시 ?�음    

    return msp.add_multi_point_linear_dim(
        base=(x1 + dis if angle == 270 else x1 - dis , y1),  #40?� 보정
        points=points,
        angle=angle,
        dimstyle=dim_style,
        discard=True,
        dxfattribs={'layer': layer},
        # 치수 문자 ?�치 조정 (0: 치수???? 1: 치수???? 'dimtmove': 1 
        override={'dimtxt': text_height, 'dimgap': text_gap, 'dimscl': 1, 'dimlfac': 1, 'dimclrt': 7, 'dimdec': dimdec,  'dimtmove': 3 , 'text' : textstr      }
    )
def dim_vertical_right_string(doc, x1, y1, x2, y2, dis, textstr, text_height=0.30,  text_gap=0.07):
    return create_vertical_dim_string(doc, x1, y1, x2, y2, dis, 270, textstr, text_height, text_gap)
def dim_vertical_left_string(doc, x1, y1, x2, y2, dis, textstr, text_height=0.30,  text_gap=0.07):
    return create_vertical_dim_string(doc, x1, y1, x2, y2, dis, 90, textstr, text_height, text_gap)
def draw_Text_direction(doc, x, y, size, text, layer=None, rotation = 90):
    # 모델 공간 (2D)??가?�옴
    msp = doc.modelspace()

    # text_style_name ='GHS'
    text_style_name =selected_dimstyle
    # MText 객체 ?�성
    mtext = msp.add_mtext(
        text,  # ?�스???�용
        dxfattribs={
            'layer': layer,  # ?�이??지??            'style': text_style_name,  # ?�스???��???지??            'char_height': size,  # 문자 ?�이 (?�기) 지??        }
    )

    # MText ?�치?� ?�전 ?�정
    mtext.set_location(insert=(x, y), attachment_point=1, rotation=rotation)

    return mtext

def draw_Text(doc, x, y, size, text, layer=None):
    # 모델 공간 (2D)??가?�옴
    msp = doc.modelspace()

    # layer가 None?�면 기본�??�정
    if layer is None:
        layer = "mydim"  # 기본 ?�이???�름 ?�정 (?�요 ??변�?가??

    # text_style_name???�정
    text_style_name = selected_dimstyle

    # ?�스??추�? �??�성??Text 객체 가?�오�?    text_entity = msp.add_text(
        text,  # ?�스???�용
        dxfattribs={
            'layer': layer,  # ?�이??지??            'style': text_style_name,  # ?�스???��???지??            'height': size,  # ?�스???�이 (?�기) 지??        }
    )

    # Text 객체???�치 ?�정
    text_entity.set_placement((x, y), align=TextEntityAlignment.BOTTOM_LEFT)

    # MIDDLE_LEFT�??�으�??�른 ?�면??붙일??문자 ?�치가 ?�라지??경우가 ?�다. 주의?�함 (?�해?�엔지 개발??발견??
def draw_circle(doc, center_x, center_y, radius, layer='0', color='7'):
    """
    ?�을 그리???�수 radius??지름으�??�도�??�정
    :param doc: ezdxf 문서 객체
    :param center_x: ?�의 중심 x 좌표
    :param center_y: ?�의 중심 y 좌표
    :param radius: ?�의 반�?름이 ?�닌 지름입??    : radius=radius/2 ?�용
    :param layer: ?�을 추�????�이???�름 (기본값�? '0')
    지름으�??�정 
    """
    msp = doc.modelspace()
    circle = msp.add_circle(center=(center_x, center_y), radius=radius/2, dxfattribs={'layer': layer, 'color' : color})
    return circle
def circle_cross(doc, center_x, center_y, radius, layer='0', color='7'):
    """
    DXF 문서???�을 그리???�수
    :param doc: ezdxf 문서 객체
    :param center_x: ?�의 중심 x 좌표
    :param center_y: ?�의 중심 y 좌표
    :param radius: ?�의 반�?�?    :param layer: ?�을 추�????�이???�름 (기본값�? '0')
    지름으�??�정 /2 ?�용
    """
    draw_circle(doc, center_x, center_y, radius, layer=layer, color=color)
    # ?�색 ??��??그려주기    
    line(doc, center_x - radius/2 - 5, center_y, center_x +  radius/2 + 5, center_y, layer="CEN" )
    line(doc, center_x , center_y - radius/2 - 5, center_x , center_y +  radius/2 + 5, layer="CEN" )
    return circle_cross        
def cross10(doc, center_x, center_y):
# 10미리 ??��???�이??만들�?  
    line(doc, center_x - 5, center_y, center_x +  5, center_y, layer="?�이?? )
    line(doc, center_x, center_y - 5, center_x , center_y + 5, layer="?�이?? )
    return cross10
def cross(doc, center_x, center_y, length, layer='?�이??):
    line(doc, center_x - length, center_y, center_x +  length, center_y, layer=layer )
    line(doc, center_x, center_y - length, center_x , center_y + length, layer=layer )
    return cross
def crossslot(doc, center_x, center_y, direction=None):
# 10미리 ??��???�이??만들�?  
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
    # ?�색 ??��??그려주기    
    line(doc, center_x - radius/2 - 5, center_y, center_x +  radius/2 + 5, center_y, layer="CEN" )
    line(doc, center_x , center_y - radius/2 - 5, center_x , center_y +  radius/2 + 5, layer="CEN" )
    return 
def extract_abs(a, b):
    return abs(a - b)   
def insert_block(doc, x, y, block_name, layer='?�이??):
    # ?�면?� ?�입    
    scale = 1
    insert_point = (x, y, scale)

    # 블록 ?�입?�는 방법           
    doc.modelspace().add_blockref(block_name, insert_point, dxfattribs={
        'xscale': scale,
        'yscale': scale,
        'rotation': 0,
        'layer': layer
    })
def insert_frame(x, y, scale, title, description, workplaceStr, sep="NOtable"):
    # issuedate ?�류 ?�인 ?�명:
    # issuedate가 None?�거?? 문자?�이지�??�식??"%Y-%m-%d %H:%M:%S"가 ?�닐 경우 datetime.strptime?�서 ?�러가 발생?�니??
    # ?�한, issuedate가 ?�예 값이 ?�거?? ?�?�이 ?�상�??��? ?�도 문제가 ?�깁?�다.
    # ?�전?�게 처리?�려�?None 체크?� 문자???�맷 ?�외처리가 ?�요?�니??

    global issuedate

    formatted_date = ""
    if issuedate is None:
        formatted_date = ""
    elif isinstance(issuedate, str):
        # 문자?�이지�??�맷???��? ???�으므�??�외처리
        try:
            date_object = datetime.strptime(issuedate, "%Y-%m-%d %H:%M:%S")
            formatted_date = date_object.strftime("%y.%m.%d")
        except Exception:
            # ?�른 ?�맷 ?�도 ?�는 그냥 ?�본 ?�용
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
    # ?�면?� ?�입
    if sep == "basic":
        block_name = "drawings_frame"
    # 2???�입 ?�면 ASSY?�면
    if sep == "column2":
        block_name = "drawings_frame_column2"
    # 7???�입 ?�면 ASSY?�면
    if sep == "column7":
        block_name = "drawings_frame_column7"
    # 1???�입 ?�면 ASSY?�면
    if sep == "NOtable":
        block_name = "drawings_frame_NOtable"

    insert_point = (x, y, scale)

    # 블록 ?�입?�는 방법
    msp.add_blockref(block_name, insert_point, dxfattribs={
        'xscale': scale,
        'yscale': scale,
        'rotation': 0
    })

    draw_Text(doc, x + (3545 + 900) * scale, y + 420 * scale, 50 * scale, str(description), '0')
    draw_Text(doc, x + (3545 + 900 + 100) * scale, y + 630 * scale, 60 * scale, f"{title}", '0')
    draw_Text(doc, x + (3545 + 900) * scale, y + 850 * scale, 100 * scale, f"?�장�?: {workplaceStr}", '0')

def envsettings():
    # ?�드?�스??고유번호�?가?�오??코드 (?�스?�에 ?�라 ?��? ???�음)
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
    # 결과�??�?�할 리스??초기??    hole_array = []

    # ?�재 ?�자�?startnum?�로 ?�정
    current_num = startnum

    # current_num??limit�?length�??��? ?�을 ?�까지 반복
    while current_num <= limit and current_num <= length:
        # 리스?�에 ?�재 ?�자 추�?
        hole_array.append(current_num)
        # ?�음 ?�자�?interval만큼 증�?
        current_num += interval
    hole_array.append(length-85)
    return hole_array
def calculate_splitholeArray(startnum, interval, limit, length):    
    hole_array = []

    # ?�재 ?�자�?startnum?�로 ?�정
    current_num = startnum

    # current_num??limit�?length�??��? ?�을 ?�까지 반복
    while current_num <= limit and current_num <= length:
        # 리스?�에 ?�재 ?�자 추�?
        hole_array.append(current_num)
        # ?�음 ?�자�?interval만큼 증�?
        current_num += interval
    hole_array.append(length)
    return hole_array

def calSplitHole(start, interval, limit):
    """
    start: ?�작�?    interval: 간격
    limit: 최�? 범위 (?? MH + HH + grounddig)
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

def draw_slot(doc, x, y, size, direction="가�?, option=None, layer='0'):
    """
    Draws a slot (?�공) with specified parameters in a DXF document.
    
    Parameters:
    doc (ezdxf.document): The DXF document to draw on.
    x (float): The x-coordinate of the slot's center.
    y (float): The y-coordinate of the slot's center.
    size (str): The size of the slot in the format 'WxH' (e.g., '8x16').
    direction (str): The direction of the slot, either '가�? (horizontal) or '?�로' (vertical). Default is '가�?.
    option (str): Optional feature for the slot. If 'cross', draws center lines extending beyond the slot. Default is None.
    layer (str): The layer to draw the slot on. Default is '0'.
    """    
    msp = doc.modelspace()

    # size�?분리?�여 ??�� ?�이 계산
    width, height = map(int, size.lower().split('x'))

    if direction == "가�?:
        slot_length = height - width
        slot_width = width
    else:  # ?�로
        slot_length = width
        slot_width = height - width

    radius = slot_width / 2

    # 중심?�을 기�??�로 ?�작?�과 ?�점 계산
    if direction == "가�?:
        start_point = (x - slot_length / 2, y)
        end_point = (x + slot_length / 2, y)
    else:  # ?�로
        start_point = (x, y - slot_length / 2)
        end_point = (x, y + slot_length / 2)

    # 직선 부�?그리�?    if direction == "가�?:
        msp.add_line((start_point[0], start_point[1] + radius), (end_point[0], end_point[1] + radius), dxfattribs={'layer': layer})
        msp.add_line((start_point[0], start_point[1] - radius), (end_point[0], end_point[1] - radius), dxfattribs={'layer': layer})
    else:  # ?�로
        msp.add_line((start_point[0] + radius, start_point[1]), (end_point[0] + radius, end_point[1]), dxfattribs={'layer': layer})
        msp.add_line((start_point[0] - radius, start_point[1]), (end_point[0] - radius, end_point[1]), dxfattribs={'layer': layer})

    # ???�의 반원 그리�?    if direction == "가�?:
        draw_arc_slot(doc, start_point, radius, 90, 270, layer)  # 반원 방향 ?�정
        draw_arc_slot(doc, end_point, radius, 270, 90, layer)  # 반원 방향 ?�정
    else:  # ?�로
        draw_arc_slot(doc, start_point, radius, 180, 360, layer)  # 반원 방향 ?�정
        draw_arc_slot(doc, end_point, radius, 0, 180, layer)  # 반원 방향 ?�정

    # ?�션??"cross"??경우 중심?�을 추�?�?그리�?    if option == "cross":
        if direction == "가�?:
            msp.add_line((x - slot_length / 2 - 8, y), (x + slot_length / 2 + 8, y), dxfattribs={'layer': 'CL'})
            msp.add_line((x, y - slot_width / 2 - 4), (x, y + slot_width / 2 + 4), dxfattribs={'layer': 'CL'})
        else:  # ?�로
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
    ?�면 ?�성 로직??처리?�는 ?�수
    """
    global jambType, floorDisplay, material, spec, vcut, OP, JE, JD, HH, MH
    global HPI_height, U, C, A, grounddig, control_width, controltopbar, controlbox, poleAngle

    # ???�이?��? ?�역 변?�로 ?�정
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
    poleAngle = round(row_data.get("poleAngle", 0), 1)  # ?�수???�째 ?�리까�? ?�한

    # ?�면 ?�성 ?�업
    # print(f"Generating drawing for jambType: {jambType}, material: {material}, spec: {spec}")
    # ?�기???�제 ?�면 ?�성 코드�?추�?

def calculate_base(height, angle_degrees, option=None):
    """
    주어�??�이?� ?�인�??????�용??밑�???계산?�니??

    # # ?�제 ?�용�?    # height = 200
    # angle_degrees = 10
    # base = calculate_base(height, angle_degrees)
    # print(f"?�이가 {height}?�고 ?�인각이 {angle_degrees}?�일 ??밑�??� {base}?�니??")    

    :param height: ?�이 (?�로 변)
    :param angle_degrees: ?�인�?(??
    :return: 밑�? 길이
    """
    if angle_degrees <= 0 or angle_degrees >= 90:
        raise ValueError("각도??0??초과 90??미만?�어???�니??")
    
    # 각도�??�디?�으�?변??    angle_radians = math.radians(angle_degrees)
    # 밑�? 계산: ?�이?� ?�젠?�의 �?    base = height * math.tan(angle_radians)
    if option=="height" :
        base = height * math.cos(angle_radians)
    return round(base, 2)

def set_point(a, index, x_value, y_value):
    """
    주어�??�셔?�리??x, y 값을 ?�정?�는 ?�수
    """
    a[f'x{index}'], a[f'y{index}'] = x_value, y_value    
def calcuteHoleArray(totalLength, startLength, holeNumber):
    """
    totalLength : ?�체 길이
    startLength : 좌우 ?�유 길이(?�작·??지?�에???�어�?거리)
    holeNumber  : ?�체?�으�??�을 ?�(구멍)??개수

    ?? totalLength=250, startLength=20, holeNumber=5
       -> 5개의 ?� ?�치가 [20, 72.5, 125.0, 177.5, 230.0]?�로 계산
    """
    # ?�시 holeNumber가 1 ?�하?�면 처리 불�??�하???�외 처리
    if holeNumber < 2:
        raise ValueError("holeNumber??최소 2 ?�상?�어???�니??")

    # 결과 리스??    holes = []

    # "?�쪽 가?�자�?�??��? ?�함?�야 ?��?�?
    # 가?�데???�어�??� ?�는 (holeNumber - 2),
    # ?��?�?간격??구할 ??(holeNumber - 1)�??�눈??
    #
    # --?�시--
    #  holeNumber=5?�면, �?5개의 ?�??찍을 것이�?
    #  간격?� 4구간(=5-1)?�로 ?�눈??
    #
    # ?�쪽 startLength�??�외???�수 중간 ?�역
    middle_length = totalLength - 2 * startLength

    # 구간 간격
    interval = middle_length / (holeNumber - 1)

    # 0�??�부??holeNumber-1�??�까�? 계산
    for i in range(holeNumber):
        pos = startLength + interval * i
        holes.append(pos)

    # 간격 보정 로직 추�?
    # 모든 간격???�수??첫째?�리까�? 반올림해??비교
    if len(holes) >= 3:
        # ?�제 간격 리스???�성
        intervals = [round(holes[i+1] - holes[i], 1) for i in range(len(holes)-1)]
        # 가??많이 ?�오??간격??기�??�로 ?�음
        from collections import Counter
        cnt = Counter(intervals)
        most_common_gap, _ = cnt.most_common(1)[0]
        # 첫번�?간격???�른 경우 보정
        if abs(intervals[0] - most_common_gap) >= 0.1:
            # 첫번�??�치�?보정
            diff = intervals[0] - most_common_gap
            # holes[0]?� startLength, holes[1]�?보정
            holes[1] = holes[0] + most_common_gap
            # ?�후 값들??간격??맞춰???�계??            for i in range(2, len(holes)):
                holes[i] = holes[i-1] + most_common_gap
            # 마�?�?값이 endLength�??�으�?마�?막만 endLength�?맞춤
            endLength = totalLength - startLength
            if abs(holes[-1] - endLength) > 0.1:
                holes[-1] = endLength

    return holes

def simulate_hole_positions_from_bottom(totalLength, target_first_distance, hole_count, poleAngle):
    """
    ?�이 ?�어가??구조�??�한 ?� ?�치 ?��??�이??    
    - �?번째 ?�: ?�직거리 52mm
    - ?� 간격: ?�선거리 64mm ?�후 (각도 고려)
    """
    holes = []
    
    # 기존 방식?�로 기본 ?� ?�치 계산 (조건 ?�서 ?�정)
    if hole_count >= 5:
        holes = calcuteHoleArray(totalLength, 48.5, 5)
    elif hole_count >= 4:
        holes = calcuteHoleArray(totalLength, 48.5, 4)
    else:
        holes = calcuteHoleArray(totalLength, 48.5, 3)
    
    # ?�전???�근: 기존 ?� ?�치�?거의 그�?�??��??�면??미세 조정�?    # 문제: offset???�수가 ?�어 ?�?�이 ?�래�??�려�?    
    # ?�재 ?�?�이 ?�상 ?�치???�는지 ?�인 ??미세 조정�??�용
    # 41.6 ??52�?증�??�키?�면 ?�???�래쪽으�??�동?�야 ??    # (?�???�래�?가�??�단까�? 거리가 증�?)
    
    # ?�전?�게 5mm�??�래쪽으�??�동 (41.6 + ??5 ??47 ?�도)
    safe_adjustment = -0.6  # ?�수 = ?�래�??�동 = ?�단 거리 증�?
    
    # 모든 ?�???�일?�게 ?�동 (간격 ?��?)
    holes = [hole + safe_adjustment for hole in holes]
    
    return holes

def ds(doc, x1, y1, x2, y2, dis, text_height=0.30, text_gap=0.05,
       direction="up", option=None, starbottomtion=None, text=None,
       dim_style=None):
    """
    문자??+ 치수???�께 구현?�는 ?�제 ?�수.

    Parameters
    ----------
    doc : ezdxf.document.Document
        ezdxf ?�큐먼트 객체(치수�?추�???DXF 문서).
    x1, y1 : float
        �?번째 ??P1 (x, y)
    x2, y2 : float
        ??번째 ??P2 (x, y)
    dis : float
        치수??문자 ?�시 ?�치)??기�??�에???�마???�어?�릴지 결정?�는 거리.
    text_height : float, default=0.30
        치수 문자 ?�이.
    text_gap : float, default=0.05
        치수 문자?� 치수???�이 간격.
    direction : {"up", "down", "left", "right"}
        치수?�을 ?�느 방향(?�평/?�직)?�로 배치?��? 결정.
        - "up"    : P1, P2�??�는 ?�분 ?�쪽???�평 치수??        - "down"  : P1, P2�??�는 ?�분 ?�래쪽에 ?�평 치수??        - "left"  : P1, P2�??�는 ?�분 ?�쪽?? ?�직 치수??        - "right" : P1, P2�??�는 ?�분 ?�른쪽에 ?�직 치수??    option : str or None
        'reverse'?�면 P1, P2�?바꾸???�집?�서) 치수 ?�시.
    text : str or None
        치수 값에 ?�서 ?�시??문구. ?? "?�판 ?�이"
    dim_style : str or None
        ?�용??치수 ?��??? None?�면 기본 "over1000dim1" ?�용.

    Returns
    -------
    ezdxf.entities.dimension.Dimension 
        ?�성??치수 ?�티??객체�?반환.
    """

    # ?�션??'reverse'??경우, ???�을 ?�로 바꿔??치수�?반�?�??�시
    if option == 'reverse':
        x1, x2 = x2, x1
        y1, y2 = y2, y1

    # 방향???�라 angle(0 or 90)�?base(치수??기�??? 계산
    if direction == "up":
        angle = 0  # ?�평 치수??        base_x = (x1 + x2) / 2
        base_y = max(y1, y2) + dis
    elif direction == "down":
        angle = 0  # ?�평 치수??        base_x = (x1 + x2) / 2
        base_y = min(y1, y2) - dis
    elif direction == "left":
        angle = 90  # ?�직 치수??        base_x = min(x1, x2) - dis
        base_y = (y1 + y2) / 2
    elif direction == "right":
        angle = 90  # ?�직 치수??        base_x = max(x1, x2) + dis
        base_y = (y1 + y2) / 2
    else:
        # ?��? 기본값을 "up"?�로 처리
        angle = 0
        base_x = (x1 + x2) / 2
        base_y = max(y1, y2) + dis

    # Model Space ?�득
    msp = doc.modelspace()

    # 치수???�성
    dim = msp.add_linear_dim(
        base=(base_x, base_y),  # 치수??기�???        p1=(x1, y1),            # �???        p2=(x2, y2),            # ?�째 ??        angle=angle,            # 0°=가�? 90°=?�로
        dimstyle=dim_style or "over1000dim1",
        override={
            # 문자??+ 치수�?: ?? "?�판 ?�이 200.0"
            'dimpost': f"{text} <>" if text else "<>",
            'dimtxt': text_height,  # ?�스???�이
            'dimdec': 1,            # ?�수???�리??(1?�리)
            'dimgap': text_gap,     # ?�스?��? 치수???�이 간격
        }
    )

    # ?�면??반영
    dim.render()
    return dim
def calculate_jb(JE, JD_plus_10):
    """
    JE (float) : 밑�? (?? 60)
    JD_plus_10 (float) : ?�이 JD + 10 (?? 300)

    Returns
    -------
    (jb_pythagoras, jb_trig)
      jb_pythagoras : ?��?고라???�리�??�용??계산??빗�?
      jb_trig       : ?�각?�수�??�용??계산??빗�?
    """

    # 1) ?��?고라???�리�?빗�? JB 구하�?    jb_pythagoras = math.sqrt(JE**2 + JD_plus_10**2)

    # 2) ?�각?�수�??�용??빗�? JB 구하�?    #    tanθ = (JD+10) / JE  ->  θ = arctan((JD+10)/JE)
    #    JB = (JE / cosθ) ?�는 (JD+10) / sinθ
    theta = math.atan(JD_plus_10 / JE)  # ?�디??radian) �?    jb_trig = JE / math.cos(theta)

    # return jb_pythagoras, jb_trig
    return jb_pythagoras
def aggregate_rows(rows_data):
    """
    rows_data: �??�의 ?�보�??��? 리스??               (?? global_data["rows_data"] ?�태)

    반환�?
      ?�일??(jambType, material, spec, vcut, OP, JE, JD, HH, MH, HPI_height,
              U, C, A, grounddig, poleAngle, FireDoor)
      ???�?�서 floorDisplay�?콤마�??�쳐 'floorDisplay'�??�고,
      surang???�쳐�?개수�??�?�한 리스?��? 반환.
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
        FireDoor    = row_data.get("FireDoor", 25)  # 기본�?25 (방화?�어)

        floorDisplay = row_data["floorDisplay"]

        # floorDisplay ?�외??모든 ?�성?�로 key 구성 (FireDoor ?�함)
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
        if jambType == '막판?? :
            HPIsurang += 1
        SU += 1

    # ?�제 aggregated -> 최종 리스???�성
    final_list = []
    for key, val in aggregated.items():
        (
            jambType, material, spec, vcut, OP, JE, JD, HH, MH,
            HPI_height, U, C, A, grounddig, pangle, FireDoor
        ) = key

        # floorDisplay�?콤마�??�결
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
    모자보강(Hat Shape) ?�면??그리???�시 ?�수.
    7�??�이 기본?�으�?(0,0)???�치?�도�?좌표�??�의????
    angle만큼 ?�계방향 ?�전 ?? 최종?�으�?(basex, basey)???�습?�다.

    Parameters
    ----------
    doc : ezdxf.document.Document
        ezdxf ?�큐먼트 객체
    basex, basey : float
        7�??�을 최종?�으�??�치?�킬 기�? 좌표
    angle : float
        ?�계방향 ?�전 각도(???�위). ?? 10
    bottomLength : float
        ?? 1??, 6?? 같�? ?�단 길이
    topLength : float
        11??0 구간 길이
    height : float
        1??1 구간 ?�이
    layer : str
        ?�면???�용???�이???�름
    """

    # -----------------------------
    # 1) 좌표 ?�의(각도=0, 7�?(0,0))
    # -----------------------------
    # ???�시?�선 ??번이 ?�점?�이�?
    # ?�머지 ?�들(1~6,8~13)???�당??배치.
    # ?�제 ?�면 ?�황??맞추??조정 가?�합?�다.

    points = {}

    t = 1.6

    # 7�??�을 ?�점(0,0)
    points[6] = (0.0, 0.0)

    # 8번�? 7????bottomLength?�면, x+ 방향??배치
    points[7] = (0.0, t)

    # 9번�? 모자 ?��?�?꺾이??부�?- ?�??(someX, someY)
    points[8] = (-bottomLength + t, t)
    
    # ?�기?�는 ?�시 좌표
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
    # 2) ?�전 변??(?�계방향 angle)
    #    7번점 (지금�? (0,0)) 기�??�로 모든 ?�을 ?�전
    # -----------------------------
    # ?�계방향 angle??= ?�반?�계 -angle?��? ?�일
    # ?�전 공식 (반시�?θ):
    #   x' = x*cosθ - y*sinθ
    #   y' = x*sinθ + y*cosθ
    # ?�기?�는 θ = -angle (degree) ??rad = -angle * ?/180
    theta = math.radians(-angle)
    cos_t = math.cos(theta)
    sin_t = math.sin(theta)

    def rotate_clockwise_about_7(pt):
        # pt??(x, y)
        x, y = pt
        # ?��? 7번점?� (0,0)?��?�? 
        # 굳이 ?�원???�동 ???�전 ???�점 복�???과정?�서
        # ?�점 ?�동???�요?�다.
        x_new = x * cos_t - y * sin_t
        y_new = x * sin_t + y * cos_t
        return (x_new, y_new)

    # 모든 ?�을 ?�전
    for i in points:
        points[i] = rotate_clockwise_about_7(points[i])

    # -----------------------------
    # 3) ?�행?�동 (7번을 (basex, basey)????
    # -----------------------------
    # ?�전 ?�에??7번�? (0,0) ?�태 ???��? (basex, basey)�??�동
    px7, py7 = points[6]
    shift_x = basex - px7
    shift_y = basey - py7

    for i in points:
        x0, y0 = points[i]
        points[i] = (x0 + shift_x, y0 + shift_y)

    # -----------------------------
    # 4) ??그리�? 1~13???�하???�서?��??�결
    #    (line ?�수???��? ?�공?�었?�고 가??
    # -----------------------------
    def draw_line(idx1, idx2):
        x1, y1 = points[idx1]
        x2, y2 = points[idx2]
        line(doc, x1, y1, x2, y2, layer=layer)

    # ?�제 모양??맞게 ?�결 (?? 1??, 2??, 6??, 7??, 8??, 9??0, 10??1, 11??2, 12??3, 13??, ...)
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

##########################################################################################################################################
# ?�넬?�작 ?�동?�도
##########################################################################################################################################
def execute_panel(): 
    """
    ?�작?�출결과 기반 ?�넬?�작 ?�면 ?�성 ?�수
    기존???�용???�수?�을 ?�용?�여 ?�작?�출결과 ?�이?��? 기반?�로 ?�면???�성?�니??
    """
    global global_data, doc, msp, pageCount
    global company, drawnby, workplace, issuedate

    # ===================== (1) 기본 ?�보 ?�팅 =====================
    company = global_data["company"]
    drawnby = global_data["drawnby"]
    workplace = global_data["workplace"]
    issuedate = global_data["issuedate"]

    # ===================== (2) ?�작?�출결과 ?�이??가?�오�?=====================
    manufacturing_data = global_data.get("manufacturing_data", [])

    # ===================== (3) ?�작?�출결과 기반 ?�넬?�작 ?�면 ?�도 =====================
    t = 1.5  # ?�께??1.5�?강제
    AbsX = 0

    # ?�작?�출결과 ?�이?��? ?�으�??��? 기반?�로 ?�면 ?�성
    if manufacturing_data:
        # 고유번호별로 그룹??        unique_ids = {}
        for panel_data in manufacturing_data:
            unique_id = panel_data.get('unique_id', 0)
            if unique_id not in unique_ids:
                unique_ids[unique_id] = []
            unique_ids[unique_id].append(panel_data)
        
        # �?고유번호별로 ?�면 ?�성
        for unique_id, panels in unique_ids.items():
            if not panels:
                continue
                
            # �?번째 ?�널?�서 기본 ?�보 가?�오�?            first_panel = panels[0]
            site_name = first_panel.get('site_name', '')
            car_width = first_panel.get('car_width', 0)
            car_depth = first_panel.get('car_depth', 0)
            car_height = first_panel.get('car_height', 0)
            manufacturing_height = first_panel.get('manufacturing_height', 0)
            manufacturing_count = first_panel.get('manufacturing_count', 0)
            
            rx, startYpos = AbsX + unique_id*15000, 3000     
            pageCount += 1   
            
            print(f"?�작?�출결과 기반 ?�면 ?�성 �?.. 고유번호: {unique_id}, ?�장: {site_name}")
            
            # 기본 ?�두�?그리�?(기존 rectangle ?�수 ?�용)
            border_width = car_width + 200  # �?가�?+ ?�백
            border_height = car_height + 200  # �??�이 + ?�백
            rectangle(doc, rx, startYpos, border_width, border_height, layer='0')
            
            # ?�장 ?�보 ?�스??출력 (기존 draw_Text ?�수 ?�용)
            draw_Text(doc, rx + 50, startYpos + border_height - 50, 40, f"?�장�? {site_name}", layer='0')
            draw_Text(doc, rx + 50, startYpos + border_height - 100, 30, f"고유번호: {unique_id}", layer='0')
            draw_Text(doc, rx + 50, startYpos + border_height - 130, 30, f"?�작 ?�?? {manufacturing_count}?�", layer='0')
            
            # �??��? 치수 ?�보 출력
            draw_Text(doc, rx + 50, startYpos + border_height - 170, 30, f"�??��? 치수: {car_width} x {car_depth} x {car_height} (mm)", layer='0')
            draw_Text(doc, rx + 50, startYpos + border_height - 200, 30, f"?�작?�이: {manufacturing_height}mm", layer='0')
            
            # ?�널 ?�보 출력
            draw_Text(doc, rx + 50, startYpos + border_height - 240, 30, f"�??�널 ?? {len(panels)}�?, layer='0')
            
            # ?�널 2번�???9번까지 ?�터�?�??�렬
            panels_2_to_9 = []
            for panel_data in panels:
                panel_number = panel_data.get('panel_number', 0)
                if 2 <= panel_number <= 9:
                    panels_2_to_9.append(panel_data)
            
            # ?�널 번호?�으�??�렬
            panels_2_to_9.sort(key=lambda x: x.get('panel_number', 0))
            
            # �??�널�??�면 그리�?(?�널 2번�???9번까지)
            panel_start_x = rx + 100
            panel_start_y = startYpos + 100
            current_x = panel_start_x
            
            for panel_data in panels_2_to_9:
                panel_number = panel_data.get('panel_number', 0)
                manufacturing_width = panel_data.get('manufacturing_width', 0)
                panel_height = manufacturing_height
                
                print(f"?�널 {panel_number} 그리�? ?�비 {manufacturing_width}mm, ?�이 {panel_height}mm")
                
                # ?�널 ?�곽??그리�?(4개의 ?�을 ?�결?�는 ?�각?? - 기존 rectangle ?�수 ?�용
                rectangle(doc, current_x, panel_start_y, current_x + manufacturing_width, panel_start_y + panel_height, layer='?�이??)
                
                # ?�널 번호?� 치수 ?�시 (기존 draw_Text ?�수 ?�용)
                draw_Text(doc, current_x + manufacturing_width/2 - 20, panel_start_y - 20, 25, f"#{panel_number}", layer='0')
                draw_Text(doc, current_x + manufacturing_width/2 - 40, panel_start_y - 45, 20, f"{manufacturing_width}mm", layer='0')
                
                # ?��??�보가 ?�으�??�시
                perforation_width = panel_data.get('perforation_width', 0)
                perforation_length = panel_data.get('perforation_length', 0)
                perforation_height = panel_data.get('perforation_height', 0)
                distance_from_entrance = panel_data.get('distance_from_entrance', 0)
                
                if perforation_width > 0 and perforation_length > 0:
                    # ?��??�치 계산 (?�널 ?�단?�서 ?�쪽?�로)
                    hole_x = current_x + distance_from_entrance
                    hole_y = panel_start_y + panel_height - perforation_height
                    
                    # ?��??�각??그리�?(기존 rectangle ?�수 ?�용)
                    rectangle(doc, hole_x, hole_y, hole_x + perforation_width, hole_y + perforation_length, layer='?�이??)
                    
                    # ?��??�보 ?�스??(기존 draw_Text ?�수 ?�용)
                    draw_Text(doc, hole_x + perforation_width/2 - 30, hole_y - 15, 15, f"?��?{perforation_width}x{perforation_length}", layer='0')
                
                # ?�음 ?�널 ?�치�??�동
                current_x += manufacturing_width + 50  # ?�널 간격 50mm
            
            # 첫페?��????�면?� ?�입 (기존 insert_frame ?�수 ?�용)
            if unique_id == list(unique_ids.keys())[0]:
                insert_frame(0, 0, 1.0, "?�작?�출결과 기반 ?�넬?�작 ?�면", f"?�장: {workplace}", f"?�작?�출결과_{workplace}")
                
                # ?�체 ?�작 ?�보 ?�약 출력
                total_panels = sum(len([p for p in panels if 2 <= p.get('panel_number', 0) <= 9]) for panels in unique_ids.values())
                
                draw_Text(doc, 100, 2500, 50, "?�작?�출결과 ?�장 ?�약", layer='0')
                draw_Text(doc, 100, 2400, 30, f"고유번호 ?? {len(unique_ids)}�?, layer='0')
                draw_Text(doc, 100, 2350, 30, f"�??�널 ??(2-9�?: {total_panels}�?, layer='0')
    else:
        print("?�작?�출결과 ?�이?��? ?�습?�다. ?�면???�성?��? ?�습?�다.")

############################################################################################################################################################
@Gooey(encoding='utf-8', program_name='?�완?�크 ?�넬?�작 ?�동?�도', tabbed_groups=True, navigation='Tabbed', show_success_modal=False,  default_size=(1200, 600))

def main():
    global args  #?�정?�때??global ?�언???�요?�다. ?�순???�기�??�때???�요?�다.        
    global exit_program, program_message, text_style_name    
    global SU
    # ?�역 ?�이???�셔?�리 초기??   
    global global_data, doc, msp

    # ?�재 ?�짜?� ?�간??가?�옵?�다.
    current_datetime = datetime.now()
    global_data["formatted_date"] = current_datetime.strftime('%Y-%m-%d')
    global_data["current_time"] = current_datetime.strftime("%H%M%S")

    # .xlsm ?�일???�을 경우 ?�류 메시지�?출력?�고 ?�행??중단
    if not xlsm_files:
        error_message = ".xlsm ?�일??excel?�일 ?�더???�습?�다. ?�인바랍?�다."
        show_custom_error(error_message)
        sys.exit(1)

    for file_path in xlsm_files:
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            error_message = f"?��? ?�일???????�습?�다: {str(e)}"
            show_custom_error(error_message)

        # ?�작?�출결과 ?�트�??�기
        try:
            sheet = workbook["?�작?�출결과"]
        except KeyError:
            error_message = "'?�작?�출결과' ?�트�?찾을 ???�습?�다."
            show_custom_error(error_message)
            sys.exit(1)

        try:
            if readfile is not None:
                doc = readfile(os.path.join(dxf_saved_file, 'style.dxf'))
                msp = doc.modelspace()
            else:
                raise AttributeError("readfile ?�수�??�용?????�습?�다.")
        except (AttributeError, FileNotFoundError) as e:
            try:
                if new is not None:
                    doc = new()
                    if readfile is not None and os.path.exists(os.path.join(dxf_saved_file, 'style.dxf')):
                        doc = readfile(os.path.join(dxf_saved_file, 'style.dxf'))
                    msp = doc.modelspace()
                else:
                    error_message = "ezdxf new ?�수�??�용?????�습?�다."
                    show_custom_error(error_message)
                    return
            except Exception as e:
                error_message = f"DXF ?�일 로드 ?�류: {str(e)}"
                show_custom_error(error_message)
                return
        except Exception as e:
            error_message = f"DXF ?�일???�을 ???�습?�다: {str(e)}"
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

        for cell_ref, var_name in variable_names.items():
            value = read_excel_value(sheet, cell_ref)
            global_data[var_name] = value
            globals()[var_name] = value

        # ?�작?�출결과 ?�이???�기
        global_data["manufacturing_data"] = read_manufacturing_results(sheet)

        thickness = float(re.sub("[A-Z]", "", global_data["thickness_string"]))
        global_data["WorkTitle"] = f"?�체�? {global_data['company']}, ?�장�? {global_data['workplace']}, thickness: {thickness}"

        execute_panel()

        invalid_chars = '<>:"/\\|?*'
        cleaned_file_name = re.sub(
            f'[{re.escape(invalid_chars)}]', '',
            f"{global_data['company']}_{global_data['workplace']}_{global_data['thickness_string']}_{global_data['current_time']}"
        )
        script_directory = os.path.dirname(os.path.abspath(__file__))
        full_file_path = os.path.join(script_directory, f"c:/python/osel/?�업?�료/{cleaned_file_name}.dxf")
        global_data["file_name"] = full_file_path

        exit_program = False
        program_message = \
            '''
        ?�로그램 ?�행결과?�니??
        -------------------------------------
        {0}
        -------------------------------------
        ?�용??주셔??감사?�니??
        '''
        args = parse_arguments()

        log_login()

        doc.saveas(global_data["file_name"])
        print(f" ?�???�일�? '{global_data['file_name']}' ?�???�료!")

if __name__ == '__main__':
    main()
            set_point(yy, 8, yy['x7']   , yy['y7'] + SideHatribLength )
            set_point(yy, 9, yy['x8']  - 23.5 , yy['y8']   )
            set_point(yy, 10, yy['x9']  - 22 , yy['y9']   )
            set_point(yy, 11, yy['x10']  - 15, yy['y10']  )
            set_point(yy, 12, yy['x11']  - 15, yy['y11']  )
            set_point(yy, 13, yy['x12']  -22 , yy['y12']   )
            set_point(yy, 14, yy['x13']  - 23.5 , yy['y13']   )

            prev_x, prev_y = yy['x1'], yy['y1']  # �?번째 ?�으�?초기??            lastNum = 14
            for i in range(1, lastNum + 1):
                cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
                line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="?�이??)
                prev_x, prev_y = cuyy_x, cuyy_y
            line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="?�이??)    

            # ?�곡??4개소
            bx1, by1, bc1, bd1 = yy['x13'], yy['y13'], yy['x2'], yy['y2'] 
            bx2, by2, bc2, bd2 = yy['x12'], yy['y12'], yy['x3'], yy['y3'] 
            bx3, by3, bc3, bd3 = yy['x5'], yy['y5'], yy['x10'], yy['y10'] 
            bx4, by4, bc4, bd4 = yy['x6'], yy['y6'], yy['x9'], yy['y9'] 
            line(doc, bx1, by1, bc1, bd1, layer="?�곡??)        
            line(doc, bx2, by2, bc2, bd2, layer="?�곡??)        
            line(doc, bx3, by3, bc3, bd3, layer="?�곡??)        
            line(doc, bx4, by4, bc4, bd4, layer="?�곡??)               

            d(doc, yy['x1'] , yy['y1'], yy['x7'] , yy['y7'], 120,  direction="down", dim_style=over1000dim_style)          
            d(doc, yy['x9'] , yy['y9'], yy['x8'] , yy['y8'], 80,  direction="up", dim_style=over1000dim_style)          
            d(doc, yy['x13'] , yy['y13'], yy['x14'] , yy['y14'], 80,  direction="up", dim_style=over1000dim_style)          
            d(doc, yy['x14'] , yy['y14'], yy['x1'] , yy['y1'],  150,  direction="left", dim_style=over1000dim_style)        

            string = f"1.6T EGI" 
            draw_Text(doc, (yy['x1'] + yy['x7'])/2 - len(string)*50/2 , yy['y14'] + 800, 50, text=string, layer='?�이??)            
            string = f"기둥 모자 보강" 
            draw_Text(doc, (yy['x1'] + yy['x7'])/2 - len(string)*60/2 , yy['y14'] + 600, 60, text=string, layer='?�이??)        
            string = f"{SU*2} EA" 
            draw_Text(doc, (yy['x1'] + yy['x7'])/2 - len(string)*90/2 , yy['y1'] - 300, 90, text=string, layer='?�이??)

           
        ######################################################################################
        # ?�로???�판+기둥 조립?? MH (270) - 5 = 265 막판 ?�???�판 보강 1
        ######################################################################################

        if jambType == '막판??:
            # 좌표 초기??            Length = MH-5

            yy = {f'x{i}': 0 for i in range(1, 31)}
            yy.update({f'y{i}': 0 for i in range(1, 31)})

            # 좌표 ?�정
            set_point(yy, 1, rx, startYpos )
            set_point(yy, 2, yy['x1'] + Length  , yy['y1'] )
            set_point(yy, 3, yy['x2'] , yy['y2'] + 48 )
            set_point(yy, 4, yy['x3'] , yy['y3'] + 26 )
            set_point(yy, 5, yy['x4'] - Length  , yy['y4'] )
            set_point(yy, 6, yy['x5'] , yy['y5'] - 26 )		

            prev_x, prev_y = yy['x1'], yy['y1']  # �?번째 ?�으�?초기??            lastNum = 6
            for i in range(1, lastNum + 1):
                cuyy_x, cuyy_y = yy[f'x{i}'], yy[f'y{i}']
                line(doc, prev_x, prev_y, cuyy_x, cuyy_y, layer="?�이??)
                prev_x, prev_y = cuyy_x, cuyy_y
            line(doc, prev_x, prev_y, yy['x1'], yy['y1'], layer="?�이??)  

            d(doc, yy['x1'] , yy['y1'], yy['x2'] , yy['y2'], 100, text_height=0.20, direction="down", dim_style=over1000dim_style)
            d(doc, yy['x6'] , yy['y6'], yy['x1'] , yy['y1'], 80, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x6'] , yy['y6'], yy['x5'] , yy['y5'], 80, text_height=0.20, direction="left", dim_style=over1000dim_style)
            d(doc, yy['x1'] , yy['y1'], yy['x5'] , yy['y5'], 150, text_height=0.20, direction="left", dim_style=over1000dim_style)

        doc.saveas(global_data["file_name"])
        print(f" 저장 파일명: '{global_data['file_name']}' 저장 완료!")

if __name__ == '__main__':
    main()
