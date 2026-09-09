# ============================================================
# Seoul_Bus_Drive_Recorder v1.19
#
# 【프로그램 설명】
#   서울시 공공데이터 API를 이용하여 특정 버스 노선의
#   운행 출발/종료 시각을 자동으로 기록하고 엑셀로 저장하는 프로그램.
#
# ★ 전체 구조 목차 ★
# ══════════════════════════════════════════════════════════
# 【1】 라이브러리 가져오기
#      ├─ 【1-1】 파이썬 기본 라이브러리 (sys, os, time, threading 등)
#      ├─ 【1-2】 PySide6 GUI 라이브러리 (창·버튼·표·그래픽 등)
#      └─ 【1-3】 외부 선택 라이브러리 (requests, pandas, openpyxl, cryptography)
# 【2】 전역 상수
#      ├─ 【2-1】 프로그램 버전 상수 (APP_VERSION)
#      ├─ 【2-2】 글꼴 이름 (FONT_FAMILY, FONT_MONO)
#      ├─ 【2-3】 아이콘/이미지 데이터 (ICON_B64, GG_IMG_B64 - Base64)
#      ├─ 【2-4】 API URL 상수 (URL_POS1 ~ URL_SRCH)
#      ├─ 【2-5】 노선 종류 이름표·색상표 (ROUTE_TYPE_LABEL, ROUTE_TYPE_COLOR)
#      └─ 【2-6】 지도 그리기 크기/위치 상수 (CELL_W, CELL_H, PAD_X 등)
# 【3】 유틸리티 함수 및 클래스
#      ├─ 【3-1】  _make_palette()         라이트/다크 색상 팔레트 생성
#      ├─ 【3-2】  darken_color()          색상 어둡게 조정
#      ├─ 【3-3】  truncate_name()         긴 정류소 이름 줄임표 처리
#      ├─ 【3-4】  fmt_bus_no()            버스 번호판 포맷
#      ├─ 【3-5】  format_remain_time()    초 → "X분 Y초" 변환
#      ├─ 【3-6】  format_hhmm()           날짜문자열 → HH:MM 추출
#      ├─ 【3-7】  format_datetm()         날짜문자열 → YYYY-MM-DD HH:MM:SS
#      ├─ 【3-8】  load_pixmap_from_b64()  Base64 → QPixmap 변환
#      ├─ 【3-9】  make_bus_pixmap()       버스 아이콘 생성
#      ├─ 【3-10】 get_app_bg_color()      앱 배경색 반환
#      ├─ 【3-11】 get_text_color()        글자색 반환
#      ├─ 【3-12】 get_base_color()        표 배경색 반환
#      ├─ 【3-13】 get_header_bg_color()   표 헤더 배경색 반환
#      ├─ 【3-14】 detect_os_dark_mode()   OS 다크모드 감지
#      ├─ 【3-15】 ElidedLabel             말줄임 처리 레이블 (로그 헤더 파일 경로용)
#      ├─ 【3-16】 _licenses_dir()         licenses 폴더 경로 반환
#      └─ 【3-17】 _open_licenses_dir()    licenses 폴더를 파일탐색기로 열기
# 【4】 RouteMapPanel 클래스 (노선 지도 패널 위젯)
#      ├─ 【4-1】  __init__()              초기화·위젯 구성·타이머 설정
#      ├─ 【4-2】  set_sect_speeds()       구간 속도 데이터 저장
#      ├─ 【4-3】  eventFilter()           마우스 이동 시 툴팁(말풍선) 처리
#      ├─ 【4-4】  resizeEvent()           창 크기 변경 감지
#      ├─ 【4-5】  _on_resize_done()       크기 변경 완료 후 지도 재그리기
#      ├─ 【4-6】  _calc_layout()          창 크기에 맞는 셀 크기·폰트 계산
#      ├─ 【4-7】  _update_bg_color()      지도 배경색 갱신
#      ├─ 【4-8】  _update_tip_style()     툴팁 스타일 갱신
#      ├─ 【4-9】  refresh_theme()         테마 전체 갱신
#      ├─ 【4-10】 _rebuild_table_colors() 표 색상 재적용
#      ├─ 【4-11】 load_route()            정류소 목록 로드 및 지도 그리기
#      ├─ 【4-12】 update_buses()          버스 위치 갱신
#      ├─ 【4-13】 _speed_color()          구간 속도 → 선 색상 반환
#      ├─ 【4-14】 _draw()           ★     지도 전체 그리기 핵심 함수
#      ├─ 【4-15】 _build_table()          종점 도착 예정 표 생성
#      ├─ 【4-16】 _tick_countdown()       1초마다 남은 시간 카운트다운
#      ├─ 【4-17】 pause_tick()            카운트다운 일시 정지
#      └─ 【4-18】 resume_tick()           카운트다운 재개
# 【5】 RecordTable 클래스 (운행 기록 표 위젯)
#      ├─ 【5-1】  __init__()              초기화·표 구성
#      └─ 【5-2】  add_row()              새 기록 행 추가
# 【6】 SeoulBusRecorder 클래스 (메인 창 - 프로그램 본체)
#      ├─ 【6-1】  __init__()              초기화·시그널·변수 선언
#      ├─ 【6-2】  _make_info_html()       메뉴바 우측 제작자 HTML 생성
#      ├─ 【6-3】  _setup_ui()             메뉴·레이아웃·위젯 전체 구성
#      ├─ 【6-4】  _toggle_log_panel()     로그창 접기/펼치기
#      ├─ 【6-5】  _open_save_file()          기록 파일 바로 열기
#      ├─ 【6-6】  _open_save_folder()        기록 파일 폴더 열기
#      ├─ 【6-7】  _apply_theme()          라이트/다크 테마 전체 적용
#      ├─ 【6-8】  _slot_log()             로그 메시지 화면 출력 [슬롯]
#      ├─ 【6-9】  _slot_record()          운행 기록 표 추가 [슬롯]
#      ├─ 【6-10】 _slot_update_map()      지도 버스 위치 갱신 [슬롯]
#      ├─ 【6-11】 _slot_clear_map()       지도 버스 제거 [슬롯]
#      ├─ 【6-12】 log()                   로그 시그널 발행
#      ├─ 【6-13】 _make_fernet()          암호화 객체 생성
#      ├─ 【6-14】 _enc_key()              API 키 암호화
#      ├─ 【6-15】 _dec_key()              API 키 복호화
#      ├─ 【6-16】 _load_config()          설정 파일에서 API 키 읽기
#      ├─ 【6-17】 _save_config()          설정 파일에 API 키 저장
#      ├─ 【6-18】 _show_key_input()       인증키 입력/검증 대화상자
#      ├─ 【6-19】 _show_route_search()    노선 검색/즐겨찾기 대화상자
#      ├─ 【6-20】 _load_route_from_search() 노선 상세 정보 API 로드
#      ├─ 【6-21】 _slot_route_loaded()    노선 로드 완료 처리 [슬롯]
#      ├─ 【6-22】 fetch_api()       ★     서울시 버스 API 공통 호출
#      ├─ 【6-23】 _fetch_first_time()     첫차 시각 조회
#      ├─ 【6-24】 _on_schedule_toggle()   예약 버튼 4가지 분기 처리
#      ├─ 【6-25】 _ask_scheduled_start()  기록 시작 예약 시각 입력 창
#      ├─ 【6-26】 _register_schedule_timer()   시작 예약 폴링 타이머
#      ├─ 【6-27】 _ask_scheduled_stop()   기록 중지 예약 시각 입력 창
#      ├─ 【6-28】 _register_stop_schedule_timer() 중지 예약 폴링 타이머
#      ├─ 【6-29】 _stop_monitoring_silent()  예약 자동 중지 (확인창 없음)
#      ├─ 【6-30】 _on_toggle()            기록 시작/중지 토글 (예약·이전기록 팝업 포함)
#      ├─ 【6-31】 _start_monitoring()     모니터링 시작 (중복 실행 가드 포함)
#      ├─ 【6-32】 _clear_recorded_data()  기록창·recorded_data 초기화
#      ├─ 【6-33】 _stop_monitoring()      모니터링 중지 (중지예약 팝업 분기 포함)
#      ├─ 【6-34】 _main_loop()            백그라운드 갱신 루프 스레드
#      ├─ 【6-35】 _refresh_data()         한 번의 데이터 갱신 실행
#      ├─ 【6-36】 _process_routes()  ★    버스 위치 분석·운행 판정 핵심
#      ├─ 【6-37】 _record()               운행 이벤트 기록
#      ├─ 【6-38】 _perform_auto_save()    자동 엑셀 저장 (락 보호)
#      ├─ 【6-39】 _core_excel_save()      엑셀 파일 저장 (날짜별 시트)
#      ├─ 【6-40】 _write_source_sheet()   엑셀 "출처 및 수집방법" 시트 생성
#      ├─ 【6-41】 _axs()                  엑셀 시트 스타일 적용
#      ├─ 【6-42】 _ask_interval()         갱신 주기 입력 대화상자
#      ├─ 【6-43】 _show_api_status()      API 호출 현황 대화상자
#      ├─ 【6-44】 _show_program_info()    프로그램 정보 대화상자
#      ├─ 【6-45】 closeEvent()            창 닫기 이벤트 처리
#      ├─ 【6-46】 _load_favorites()       즐겨찾기 불러오기
#      ├─ 【6-47】 _save_favorites()       즐겨찾기 저장
#      └─ 【6-48】 _cleanup_search_signal() 검색 시그널 연결 해제
# 【7】 프로그램 진입점 (if __name__ == "__main__")
# ══════════════════════════════════════════════════════════
# ============================================================
# Seoul_Bus_Drive_Recorder_v1.19
# ============================================================

# ══════════════════════════════════════════════════════════
# 【1-1】 파이썬 기본 라이브러리
#   파이썬 설치 시 기본 포함 - 별도 설치 불필요
# ══════════════════════════════════════════════════════════
import subprocess
import sys# sys        : 파이썬 인터프리터 제어 (sys.exit 강제 종료, sys.platform OS 판별 등)
import os
# os         : 운영체제 인터페이스 (파일 경로 조작, 파일 존재 확인, 디렉터리 등)
import time
# time       : 시간 제어 (time.sleep 대기, time.time 현재 시각 숫자로 반환)
import threading
# threading  : 멀티스레드 지원 - API 호출을 백그라운드에서 실행해
#              UI가 멈추지 않도록 한다. threading.Thread + threading.Lock 사용.
import colorsys
# colorsys   : 색상 공간 변환 (RGB ↔ HSV). 선 색상 어둡게 만들 때 사용.
import base64
# base64     : 이진 데이터↔ASCII 텍스트 인코딩. 이미지를 코드에 내장할 때 사용.
import re
# re         : 정규 표현식. 버스 번호판 "1234사5678" 형식 파악 등에 사용.
import json
# json       : JSON 직렬화/역직렬화. 즐겨찾기 노선을 설정 파일에 저장할 때 사용.
import configparser
# configparser: INI 형식 설정 파일 읽고 쓰기. API 키 보관 파일(*.ini)에 사용.
from collections import defaultdict
# defaultdict: 키가 없어도 기본값을 자동 생성하는 딕셔너리.
#              defaultdict(list) → 없는 키 접근 시 빈 [] 반환.
#              지도 위 버스 레이블 겹침 방지용 그룹핑에 사용.
from datetime import datetime, timedelta
# datetime   : 날짜·시간 클래스.
#   datetime  → 연·월·일·시·분·초를 담는 객체
#   timedelta → 시간 간격 (예: timedelta(minutes=5) = 5분)
import xml.etree.ElementTree as ET
# xml.etree.ElementTree (ET): XML 파싱 라이브러리.
#   서울시 버스 API 응답이 XML 형식이므로
#   ET.fromstring(text)으로 트리 구조로 변환하고
#   root.findtext(".//태그명")으로 원하는 값 추출.
from urllib.parse import unquote
# urllib.parse.unquote: URL 인코딩(%XX)을 원래 문자로 복원.
#   공공데이터 API 키가 URL 인코딩 형태로 올 수 있어 복호화 전 unquote 적용.

# ══════════════════════════════════════════════════════════
# 【1-2】 PySide6 화면(GUI) 라이브러리
#   Qt 프레임워크의 파이썬 공식 바인딩.
#   창·버튼·표·그래픽 등 모든 화면 요소를 만들고 제어한다.
# ══════════════════════════════════════════════════════════
# ─ QtWidgets ─
#   QApplication    : Qt 앱 뼈대. 프로그램 시작 시 제일 먼저 생성해야 함.
#   QMainWindow     : 메뉴바·상태바를 갖춘 메인 창 클래스.
#   QWidget         : 모든 위젯(화면 요소)의 기본 클래스.
#   QVBoxLayout     : 자식 위젯을 위→아래로 순서대로 배치.
#   QHBoxLayout     : 자식 위젯을 왼쪽→오른쪽으로 배치.
#   QSplitter       : 두 영역 사이의 드래그 가능한 구분선 (크기 조절).
#   QLabel          : 텍스트·이미지 표시용 라벨 (클릭 불가).
#   QPushButton     : 클릭 가능한 일반 버튼.
#   QTableWidget    : 행·열로 구성된 스프레드시트형 표.
#   QTableWidgetItem: 표의 각 셀에 들어가는 항목.
#   QHeaderView     : 표 헤더(머리줄) 크기 조절 방식 설정.
#   QPlainTextEdit  : 여러 줄 텍스트 표시·편집 (로그 창에 사용).
#   QDialog         : 팝업 대화 상자 기본 클래스.
#   QDialogButtonBox: OK·Cancel 등 표준 버튼 묶음.
#   QLineEdit       : 한 줄 텍스트 입력창.
#   QMessageBox     : 알림·경고·확인 팝업 메시지 창.
#   QGraphicsView   : 2D 도형·이미지를 표시하는 뷰 (노선 지도).
#   QGraphicsScene  : QGraphicsView에 그릴 객체들을 담는 "무대".
#   QAbstractItemView: 리스트·표·트리 공통 기능 (예: 편집 금지 설정).
#   QStackedWidget  : 여러 위젯을 쌓아두고 하나씩 보여줌 (노선별 탭 대체).
#   QStyleFactory   : 스타일 이름("Fusion")으로 스타일 객체 생성.
#   QTreeWidget     : 계층(트리) 구조 목록 위젯 (노선 검색 결과).
#   QTreeWidgetItem : 트리 위젯의 각 항목.
#   QTabWidget      : 탭 버튼으로 여러 페이지를 전환.
#   QSpinBox        : 위아래 화살표로 숫자를 조절하는 입력창.
# ─ QtCore ─
#   Qt      : 전역 상수 모음 (Qt.AlignCenter, Qt.Horizontal 등).
#   QTimer  : 일정 시간(ms)마다 함수를 자동 호출하는 타이머.
#   Signal  : 이벤트 발신 통로 정의. 백그라운드→UI 스레드 간 안전한 통신.
#   Slot    : 시그널을 받아 처리하는 함수 표시 데코레이터.
# ─ QtGui ─
#   QColor   : 색상 표현 (RGB, 헥사코드 "#RRGGBB", 색상 이름 "red" 등).
#   QPen     : 선 그리기 도구 (색·두께·스타일).
#   QBrush   : 도형 채우기 도구 (색·패턴).
#   QFont    : 글꼴 설정 (이름·크기·굵기).
#   QPixmap  : 픽셀 기반 이미지 (PNG/JPEG 로드·화면 표시).
#   QPainter : 도형·텍스트·이미지를 실제로 그리는 "화가" 클래스.
#   QIcon    : 창 제목줄·메뉴에 표시되는 아이콘.
#   QAction  : 메뉴·툴바에 추가되는 동작(이름+클릭 이벤트).
#   QPalette : 위젯 전체 색상 팔레트 (배경·글자·버튼·강조색 일괄 설정).
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QLabel, QPushButton, QSizePolicy,
    QTableWidget, QTableWidgetItem, QHeaderView, QPlainTextEdit,
    QDialog, QDialogButtonBox, QLineEdit, QMessageBox,
    QGraphicsView, QGraphicsScene,
    QAbstractItemView, QStackedWidget, QStyleFactory,
    QTreeWidget, QTreeWidgetItem, QTabWidget, QSpinBox,
    QScrollArea
)
from PySide6.QtCore import Qt, QTimer, Signal, Slot
from PySide6.QtGui import (
    QColor, QPen, QBrush, QFont, QPixmap, QPainter,
    QIcon, QAction, QPalette
)

# ══════════════════════════════════════════════════════════
# 【1-3】 외부 선택 라이브러리 (pip install로 따로 설치 필요)
#   try~except 구조: 설치되어 있으면 가져오고,
#   없으면 설치 안내 메시지를 출력한 뒤 sys.exit(1)으로 즉시 종료.
# ══════════════════════════════════════════════════════════
# requests   : HTTP 요청 라이브러리.
#              requests.get(url, params={...}, timeout=10)으로
#              서울시 버스 API를 호출하고 XML 응답을 받아온다.
# pandas(pd) : 표(DataFrame) 형태로 데이터를 다루는 라이브러리.
#              pd.DataFrame(...)으로 운행 기록을 표로 만들고
#              df.to_excel(...)으로 엑셀 파일에 저장.
# openpyxl   : 엑셀(.xlsx) 파일 직접 제어 라이브러리.
#   XlFont    : 엑셀 셀 글꼴 스타일 (Font와 이름 충돌 방지를 위해 별칭 사용).
#   PatternFill: 엑셀 셀 배경을 특정 색으로 채우기.
#   Alignment  : 엑셀 셀 텍스트 정렬 (가운데·왼쪽 등).
try:
    import requests
    import pandas as pd
    from openpyxl.styles import Font as XlFont, PatternFill, Alignment
except ImportError as e:
    print(f"필수 패키지 누락: {e}\npip install requests pandas openpyxl PySide6")
    sys.exit(1)

# ══════════════════════════════════════════════════════════
# 【1-3 계속】 암호화 라이브러리 (선택적 - 없어도 실행 가능)
#   있으면 API 키를 암호화해서 저장, 없으면 평문으로 저장.
# ══════════════════════════════════════════════════════════
# cryptography.Fernet: 대칭키 암호화. 같은 키로 암호화·복호화.
#   API 키를 설정 파일에 그냥 저장하면 다른 사람이 볼 수 있으므로
#   암호화해서 보관.
# hashlib: 해시 함수 모음.
#   pbkdf2_hmac()으로 _SECRET 문장을 32바이트 암호화 키로 변환.
# _CRYPTO_OK: True이면 암호화 사용 가능, False이면 평문 저장.
try:
    from cryptography.fernet import Fernet as _Fernet
    import hashlib as _hl
    _CRYPTO_OK = True
except ImportError:
    _CRYPTO_OK = False
    _Fernet = None
    _hl = None

# ══════════════════════════════════════════════════════════
# 【2-1】 프로그램 버전 상수
#   창 제목·정보창·엑셀 출처 시트 등에서 공통으로 참조하는
#   단일 버전 문자열. 버전 변경 시 이 값만 수정하면 된다.
# ══════════════════════════════════════════════════════════
APP_VERSION = "1.19"

# ══════════════════════════════════════════════════════════
# 【2-2】 글꼴(폰트) 이름 상수
#   프로그램 전체에서 일관된 글꼴을 유지하기 위해 상수로 정의.
# ══════════════════════════════════════════════════════════
# FONT_FAMILY: 일반 텍스트용 폰트 - "맑은 고딕" (한글이 깔끔한 Windows 기본 폰트)
# FONT_MONO  : 로그창 등 코드형 텍스트용 고정폭(등폭) 폰트 - "Consolas"
FONT_FAMILY = "맑은 고딕"
FONT_MONO = "Consolas"

# ══════════════════════════════════════════════════════════
# 【2-3】 아이콘/이미지 데이터 (Base64 인코딩)
#   외부 이미지 파일 없이 코드 안에 이미지를 텍스트로 내장.
#   load_pixmap_from_b64()로 복원해서 사용.
# ══════════════════════════════════════════════════════════
# ICON_B64   : 프로그램 창 아이콘 (버스 모양 PNG). 박국환 직접 제작 (2026),
#              제3자 권리 없음. assets/bus.ico·assets/bus.icns(빌드용)와 동일 디자인.
# GG_IMG_B64 : 공공누리 제1유형 마크 이미지 (프로그램 정보 창에 원본 크기 그대로 표시).
#              변형(색상·비율·문구) 없이 원본 그대로 내장 — scaled() 사용 금지.
# ── 이미지 Base64 ──
ICON_B64 = "iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAYAAAD0eNT6AAAACXBIWXMAAAsTAAALEwEAmpwYAAAgAElEQVR4nOy9f/BtyVEf1nO+72m1knbf6kcEAoHQSjIFxGBjwDGG2CrAjhNcQBKTkIQ/bJwYiI3tcpVd0Q+E0Ao7lZCY/DI4dlKVuFIhdsWQFBVIYew4hZNQKUISQhmkt8JRDLiQQbur1Wr13j2TP+b0nZ7P+XTP3Lf3fu/9vj1d9d73nnNmerp7Zro/PeecOSIbbbTRRhtttNFGG2200UYbbbTRRhtttNFGG2200UYbbbTRRhtttNFGG2200UYbbbTRRhtttNFGG2200UYbbbTRRhtttNFGG2200UYbbbTRRhtttNFGG2200UYbbbTRRhtttNFGG2200UYbbbTRCSmdW4CLo+/5+Vc8/sgrP3fa5S+SPL1OprwTEZEsSZLkOU95SnOa85Sbekmy5JxF0nQOsdeU5/o7TVU+PZWS5HX/T2lOIiIr/UbLJMnIV8tbivizNqjN96S6gu1R56hvmrJqr8PmR2SXEbseSsxGIiKzTHOrNyGn/0fao206x1ivqe/Y+BS2OiZ58jH7tLSM04Ntn+d9naFxacoH8jLq91Wda1h2qL8OmZOOPF5btGyR/Jn5lvz8s5/66C/L+995vyvjy4g2ALDQm973D1/1/Cs+9c5plrfnJL8lZfmSnORxEdktRZKIZJGURXIqf5FyFkkXYtNmkiX/3Krecp7pN1KG2SCTtiL+rA3P5tqmCGkXyvf6xuMzSqFdBuz6UttTG3WCf6nzADrafrRtRsdYz9b35sspbHVM8uRj9mmuP+D4svVGfAy2c4g9e33VyJLW5br8D52TzvVojjU0SU7Py5R/Luf89yXLLz17/6M/uQGBQhcSrM5E3/Pzr3j89iv/+STpD0mWrxKRN55bpI022mijjU5KH5UsP52n/F88++63/w/nFuac9LIEAK/9C3fvzLv8b0hO/6aIvOXc8my00UYbbXT9lER+QbL8lY/ff/IH5P1p7td4uOjlBQDel6fHb3/kO5Pk94jIZ5xbnI022mijjS6Cfjmn9N5n3/3kXzu3INdJLxsA8PgH7r5DRP5qSvLV55Zlo4022mijS6T04/fvpz/8/Pe89dfOLcl10MsCADzxgQ9/Y07pb4jI1bll2WijjTba6KLpeUnyDc+8+21/69yCnJou5JW109HjTz39x3NKf1O24L/RRhtttFGfXi1ZfvKJpz78TecW5NT0UAOAOx+4+xeS5P/w3HJstNFGG210syhL+m9f+9TTf/DccpySHtpbAE988O4fyVn+6rnl2GijjTba6MZSlpy/9pn3vv2nzi3IKeihBACPfeDu109J/vtzy7HRRhtttNGNpxfyPL3z2e9+6/92bkGOTQ8dAHjtB5/+4jnnvycirz63LBtttNFGG918SiK/ups+/Vufe9cX/ONzy3JMerieAXhfnnLOf0m24L/RRhtttNGRKIu8aZpf8UPnluPY9FABgDu3734gi3zlueXYaKONNtrooaN/4YkP3v0j5xbimPTQ3AJ43fd+6It20/RzInLr3LJstNFGG230UNKv7ab5iz/xrnf8+rkFOQY9NCsAu2n647IF/4022mijjU5Hn3mVp4dmFeChAACv+cAvf4GIfNu55dhoo4022ughpyzfcm4RjkUPBQCY0u4Picjtc8ux0UYbbbTRQ09f8rBsEPRQAICU5OvPLcNGG2200UYvD8op//5zy3AMuvEA4M5Td79Msnz5ueXYaKONNtro5UE5y9eeW4Zj0I0HACLp955bgo022mijjV5W9PmPfd+Hfte5hXipdPMBQMpPnluEjTbaaKONXl50NU8bADg7zem3nVuEjTbaaKONXmaU5MYnnzcbALzvb9+SJJ9xbjE22mijjTZ6eVEWedO5ZXipdKMBwJvvvP22SH7VueXYaKONNtroZUY5ve7cIrxUutEA4MVnn7kSkUfOLcdGG2200UYvN8qvObcEL5VuNABY6KH5nsFGG2200UYbXRfdaACQH72VZAMAG2200UYbXTulq3NL8FLpRgOA9ML9LCLzueXYaKONNtpoo5tGNxoATK9+QxZJGwDYaKONNtpoowPpRgOA+fmPJZG83QLYaKONNtpoowPpRgOAhR4GHTbaaKONNrpRdPOTzy14brTRRhtttNHhdOPj541XYKONNtpoo43OQDf++bMNAGy00UYbbbTR4bQBgAugfG4BNtpoo402etnRjY89DwMAuPEPYmy00UYbbXTjaNsIaKONNtpoo41ehrTdAthoo4022mijlyFtAGCjjTbaaKONNrp5tAGAjTbaaKONNjqcbvzzZ7fOLcBNpSQiuyzywv1Z7u/qSLipj4WmRYE8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"
GG_IMG_B64 = "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/0sIiUQiMtbIGItEIsj2p3bK87Bhw1asWLF9+/Y77ryza9eu69atw22XX365qmmbN28WMkFDVcE3EYzctGkTBEMikYBZAbIXdjBMujkZY6WlpV6v1+12C86Li4uR3KkoisvlKi4uJoREo9H333/fNM1rr732yu7dJYpOGDcuKyvr8OHDyCI1TTMYDCYSiUAgUFtb6/f758+fn4jFrGSSm6bb7ebHJn1LSXNazPjT75OXBI/1rV69GqIG3vVOnTq1bdsWqWCRSOTIkSOfffbZtm3bAoGA1+slhGiatnr1auQVeb3eiy666MILL2zRooWmaYKSeDS2Z9/ejes3fLXjaytpwsfWv3//a6+7jhB7JIgQAmO6TwpCuFyuyZMnm6apahoRAg5kl8sVjUYzMjI0h8NlGJ06d27evLlsk5jKa0+975qmZWVl1dbWXnjhhQMHDuzatWtOTg7kTzKZLC8v37p168KFC4uKisCVGCHwJoTD4SuuuGLgwIGdO3du3ry5TOQ+evTohg0bFi9e/Mknn6Snp1vJpNPprK6uLigomDVrVlZ29rAnnjh06NCBAwfGjRs3YsSIJUuWPProo1988QVPqZwBnm3cuLFly5b33XefZVdxwepBOACcbu3atcLOckEgrV+/flAycbKfbd1KKR0wYEDSMj0eT/Oc3JUrVy5duhQPBJbjA6U0FAq1bt16+owZHfPyxo4Zs23bNiHEU0899eKLLy5YsGDo0KGrVq1SFEUOcGKMlZeXb9myBVpJx44dO1144aaNG/ft2wfr4Oqrr27UqNGqVasCgQDCMTU1NS6X6+OPPz565MiQIUNWrVq1fv36Dh06/OlPf4pFox9++KHMV5fxAofDsX79+vXr1yP+ct5553Xu3LlNmzZer9flclVVVR09erS4uHjNmjVVVVWItANja2pqKGMXX3wxdgM5iN26dSNCHD16FI5kBJik22LNmjUejyc/Px/anKqqcNnm5OSAAX366acbN26EhxIWk5SaEtnYaer+etqonds1QPhfuTjIzJycnLFjx1511VWgw+OaXm3ZvPm5557bsGGD1+slnGuGYSYSl3fvPnr06PMvuIAwVpeCRuromSeTH61YMWrEiLKyMsPn271r19jnnkv1bSSTya5du15z7bU1lZWzZ8+GglDnTNa0PXv21NlUQhBKR4waVVtbCzM4PT3d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"

# STATION_INFO: 정류소 정보 캐시 딕셔너리 (현재는 빈 딕셔너리로 초기화만 됨).
#               향후 정류소 정보를 로컬에 캐싱할 때 사용하기 위해 예약.
STATION_INFO = {}

# ══════════════════════════════════════════════════════════
# 【2-4】 서울시 버스 공공데이터 API 주소(URL) 상수
#   공공데이터포털(data.go.kr) 서울시 버스 OpenAPI 엔드포인트.
#   fetch_api(url, params)를 통해 호출되며,
#   ServiceKey + 파라미터를 GET으로 전송하면 XML 형태로 응답.
# ══════════════════════════════════════════════════════════
# URL_POS1 (getBusPosByRtid):    노선 ID → 해당 노선 모든 버스 현재 위치 조회.
#                                 출발 감지(첫·두 번째 정류소 통과 판정)에 사용.
# URL_POS2 (getBusPosByRouteSt): 노선의 특정 구간(startOrd~endOrd) 버스 위치.
#                                 종점 도착 판정에 사용.
# URL_SLST (getStaionByRoute):   노선의 전체 정류소 목록 (ID·이름·순번·회차지 등).
#                                 지도 그리기 및 구간 속도 수집에 사용.
# URL_RINF (getRouteInfo):       노선 기본 정보 (운수사·첫차·막차·길이 등).
# URL_SRCH (getBusRouteList):    노선 번호 문자열로 노선 목록 검색.
#                                 노선 검색 대화상자에서 사용.
URL_POS1 = "http://ws.bus.go.kr/api/rest/buspos/getBusPosByRtid"
URL_POS2 = "http://ws.bus.go.kr/api/rest/buspos/getBusPosByRouteSt"
URL_SLST = "http://ws.bus.go.kr/api/rest/busRouteInfo/getStaionByRoute"
URL_RINF = "http://ws.bus.go.kr/api/rest/busRouteInfo/getRouteInfo"
URL_SRCH = "http://ws.bus.go.kr/api/rest/busRouteInfo/getBusRouteList"

# ══════════════════════════════════════════════════════════
# 【2-5】 노선 종류 이름표(LABEL) · 색상표(COLOR)
#   API routeType 코드(숫자 문자열)를 이름·색상으로 매핑.
# ══════════════════════════════════════════════════════════
# 코드 → 이름   → 지도 선 색상(대략)
#  "1" → 공항버스  → 황금색    "#C8A000"
#  "2" → 마을버스  → 연초록    "#6DBF67"
#  "3" → 간선버스  → 파란색    "#1E6FD9"  (서울 주요 간선)
#  "4" → 지선버스  → 짙은 초록 "#005A00"  (간선 보조)
#  "5" → 순환버스  → 황금색    "#E0B800"
#  "6" → 광역버스  → 짙은 빨강 "#5A0E11"  (수도권 광역)
#  "7" → 인천버스  → 청록색    "#20B2AA"
#  "8" → 경기버스  → 청록색    "#20B2AA"
#  "9" → 폐지노선  → 회색      "#888888"
#  "0" → 공용      → 어두운 회색 "#333333"
# DEFAULT_LINE_COLOR: 코드 없을 때 사용하는 기본 선 색상 "#333333"
ROUTE_TYPE_LABEL = {
    "1": "공항", "2": "마을", "3": "간선", "4": "지선",
    "5": "순환", "6": "광역", "7": "인천", "8": "경기",
    "9": "폐지", "0": "공용",
}
ROUTE_TYPE_COLOR = {
    "1": "#C8A000", "2": "#6DBF67", "3": "#1E6FD9", "4": "#005A00",
    "5": "#E0B800", "6": "#5A0E11", "7": "#20B2AA", "8": "#20B2AA",
    "9": "#888888", "0": "#333333",
}
DEFAULT_LINE_COLOR = "#333333"

# ══════════════════════════════════════════════════════════
# 【2-6】 노선 지도 그리기용 크기/위치 상수
#   QGraphicsScene에 정류소·선을 그릴 때 사용하는 기본 크기값.
#   창 크기에 따라 _calc_layout()에서 동적으로 조정된다.
# ══════════════════════════════════════════════════════════
# STOPS_PER_ROW  : 한 행에 배치할 정류소 수 기본값 (15개)
# CELL_W / CELL_H: 정류소 하나가 차지하는 셀의 가로(42)/세로(70) 픽셀
# PAD_X / PAD_Y  : 지도 좌우(55)/상단(110) 여백. PAD_Y는 정보 텍스트 여백 포함.
# CIRCLE_R       : 일반 정류소 원 반지름 (8픽셀)
# CIRCLE_R_SPECIAL: 출발·회차·종점 정류소 원 반지름 (13픽셀, 더 크게)
# TEXT_X_OFFSET  : 정류소 이름 텍스트 X축 미세 조정 (-1픽셀)
# FONT_CAP       : 셀 너비가 이 값(150) 이상이면 폰트 크기 최대(10pt) 고정
STOPS_PER_ROW = 15
CELL_W = 42
CELL_H = 70
PAD_X = 55
PAD_Y = 110
CIRCLE_R = 8
CIRCLE_R_SPECIAL = 13
TEXT_X_OFFSET = -1
FONT_CAP = 150

# ══════════════════════════════════════════════════════════
# 【3-1】 _make_palette(mode) → QPalette
#   라이트/다크 모드에 맞는 Qt 색상 팔레트를 만들어 반환.
#   QPalette는 창 배경·글자·버튼·강조색 등을 한꺼번에 설정하는 객체.
#   앱 전체에 app.setPalette(p)로 적용하면 모든 위젯의 색상이 바뀐다.
#
#   @param mode : "dark" → 다크 테마, 그 외 → 라이트 테마
#   @return     : 색상이 설정된 QPalette 객체
#
#   다크 모드 주요 색상:
#     Window(배경) = #353535, WindowText(글자) = #DCDCDC
#     Base(표/입력창) = #232323, Highlight(선택) = #2A82DA
#   라이트 모드 주요 색상:
#     Window(배경) = #F0F0F0, WindowText(글자) = #000000
#     Base(표/입력창) = #FFFFFF, Highlight(선택) = #0078D7
# ══════════════════════════════════════════════════════════
def _make_palette(mode):
    p = QPalette()
    if mode == "dark":
        p.setColor(QPalette.Window, QColor(53, 53, 53))
        p.setColor(QPalette.WindowText, QColor(220, 220, 220))
        p.setColor(QPalette.Base, QColor(35, 35, 35))
        p.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        p.setColor(QPalette.Text, QColor(220, 220, 220))
        p.setColor(QPalette.Button, QColor(53, 53, 53))
        p.setColor(QPalette.ButtonText, QColor(220, 220, 220))
        p.setColor(QPalette.Highlight, QColor(42, 130, 218))
        p.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        p.setColor(QPalette.ToolTipBase, QColor(53, 53, 53))
        p.setColor(QPalette.ToolTipText, QColor(220, 220, 220))
        p.setColor(QPalette.Link, QColor(100, 180, 255))
        p.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(90, 90, 90))
        p.setColor(QPalette.Disabled, QPalette.WindowText, QColor(90, 90, 90))
        p.setColor(QPalette.Disabled, QPalette.Text, QColor(90, 90, 90))
    else:
        p.setColor(QPalette.Window, QColor(240, 240, 240))
        p.setColor(QPalette.WindowText, QColor(0, 0, 0))
        p.setColor(QPalette.Base, QColor(255, 255, 255))
        p.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
        p.setColor(QPalette.Text, QColor(0, 0, 0))
        p.setColor(QPalette.Button, QColor(240, 240, 240))
        p.setColor(QPalette.ButtonText, QColor(0, 0, 0))
        p.setColor(QPalette.Highlight, QColor(0, 120, 215))
        p.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        p.setColor(QPalette.ToolTipBase, QColor(255, 255, 220))
        p.setColor(QPalette.ToolTipText, QColor(0, 0, 0))
        p.setColor(QPalette.Link, QColor(30, 100, 220))
        p.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(170, 170, 170))
        p.setColor(QPalette.Disabled, QPalette.WindowText, QColor(170, 170, 170))
        p.setColor(QPalette.Disabled, QPalette.Text, QColor(170, 170, 170))
    return p

# ══════════════════════════════════════════════════════════
# 【3-2】 darken_color(hex_color, factor=0.60) → str
#   헥사 색상 코드("#RRGGBB")를 받아 factor 비율만큼 밝기를 낮춰 반환.
#   노선 지도에서 회차 이후 구간 선을 원래 색보다 어둡게 표시할 때 사용.
#
#   처리 과정:
#   1) 헥사 코드 → RGB (0.0~1.0 범위)
#   2) RGB → HSV (색조 H, 채도 S, 명도 V) 변환
#   3) 명도 V에 factor 곱해 낮춤 (0.6이면 40% 어두워짐)
#   4) HSV → RGB → 헥사 코드로 재변환 후 반환
#
#   @param hex_color : 입력 색상 (예: "#1E6FD9")
#   @param factor    : 밝기 비율 (0.0=완전검정, 1.0=원본, 0.6=40% 어둡게)
#   @return          : 어두워진 헥사 색상 (예: "#123F7A")
# ══════════════════════════════════════════════════════════
def darken_color(hex_color, factor=0.60):
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i+2], 16) / 255.0 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    r2, g2, b2 = colorsys.hsv_to_rgb(h, s, max(0.0, v * factor))
    return "#{:02X}{:02X}{:02X}".format(int(r2 * 255), int(g2 * 255), int(b2 * 255))

# ══════════════════════════════════════════════════════════
# 【3-3】 truncate_name(name, max_len=8) → str
#   정류소 이름이 max_len을 초과하면 앞 7글자만 남기고 ".."을 붙여 반환.
#   지도 위 좁은 공간에 이름을 표시할 때 공간이 부족하기 때문.
#   예: "서울역버스환승센터" (9글자) → "서울역버스환승.."
# ══════════════════════════════════════════════════════════
def truncate_name(name, max_len=8):
    return name if len(name) <= max_len else name[:7] + ".."

# ══════════════════════════════════════════════════════════
# 【3-4】 fmt_bus_no(raw) → str
#   API에서 받은 버스 번호판 문자열을 보기 좋게 포맷.
#   전세버스 형식 "1234사5678" → "서울 1234 사 5678"
#   일반 형식    "1234"       → "서울 1234"
#   정규식 r'^(\d+)(사)(\d+)$'으로 전세버스 형식을 구분.
# ══════════════════════════════════════════════════════════
def fmt_bus_no(raw):
    m = re.match(r'^(\d+)(사)(\d+)$', raw.strip())
    if m:
        return f"서울 {m.group(1)} {m.group(2)} {m.group(3)}"
    return f"서울 {raw}"

# ══════════════════════════════════════════════════════════
# 【3-5】 format_remain_time(seconds) → str
#   남은 시간(초)을 사람이 읽기 좋은 문자열로 변환.
#   3661 → "1시간 1분 1초",  125 → "2분 5초",  45 → "45초"
# ══════════════════════════════════════════════════════════
def format_remain_time(seconds):
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    if h > 0:
        return f"{h}시간 {m}분 {s}초"
    elif m > 0:
        return f"{m}분 {s}초"
    return f"{s}초"

# ══════════════════════════════════════════════════════════
# 【3-6】 format_hhmm(raw) → str
#   API 반환 날짜+시각 문자열에서 "HH:MM" 부분만 추출.
#   "20240101042500" (14자) → "04:25"  (8~10번째=시, 10~12번째=분)
#   "0425"           (4자)  → "04:25"  (앞 2자=시, 뒤 2자=분)
# ══════════════════════════════════════════════════════════
def format_hhmm(raw):
    if not raw or len(raw) < 4:
        return raw
    return f"{raw[8:10]}:{raw[10:12]}" if len(raw) >= 14 else f"{raw[:2]}:{raw[2:4]}"

# ══════════════════════════════════════════════════════════
# 【3-7】 format_datetm(raw) → str
#   14자리·12자리 숫자 형식 날짜+시각 문자열을 "YYYY-MM-DD HH:MM:SS"로 변환.
#   raw가 None이거나 파싱 실패 시 → 현재 시각을 포맷해서 반환.
#   예: "20240101042500" → "2024-01-01 04:25:00"
# ══════════════════════════════════════════════════════════
def format_datetm(raw):
    if not raw:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    raw = str(raw).strip()
    try:
        if len(raw) == 14:
            return f"{raw[0:4]}-{raw[4:6]}-{raw[6:8]} {raw[8:10]}:{raw[10:12]}:{raw[12:14]}"
        if len(raw) == 12:
            return f"{raw[0:4]}-{raw[4:6]}-{raw[6:8]} {raw[8:10]}:{raw[10:12]}:00"
    except:
        pass
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ══════════════════════════════════════════════════════════
# 【3-8】 load_pixmap_from_b64(b64_str) → QPixmap
#   Base64 인코딩된 이미지 문자열을 QPixmap 이미지 객체로 변환.
#   실패 시 빈(null) QPixmap 반환.
#   처리: base64.b64decode() → 이진 데이터 → QPixmap.loadFromData()
# ══════════════════════════════════════════════════════════
def load_pixmap_from_b64(b64_str):
    try:
        raw = base64.b64decode(b64_str)
        pm = QPixmap()
        pm.loadFromData(raw)
        return pm
    except:
        return QPixmap()

# ══════════════════════════════════════════════════════════
# 【3-9】 make_bus_pixmap() → QPixmap
#   지도에 버스 위치 표시용 25×25픽셀 아이콘을 만들어 반환.
#   ICON_B64 복원 실패 시 파란 사각형(#1E6FD9)으로 대체.
#   Qt.KeepAspectRatio + Qt.SmoothTransformation으로 비율 유지하며 축소.
# ══════════════════════════════════════════════════════════
def make_bus_pixmap():
    pm = load_pixmap_from_b64(ICON_B64)
    if pm.isNull():
        pm = QPixmap(25, 25)
        pm.fill(QColor("#1E6FD9"))
    else:
        pm = pm.scaled(25, 25, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    return pm

# ══════════════════════════════════════════════════════════
# 【3-10~13】 현재 테마 색상 getter 함수들
#   테마가 바뀌어도 항상 현재 팔레트의 올바른 색상을 반환.
# ══════════════════════════════════════════════════════════
# get_app_bg_color()   → QPalette.Window      (창 배경색)
# get_text_color()     → QPalette.WindowText  (기본 글자색)
# get_base_color()     → QPalette.Base        (표·입력창 배경색)
# get_header_bg_color()→ QPalette.Button      (표 헤더 배경색)
def get_app_bg_color():
    return QApplication.palette().color(QPalette.Window)

def get_text_color():
    return QApplication.palette().color(QPalette.WindowText)

def get_base_color():
    return QApplication.palette().color(QPalette.Base)

def get_header_bg_color():
    return QApplication.palette().color(QPalette.Button)

# ══════════════════════════════════════════════════════════
# 【3-14】 detect_os_dark_mode() → bool
#   OS의 다크 모드 설정을 감지. 프로그램 시작 시 테마 자동 선택에 사용.
#   Windows: 레지스트리 "AppsUseLightTheme" 값 확인 (0=다크)
#   macOS  : "defaults read -g AppleInterfaceStyle" 명령 실행 (결과 "Dark"=다크)
#   기타 OS: 항상 False(라이트 모드) 반환
# ══════════════════════════════════════════════════════════
def detect_os_dark_mode():
    try:
        if sys.platform == "win32":
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            winreg.CloseKey(key)
            return val == 0
        elif sys.platform == "darwin":
            import subprocess
            result = subprocess.run(
                ["defaults", "read", "-g", "AppleInterfaceStyle"],
                capture_output=True, text=True, timeout=3
            )
            return result.stdout.strip().lower() == "dark"
    except:
        pass
    return False

# ══════════════════════════════════════════════════════════
# 【3-15】 ElidedLabel — 말줄임 처리 QLabel
#   창 너비가 좁아질 때 텍스트를 "…"으로 축소하는 레이블.
#   로그 헤더의 기록 파일 경로 표시에 사용.
#   크기 정책: Preferred — sizeHint()로 전체 경로 너비를 레이아웃에 요청.
#   공간이 충분하면 전체 경로를 표시하고, 창이 좁아지면 ElideMiddle(경로 중간 줄임)로 압축.
#   minimumSizeHint()를 0으로 반환해 최소한까지 압축 허용.
# ══════════════════════════════════════════════════════════
class ElidedLabel(QLabel):
    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self._full_text = text
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        self.setMinimumWidth(0)

    def sizeHint(self):
        # 레이아웃에 전체 경로 너비를 요청 → 공간이 충분하면 전체 경로 표시
        hint = super().sizeHint()
        hint.setWidth(self.fontMetrics().horizontalAdvance(self._full_text))
        return hint

    def minimumSizeHint(self):
        # 최소 너비 0 → 창이 좁아지면 줄임표로 압축 허용
        hint = super().minimumSizeHint()
        hint.setWidth(0)
        return hint

    def set_full_text(self, text):
        self._full_text = text
        self.updateGeometry()   # 레이아웃에 sizeHint 재계산 요청
        self._update_elided()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_elided()

    def _update_elided(self):
        fm = self.fontMetrics()
        elided = fm.elidedText(self._full_text, Qt.ElideMiddle, self.width())
        super().setText(elided)
        self.setToolTip(self._full_text)

# ══════════════════════════════════════════════════════════
# 【3-16】 _licenses_dir() → str
#   배포 ZIP에 실행파일과 나란히 둔 licenses 폴더의 경로를 반환.
#   onefile 빌드는 --add-data 로 넣은 자원이 임시 폴더(_MEIPASS)에
#   풀리므로, 실행파일 옆(side)의 licenses 폴더를 우선 사용하고,
#   없으면 _MEIPASS(또는 스크립트 폴더) 기준 경로로 대체한다.
# ══════════════════════════════════════════════════════════
def _licenses_dir() -> str:
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))
    side = os.path.join(exe_dir, "licenses")
    if os.path.isdir(side):
        return side
    base = getattr(sys, "_MEIPASS", exe_dir)
    return os.path.join(base, "licenses")

# ══════════════════════════════════════════════════════════
# 【3-17】 _open_licenses_dir(parent=None)
#   [라이선스 전문 폴더 열기] 버튼 클릭 시 호출.
#   OS별 파일탐색기로 licenses 폴더를 연다.
#   Windows: os.startfile / macOS: open 명령 / Linux: xdg-open
# ══════════════════════════════════════════════════════════
def _open_licenses_dir(parent=None):
    d = _licenses_dir()
    if not os.path.isdir(d):
        return
    if sys.platform.startswith("win"):
        os.startfile(d)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", d])
    else:
        subprocess.Popen(["xdg-open", d])

# ══════════════════════════════════════════════════════════
# 【4】 RouteMapPanel 클래스 (노선 지도 패널 위젯)
#   버스 노선의 정류소를 그래픽으로 표시하고
#   운행 중인 버스 아이콘을 실시간 갱신하는 패널 위젯.
#
#   화면 구조:
#   ┌─────────────────────────────────────────────┐
#   │ [간선] 101 | 총 35개 정류소 | 약 35.2km ...  │ ← 노선 정보 텍스트
#   │  ①─②─③─ … ─⑮                             │
#   │  ↓                                          │ ← QGraphicsScene
#   │  ㉚─㉙─ … ─⑯  (뱀 형태 지그재그 배치)       │   (정류소 + 버스 아이콘)
#   ├─────────────────────────────────────────────┤
#   │ 종점도착예정: [1234] 5분20초 | [5678] 8분45초 │ ← _table (도착 예정 표)
#   └─────────────────────────────────────────────┘
#
#   @param route_rnm: 이 패널이 담당하는 노선 번호/이름 (예: "101")
# ══════════════════════════════════════════════════════════
class RouteMapPanel(QWidget):
# ──────────────────────────────────────────────────────────
# 【4-1】 __init__(route_rnm, parent=None)
#   패널 위젯을 초기화. 인스턴스 변수 선언 및 화면 요소 생성.
#
#   주요 인스턴스 변수:
#   _sect_speeds   : {seq번호: 속도(km/h)} 구간별 속도 딕셔너리
#   _stations      : [{station, seq, name, transYn, ...}] 정류소 정보 리스트
#   _current_rtype : 현재 노선 종류 코드 (기본값 "3" = 간선)
#   _current_length: 노선 총 길이 문자열
#   _last_buses    : 최근 그린 버스 목록 (창 크기 변경 시 재그리기용)
#   _bus_seconds   : {b_idx: 남은초} 버스별 종점까지 남은 시간
#   _table_bus_info: [{col, b_idx}] 도착 예정 표의 버스 정보
#
#   주요 화면 요소:
#   _scene     : 정류소·버스를 그리는 QGraphicsScene (무대)
#   _view      : _scene을 화면에 표시하는 QGraphicsView (스크롤 가능)
#   _table     : 버스 도착 예정 시간 표 (2행 × N열, 최대높이 65px)
#   _tip_label : 마우스 오버 시 표시되는 툴팁 QLabel (ToolTip 플래그)
#   _tick_timer: 1000ms 간격으로 _tick_countdown을 호출하는 타이머
#   _resize_timer: 창 크기 변경 후 300ms 뒤 지도 재그리기 (단발성 타이머)
# ──────────────────────────────────────────────────────────
    def __init__(self, route_rnm, parent=None):
        super().__init__(parent)
        self.route_rnm = route_rnm
        self._sect_speeds = {} 
        self._stations = []
        self._current_rtype = "3"
        self._current_length = ""
        self._last_buses = None
        self._bus_seconds = {}
        self._table_bus_info = []
        self._min_map_width = 600

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._scene = QGraphicsScene(self)
        self._view = QGraphicsView(self._scene, self)
        self._view.setRenderHint(QPainter.Antialiasing)
        self._view.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._view.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self._view.setMouseTracking(True)
        self._view.viewport().setMouseTracking(True)
        self._update_bg_color()
        layout.addWidget(self._view, stretch=1)

        self._table = QTableWidget(2, 1)
        self._table.setMaximumHeight(65)
        self._table.setMinimumHeight(65)
        self._table.verticalHeader().setVisible(False)
        self._table.horizontalHeader().setVisible(False)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setSelectionMode(QAbstractItemView.NoSelection)
        self._table.setShowGrid(True)
        layout.addWidget(self._table)

        t = self._scene.addText("노선 데이터 로딩 중...", QFont(FONT_FAMILY, 12))
        t.setDefaultTextColor(QColor("#777777"))
        t.setPos(200, 100)

        self._tip_label = QLabel(self)
        self._tip_label.setWindowFlags(Qt.ToolTip)
        self._tip_label.setFont(QFont(FONT_FAMILY, 9))
        self._update_tip_style()
        self._tip_label.hide()
        self._view.viewport().installEventFilter(self)
        self._view.setMinimumWidth(self._min_map_width)

        self._tick_timer = QTimer(self)
        self._tick_timer.setInterval(1000)
        self._tick_timer.timeout.connect(self._tick_countdown)
        self._tick_timer.start()

        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.setInterval(300)
        self._resize_timer.timeout.connect(self._on_resize_done)
        
# ──────────────────────────────────────────────────────────
# 【4-2】 set_sect_speeds(speeds)
#   구간별 속도 데이터를 저장. _speed_color()에서 선 색상 결정 시 참조.
#   @param speeds: {seq번호(int): 속도(float km/h)} dict, None이면 빈 dict로 초기화
# ──────────────────────────────────────────────────────────
    def set_sect_speeds(self, speeds):
        self._sect_speeds = speeds if speeds else {}

# ──────────────────────────────────────────────────────────
# 【4-3】 eventFilter(obj, event) → bool
#   지도 뷰포트에서 발생하는 마우스 이벤트를 가로채어 툴팁을 제어.
#   __init__에서 installEventFilter(self)로 등록되어 자동 호출됨.
#
#   MouseMove: 마우스가 올라간 아이템의 data(0) 값(툴팁 문자열)을 읽어
#              _tip_label을 마우스 근처에 표시. 아이템 없으면 숨김.
#   Leave    : 마우스가 지도 영역을 벗어나면 툴팁 숨김.
#   @return  : False (이벤트를 소비하지 않고 원래 처리도 계속 허용)
# ──────────────────────────────────────────────────────────
    def eventFilter(self, obj, event):
        if obj is self._view.viewport():
            if event.type() == event.Type.MouseMove:
                pos = self._view.mapToScene(event.position().toPoint())
                item = self._scene.itemAt(pos, self._view.transform())
                tip = item.data(0) if item else None
                if tip:
                    self._tip_label.setText(tip)
                    self._tip_label.adjustSize()
                    gp = self._view.viewport().mapToGlobal(event.position().toPoint())
                    self._tip_label.move(gp.x() + 12, gp.y() + 12)
                    if not self._tip_label.isVisible():
                        self._tip_label.show()
                else:
                    self._tip_label.hide()
                return False
            elif event.type() == event.Type.Leave:
                self._tip_label.hide()
                return False
        return super().eventFilter(obj, event)

# ──────────────────────────────────────────────────────────
# 【4-4】 resizeEvent(event)
#   창 크기 변경 시 Qt가 자동 호출. 성능을 위해 즉시 재그리기 대신
#   _resize_timer를 300ms로 (재)시작하여 변경이 끝난 후 한 번만 재그림.
# ──────────────────────────────────────────────────────────
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._stations:
            self._resize_timer.start()

# ──────────────────────────────────────────────────────────
# 【4-5】 _on_resize_done()
#   _resize_timer 만료(300ms 경과) 후 호출. 정류소 데이터가 있을 때만
#   _last_buses와 함께 _draw()를 다시 호출해 지도 재그리기.
# ──────────────────────────────────────────────────────────
    def _on_resize_done(self):
        if self._stations:
            self._draw(self._last_buses)

# ──────────────────────────────────────────────────────────
# 【4-6】 _calc_layout(n) → (spr, cw, ch, font_size)
#   현재 뷰포트 크기와 정류소 수를 기반으로 레이아웃 값 계산.
#
#   계산 흐름:
#   1) usable_w/h = 뷰포트 크기 - 패딩
#   2) spr(한 행 정류소 수) = usable_w / min_cw, 범위 5~15
#   3) cw(셀 너비) = usable_w / spr
#   4) rows = ceil(n / spr),  ch = usable_h / rows, 최소 CELL_H 보장
#   5) font_size: cw에 비례하여 6pt~10pt 선형 보간
#
#   @param n   : 정류소 총 수
#   @return    : (한행정류소수, 셀너비, 셀높이, 폰트크기) 튜플
# ──────────────────────────────────────────────────────────
    def _calc_layout(self, n):
        view_w = max(self._view.viewport().width(), self._min_map_width)
        view_h = self._view.viewport().height()
        usable_w = view_w - PAD_X * 2 - 20
        usable_h = view_h - PAD_Y - 60
        min_cw = CELL_W
        min_ch = CELL_H
        min_fs = 6
        max_fs = 10

        if n <= 0 or usable_w <= 0 or usable_h <= 0:
            return STOPS_PER_ROW, CELL_W, CELL_H, min_fs

        spr = max(5, int(usable_w // min_cw))
        spr = min(spr, 15)
        if spr > n:
            spr = n

        cw = usable_w / spr

        rows = (n + spr - 1) // spr
        if rows > 0 and usable_h > 0:
            ch = usable_h / rows
            ch = min(ch, cw)
            ch = max(min_ch, ch)
        else:
            ch = min_ch

        if cw <= min_cw:
            font_size = min_fs
        elif cw >= FONT_CAP:
            font_size = max_fs
        else:
            ratio = (cw - min_cw) / (FONT_CAP - min_cw)
            font_size = int(min_fs + ratio * (max_fs - min_fs))

        return spr, cw, ch, font_size

# ──────────────────────────────────────────────────────────
# 【4-7】 _update_bg_color()
#   지도 뷰(_view)의 배경색을 현재 앱 테마의 창 배경색으로 갱신.
#   테마 변경 시 refresh_theme()에서 호출.
# ──────────────────────────────────────────────────────────
    def _update_bg_color(self):
        self._view.setBackgroundBrush(QBrush(get_app_bg_color()))
        
# ──────────────────────────────────────────────────────────
# 【4-8】 _update_tip_style()
#   툴팁 라벨(_tip_label)의 CSS 스타일을 현재 테마에 맞게 갱신.
#   다크: 어두운 배경(#1E1E1E) + 밝은 글자(#DCDCDC)
#   라이트: 밝은 배경(#FDFDFD) + 검정 글자(#000000)
# ──────────────────────────────────────────────────────────
    def _update_tip_style(self):
        if get_app_bg_color().lightness() < 128:
            self._tip_label.setStyleSheet(
                "QLabel { background-color: #1E1E1E; color: #DCDCDC; "
                "border: 1px solid #555; padding: 4px; }")
        else:
            self._tip_label.setStyleSheet(
                "QLabel { background-color: #FDFDFD; color: #000000; "
                "border: 1px solid #AAAAAA; padding: 4px; }")

# ──────────────────────────────────────────────────────────
# 【4-9】 refresh_theme()
#   테마 변경 시 이 패널 전체를 새 테마로 갱신.
#   _apply_theme()에서 모든 RouteMapPanel에 대해 호출됨.
#   처리: 배경색 → 툴팁 스타일 → 지도 재그리기 → 표 색상 재적용
# ──────────────────────────────────────────────────────────
    def refresh_theme(self):
        self._update_bg_color()
        self._update_tip_style()
        if self._stations:
            self._draw(self._last_buses)
        self._rebuild_table_colors()

# ──────────────────────────────────────────────────────────
# 【4-10】 _rebuild_table_colors()
#   도착 예정 표(_table)의 모든 셀 배경색·글자색을 현재 테마로 재적용.
#   0번 열(라벨) 또는 0번 행(헤더)은 header_bg, 나머지는 base_bg 적용.
# ──────────────────────────────────────────────────────────
    def _rebuild_table_colors(self):
        header_bg = get_header_bg_color()
        base_bg = get_base_color()
        txt_color = get_text_color()
        for row in range(self._table.rowCount()):
            for col in range(self._table.columnCount()):
                item = self._table.item(row, col)
                if item:
                    item.setForeground(QBrush(txt_color))
                    if col == 0 or row == 0:
                        item.setBackground(QBrush(header_bg))
                    else:
                        item.setBackground(QBrush(base_bg))

# ──────────────────────────────────────────────────────────
# 【4-11】 load_route(stations, rtype="3", length="")
#   새 노선의 정류소 목록을 받아 저장하고 지도를 처음 그림 (buses=None).
#   _slot_route_loaded()에서 패널 생성 후 바로 호출됨.
#
#   @param stations: [{station, seq, name, arsId, transYn, fullSectDist}, ...] 리스트
#   @param rtype   : 노선 종류 코드
#   @param length  : 노선 총 길이 문자열
# ──────────────────────────────────────────────────────────
    def load_route(self, stations, rtype="3", length=""):
        self._stations = stations
        self._current_rtype = rtype
        self._current_length = length
        self._draw(None)

# ──────────────────────────────────────────────────────────
# 【4-12】 update_buses(buses)
#   현재 운행 중인 버스 목록을 받아 지도를 갱신.
#   _slot_update_map 시그널 슬롯에서 호출됨.
#   정류소 데이터가 없으면 즉시 반환.
#
#   @param buses: [{vehId, plainNo, busType, lastStnId, sectDist, ...}, ...] 리스트
# ──────────────────────────────────────────────────────────
    def update_buses(self, buses):
        if not self._stations:
            return
        self._last_buses = buses
        self._draw(buses)
        
# ──────────────────────────────────────────────────────────
# 【4-13】 _speed_color(seq_idx) → QColor or None
#   구간 순번에 해당하는 속도를 조회하여 색상 반환.
#     0~9 km/h   → 빨강  QColor(255, 0, 0)   (정체/서행)
#    10~19 km/h  → 노랑  QColor(255, 255, 0)
#    20 km/h 이상 → 초록 QColor(0, 255, 0)   (원활)
#    데이터 없음   → None (기본 노선 색 사용)
# ──────────────────────────────────────────────────────────
    def _speed_color(self, seq_idx):
        spd = self._sect_speeds.get(seq_idx, -1)
        if spd < 0:
            return None
        if spd < 10:
            return QColor(255, 0, 0)
        elif spd < 20:
            return QColor(255, 255, 0)
        else:
            return QColor(0, 255, 0)

# ──────────────────────────────────────────────────────────
# 【4-14】 _draw(buses)  ★ 지도 그리기 핵심 함수 ★
#   QGraphicsScene을 clear하고 모든 요소를 처음부터 재그림.
#
#   그리는 순서:
#   ① scene 초기화 & 변수 준비
#   ② 상단 노선 정보 텍스트 (노선 종류·정류소 수·길이·운행 버스 수)
#   ③ 뱀 형태(지그재그) 좌표 계산 (coords 배열)
#      → 짝수 행: 왼→오른쪽,  홀수 행: 오른→왼쪽
#   ④ 정류소 간 연결선 (구간 속도 색 또는 노선 색, 회차 이후는 어두운 색)
#   ⑤ 정류소 원 + 번호/이름 텍스트
#      출발·회차·종점: 큰 원(CIRCLE_R_SPECIAL) + 굵은 레이블
#      일반 정류소  : 작은 원(CIRCLE_R) + 순번 숫자
#   ⑥ 버스 아이콘 (buses가 있을 때)
#      - lastStnId + sectDist 비율로 정확한 위치 보간
#      - 같은 Y대에 겹치는 레이블은 X 오프셋으로 분리
#      - 저상 버스 "저상" 텍스트, 막차 "[막차]" 텍스트 표시
#   ⑦ 행 연결 화살표(↓) 그리기
#   ⑧ sceneRect 설정 + _build_table() 호출
#
#   @param buses: 버스 정보 리스트 (None이면 버스 아이콘 없이 그림)
# ──────────────────────────────────────────────────────────
    def _draw(self, buses):
        self._scene.clear()
        self._bus_seconds.clear()
        stations = self._stations
        n = len(stations)

        theme_text_color = get_text_color()
        is_dark = get_app_bg_color().lightness() < 128

        line_color_hex = ROUTE_TYPE_COLOR.get(self._current_rtype, DEFAULT_LINE_COLOR)
        line_dark_hex = darken_color(line_color_hex, 0.58)
        line_color = QColor(line_color_hex)
        line_dark = QColor(line_dark_hex)

        turn_idx = next((i for i, st in enumerate(stations) if st.get("transYn") == "Y"), None)

        stops_per_row, cell_w, cell_h, dyn_font = self._calc_layout(n)

        typ_str = ROUTE_TYPE_LABEL.get(self._current_rtype, "기타")
        dist_str = self._current_length if self._current_length else "?"
        bus_cnt = len(buses) if buses else 0
        info_str = (f"[{typ_str}] {self.route_rnm} | "
                    f"총 {n}개 정류소 | 노선 총 길이 약 {dist_str}km | 현재 운행중인 버스 {bus_cnt}대")
        total_w = PAD_X * 2 + (stops_per_row - 1) * cell_w + 20

        info_item = self._scene.addText(info_str, QFont(FONT_FAMILY, 11, QFont.Bold))
        info_item.setDefaultTextColor(theme_text_color)
        info_item.setPos((total_w - info_item.boundingRect().width()) / 2, 10)

        if n == 0:
            t = self._scene.addText("정류소 정보가 없습니다.", QFont(FONT_FAMILY, 11))
            t.setDefaultTextColor(theme_text_color)
            t.setPos(total_w / 2 - 80, 110)
            self._scene.setSceneRect(0, 0, total_w, 200)
            self._build_table([])
            return

        rows = (n + stops_per_row - 1) // stops_per_row
        coords = []
        for i in range(n):
            ri = i // stops_per_row
            ci = i % stops_per_row
            x = PAD_X + ci * cell_w if ri % 2 == 0 else PAD_X + (stops_per_row - 1 - ci) * cell_w
            coords.append((x, PAD_Y + ri * cell_h))

        for i in range(n - 1):
            x1, y1 = coords[i]
            x2, y2 = coords[i + 1]
            spd_color = self._speed_color(i + 2)
            if spd_color is not None:
                sc = spd_color
            else:
                sc = line_dark if (turn_idx is not None and i >= turn_idx) else line_color
            pen = QPen(sc, 3)
            pen.setCapStyle(Qt.RoundCap)
            if (i // stops_per_row) == ((i + 1) // stops_per_row):
                self._scene.addLine(x1, y1, x2, y2, pen)
            else:
                self._scene.addLine(x1, y1, x1, y2, pen)

        for i, (st, (x, y)) in enumerate(zip(stations, coords)):
            is_first = (i == 0)
            is_last = (i == n - 1)
            is_turn = (st.get("transYn") == "Y")
            after_turn = (turn_idx is not None and i > turn_idx)
            nd = truncate_name(st.get("name", "?"))

            if is_turn:
                cr, cf, co = CIRCLE_R_SPECIAL, QColor("white"), line_color
                li, tc, fs, fb = "회차", line_color, dyn_font + 1, True
            elif is_first:
                cr, cf, co = CIRCLE_R_SPECIAL, QColor("white"), line_color
                li, tc, fs, fb = "출발", line_color, dyn_font + 1, True
            elif is_last:
                cr, cf, co = CIRCLE_R_SPECIAL, QColor("white"), line_color
                li, tc, fs, fb = "종점", line_color, dyn_font + 1, True
            else:
                sc2 = line_dark if after_turn else line_color
                if is_dark:
                    cr, cf, co = CIRCLE_R, QColor(60, 60, 60), sc2
                    li, tc, fs, fb = str(i + 1), QColor(200, 200, 200), dyn_font, False
                else:
                    cr, cf, co = CIRCLE_R, QColor("white"), sc2
                    li, tc, fs, fb = str(i + 1), theme_text_color, dyn_font, False

            self._scene.addEllipse(x - cr, y - cr, cr * 2, cr * 2, QPen(co, 2), QBrush(cf))
            fi = QFont(FONT_FAMILY, fs, QFont.Bold if fb else QFont.Normal)
            it = self._scene.addText(li, fi)
            it.setDefaultTextColor(tc)
            br = it.boundingRect()
            it.setPos(x - br.width() / 2, y - br.height() / 2)

            if is_first or is_last or is_turn:
                nc = tc
            else:
                nc = QColor(140, 140, 140) if is_dark else QColor(130, 130, 130)
            nt = self._scene.addText(nd, QFont(FONT_FAMILY, fs, QFont.Bold if fb else QFont.Normal))
            nt.setDefaultTextColor(nc)
            nt.setRotation(-45)
            nt.setPos(x + TEXT_X_OFFSET, y - cr - 2)

        table_buses = []
        if buses:
            bus_pm = make_bus_pixmap()
            bus_draw_list = []
            for b_idx, bus in enumerate(buses):
                lsi = bus.get("lastStnId", "")
                rpn = bus.get("plainNo", "").replace("서울", "").strip()
                ilb = (bus.get("islastyn") == "1")
                try:
                    sm = float(bus.get("sectDist", "0")) * 1000.0
                except:
                    sm = 0.0
                fi2 = next((i for i, st in enumerate(stations) if st.get("station", "") == lsi), None)
                if fi2 is None:
                    continue
                ti = fi2 + 1 if fi2 + 1 < n else fi2
                try:
                    sdm = float(stations[ti].get("fullSectDist", "0"))
                except:
                    sdm = 0.0
                ratio = min(sm / sdm, 1.0) if sdm > 0 and ti != fi2 else 0.0
                x1, y1 = coords[fi2]
                x2, y2 = coords[ti]
                if (fi2 // stops_per_row) == (ti // stops_per_row):
                    bx, by = x1 + (x2 - x1) * ratio, y1
                else:
                    bx, by = x1, y1 + (y2 - y1) * ratio
                try:
                    secs = int(bus.get("lastStTm", 0))
                except:
                    secs = 0
                bus_draw_list.append({
                    "b_idx": b_idx, "bus": bus, "bx": bx, "by": by,
                    "rpn": rpn, "ilb": ilb, "secs": secs,
                    "fi2": fi2, "ratio": ratio
                })

            Y_THRESH = 12 + dyn_font
            SEP = 15 + (dyn_font - 6) * 3
            groups = defaultdict(list)
            for bd in bus_draw_list:
                gk = round(bd["by"] / Y_THRESH) * Y_THRESH
                groups[gk].append(bd)
            label_offsets = {}
            for gk in sorted(groups.keys()):
                grp = sorted(groups[gk], key=lambda b: b["bx"])
                prev_label_x = None
                for bd in grp:
                    offset_x = 0
                    if prev_label_x is not None:
                        dist = bd["bx"] - prev_label_x
                        if dist < SEP:
                            offset_x = SEP - dist
                    label_offsets[id(bd)] = offset_x
                    prev_label_x = bd["bx"] + offset_x

            for bd in bus_draw_list:
                bus = bd["bus"]
                bx = bd["bx"]
                by = bd["by"]
                rpn = bd["rpn"]
                ilb = bd["ilb"]
                secs = bd["secs"]
                b_idx = bd["b_idx"]
                fi2 = bd["fi2"]
                ratio = bd["ratio"]
                offset_x = label_offsets.get(id(bd), 0)

                cc = bus.get("congestion", "0")
                _CONG = {"3": "여유", "4": "보통", "5": "혼잡"}
                cs = str(cc)
                if self._current_rtype == "6":
                    if cs in ("99", "0"):
                        cl = ""
                    else:
                        try:
                            cl = f"잔여{int(cs)}석"
                        except:
                            cl = ""
                else:
                    cl = _CONG.get(cs, "")
                is_low = bus.get("busType") == "1"
                parts = fmt_bus_no(rpn)
                if cl:
                    parts += f"  {cl}"
                if is_low:
                    parts += "  ( 저상 )"
                if ilb:
                    parts += "  [ 막차 ]"
                tip = f"{parts}\n종점도착까지 남은 시간 : {format_remain_time(secs)}"

                pi = self._scene.addPixmap(bus_pm)
                pi.setPos(bx - 12, by - 12)
                pi.setZValue(5)
                pi.setData(0, tip)

                hit = self._scene.addRect(bx - 15, by - 15, 30, 30,
                                        QPen(Qt.NoPen), QBrush(QColor(0, 0, 0, 0)))
                hit.setData(0, tip)
                hit.setZValue(8)

                bt = self._scene.addText(rpn, QFont(FONT_FAMILY, dyn_font + 3, QFont.Bold))
                bt.setDefaultTextColor(theme_text_color)
                bt.setRotation(-45)
                bt.setPos(bx - 2 + offset_x, by - 18)
                bt.setZValue(6)
                bt.setData(0, tip)

                if ilb:
                    lt = self._scene.addText("[막차]", QFont(FONT_FAMILY, dyn_font + 2, QFont.Bold))
                    lt.setDefaultTextColor(theme_text_color)
                    lt.setPos(bx - 18 + offset_x, by + 7)
                    lt.setZValue(6)

                if is_low:
                    lw = self._scene.addText("저상", QFont(FONT_FAMILY, 7, QFont.Bold))
                    lw.setDefaultTextColor(QColor(0, 0, 0))
                    lw.setPos(bx - 12, by - 14)
                    lw.setZValue(7)

                table_buses.append({
                    "raw_plain_no": rpn, "is_last_bus": ilb,
                    "seconds": secs, "from_idx": fi2, "ratio": ratio, "b_idx": b_idx
                })
                self._bus_seconds[b_idx] = secs

        for ri in range(rows - 1):
            ry = PAD_Y + ri * cell_h
            ny = PAD_Y + (ri + 1) * cell_h
            ax = (PAD_X + (stops_per_row - 1) * cell_w - 3) if ri % 2 == 0 else (PAD_X - 22)
            lir = min((ri + 1) * stops_per_row - 1, n - 1)
            spd_color = self._speed_color(lir + 2)
            if spd_color is not None:
                ac = spd_color
            else:
                ac = line_dark if (turn_idx is not None and lir >= turn_idx) else line_color
            at = self._scene.addText("↓", QFont(FONT_FAMILY, 11, QFont.Bold))
            at.setDefaultTextColor(ac)
            at.setPos(ax, (ry + ny) / 2 - 10)
            at.setZValue(-1)

        self._scene.setSceneRect(0, 0, total_w, PAD_Y + (rows - 1) * cell_h + 60)
        self._view.resetTransform()
        self._view.ensureVisible(0, 0, 10, 10)
        self._build_table(table_buses)

# ──────────────────────────────────────────────────────────
# 【4-15】 _build_table(table_buses)
#   지도 하단의 "종점 도착 예정 버스" 표(2행 × N열)를 구성.
#
#   표 구조:
#   ┌──────────────────────┬─────────┬─────────┬──────────┐
#   │ 종점 도착 예정 버스번호│  1234   │  5678   │ 9012(막차)│
#   │ 종점 도착까지 남은 시간│ 5분 20초│ 8분 45초│  2분 10초 │
#   └──────────────────────┴─────────┴─────────┴──────────┘
#
#   정렬: 종점에 가장 가까운 버스(from_idx 높은 순)가 왼쪽
#   막차 버스는 항상 맨 오른쪽
#   표 너비에 따라 보여줄 버스 열 수 동적 결정
#
#   @param table_buses: [{raw_plain_no, is_last_bus, seconds, from_idx, ratio, b_idx}] 리스트
# ──────────────────────────────────────────────────────────
    def _build_table(self, table_buses):
        self._table_bus_info = []
        header_bg = get_header_bg_color()
        base_bg = get_base_color()
        txt_color = get_text_color()

        if not table_buses:
            self._table.clear()
            self._table.setColumnCount(1)
            self._table.setRowCount(1)
            it = QTableWidgetItem("")
            it.setBackground(QBrush(base_bg))
            it.setForeground(QBrush(txt_color))
            self._table.setItem(0, 0, it)
            return

        sb = sorted(table_buses, key=lambda x: (x['from_idx'], x['ratio']), reverse=True)
        lb = next((b for b in sb if b['is_last_bus']), None)
        normal = [b for b in sb if not b['is_last_bus']]

        label_col_w = 160
        bus_col_w = 100
        table_w = self._table.width()
        if table_w < 200:
            table_w = 700
        available_w = table_w - label_col_w
        if available_w < bus_col_w:
            available_w = bus_col_w
        max_bus_cols = max(1, available_w // bus_col_w)

        if lb:
            cols = normal[:max_bus_cols - 1] + [lb]
        else:
            cols = normal[:max_bus_cols]

        if not cols:
            self._table.clear()
            self._table.setColumnCount(1)
            self._table.setRowCount(1)
            return

        cc = 1 + len(cols)
        self._table.setColumnCount(cc)
        self._table.setRowCount(2)
        hf = QFont(FONT_FAMILY, 8, QFont.Bold)
        cf = QFont(FONT_FAMILY, 8)

        def _mi(txt, font, is_header=False):
            it = QTableWidgetItem(txt)
            it.setFont(font)
            it.setTextAlignment(Qt.AlignCenter)
            it.setBackground(QBrush(header_bg if is_header else base_bg))
            it.setForeground(QBrush(txt_color))
            return it

        self._table.setItem(0, 0, _mi("종점 도착 예정 버스 번호", hf, True))
        self._table.setItem(1, 0, _mi("종점 도착까지 남은 시간", hf, True))

        for ci, bus in enumerate(cols):
            title = f"{bus['raw_plain_no']} (막차)" if bus['is_last_bus'] else bus['raw_plain_no']
            self._table.setItem(0, ci + 1, _mi(title, hf, True))
            self._table.setItem(1, ci + 1, _mi(format_remain_time(bus['seconds']), cf, False))
            self._table_bus_info.append({"col": ci + 1, "b_idx": bus['b_idx']})

        self._table.horizontalHeader().setDefaultSectionSize(bus_col_w)
        self._table.setColumnWidth(0, label_col_w)

# ──────────────────────────────────────────────────────────
# 【4-16】 _tick_countdown()
#   _tick_timer에 의해 1초마다 자동 호출.
#   _bus_seconds의 각 버스 남은 시간을 1씩 감소시키고
#   표(_table) "남은 시간" 셀을 즉시 갱신.
#   → API 갱신 주기 사이에도 카운트다운이 부드럽게 작동.
# ──────────────────────────────────────────────────────────
    def _tick_countdown(self):
        txt_color = get_text_color()
        for k in list(self._bus_seconds.keys()):
            if self._bus_seconds[k] > 0:
                self._bus_seconds[k] -= 1
        for info in self._table_bus_info:
            sec = self._bus_seconds.get(info['b_idx'], 0)
            item = self._table.item(1, info['col'])
            if item:
                item.setText(format_remain_time(sec))
                item.setForeground(QBrush(txt_color))

# ──────────────────────────────────────────────────────────
# 【4-17】 pause_tick()  : _tick_timer 정지 (기록 중지 시 호출)
# 【4-18】 resume_tick() : _tick_timer 재시작 (기록 시작 시 호출)
# ──────────────────────────────────────────────────────────
    def pause_tick(self):
        self._tick_timer.stop()

    def resume_tick(self):
        self._tick_timer.start()


# ══════════════════════════════════════════════════════════
# 【5】 RecordTable 클래스
#   버스 운행 기록을 표 형태로 보여주는 위젯.
#   메인 창 우측에 "운행 출발 시각 기록"과 "운행 종료 시각 기록"
#   두 개가 세로로 배치됨.
#   컬럼: [시각(145px)] [노선(80px)] [차량번호(95px)] [상태(나머지)]
# ══════════════════════════════════════════════════════════
class RecordTable(QWidget):
# ──────────────────────────────────────────────────────────
# 【5-1】 __init__(title, parent=None)
#   표 위에 제목 라벨을 배치하고, 컬럼 너비·정렬·스타일을 설정.
#   - 편집 불가(NoEditTriggers): 실수 수정 방지
#   - 행 선택(SelectRows): 클릭 시 한 행 전체 선택
#   - 교대 배경색(AlternatingRowColors): 홀짝 행 배경이 달라 읽기 편함
#   - 세로 헤더 숨김: 줄 번호 표시 안 함
# ──────────────────────────────────────────────────────────
    def __init__(self, title, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(3, 3, 3, 3)
        lbl = QLabel(f"  {title}  ")
        lbl.setFont(QFont(FONT_FAMILY, 10, QFont.Bold))
        layout.addWidget(lbl)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["시각", "노선", "차량번호", "상태"])
        self.table.verticalHeader().setDefaultSectionSize(15)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.Fixed)
        hdr.setSectionResizeMode(2, QHeaderView.Fixed)
        hdr.setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 145)
        self.table.setColumnWidth(1, 80)
        self.table.setColumnWidth(2, 95)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

# ──────────────────────────────────────────────────────────
# 【5-2】 add_row(values)
#   표 맨 아래에 새 행을 삽입하고 scrollToBottom()으로 자동 스크롤.
#   시각·노선·차량번호(col 0~2)는 AlignCenter, 상태(col 3)는 기본 정렬.
#   @param values: [시각, 노선, 차량번호, 상태] 순서 리스트
# ──────────────────────────────────────────────────────────
    def add_row(self, values):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col, val in enumerate(values):
            it = QTableWidgetItem(str(val))
            if col < 3:
                it.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, col, it)
        self.table.scrollToBottom()

# ══════════════════════════════════════════════════════════
# 【6】 SeoulBusRecorder 클래스  ★ 프로그램의 핵심 ★
#   메인 창(QMainWindow)이자 프로그램 전체를 총괄하는 클래스.
#   메뉴, 노선 지도 패널, 운행 기록 표, 로그 창을 모두 포함하며
#   API 호출·운행 판정·엑셀 저장을 관리.
#
#   시그널(Signal) 목록 - 스레드 간 안전한 UI 업데이트용:
#   sig_log          (str)         → 로그 메시지 전달
#   sig_record       (int, tuple)  → 운행 기록 표 추가
#   sig_update_map   (str, list)   → 지도 버스 위치 갱신
#   sig_clear_map    (str)         → 지도 버스 제거
#   sig_search_done  (list)        → 노선 검색 결과 전달
#   sig_route_loaded (dict)        → 노선 정보 로드 완료
#   sig_key_verified (str,str,str,str) → API 키 검증 완료
#
#   ※ 왜 시그널을 사용하나?
#   API 호출은 백그라운드 스레드에서 실행되지만,
#   Qt UI 업데이트는 반드시 메인(GUI) 스레드에서만 해야 한다.
#   시그널-슬롯이 이 스레드 간 안전한 통신을 자동으로 처리해준다.
# ══════════════════════════════════════════════════════════
class SeoulBusRecorder(QMainWindow):
    sig_log = Signal(str)
    sig_record = Signal(int, tuple)
    sig_update_map = Signal(str, list)
    sig_clear_map = Signal(str)
    sig_search_done = Signal(list)
    sig_route_loaded = Signal(dict)
    sig_key_verified = Signal(str, str, str, str)

    _SECRET = "l'existence précède l'essence"

# ──────────────────────────────────────────────────────────
# 【6-1】 __init__()
#   메인 창 초기화. 처리 순서:
#   1) 창 제목·크기 설정, 아이콘 적용
#   2) current_dir 결정 (실행파일/스크립트/macOS .app 각각 처리)
#   3) 설정 파일 경로 결정 (Seoul_Bus_Config.ini)
#   4) 인스턴스 변수 초기화:
#      api_key_main/back     : API 인증키 (빈 문자열로 시작)
#      is_monitoring         : 모니터링 중 여부 (False)
#      routes                : 모니터링할 노선 데이터 리스트
#      refresh_interval      : API 갱신 주기 초 (기본 25)
#      recorded_data         : [(시각,구분,정류소,노선,차번), ...] 기록 리스트
#      _saved_record_count   : 마지막 저장 시점 기록 수 (중복 저장 방지)
#      last_arrival_logs     : {(방향,rid,vn): timestamp} 최근 기록 시각
#                              → 같은 이벤트를 40분(2400초) 내 중복 기록 방지
#      departed_vehicles     : {(rid,vn): timestamp} 출발 감지된 버스 목록
#                              → 6시간(21600초) 후 자동 삭제
#      pos_suspend_until     : {rid: datetime} 운행 종료 후 API 호출 중지 시각
#      _save_lock            : 엑셀 저장 동시 접근 방지 스레드 잠금
#      _refresh_lock         : 갱신 중복 실행 방지 스레드 잠금
#      api_stats_today/yesterday: 오늘/어제 API 호출 횟수 통계
#      route_map_panels      : {노선명: RouteMapPanel} 딕셔너리
#   5) 시그널-슬롯 연결
#   6) _load_config() → _setup_ui() → 초기 로그 출력
# ──────────────────────────────────────────────────────────
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"서울시내버스 노선 운행기록 수집 프로그램 v{APP_VERSION}")
        self.setMinimumSize(800, 500)
        self.resize(1350, 1000)
        try:
            pm = load_pixmap_from_b64(ICON_B64)
            if not pm.isNull():
                self.setWindowIcon(QIcon(pm))
        except:
            pass

        if getattr(sys, 'frozen', False):
            _exe = os.path.abspath(sys.executable)
            if sys.platform == "darwin" and ".app/Contents/MacOS" in _exe:
                self.current_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(_exe))))
            else:
                self.current_dir = os.path.dirname(_exe)
        else:
            self.current_dir = os.path.dirname(os.path.abspath(__file__))

        self._CFG_FILE = os.path.join(self.current_dir, "Seoul_Bus_Config.ini")

        self.api_key_main = ""
        self.api_key_back = ""

        self.is_monitoring = False
        self.routes = []
        self.routes_ready = True
        self.refresh_interval = 25
        self.recorded_data = []
        self._saved_record_count = 0
        self.last_arrival_logs = {}
        self.departed_vehicles = {}
        self.pos_suspend_until = {}
        self.pos_resume_logged = set()
        self.temp_pos1_data = {}
        self.temp_pos2_data = {}
        self._key_limit_notified = False
        self._last_date = datetime.now().date()
        self._completed_dates_saved = set()
        self._save_lock = threading.Lock()
        self._refresh_lock = threading.Lock()
        self.auto_save_path = None
        self.can_auto_save = False
        self.api_stats_today = {"POS1": 0, "POS2": 0, "SLST": 0, "RINF": 0, "SRCH": 0, "기타": 0}
        self.api_stats_yesterday = {"POS1": 0, "POS2": 0, "SLST": 0, "RINF": 0, "SRCH": 0, "기타": 0}
        self.route_map_panels = {}
        self.route_map_current = None
        self._schedule_timer = None   # 예약 기록 시작: 1초 폴링 QTimer
        self._scheduled_dt = None     # 예약 기록 시작: 목표 datetime
        self._stop_schedule_timer = None  # 예약 기록 중지: 1초 폴링 QTimer
        self._stop_scheduled_dt = None    # 예약 기록 중지: 목표 datetime

        self.sig_log.connect(self._slot_log)
        self.sig_record.connect(self._slot_record)
        self.sig_update_map.connect(self._slot_update_map)
        self.sig_clear_map.connect(self._slot_clear_map)
        self.sig_route_loaded.connect(self._slot_route_loaded)

        self._load_config()
        self._setup_ui()
        self.log("노선 검색 버튼으로 노선을 추가하세요.")

# ──────────────────────────────────────────────────────────
# 【6-2】 _make_info_html(link_color=None) → str
#   메뉴바 우측 모서리에 표시되는 제작자·데이터출처 정보 HTML 문자열 생성.
#   link_color: 현재 팔레트의 링크 색상. 테마 변경 시 함께 업데이트됨.
# ──────────────────────────────────────────────────────────
    def _make_info_html(self, link_color=None):
        if link_color is None:
            link_color = QApplication.palette().color(QPalette.Link).name()
        return (
            f'● 만든이 : 박 국 환 ( '
            f'<a href="mailto:ggoyong2@naver.com" style="color:{link_color};">ggoyong2@naver.com</a>'
            f' )   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;● 데이터 출처 : 공공데이터포털 Open API ( '
            f'<a href="https://www.data.go.kr" style="color:{link_color};">https://www.data.go.kr</a> )')

# ──────────────────────────────────────────────────────────
# 【6-3】 _setup_ui()
#   메인 창의 메뉴바·레이아웃·위젯 전체를 구성.
#
#   메뉴 구성:
#   [메뉴]      : 인증키 입력 / 갱신주기 입력 / 기록 시작 / 프로그램 종료
#   [노선 검색] : 노선 검색 열기
#   [테마]      : 라이트모드 / 다크모드
#   [API 현황]  : API 호출 현황 보기
#   [프로그램 정보]: 프로그램 정보 보기
#   (우측 모서리): 제작자 정보 HTML 라벨
#
#   레이아웃 구조:
#   QVBoxLayout(메인)
#   └─ outer_splitter (세로 분할자)
#      ├─ mid_splitter (가로 분할자)
#      │  ├─ map_stack      ← RouteMapPanel들이 쌓이는 스택 위젯
#      │  └─ right_widget
#      │     ├─ table_depart  ← 출발 기록 표 (RecordTable)
#      │     └─ table_arrive  ← 종료 기록 표 (RecordTable)
#      └─ log_container
#         ├─ log_header ("로그" 라벨 + "접기" 버튼)
#         └─ log_text  (QPlainTextEdit, 최대 3000줄)
# ──────────────────────────────────────────────────────────
    def _setup_ui(self):
        menubar = self.menuBar()
        menubar.setStyleSheet("QMenu::item:disabled { color: #AAAAAA; }")

        menu_main = menubar.addMenu("메뉴")
        menu_main.addAction("인증키 입력").triggered.connect(self._show_key_input)
        menu_main.addAction("갱신주기 입력").triggered.connect(self._ask_interval)
        self.act_toggle = menu_main.addAction("기록 시작")
        self.act_toggle.setEnabled(False)
        self.act_toggle.triggered.connect(self._on_toggle)
        self.act_schedule = menu_main.addAction("예약 기록 시작")
        self.act_schedule.setEnabled(False)
        self.act_schedule.triggered.connect(self._on_schedule_toggle)
        menu_main.addSeparator()
        menu_main.addAction("프로그램 종료").triggered.connect(self.close)

        menu_search = menubar.addMenu("노선 검색")
        self.act_search = menu_search.addAction("노선 검색 열기")
        self.act_search.triggered.connect(self._show_route_search)

        menu_theme = menubar.addMenu("테마")
        menu_theme.addAction("라이트모드").triggered.connect(lambda: self._apply_theme("light"))
        menu_theme.addAction("다크모드").triggered.connect(lambda: self._apply_theme("dark"))

        menu_api = menubar.addMenu("API 현황")
        menu_api.addAction("API 현황 보기").triggered.connect(self._show_api_status)

        menu_info = menubar.addMenu("프로그램 정보")
        menu_info.addAction("프로그램 정보 보기").triggered.connect(self._show_program_info)

        # 메뉴바 오른쪽에 정보 라벨 삽입
        self._info_lbl = QLabel(self._make_info_html())
        self._info_lbl.setFont(QFont(FONT_FAMILY, 8))
        self._info_lbl.setOpenExternalLinks(True)
        self._info_lbl.setTextFormat(Qt.RichText)
        self._info_lbl.setContentsMargins(20, 0, 10, 0)
        menubar.setCornerWidget(self._info_lbl, Qt.TopRightCorner)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(6, 6, 6, 6)

        outer_splitter = QSplitter(Qt.Vertical)
        main_layout.addWidget(outer_splitter)
        mid_splitter = QSplitter(Qt.Horizontal)
        outer_splitter.addWidget(mid_splitter)

        self.map_stack = QStackedWidget()
        mid_splitter.addWidget(self.map_stack)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(2)
        self.table_depart = RecordTable("운행 출발 시각 기록")
        self.table_arrive = RecordTable("운행 종료 시각 기록")
        right_layout.addWidget(self.table_depart)
        right_layout.addWidget(self.table_arrive)
        mid_splitter.addWidget(right_widget)

        mid_splitter.setSizes([740, 560])
        mid_splitter.setStretchFactor(0, 1)
        mid_splitter.setStretchFactor(1, 1)

        log_container = QWidget()
        log_container_layout = QVBoxLayout(log_container)
        log_container_layout.setContentsMargins(0, 0, 0, 0)
        log_container_layout.setSpacing(0)

        log_header = QHBoxLayout()
        log_header.setContentsMargins(6, 0, 6, 0)
        log_header.setSpacing(4)
        log_title_btn = QLabel("  로그  ")
        log_title_btn.setFont(QFont(FONT_FAMILY, 9, QFont.Bold))
        log_header.addWidget(log_title_btn)
        log_header.addStretch(1)
        # ── 기록 파일 경로 표시 (기록 중일 때만 보임) ──────────────
        self._log_file_prefix = QLabel("현재 기록 중인 파일 :")
        self._log_file_prefix.setFont(QFont(FONT_FAMILY, 8))
        self._log_file_prefix.setVisible(False)
        log_header.addWidget(self._log_file_prefix)
        self._log_file_label = ElidedLabel("")
        self._log_file_label.setFont(QFont(FONT_FAMILY, 8))
        self._log_file_label.setVisible(False)
        log_header.addWidget(self._log_file_label)
        self._btn_open_file = QPushButton("파일열기")
        self._btn_open_file.setFixedHeight(16)
        self._btn_open_file.setFont(QFont(FONT_FAMILY, 7))
        self._btn_open_file.setVisible(False)
        self._btn_open_file.clicked.connect(self._open_save_file)
        log_header.addWidget(self._btn_open_file)
        self._btn_open_folder = QPushButton("폴더열기")
        self._btn_open_folder.setFixedHeight(16)
        self._btn_open_folder.setFont(QFont(FONT_FAMILY, 7))
        self._btn_open_folder.setVisible(False)
        self._btn_open_folder.clicked.connect(self._open_save_folder)
        log_header.addWidget(self._btn_open_folder)
        log_header.addSpacing(8)
        # ────────────────────────────────────────────────────────────
        self._log_toggle_btn = QPushButton("접기")
        self._log_toggle_btn.setFixedSize(45, 16)
        self._log_toggle_btn.setFont(QFont(FONT_FAMILY, 7))
        self._log_toggle_btn.clicked.connect(self._toggle_log_panel)
        log_header.addWidget(self._log_toggle_btn)
        log_container_layout.addLayout(log_header)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont(FONT_MONO, 8))
        self.log_text.setMaximumBlockCount(3000)
        self._log_text_widget = self.log_text
        log_container_layout.addWidget(self.log_text)

        outer_splitter.addWidget(log_container)
        outer_splitter.setSizes([700, 154])
        self._outer_splitter = outer_splitter
        self._log_container = log_container
        self._log_expanded = True
        self._log_last_height = 200

        line_h = self.log_text.fontMetrics().lineSpacing()
        self._log_min_height = line_h * 5 + 30
        self.log_text.setMinimumHeight(line_h * 5 + 10)
        log_container.setMinimumHeight(self._log_min_height)
        outer_splitter.setCollapsible(outer_splitter.indexOf(log_container), False)

# ──────────────────────────────────────────────────────────
# 【6-4】 _toggle_log_panel()
#   로그 창 접기/펼치기 토글.
#   접기  : log_text를 숨기고, 로그 컨테이너를 24px(헤더만)으로 줄임.
#   펼치기: log_text를 보여주고, 이전 저장 높이로 복원.
#   outer_splitter.setSizes([위 크기, 아래 크기])로 영역 재배분.
# ──────────────────────────────────────────────────────────
    def _open_save_file(self):
# ──────────────────────────────────────────────────────────
# 【6-5】 _open_save_file()
#   로그 헤더 [파일열기] 버튼 클릭 시 호출.
#   현재 기록 중인 엑셀 파일을 기본 연결 프로그램(Excel 등)으로 바로 열기.
#   Windows: os.startfile / macOS: open 명령 / Linux: xdg-open
# ──────────────────────────────────────────────────────────
        if not self.auto_save_path or not os.path.exists(self.auto_save_path):
            QMessageBox.warning(self, "알림", "파일을 찾을 수 없습니다.")
            return
        try:
            if sys.platform == "win32":
                os.startfile(self.auto_save_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self.auto_save_path])
            else:
                subprocess.Popen(["xdg-open", self.auto_save_path])
        except Exception as e:
            QMessageBox.warning(self, "오류", f"파일 열기 실패: {e}")

    def _open_save_folder(self):
# ──────────────────────────────────────────────────────────
# 【6-6】 _open_save_folder()
#   로그 헤더 [폴더열기] 버튼 클릭 시 호출.
#   현재 기록 중인 엑셀 파일이 있는 폴더를 탐색기로 열고 파일을 선택 상태로 표시.
#   Windows: explorer /select / macOS: open -R / Linux: xdg-open (폴더)
# ──────────────────────────────────────────────────────────
        if not self.auto_save_path:
            QMessageBox.warning(self, "알림", "파일 경로를 확인할 수 없습니다.")
            return
        try:
            if sys.platform == "win32":
                subprocess.Popen(["explorer", "/select,", self.auto_save_path])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", self.auto_save_path])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(self.auto_save_path)])
        except Exception as e:
            QMessageBox.warning(self, "오류", f"폴더 열기 실패: {e}")

    def _toggle_log_panel(self):
        if self._log_expanded:
            sizes = self._outer_splitter.sizes()
            self._log_last_height = sizes[1]
            self._log_text_widget.hide()
            self._log_container.setMinimumHeight(0)
            self.log_text.setMinimumHeight(0)
            self._outer_splitter.setSizes([sizes[0] + sizes[1] - 24, 24])
            self._log_toggle_btn.setText("펼치기")
            self._log_expanded = False
        else:
            self._log_text_widget.show()
            self._log_container.setMinimumHeight(self._log_min_height)
            self.log_text.setMinimumHeight(self._log_min_height - 60)
            sizes = self._outer_splitter.sizes()
            restore_h = max(self._log_last_height, self._log_min_height)
            self._outer_splitter.setSizes([sizes[0] + sizes[1] - restore_h, restore_h])
            self._log_toggle_btn.setText("접기")
            self._log_expanded = True

# ──────────────────────────────────────────────────────────
# 【6-7】 _apply_theme(mode)
#   라이트/다크 테마를 앱 전체에 적용.
#   ① QApplication에 새 팔레트 적용
#   ② allWidgets()로 모든 위젯에 팔레트 강제 적용
#   ③ 메뉴바 우측 정보 라벨 링크 색상 갱신
#   ④ 모든 RouteMapPanel.refresh_theme() 호출
# ──────────────────────────────────────────────────────────
    def _apply_theme(self, mode):
        app = QApplication.instance()
        app.setStyle(QStyleFactory.create("Fusion"))
        p = _make_palette(mode)
        app.setPalette(p)
        for widget in app.allWidgets():
            widget.setPalette(p)
        if hasattr(self, '_info_lbl'):
            self._info_lbl.setText(self._make_info_html(p.color(QPalette.Link).name()))
        for panel in self.route_map_panels.values():
            panel.refresh_theme()

# ──────────────────────────────────────────────────────────
# 【6-8】 _slot_log(msg) [슬롯]
#   sig_log 수신 → log_text에 "[HH:MM:SS] 메시지" 형식으로 추가.
#
# 【6-9】 _slot_record(idx, entry) [슬롯]
#   sig_record 수신 → idx=0이면 table_depart, 1이면 table_arrive에 추가.
#
# 【6-10】 _slot_update_map(rnm, buses) [슬롯]
#   sig_update_map 수신 → 해당 노선 RouteMapPanel의 update_buses() 호출.
#
# 【6-11】 _slot_clear_map(rnm) [슬롯]
#   sig_clear_map 수신 → 해당 노선 지도에 빈 리스트 전달하여 버스 제거.
#
# 【6-12】 log(msg)
#   sig_log를 emit하는 헬퍼. 백그라운드 스레드에서 UI를 안전하게 업데이트.
# ──────────────────────────────────────────────────────────
    @Slot(str)
    def _slot_log(self, msg):
        self.log_text.appendPlainText(datetime.now().strftime("[%H:%M:%S] ") + msg)

    @Slot(int, tuple)
    def _slot_record(self, idx, entry):
        (self.table_depart if idx == 0 else self.table_arrive).add_row(entry)

    @Slot(str, list)
    def _slot_update_map(self, rnm, buses):
        if rnm in self.route_map_panels:
            self.route_map_panels[rnm].update_buses(buses)

    @Slot(str)
    def _slot_clear_map(self, rnm):
        if rnm in self.route_map_panels:
            self.route_map_panels[rnm].update_buses([])

    def log(self, msg):
        self.sig_log.emit(msg)

# ──────────────────────────────────────────────────────────
# 【6-13】 _make_fernet() → Fernet or None
#   _SECRET 문장을 pbkdf2_hmac(SHA-256, 100000회 반복)으로
#   32바이트 키로 변환 후 urlsafe_b64encode하여 Fernet 암호화 객체 생성.
#   cryptography 라이브러리 없으면 None 반환.
#
# 【6-14】 _enc_key(raw) → str
#   평문 API 키를 Fernet으로 암호화. 실패 시 원본 반환.
#
# 【6-15】 _dec_key(enc) → str
#   암호화된 API 키를 복호화. 실패 시 원본 반환.
# ──────────────────────────────────────────────────────────
    def _make_fernet(self):
        if not _CRYPTO_OK:
            return None
        dk = _hl.pbkdf2_hmac(
            "sha256", self._SECRET.encode("utf-8"),
            b"SeoulBusSalt2025", iterations=100_000, dklen=32)
        return _Fernet(base64.urlsafe_b64encode(dk))

    def _enc_key(self, raw):
        if not raw or not _CRYPTO_OK:
            return raw
        try:
            return self._make_fernet().encrypt(raw.encode("utf-8")).decode("ascii")
        except:
            return raw

    def _dec_key(self, enc):
        if not enc or not _CRYPTO_OK:
            return enc
        try:
            return self._make_fernet().decrypt(enc.encode("ascii")).decode("utf-8")
        except:
            return enc

# ──────────────────────────────────────────────────────────
# 【6-16】 _load_config()
#   Seoul_Bus_Config.ini 파일에서 암호화된 API 키를 읽어 복호화 후 저장.
#   INI 구조: [keys] main_key = (암호화값) / back_key = (암호화값)
#   파일 없으면 조용히 리턴.
#
# 【6-17】 _save_config()
#   현재 API 키를 암호화하여 설정 파일에 저장.
#   기존 파일의 favorites 섹션 등 다른 설정은 보존.
# ──────────────────────────────────────────────────────────
    def _load_config(self):
        cfg = configparser.ConfigParser()
        if not os.path.exists(self._CFG_FILE):
            return
        try:
            cfg.read(self._CFG_FILE, encoding="utf-8")
            enc_m = cfg.get("keys", "main_key", fallback="")
            enc_b = cfg.get("keys", "back_key", fallback="")
            if enc_m:
                self.api_key_main = self._dec_key(enc_m)
            if enc_b:
                self.api_key_back = self._dec_key(enc_b)
        except:
            pass

    def _save_config(self):
        cfg = configparser.ConfigParser()
        if os.path.exists(self._CFG_FILE):
            try:
                cfg.read(self._CFG_FILE, encoding="utf-8")
            except:
                pass
        if "keys" not in cfg:
            cfg["keys"] = {}
        cfg["keys"]["main_key"] = self._enc_key(self.api_key_main)
        cfg["keys"]["back_key"] = self._enc_key(self.api_key_back)
        try:
            with open(self._CFG_FILE, "w", encoding="utf-8") as f:
                cfg.write(f)
        except:
            pass

# ──────────────────────────────────────────────────────────
# 【6-18】 _show_key_input()
#   API 인증키를 입력·검증하는 대화상자.
#
#   구성: 메인키(64자) 입력창 + 보조키(64자 또는 공란) 입력창
#         + 상태 라벨 + [입력][취소] 버튼
#
#   검증 과정:
#   ① 입력 중 실시간으로 글자 수 확인 → 64자이면 [입력] 버튼 활성화
#   ② [입력] 클릭 → 백그라운드 스레드에서 실제 API 호출(_test_key)
#   ③ sig_key_verified 시그널로 결과를 메인 스레드에 전달
#   ④ 결과 처리:
#      "REJECTED" → 등록되지 않은 키 메시지 (빨간 글자)
#      "OK"/"LIMIT" → 성공 → 키 저장 후 대화상자 닫기
#      "LIMIT" → 한도 초과 경고 후 저장 (여전히 유효한 키)
#      기타 오류 → 오류 메시지 표시
# ──────────────────────────────────────────────────────────
    def _show_key_input(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("인증키 입력")
        dlg.setFixedSize(560, 260)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel("메인 인증키 (64자리, 필수)"))
        edt_main = QLineEdit(self.api_key_main)
        edt_main.setFont(QFont(FONT_MONO, 8))
        lay.addWidget(edt_main)
        lay.addWidget(QLabel("보조 인증키 (64자리 또는 공란)"))
        edt_back = QLineEdit(self.api_key_back)
        edt_back.setFont(QFont(FONT_MONO, 8))
        lay.addWidget(edt_back)
        lbl_status = QLabel("")
        lbl_status.setStyleSheet("color: gray;")
        lay.addWidget(lbl_status)
        btn_bar = QHBoxLayout()
        btn_ok = QPushButton("입력")
        btn_ok.setEnabled(False)
        btn_cancel = QPushButton("취소")
        btn_bar.addStretch()
        btn_bar.addWidget(btn_ok)
        btn_bar.addWidget(btn_cancel)
        lay.addLayout(btn_bar)
        btn_cancel.clicked.connect(dlg.reject)

        def _check_len():
            km = len(edt_main.text().strip())
            kb = len(edt_back.text().strip())
            btn_ok.setEnabled(km == 64 and (kb == 0 or kb == 64))
        edt_main.textChanged.connect(lambda: _check_len())
        edt_back.textChanged.connect(lambda: _check_len())
        _check_len()

        def _do_ok():
            km = edt_main.text().strip()
            kb = edt_back.text().strip()
            lbl_status.setText("인증키 검증 중...")
            lbl_status.setStyleSheet("color: gray;")
            btn_ok.setEnabled(False)

            def _test_key(k):
                if not k:
                    return None
                try:
                    resp = requests.get(URL_SRCH, params={"ServiceKey": k, "strSrch": "13"}, timeout=8)
                    self.api_stats_today["SRCH"] += 1
                    raw_up = resp.text.upper()
                    if any(p in raw_up for p in ("TOO MANY", "LIMITED NUMBER", "LIMITED_NUMBER", "RATE LIMIT")):
                        return "LIMIT"
                    if any(p in raw_up for p in ("NOT REGISTERED", "UNREGISTERED", "SERVICE KEY IS NOT")):
                        return "REJECTED"
                    root = ET.fromstring(resp.text)
                    code = (root.findtext(".//headerCd") or "").strip()
                    if code in ("0", "3", "4"):
                        return "OK"
                    return f"오류코드: {code}"
                except Exception as e:
                    return f"오류: {e}"

            _results = [None, None]

            def _key_run(): 
                _results[0] = _test_key(km)
                _results[1] = _test_key(kb)
                self.sig_key_verified.emit(
                    str(_results[0] or ""),
                    str(_results[1] or ""),
                    km, kb)

            threading.Thread(target=_key_run, daemon=True).start()

        def _on_verified(rm_str, rb_str, key_m, key_b):
            # _do_ok 바깥(_show_key_input 스코프)에 정의 →
            # dlg.exec() 반환 후에도 disconnect 가 정상 참조 가능
            if not dlg.isVisible():
                return
            rm = None if rm_str == "" else rm_str
            rb = None if rb_str == "" else rb_str
            if rm == "REJECTED":
                lbl_status.setText("메인키: 등록되지 않은 인증키")
                lbl_status.setStyleSheet("color: red;")
                btn_ok.setEnabled(True)
            elif rm not in ("OK", "LIMIT", None):
                lbl_status.setText(f"메인키: {rm}")
                lbl_status.setStyleSheet("color: red;")
                btn_ok.setEnabled(True)
            elif rb == "REJECTED":
                lbl_status.setText("보조키: 등록되지 않은 인증키")
                lbl_status.setStyleSheet("color: red;")
                btn_ok.setEnabled(True)
            elif rb not in ("OK", "LIMIT", None) and key_b:
                lbl_status.setText(f"보조키: {rb}")
                lbl_status.setStyleSheet("color: red;")
                btn_ok.setEnabled(True)
            else:
                self.api_key_main = key_m
                self.api_key_back = key_b
                self._save_config()
                if self.routes:
                    self.act_toggle.setEnabled(True)
                    self.act_schedule.setEnabled(True)
                self.log(f"인증키 등록 완료 (메인: 설정됨, 보조: {'설정됨' if key_b else '없음'})")
                if rm == "LIMIT" or rb == "LIMIT":
                    self.log("⚠ 일부 키가 한도초과 상태입니다 (유효)")
                dlg.accept()

        btn_ok.clicked.connect(_do_ok)
        self.sig_key_verified.connect(_on_verified)
        dlg.exec()
        try:
            self.sig_key_verified.disconnect(_on_verified)
        except (TypeError, RuntimeError):
            pass
        
# ──────────────────────────────────────────────────────────
# 【6-19】 _show_route_search()
#   노선 검색 및 즐겨찾기 관리 대화상자.
#
#   구성:
#   - 검색어 입력창 + [검색] 버튼
#   - 검색 결과 트리위젯 (노선번호·유형·기점·종점·첫차·막차·간격·길이)
#   - ★ 즐겨찾기 목록 트리위젯 (동일 구조, 최대 150px)
#   - [★즐겨찾기 추가][✕삭제][노선 선택하기][닫기] 버튼
#
#   주요 동작:
#   - 검색: 백그라운드 스레드 → URL_SRCH API 호출 → sig_search_done → 결과 표시
#   - 선택: [노선 선택하기] 또는 더블클릭/엔터
#           → _load_route_from_search() 스레드 실행
#           → sig_route_loaded → _slot_route_loaded
#   - 즐겨찾기: 추가/삭제 후 _save_favorites() 로 설정 파일에 저장
#   - 창 크기 변경 시 컬럼 너비 비례 재배분 (지연 타이머 사용)
# ──────────────────────────────────────────────────────────
    def _show_route_search(self):
        if hasattr(self, '_search_dlg') and self._search_dlg and self._search_dlg.isVisible():
            self._search_dlg.raise_()
            self._search_dlg.activateWindow()
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("노선 검색")
        dlg.resize(875, 580)
        dlg.setMinimumWidth(875)
        dlg.setModal(True)
        self._search_dlg = dlg
        lay = QVBoxLayout(dlg)

        top = QHBoxLayout()
        edt = QLineEdit()
        edt.setPlaceholderText("노선번호 입력 (예: 101, 7016, N37, 알파벳은 대문자로 검색)")
        edt.setFont(QFont(FONT_FAMILY, 11))
        btn_search = QPushButton("검색")
        btn_search.setEnabled(False)
        top.addWidget(edt)
        top.addWidget(btn_search)
        lay.addLayout(top)

        hint = QLabel("")
        hint.setFont(QFont(FONT_FAMILY, 9))
        hint.setAlignment(Qt.AlignCenter)
        lay.addWidget(hint)

        col_labels = ["노선번호", "노선유형", "기점", "종점",
                      "첫차출발시각", "막차출발시각", "간격(분)", "노선길이(km)"]

        base_col_widths = [103, 63, 180, 180, 95, 95, 50, 72]
        base_total = sum(base_col_widths)

        def _apply_proportional_widths(tw, available_width):
            """트리위젯의 컬럼을 available_width에 비례하여 배분"""
            if available_width < base_total:
                available_width = base_total
            for i, bw in enumerate(base_col_widths):
                new_w = int(available_width * bw / base_total)
                tw.setColumnWidth(i, new_w)

        tree = QTreeWidget()
        tree.setHeaderLabels(col_labels)
        tree.setFont(QFont(FONT_FAMILY, 9))
        tree.setRootIsDecorated(False)
        tree.setAlternatingRowColors(True)
        hdr = tree.header()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.setDefaultAlignment(Qt.AlignCenter)
        for i, w in enumerate(base_col_widths):
            tree.setColumnWidth(i, w)
        lay.addWidget(tree)

        fav_tree = QTreeWidget()
        fav_tree.setHeaderLabels(col_labels)
        fav_tree.setFont(QFont(FONT_FAMILY, 9))
        fav_tree.setRootIsDecorated(False)
        fav_tree.setAlternatingRowColors(True)
        fav_hdr = fav_tree.header()
        fav_hdr.setSectionResizeMode(QHeaderView.Interactive)
        fav_hdr.setDefaultAlignment(Qt.AlignCenter)
        for i, w in enumerate(base_col_widths):
            fav_tree.setColumnWidth(i, w)
        fav_tree.setMaximumHeight(150)

        fav_label = QLabel("  ★ 즐겨찾기 노선  ")
        fav_label.setFont(QFont(FONT_FAMILY, 9, QFont.Bold))
        lay.addWidget(fav_label)
        lay.addWidget(fav_tree)

        btn_bar = QHBoxLayout()
        btn_fav_add = QPushButton("★ 즐겨찾기 추가")
        btn_fav_del = QPushButton("✕ 즐겨찾기 삭제")
        btn_load = QPushButton("노선 선택하기")
        btn_load.setEnabled(False)
        btn_close = QPushButton("닫기")
        btn_bar.addWidget(btn_fav_add)
        btn_bar.addWidget(btn_fav_del)
        btn_bar.addStretch()
        btn_bar.addWidget(btn_load)
        btn_bar.addWidget(btn_close)
        lay.addLayout(btn_bar)
        btn_close.clicked.connect(dlg.close)

        favorites = self._load_favorites()
        _active_tree = [None]

        _resize_timer = QTimer(dlg)
        _resize_timer.setSingleShot(True)
        _resize_timer.setInterval(100)

        def _on_dlg_resize():
            available = tree.viewport().width()
            if available > base_total:
                _apply_proportional_widths(tree, available)
                _apply_proportional_widths(fav_tree, available)

        _resize_timer.timeout.connect(_on_dlg_resize)

        _orig_resize = dlg.resizeEvent
        def _dlg_resize_event(event):
            if _orig_resize:
                _orig_resize(event)
            _resize_timer.start()
        dlg.resizeEvent = _dlg_resize_event

        def _fmt_time(raw):
            digits = ''.join(c for c in str(raw).strip() if c.isdigit())
            if len(digits) >= 12:
                return f"{int(digits[8:10])}시 {int(digits[10:12])}분"
            if len(digits) >= 4:
                return f"{int(digits[:2])}시 {int(digits[2:4])}분"
            return raw

        def _fmt_length(raw):
            raw = str(raw).strip()
            if not raw:
                return raw
            try:
                val = float(raw)
                if val >= 100:
                    return f"{val / 1000:.1f}"
                else:
                    return f"{val:.1f}"
            except ValueError:
                return raw

        def _make_item(r):
            texts = [
                r.get("name", ""), r.get("type", ""),
                r.get("st", ""), r.get("ed", ""),
                _fmt_time(r.get("first", "")),
                _fmt_time(r.get("last", "")),
                r.get("term", ""),
                _fmt_length(r.get("length", ""))]
            item = QTreeWidgetItem(texts)
            for i in range(len(texts)):
                item.setTextAlignment(i, Qt.AlignCenter)
            return item

        def _populate_fav():
            fav_tree.clear()
            for fav in favorites:
                item = _make_item(fav)
                item.setData(0, Qt.UserRole, fav.get("route_id", ""))
                item.setData(0, Qt.UserRole + 1, fav)
                fav_tree.addTopLevelItem(item)

        _populate_fav()

        def _on_search_text_changed(text):
            btn_search.setEnabled(len(text.strip()) > 0)

        edt.textChanged.connect(_on_search_text_changed)

        def _do_search():
            q = edt.text().strip()
            if not q:
                return
            if not self.api_key_main:
                hint.setText("인증키를 먼저 입력하세요.")
                return
            hint.setText("검색중...")
            btn_search.setEnabled(False)
            tree.clear()

            def _search_run():
                rows = []
                try:
                    p = {"ServiceKey": unquote(self.api_key_main), "strSrch": q}
                    resp = requests.get(URL_SRCH, params=p, timeout=10)
                    self.api_stats_today["SRCH"] += 1
                    root = ET.fromstring(resp.content)
                    for it in root.iter("itemList"):
                        rows.append({
                            "route_id": (it.findtext("busRouteId") or "").strip(),
                            "name": (it.findtext("busRouteNm") or "").strip(),
                            "type": ROUTE_TYPE_LABEL.get((it.findtext("routeType") or "").strip(), "기타"),
                            "route_type_raw": (it.findtext("routeType") or "").strip(),
                            "st": (it.findtext("stStationNm") or "").strip(),
                            "ed": (it.findtext("edStationNm") or "").strip(),
                            "first": (it.findtext("firstBusTm") or "").strip(),
                            "last": (it.findtext("lastBusTm") or "").strip(),
                            "term": (it.findtext("term") or "").strip(),
                            "length": (it.findtext("length") or "").strip(),
                        })
                except:
                    pass
                self.sig_search_done.emit(rows)

            threading.Thread(target=_search_run, daemon=True).start()

        def _on_search_done(rows):
            btn_search.setEnabled(len(edt.text().strip()) > 0)
            if not dlg.isVisible():
                return
            tree.clear()
            if not rows:
                hint.setText("검색 결과가 없습니다.")
                return
            for r in rows:
                item = _make_item(r)
                item.setData(0, Qt.UserRole, r["route_id"])
                item.setData(0, Qt.UserRole + 1, r)
                tree.addTopLevelItem(item)
            hint.setText(f"검색 결과 {len(rows)}건")

        self.sig_search_done.connect(_on_search_done)

        def _update_load_btn():
            cur = _get_selected()
            btn_load.setEnabled(cur is not None)

        def _on_tree_sel():
            _active_tree[0] = tree
            fav_tree.clearSelection()
            _update_load_btn()

        def _on_fav_sel():
            _active_tree[0] = fav_tree
            tree.clearSelection()
            _update_load_btn()

        tree.currentItemChanged.connect(_on_tree_sel)
        fav_tree.currentItemChanged.connect(_on_fav_sel)

        def _get_selected():
            t = _active_tree[0]
            if t and t.currentItem():
                return t.currentItem()
            return None

        def _add_fav():
            cur = tree.currentItem()
            if not cur:
                return
            rid = cur.data(0, Qt.UserRole)
            for f in favorites:
                if f.get("route_id") == rid:
                    hint.setText("이미 즐겨찾기에 있습니다.")
                    return
            rdata = cur.data(0, Qt.UserRole + 1)
            if rdata:
                favorites.append(rdata)
                self._save_favorites(favorites)
                _populate_fav()
                hint.setText(f"'{cur.text(0)}' 즐겨찾기 추가됨")

        def _del_fav():
            cur = fav_tree.currentItem()
            if not cur:
                return
            rid = cur.data(0, Qt.UserRole)
            before = len(favorites)
            favorites[:] = [f for f in favorites if f.get("route_id") != rid]
            if len(favorites) < before:
                self._save_favorites(favorites)
                _populate_fav()
                hint.setText("즐겨찾기에서 삭제됨")

        btn_fav_add.clicked.connect(_add_fav)
        btn_fav_del.clicked.connect(_del_fav)

        _loading = [False]  # 중복 실행 방지 플래그

        def _do_load():
            if _loading[0]:
                return
            _loading[0] = True
            cur = _get_selected()
            if not cur:
                _loading[0] = False
                return
            rid = cur.data(0, Qt.UserRole)
            rdata = cur.data(0, Qt.UserRole + 1)
            rname = cur.text(0)
            rtype = rdata.get("route_type_raw", "3") if rdata else "3"
            hint.setText(f"'{rname}' 노선 정보 불러오는 중...")
            btn_load.setEnabled(False)
            def _load_run():
                self._load_route_from_search(rid, rname, rtype, dlg)
            threading.Thread(target=_load_run, daemon=True).start()
            QTimer.singleShot(1500, lambda: _loading.__setitem__(0, False))

        btn_load.clicked.connect(_do_load)
        btn_search.clicked.connect(_do_search)
        edt.returnPressed.connect(_do_search)

        tree.itemActivated.connect(lambda item, col: _do_load())
        fav_tree.itemActivated.connect(lambda item, col: _do_load())

        dlg.exec()
        try:
            self.sig_search_done.disconnect(_on_search_done)
        except:
            pass
        self._search_dlg = None

# ──────────────────────────────────────────────────────────
# 【6-20】 _load_route_from_search(route_id, route_name, route_type, dlg)
#   백그라운드 스레드에서 실행. 선택 노선의 상세 정보를 API로 수집.
#
#   API 호출:
#   ① URL_SLST → 정류소 목록 (seq 순 정렬)
#   ② URL_RINF → 회차 순번·운수사·첫차·막차·노선 길이 추출
#
#   route_data 주요 키:
#   stops, turn_index, first/second/last _st_id/nm/ars,
#   corp_nm, first/last_bus_tm, rlength, rtype
#
#   완성된 route_data를 sig_route_loaded.emit()으로 메인 스레드에 전달.
# ──────────────────────────────────────────────────────────
    def _load_route_from_search(self, route_id, route_name, route_type, dlg):
        root_slst = self.fetch_api(URL_SLST, {"busRouteId": route_id})
        if root_slst is None or isinstance(root_slst, tuple):
            self.log(f"[오류] '{route_name}' 정류소 목록 API 실패")
            self.sig_route_loaded.emit({})
            return

        root_rinf = self.fetch_api(URL_RINF, {"busRouteId": route_id})
        turn_seq = "0"
        corp_nm = "서울시내버스"
        first_bus_tm = None
        last_bus_tm = None
        rlength = ""
        if root_rinf is not None and not isinstance(root_rinf, tuple):
            for it in root_rinf.iter("itemList"):
                ts = (it.findtext("turnSeq") or "0").strip()
                if ts and ts != "0":
                    turn_seq = ts
                cn = (it.findtext("corpNm") or "").strip()
                if cn:
                    corp_nm = cn
                ft = (it.findtext("firstBusTm") or "").strip()
                lt = (it.findtext("lastBusTm") or "").strip()
                if ft:
                    first_bus_tm = format_hhmm(ft)
                if lt:
                    last_bus_tm = format_hhmm(lt)
                rl = (it.findtext("length") or "").strip()
                if rl:
                    rlength = rl
                rt = (it.findtext("routeType") or "").strip()
                if rt:
                    route_type = rt
                break

        stops = []
        for it in root_slst.iter("itemList"):
            sid = (it.findtext("station") or "").strip()
            seq_raw = (it.findtext("seq") or "0").strip()
            nm = (it.findtext("stationNm") or "").strip()
            arsid = (it.findtext("arsId") or "").strip()
            trans = (it.findtext("transYn") or "N").strip().upper()
            full_sect = (it.findtext("fullSectDist") or "0").strip()
            stops.append({
                "station": sid, "seq": seq_raw, "name": nm,
                "arsId": arsid, "transYn": trans, "fullSectDist": full_sect,
            })
        stops.sort(key=lambda s: int(s["seq"]) if s["seq"].isdigit() else 0)

        if not stops:
            self.log(f"[오류] '{route_name}' 정류소 없음")
            self.sig_route_loaded.emit({})
            return

        turn_idx = 0
        try:
            ti = int(turn_seq)
            if 0 < ti <= len(stops):
                turn_idx = ti - 1
        except:
            turn_idx = len(stops) // 2

        color = ROUTE_TYPE_COLOR.get(route_type, DEFAULT_LINE_COLOR)
        rkey = route_name
        n = len(stops)

        route_data = {
            "key": rkey, "route_id": route_id, "route_name": route_name,
            "route_type": route_type, "route_color": color,
            "stops": stops, "turn_index": turn_idx,
            "rnm": route_name, "rid": route_id, "st_cnt": n,
            "first_st_id": stops[0]["station"] if stops else "",
            "second_st_id": stops[1]["station"] if n >= 2 else "",
            "last_st_id": stops[-1]["station"] if stops else "",
            "first_nm": stops[0]["name"] if stops else "?",
            "second_nm": stops[1]["name"] if n >= 2 else "?",
            "last_nm": stops[-1]["name"] if stops else "?",
            "first_ars": stops[0]["arsId"] if stops else "?",
            "second_ars": stops[1]["arsId"] if n >= 2 else "?",
            "last_ars": stops[-1]["arsId"] if stops else "?",
            "corp_nm": corp_nm, "first_bus_tm": first_bus_tm,
            "last_bus_tm": last_bus_tm, "rtype": route_type,
            "rlength": rlength, "stations": stops, "dlg": dlg,
        }
        self.sig_route_loaded.emit(route_data)

# ──────────────────────────────────────────────────────────
# 【6-21】 _slot_route_loaded(data) [슬롯]
#   sig_route_loaded 수신 → 새 노선을 메인 창에 등록.
#   ① routes 리스트 교체
#   ② 해당 노선 RouteMapPanel 없으면 생성 후 map_stack에 추가
#   ③ panel.load_route() → map_stack.setCurrentWidget()
#   ④ API 키+노선 모두 있으면 [기록 시작] 버튼 활성화
#   ⑤ 노선 정보 로그 출력
#   ⑥ 검색 대화상자 닫기 (dlg.accept)
# ──────────────────────────────────────────────────────────
    @Slot(dict)
    def _slot_route_loaded(self, data):
        if not data:
            return
        rnm = data.get("rnm") or data.get("route_name", "")
        route_name = data.get("route_name", rnm)
        stops = data["stops"]
        turn_idx = data["turn_index"]
        dlg = data.get("dlg")
        n = len(stops)

        self.routes = [data]

        if rnm not in self.route_map_panels:
            panel = RouteMapPanel(rnm)
            self.map_stack.addWidget(panel)
            self.route_map_panels[rnm] = panel

        panel = self.route_map_panels[rnm]
        panel.load_route(stops, data.get("rtype", "3"), data.get("rlength", ""))
        self.map_stack.setCurrentWidget(panel)

        if self.api_key_main and self.routes:
            self.act_toggle.setEnabled(True)
            self.act_schedule.setEnabled(True)

        self.log(f"[{route_name}] 정류소 {n}개 · "
                 f"첫정류소: {data.get('first_nm','?')} / "
                 f"두번째: {data.get('second_nm','?')} / 종점: {data.get('last_nm','?')}")
        self.log(f"[{route_name}] 노선 등록 완료 · "
                 f"운수사: {data.get('corp_nm','')} / "
                 f"첫차: {data.get('first_bus_tm','')} / 막차: {data.get('last_bus_tm','')}")

        if dlg and dlg.isVisible():
            try:
                dlg.accept()
            except:
                pass

# ──────────────────────────────────────────────────────────
# 【6-22】 fetch_api(url, params) → XML root | tuple | None  ★ API 공통 호출 ★
#   메인 키 → 보조 키 순으로 API 호출 시도.
#
#   반환값:
#   XML root Element  : 정상 응답 (headerCd == "0")
#   ('NO_BUS', 첫차시각) : 결과 없음 (데이터 없음 메시지)
#   None              : 모든 키 실패
#
#   처리 흐름 (각 키에 대해):
#   ① requests.get(url, params={ServiceKey:key, ...}, timeout=10)
#   ② 한도 초과 패턴 감지 → lh(한도초과 횟수) 증가 후 다음 키
#      패턴: "TOO MANY", "LIMITED NUMBER", "RATE LIMIT", "<headerCd>22</headerCd>"
#   ③ 미등록 키 감지 → 다음 키
#   ④ HTTP 상태코드 != 200 → 다음 키
#   ⑤ XML 파싱 → headerCd "0" → 성공 반환
#   ⑥ "결과가 없습니다"/"NODATA" → ('NO_BUS', ...) 반환
#   ⑦ 모든 키 한도 초과 시 경고 로그 (1회 출력 후 _key_limit_notified=True)
#   ⑧ API 호출 통계(api_stats_today) 갱신
# ──────────────────────────────────────────────────────────
    def fetch_api(self, url, params):
        lh = 0
        kc = sum(1 for k in [self.api_key_main, self.api_key_back] if k)
        for key in [self.api_key_main, self.api_key_back]:
            if not key:
                continue
            try:
                p = dict(params)
                p['ServiceKey'] = unquote(key)
                resp = requests.get(url, params=p, timeout=10)
                LP = ("TOO MANY", "LIMITED NUMBER", "LIMITED_NUMBER", "RATE LIMIT", "<headerCd>22</headerCd>")
                if any(pt in resp.text.upper() for pt in LP):
                    lh += 1
                    continue
                if "SERVICE KEY IS NOT REGISTERED" in resp.text or "UNREGISTERED_KEY" in resp.text:
                    continue
                if resp.status_code != 200:
                    continue
                root = ET.fromstring(resp.text)
                hc = root.findtext(".//headerCd") or ""
                em = root.findtext(".//headerMsg") or ""
                if URL_POS1 in url:
                    self.api_stats_today["POS1"] += 1
                elif URL_POS2 in url:
                    self.api_stats_today["POS2"] += 1
                elif URL_SLST in url:
                    self.api_stats_today["SLST"] += 1
                elif URL_RINF in url:
                    self.api_stats_today["RINF"] += 1
                elif URL_SRCH in url:
                    self.api_stats_today["SRCH"] += 1
                else:
                    self.api_stats_today["기타"] += 1
                if hc == "0":
                    return root
                nd = "결과가 없습니다" in em or "NODATA" in em
                if nd:
                    if URL_POS2 in url:
                        return ('NO_BUS', self._fetch_first_time(params.get('busRouteId', '')))
                    else:
                        return ('NO_BUS', None)
            except:
                continue
        if lh > 0 and lh >= kc and kc > 0 and not self._key_limit_notified:
            self._key_limit_notified = True
            self.log("⚠ 모든 인증키의 호출 한도가 초과되었습니다.")
        return None

# ──────────────────────────────────────────────────────────
# 【6-23】 _fetch_first_time(rid) → "HH:MM" or None
#   URL_RINF API로 특정 노선의 첫차 출발 시각 조회.
#   운행 종료 감지 후 다음 첫차까지 API 호출을 일시 중지할 때 필요.
# ──────────────────────────────────────────────────────────
    def _fetch_first_time(self, rid):
        if not rid:
            return None
        try:
            r = self.fetch_api(URL_RINF, {'busRouteId': rid})
            if r is None or isinstance(r, tuple):
                return None
            raw = r.findtext(".//firstBusTm") or ""
            if raw:
                return format_hhmm(raw)
        except:
            pass
        return None

    def _on_schedule_toggle(self):
# ──────────────────────────────────────────────────────────
# 【6-24】 _on_schedule_toggle()
#   act_schedule 버튼 클릭 시 호출. 현재 상태에 따라 4가지로 분기:
#
#   ① 시작 예약 대기 중  → 시작 예약 취소
#   ② 기록 중 + 중지 예약 대기 중 → 중지 예약 취소
#   ③ 기록 중 + 중지 미예약  → 중지 시각 입력 창 (_ask_scheduled_stop)
#   ④ 대기 중 (기본 상태) → 시작 시각 입력 창 (_ask_scheduled_start)
# ──────────────────────────────────────────────────────────
        # ① 시작 예약 대기 중 → 취소
        if self._schedule_timer is not None:
            self._schedule_timer.stop()
            self._schedule_timer = None
            self._scheduled_dt = None
            self.act_schedule.setText("예약 기록 시작")
            self.log("⏰ 시작 예약이 취소되었습니다.")
            return
        # ② 기록 중 + 중지 예약 대기 중 → 중지 예약 취소
        if self.is_monitoring and self._stop_schedule_timer is not None:
            self._stop_schedule_timer.stop()
            self._stop_schedule_timer = None
            self._stop_scheduled_dt = None
            self.act_schedule.setText("예약 기록 중지")
            self.log("⏰ 중지 예약이 취소되었습니다.")
            return
        # ③ 기록 중 + 중지 미예약 → 중지 시각 입력 창
        if self.is_monitoring:
            self._ask_scheduled_stop()
            return
        # ④ 대기 중 → 시작 시각 입력 창
        self._ask_scheduled_start()

    def _ask_scheduled_start(self):
# ──────────────────────────────────────────────────────────
# 【6-25】 _ask_scheduled_start()
#   기록을 시작할 예약 시각(년·월·일·시·분)을 입력 받는 대화상자.
#   갱신주기 입력 대화상자(_ask_interval)와 동일한 스타일.
#
#   UI 구성:
#   - 현재 시각 표시 레이블
#   - 년 / 월 / 일 / 시 / 분  QLineEdit 5개 (현재 시각으로 초기값)
#   - OK / Cancel 버튼
#
#   검증 규칙:
#   - 숫자 이외 입력 또는 유효하지 않은 날짜이면 경고
#   - 현재 시각 이전이면 경고
#
#   OK 확인 후:
#   - self._scheduled_dt 에 목표 datetime 저장
#   - _register_schedule_timer() 호출 → 1초 폴링 타이머 시작
#   - 메뉴 텍스트를 "예약 취소 (YYYY-MM-DD HH:MM 예약됨)"으로 변경
#   - 로그에 예약 시각 기록
# ──────────────────────────────────────────────────────────
        now = datetime.now()
        dlg = QDialog(self)
        dlg.setWindowTitle("예약 기록 시작")
        dlg.setFixedSize(340, 200)
        lo = QVBoxLayout(dlg)

        lbl_now = QLabel(f"현재 시각: {now.strftime('%Y-%m-%d %H:%M:%S')}")
        lo.addWidget(lbl_now)
        lo.addWidget(QLabel("기록을 시작할 시각을 입력하세요:"))
        _tick_start = QTimer(dlg)
        _tick_start.setInterval(1000)
        _tick_start.timeout.connect(lambda: lbl_now.setText(f"현재 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"))
        _tick_start.start()

        row = QHBoxLayout()
        ed_year  = QLineEdit(str(now.year))
        ed_month = QLineEdit(f"{now.month:02d}")
        ed_day   = QLineEdit(f"{now.day:02d}")
        ed_hour  = QLineEdit(f"{now.hour:02d}")
        ed_min   = QLineEdit(f"{now.minute:02d}")
        for ed in [ed_year, ed_month, ed_day, ed_hour, ed_min]:
            ed.setAlignment(Qt.AlignCenter)
        ed_year.setFixedWidth(52)
        ed_month.setFixedWidth(32)
        ed_day.setFixedWidth(32)
        ed_hour.setFixedWidth(32)
        ed_min.setFixedWidth(32)
        row.addWidget(ed_year);  row.addWidget(QLabel("년"))
        row.addWidget(ed_month); row.addWidget(QLabel("월"))
        row.addWidget(ed_day);   row.addWidget(QLabel("일"))
        row.addWidget(ed_hour);  row.addWidget(QLabel("시"))
        row.addWidget(ed_min);   row.addWidget(QLabel("분"))
        row.addStretch()
        lo.addLayout(row)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        lo.addWidget(bb)

        def ok():
            try:
                y  = int(ed_year.text())
                mo = int(ed_month.text())
                d  = int(ed_day.text())
                h  = int(ed_hour.text())
                mi = int(ed_min.text())
                target = datetime(y, mo, d, h, mi, 0)
            except ValueError:
                QMessageBox.warning(dlg, "알림", "올바른 날짜/시간을 입력해주세요.")
                return
            if target <= datetime.now():
                QMessageBox.warning(dlg, "알림", "현재 시각 이후로 설정해야 합니다.")
                return
            # 이전 기록이 있으면 확인 팝업 (예약 시각 도달 시 자동 초기화됨을 안내)
            if self.recorded_data:
                ret = QMessageBox.question(
                    dlg, "이전 기록 있음",
                    "이전 기록이 있습니다.\n기록을 새로 시작하면 화면의 이전 기록이 지워지고,\n새로운 파일에 기록됩니다.\n(저장된 엑셀 파일은 유지됩니다)",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if ret != QMessageBox.Yes:
                    return
            self._scheduled_dt = target
            dlg.accept()
            self._register_schedule_timer()
            ts = target.strftime('%Y-%m-%d %H:%M')
            self.act_schedule.setText(f"예약 취소 ({ts} 예약됨)")
            self.log(f"⏰ 기록 예약 완료 — {ts} 에 자동으로 기록을 시작합니다.")

        bb.accepted.connect(ok)
        bb.rejected.connect(dlg.reject)
        dlg.exec()

    def _register_schedule_timer(self):
# ──────────────────────────────────────────────────────────
# 【6-26】 _register_schedule_timer()
#   예약 시각을 1초 간격으로 감시하는 QTimer를 시작.
#   목표 시각(self._scheduled_dt)에 도달하면:
#   ① 타이머 중지 및 예약 변수 초기화
#   ② 메뉴 텍스트 원복
#   ③ 아직 기록 중이 아닐 때만 _start_monitoring() 호출
#      (사용자가 수동으로 먼저 시작한 경우 중복 실행 방지)
# ──────────────────────────────────────────────────────────
        t = QTimer(self)
        t.setInterval(1000)

        def _check():
            if self._scheduled_dt is None:
                t.stop()
                return
            if datetime.now() >= self._scheduled_dt:
                t.stop()
                self._schedule_timer = None
                self._scheduled_dt = None
                if not self.is_monitoring:
                    self.act_schedule.setText("예약 기록 시작")
                    if self.recorded_data:
                        self._clear_recorded_data()
                    self.log("⏰ 예약 시각이 되었습니다. 기록을 시작합니다.")
                    self._start_monitoring()
                else:
                    self.act_schedule.setText("예약 기록 중지")
                    self.log("⏰ 예약 시각이 되었으나 이미 기록 중이므로 시작 예약을 해제합니다.")

        t.timeout.connect(_check)
        t.start()
        self._schedule_timer = t

    def _ask_scheduled_stop(self):
# ──────────────────────────────────────────────────────────
# 【6-27】 _ask_scheduled_stop()
#   기록을 중지할 예약 시각(년·월·일·시·분)을 입력 받는 대화상자.
#   _ask_scheduled_start() 와 동일한 UI 구조.
#
#   검증 규칙:
#   - 숫자 이외 입력 또는 유효하지 않은 날짜이면 경고
#   - 현재 시각 이전이면 경고
#
#   OK 확인 후:
#   - self._stop_scheduled_dt 에 목표 datetime 저장
#   - _register_stop_schedule_timer() 호출 → 1초 폴링 타이머 시작
#   - 메뉴 텍스트를 "예약 중지 취소 (YYYY-MM-DD HH:MM 예약됨)"으로 변경
#   - 로그에 예약 시각 기록
# ──────────────────────────────────────────────────────────
        now = datetime.now()
        dlg = QDialog(self)
        dlg.setWindowTitle("예약 기록 중지")
        dlg.setFixedSize(340, 200)
        lo = QVBoxLayout(dlg)

        lbl_now = QLabel(f"현재 시각: {now.strftime('%Y-%m-%d %H:%M:%S')}")
        lo.addWidget(lbl_now)
        lo.addWidget(QLabel("기록을 중지할 시각을 입력하세요:"))
        _tick_stop = QTimer(dlg)
        _tick_stop.setInterval(1000)
        _tick_stop.timeout.connect(lambda: lbl_now.setText(f"현재 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"))
        _tick_stop.start()

        row = QHBoxLayout()
        ed_year  = QLineEdit(str(now.year))
        ed_month = QLineEdit(f"{now.month:02d}")
        ed_day   = QLineEdit(f"{now.day:02d}")
        ed_hour  = QLineEdit(f"{now.hour:02d}")
        ed_min   = QLineEdit(f"{now.minute:02d}")
        for ed in [ed_year, ed_month, ed_day, ed_hour, ed_min]:
            ed.setAlignment(Qt.AlignCenter)
        ed_year.setFixedWidth(52)
        ed_month.setFixedWidth(32)
        ed_day.setFixedWidth(32)
        ed_hour.setFixedWidth(32)
        ed_min.setFixedWidth(32)
        row.addWidget(ed_year);  row.addWidget(QLabel("년"))
        row.addWidget(ed_month); row.addWidget(QLabel("월"))
        row.addWidget(ed_day);   row.addWidget(QLabel("일"))
        row.addWidget(ed_hour);  row.addWidget(QLabel("시"))
        row.addWidget(ed_min);   row.addWidget(QLabel("분"))
        row.addStretch()
        lo.addLayout(row)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        lo.addWidget(bb)

        def ok():
            try:
                y  = int(ed_year.text())
                mo = int(ed_month.text())
                d  = int(ed_day.text())
                h  = int(ed_hour.text())
                mi = int(ed_min.text())
                target = datetime(y, mo, d, h, mi, 0)
            except ValueError:
                QMessageBox.warning(dlg, "알림", "올바른 날짜/시간을 입력해주세요.")
                return
            if target <= datetime.now():
                QMessageBox.warning(dlg, "알림", "현재 시각 이후로 설정해야 합니다.")
                return
            self._stop_scheduled_dt = target
            dlg.accept()
            self._register_stop_schedule_timer()
            ts = target.strftime('%Y-%m-%d %H:%M')
            self.act_schedule.setText(f"예약 중지 취소 ({ts} 예약됨)")
            self.log(f"⏰ 중지 예약 완료 — {ts} 에 자동으로 기록을 중지합니다.")

        bb.accepted.connect(ok)
        bb.rejected.connect(dlg.reject)
        dlg.exec()

    def _register_stop_schedule_timer(self):
# ──────────────────────────────────────────────────────────
# 【6-28】 _register_stop_schedule_timer()
#   중지 예약 시각을 1초 간격으로 감시하는 QTimer를 시작.
#   목표 시각(self._stop_scheduled_dt)에 도달하면:
#   ① 타이머 중지 및 예약 변수 초기화
#   ② 아직 기록 중일 때만 _stop_monitoring_silent() 호출
#      (사용자가 수동으로 먼저 중지한 경우 중복 실행 방지)
# ──────────────────────────────────────────────────────────
        t = QTimer(self)
        t.setInterval(1000)

        def _check():
            if self._stop_scheduled_dt is None:
                t.stop()
                return
            if datetime.now() >= self._stop_scheduled_dt:
                t.stop()
                self._stop_schedule_timer = None
                self._stop_scheduled_dt = None
                if self.is_monitoring:
                    self.log("⏰ 예약 시각이 되었습니다. 기록을 중지합니다.")
                    self._stop_monitoring_silent()
                else:
                    self.act_schedule.setText("예약 기록 시작")
                    self.log("⏰ 예약 시각이 되었으나 이미 기록이 중지되어 중지 예약을 해제합니다.")

        t.timeout.connect(_check)
        t.start()
        self._stop_schedule_timer = t

    def _stop_monitoring_silent(self):
# ──────────────────────────────────────────────────────────
# 【6-29】 _stop_monitoring_silent()
#   예약 중지 시각 도달 시 확인 다이얼로그 없이 자동으로 기록을 중지.
#   _stop_monitoring() 과 동일한 처리를 하되 QMessageBox 생략.
# ──────────────────────────────────────────────────────────
        self.is_monitoring = False
        self.act_toggle.setText("기록 시작")
        self.act_schedule.setText("예약 기록 시작")
        self.act_search.setEnabled(True)
        self._log_file_prefix.setVisible(False)
        self._log_file_label.setVisible(False)
        self._btn_open_file.setVisible(False)
        self._btn_open_folder.setVisible(False)
        for p in self.route_map_panels.values():
            p.pause_tick()
        if self.recorded_data and self.auto_save_path and self.can_auto_save and len(self.recorded_data) > self._saved_record_count:
            self._perform_auto_save()

# ──────────────────────────────────────────────────────────
# 【6-30】 _on_toggle()
#   [기록 시작]/[기록 중지] 메뉴 클릭 시 호출. 현재 상태에 따라 분기.
#
#   기록 중지 상태에서 클릭 시:
#   ① 예약 대기 중이면 팝업 →
#      [바로 시작]: 예약 취소 + 이전 기록 확인 후 시작
#      [취소]: 아무것도 안 함
#   ② 이전 기록이 있으면 팝업 →
#      [새로 시작]: 기록창 초기화 후 시작
#      [취소]: 아무것도 안 함
#   ③ 이전 기록 없으면 바로 시작
#
#   기록 중 상태에서 클릭 시:
#   → _stop_monitoring() 호출
# ──────────────────────────────────────────────────────────
    def _on_toggle(self):
        if self.is_monitoring:
            self._stop_monitoring()
        else:
            # ① 예약 대기 중이면 예약 취소 팝업
            if self._schedule_timer is not None:
                ret = QMessageBox.question(
                    self, "예약 확인",
                    "예약이 설정되어 있습니다.\n지금 바로 기록을 시작하면 예약이 취소됩니다.\n바로 시작하시겠습니까?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if ret != QMessageBox.Yes:
                    return
                self._schedule_timer.stop()
                self._schedule_timer = None
                self._scheduled_dt = None
                self.act_schedule.setText("예약 기록 시작")
                self.log("⏰ 예약이 취소되고 기록을 시작합니다.")
            # ② 이전 기록 있으면 초기화 팝업
            if self.recorded_data:
                ret = QMessageBox.question(
                    self, "이전 기록 있음",
                    "이전 기록이 있습니다.\n기록을 새로 시작하면 화면의 이전 기록이 지워지고,\n새로운 파일에 기록됩니다.\n(저장된 엑셀 파일은 유지됩니다)",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if ret != QMessageBox.Yes:
                    return
                self._clear_recorded_data()
            self._start_monitoring()

# ──────────────────────────────────────────────────────────
# 【6-31】 _start_monitoring()
#   모니터링 시작:
#   ① 이미 기록 중이면 중복 실행 방지 가드
#   ② 노선 등록 확인 → 자동 저장 파일 생성 (운행기록_YYYYMMDD_HHMMSS.xlsx)
#   ③ is_monitoring=True, 메뉴 "기록 중지"로 변경, [노선 검색] 비활성화
#   ④ resume_tick() 호출, _main_loop 데몬 스레드 시작
# ──────────────────────────────────────────────────────────
    def _start_monitoring(self):
        if self.is_monitoring:      # 중복 실행 방지 가드
            return
        if not self.routes:
            QMessageBox.warning(self, "알림", "불러온 노선이 없습니다. 먼저 노선을 검색하세요.")
            return
        fn = f"운행기록_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        self.auto_save_path = os.path.join(self.current_dir, fn)
        try:
            pd.DataFrame(columns=["데이터시각", "운행시작/종료", "정류소이름(번호)", "노선", "차량번호"]).to_excel(self.auto_save_path, index=False)
            self.can_auto_save = True
            self.log(f"자동 저장 파일: {self.auto_save_path}")
            self._log_file_label.set_full_text(self.auto_save_path)
            self._log_file_prefix.setVisible(True)
            self._log_file_label.setVisible(True)
            self._btn_open_file.setVisible(True)
            self._btn_open_folder.setVisible(True)
        except Exception as e:
            QMessageBox.warning(self, "오류", f"엑셀 파일 생성 실패: {e}")
            return
        self.is_monitoring = True
        self.act_toggle.setText("기록 중지")
        self.act_schedule.setText("예약 기록 중지")
        self.act_search.setEnabled(False)
        for p in self.route_map_panels.values():
            p.resume_tick()
        threading.Thread(target=self._main_loop, daemon=True).start()
        self.log(f"▶ 자동 기록을 시작합니다. (주기: {self.refresh_interval}초)")

# ──────────────────────────────────────────────────────────
# 【6-32】 _clear_recorded_data()
#   기록창(table_depart, table_arrive)과 recorded_data, _saved_record_count를 초기화.
#   기록 시작 전 이전 기록을 지울 때 호출.
# ──────────────────────────────────────────────────────────
    def _clear_recorded_data(self):
        self.recorded_data = []
        self._saved_record_count = 0
        self.table_depart.table.setRowCount(0)
        self.table_arrive.table.setRowCount(0)

# ──────────────────────────────────────────────────────────
# 【6-33】 _stop_monitoring()
#   모니터링 중지:
#   ① 중지 예약 중이면 별도 팝업 →
#      [바로 중지]: 중지 예약 취소 + 즉시 중지
#      [취소]: 아무것도 안 함
#   ② 중지 예약 없으면 일반 확인 팝업
#   ③ is_monitoring=False, 메뉴 원복, 파일 경로 라벨 숨김
#   ④ pause_tick() 호출, 미저장 기록 있으면 즉시 저장
# ──────────────────────────────────────────────────────────
    def _stop_monitoring(self):
        if self._stop_schedule_timer is not None:
            ret = QMessageBox.question(
                self, "중지 확인",
                "중지 예약이 설정되어 있습니다.\n지금 바로 중지하면 예약이 취소됩니다.\n바로 중지하시겠습니까?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret != QMessageBox.Yes:
                return
            self._stop_schedule_timer.stop()
            self._stop_schedule_timer = None
            self._stop_scheduled_dt = None
            self.log("⏰ 중지 예약이 취소되고 기록을 중지합니다.")
        else:
            if QMessageBox.question(self, "중지 확인", "정말 기록을 중지하시겠습니까?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
                return
            self.log("■ 기록을 중지합니다.")
        self.is_monitoring = False
        self.act_toggle.setText("기록 시작")
        self.act_schedule.setText("예약 기록 시작")
        self.act_search.setEnabled(True)
        self._log_file_prefix.setVisible(False)
        self._log_file_label.setVisible(False)
        self._btn_open_file.setVisible(False)
        self._btn_open_folder.setVisible(False)
        for p in self.route_map_panels.values():
            p.pause_tick()
        if self.recorded_data and self.auto_save_path and self.can_auto_save and len(self.recorded_data) > self._saved_record_count:
            self._perform_auto_save()

# ──────────────────────────────────────────────────────────
# 【6-34】 _main_loop()
#   백그라운드 데몬 스레드에서 실행되는 갱신 루프.
#   is_monitoring이 True인 동안 반복:
#   ① _refresh_data() 호출
#   ② 다음 실행 시각(nc) = 현재 + refresh_interval
#   ③ 남은 시간을 0.1초 단위로 나눠 sleep (중지 명령에 빠르게 반응하기 위해)
#   ④ API 호출이 갱신 주기를 초과하면 nc를 현재 시각으로 리셋 (지연 누적 방지)
# ──────────────────────────────────────────────────────────
    def _main_loop(self):
        nc = time.time()
        while self.is_monitoring:
            self._refresh_data()
            nc += self.refresh_interval
            st = nc - time.time()
            if st < 0:
                nc = time.time()
                st = 0
            for _ in range(int(st * 10)):
                if not self.is_monitoring:
                    break
                time.sleep(0.1)

# ──────────────────────────────────────────────────────────
# 【6-35】 _refresh_data()
#   한 번의 갱신 실행. _refresh_lock으로 동시 실행 방지.
#   (이미 갱신 중이면 blocking=False로 잠금 실패 → 즉시 리턴)
#   ① _process_routes() → ② 미저장 기록 있으면 _perform_auto_save()
# ──────────────────────────────────────────────────────────
    def _refresh_data(self):
        if not self._refresh_lock.acquire(blocking=False):
            return
        try:
            self._process_routes()
            if self.recorded_data and self.auto_save_path and self.can_auto_save and len(self.recorded_data) > self._saved_record_count:
                self._perform_auto_save()
        finally:
            self._refresh_lock.release()

# ──────────────────────────────────────────────────────────
# 【6-36】 _process_routes()  ★ 운행 기록 판정 핵심 로직 ★
#
#   ① 날짜 변경 감지: 자정 이후이면 통계 초기화·pos_suspend_until 초기화
#   ② 6시간 이상 지난 departed_vehicles 항목 자동 삭제
#   ③ 각 노선에 대해:
#      [출발 판정]
#      URL_POS1 호출 → 버스 없음(운행 종료 시간대)이면:
#        - 지도 버스 제거 (sig_clear_map)
#        - 첫차 시각까지 API 호출 중지 (pos_suspend_until 설정)
#      버스 있으면:
#        - lastStnId가 첫·두 번째 정류소이면 출발 기록 (_record)
#          (40분=2400초 내 중복 방지, departed_vehicles에 등록)
#        - URL_SLST로 구간 속도 수집 → 지도에 _sect_speeds 갱신
#        - sig_update_map으로 지도 버스 위치 갱신
#      [도착 판정]
#      URL_POS2 호출 → lastStnId가 종점이고 departed_vehicles에 있는 버스면:
#        - 도착 기록 (_record)
#        - departed_vehicles에서 해당 버스 삭제
# ──────────────────────────────────────────────────────────
    def _process_routes(self):
        _now = time.time()
        now_dt = datetime.now()
        today = now_dt.date()
        if today != self._last_date:
            self._last_date = today
            self.pos_suspend_until.clear()
            self.api_stats_yesterday = self.api_stats_today.copy()
            for k in self.api_stats_today:
                self.api_stats_today[k] = 0
            self._key_limit_notified = False
            self.log("📅 날짜가 바뀌었습니다. 통계를 초기화합니다.")
        for k in [k for k, ts in self.departed_vehicles.items() if _now - ts >= 21600]:
            del self.departed_vehicles[k]
        self.temp_pos1_data = {}
        self.temp_pos2_data = {}
        for route in self.routes:
            rnm = route['rnm']
            rid = route['rid']
            sc = route['st_cnt']
            if rid in self.pos_suspend_until:
                if now_dt < self.pos_suspend_until[rid]:
                    continue
                else:
                    del self.pos_suspend_until[rid]
                    self.pos_resume_logged.discard(rid)
            if rid in self.temp_pos1_data:
                rp1 = self.temp_pos1_data[rid]
            else:
                rp1 = self.fetch_api(URL_POS1, {'busRouteId': rid})
                if rp1 is not None and not isinstance(rp1, tuple):
                    self.temp_pos1_data[rid] = rp1
            if rp1 is None:
                continue
            p1b = not isinstance(rp1, tuple) and len(rp1.findall(".//itemList")) > 0
            if not p1b:
                ins = False
                ft = route.get('first_bus_tm')
                lt = route.get('last_bus_tm')
                if ft and lt:
                    try:
                        fh, fm = map(int, ft.split(":"))
                        lh, lm = map(int, lt.split(":"))
                        ps = now_dt.replace(hour=fh, minute=fm, second=0, microsecond=0) - timedelta(minutes=5)
                        pe = now_dt.replace(hour=lh, minute=lm, second=0, microsecond=0) + timedelta(minutes=30)
                        if pe < ps:
                            pe += timedelta(days=1)
                        ins = ps <= now_dt <= pe
                    except:
                        pass
                self.sig_clear_map.emit(rnm)
                if not ins:
                    fs = ft or self._fetch_first_time(rid)
                    if fs:
                        try:
                            fhm = datetime.strptime(fs, "%H:%M")
                            base = now_dt.replace(hour=fhm.hour, minute=fhm.minute, second=0, microsecond=0)
                            if base <= now_dt:
                                base += timedelta(days=1)
                            resume = base - timedelta(minutes=5)
                            self.pos_suspend_until[rid] = resume if resume > now_dt else now_dt + timedelta(minutes=1)
                            if resume > now_dt:
                                self.log(f"💤 {rnm} 운행 종료 → 첫차 {fs} 5분 전까지 POS 정지")
                        except:
                            self.pos_suspend_until[rid] = now_dt + timedelta(minutes=30)
                    else:
                        self.pos_suspend_until[rid] = now_dt + timedelta(minutes=30)
            if p1b:
                for bus in rp1.findall(".//itemList"):
                    ls = bus.findtext("lastStnId") or ""
                    vn = bus.findtext("plainNo") or ""
                    try:
                        stm = int(bus.findtext("lastStTm") or "0")
                    except:
                        stm = 0
                    if stm <= 0:
                        continue
                    if (rid, vn) not in self.departed_vehicles:
                        self.departed_vehicles[(rid, vn)] = _now
                    ifs = route['first_st_id'] and ls == route['first_st_id']
                    iss = route['second_st_id'] and ls == route['second_st_id']
                    if not (ifs or iss):
                        continue
                    k0 = (0, rid, vn)
                    if k0 not in self.last_arrival_logs or _now - self.last_arrival_logs[k0] >= 2400:
                        ft2 = format_datetm(bus.findtext("dataTm"))
                        if ifs:
                            dn, da = route['first_nm'], route['first_ars']
                            st = f"[{dn}({da}) 출발]"
                        else:
                            dn, da = route['second_nm'], route['second_ars']
                            st = f"[{dn}({da}) 출발 - 2번째 정류소 감지]"
                        self._record(0, ft2, rnm, vn, dn, da, st)
                        self.last_arrival_logs[k0] = _now
                        self.departed_vehicles[(rid, vn)] = _now

                sect_speeds = {}
                root_slst_spd = self.fetch_api(URL_SLST, {'busRouteId': rid})
                if root_slst_spd is not None and not isinstance(root_slst_spd, tuple):
                    for item in root_slst_spd.findall(".//itemList"):
                        seq_str = (item.findtext("seq") or "0").strip()
                        spd_str = (item.findtext("sectSpd") or "").strip()
                        try:
                            seq_val = int(seq_str)
                            spd_val = float(spd_str) if spd_str else -1
                            sect_speeds[seq_val] = spd_val
                        except ValueError:
                            pass

                if rnm in self.route_map_panels:
                    self.route_map_panels[rnm]._sect_speeds = sect_speeds

                bfm = []
                for item in rp1.findall(".//itemList"):
                    bfm.append({
                        "vehId": item.findtext("vehId") or "",
                        "plainNo": item.findtext("plainNo") or "",
                        "busType": item.findtext("busType") or "",
                        "lastStnId": item.findtext("lastStnId") or "",
                        "sectDist": item.findtext("sectDist") or "0",
                        "islastyn": item.findtext("islastyn") or "",
                        "lastStTm": item.findtext("lastStTm") or "0",
                        "congestion": item.findtext("congetion") or "0"
                    })
                self.sig_update_map.emit(rnm, bfm)

            if rid in self.temp_pos2_data:
                rp2 = self.temp_pos2_data[rid]
            else:
                r2 = self.fetch_api(URL_POS2, {'busRouteId': rid, 'startOrd': '1', 'endOrd': str(sc)})
                if isinstance(r2, tuple) and r2[0] == 'NO_BUS':
                    continue
                rp2 = r2
                if rp2 is not None:
                    self.temp_pos2_data[rid] = rp2
            if rp2 is None:
                continue
            for bus in rp2.findall(".//itemList"):
                ls = bus.findtext("lastStnId") or ""
                vn = bus.findtext("plainNo") or ""
                if not (route['last_st_id'] and ls == route['last_st_id']):
                    continue
                if (rid, vn) not in self.departed_vehicles:
                    continue
                k1 = (1, rid, vn)
                if k1 not in self.last_arrival_logs or _now - self.last_arrival_logs[k1] >= 2400:
                    ft2 = format_datetm(bus.findtext("dataTm"))
                    st = f"[{route['last_nm']}({route['last_ars']}) 도착]"
                    self._record(1, ft2, rnm, vn, route['last_nm'], route['last_ars'], st)
                    self.last_arrival_logs[k1] = _now
                    del self.departed_vehicles[(rid, vn)]

# ──────────────────────────────────────────────────────────
# 【6-37】 _record(idx, ft, rnm, vn, sn, sa, status)
#   운행 이벤트(출발/도착)를 기록하는 함수.
#   ① recorded_data에 (시각, "운행시작"/"운행종료", "정류소(ARS)", 노선, 차번) 추가
#   ② _perform_auto_save() 즉시 저장
#   ③ log() 로그 출력
#   ④ sig_record 시그널 발행 → UI 표에 반영
# ──────────────────────────────────────────────────────────
    def _record(self, idx, ft, rnm, vn, sn, sa, status):
        op = "운행시작" if idx == 0 else "운행종료"
        self.recorded_data.append((ft, op, f"{sn} ({sa})", rnm, vn))
        self._perform_auto_save()
        self.log(f"★ {rnm} {vn} → {status}")
        self.sig_record.emit(idx, (ft, rnm, vn, status))

# ──────────────────────────────────────────────────────────
# 【6-38】 _perform_auto_save()
#   _save_lock으로 동시 저장 방지 후 _core_excel_save() 호출.
#   저장 조건 모두 충족해야 실행: recorded_data 비어있지 않음 +
#   auto_save_path 설정됨 + can_auto_save=True
# ──────────────────────────────────────────────────────────
    def _perform_auto_save(self):
        if not self.recorded_data or not self.auto_save_path or not self.can_auto_save:
            return
        if not self._save_lock.acquire(blocking=False):
            return
        try:
            self._core_excel_save(self.auto_save_path, True)
        finally:
            self._save_lock.release()

# ──────────────────────────────────────────────────────────
# 【6-39】 _core_excel_save(tp, sc=False)
#   운행 기록을 날짜별 시트로 분리하여 엑셀 저장.
#
#   날짜 분리 기준: 새벽 3시 이전(0~2시)은 전날 운행으로 간주
#   (심야 버스 등 자정 이후 운행을 같은 영업일로 처리)
#
#   처리 순서:
#   ① DataFrame 생성 → gbd()로 각 데이터의 영업일(BizDate) 계산
#   ② 현재 날짜(cb) 결정 (새벽 3시 이전이면 전날)
#   ③ 완결 날짜(cd) = 오늘·Unknown 제외한 모든 날짜
#   ④ ExcelWriter로 날짜별 시트에 저장 + _axs()로 스타일 적용
#   ⑤ sc=True이면 완결 날짜의 별도 완료 파일 생성
#      → "운행기록_YYYYMMDD_완료.xlsx" (_completed_dates_saved로 중복 방지)
#   ⑥ _saved_record_count 갱신
# ──────────────────────────────────────────────────────────
    def _core_excel_save(self, tp, sc=False):
        try:
            cols = ["데이터시각", "운행시작/종료", "정류소이름(번호)", "노선", "차량번호"]
            df = pd.DataFrame(self.recorded_data, columns=cols)
            def gbd(ds):
                try:
                    d = datetime.strptime(ds, "%Y-%m-%d %H:%M:%S")
                    if d.hour < 3:
                        d -= timedelta(days=1)
                    return d.strftime("%Y-%m-%d")
                except:
                    return "Unknown"
            df['BizDate'] = df['데이터시각'].apply(gbd)
            now = datetime.now()
            cb = (now - timedelta(days=1)).strftime("%Y-%m-%d") if now.hour < 3 else now.strftime("%Y-%m-%d")
            cd = set(df['BizDate'].unique()) - {cb, "Unknown"}
            with pd.ExcelWriter(tp, engine='openpyxl') as w:
                for bd, g in df.groupby('BizDate'):
                    sd = g.drop(columns=['BizDate'])
                    sd.to_excel(w, sheet_name=bd, index=False)
                    self._axs(w.sheets[bd], sd)
                self._write_source_sheet(w.book)
            if sc:
                for bd in sorted(cd - self._completed_dates_saved):
                    dd = df[df['BizDate'] == bd].drop(columns=['BizDate'])
                    if dd.empty:
                        continue
                    sd2 = bd.replace("-", "")
                    cp = os.path.join(self.current_dir, f"운행기록_{sd2}_완료.xlsx")
                    with pd.ExcelWriter(cp, engine='openpyxl') as cw:
                        dd.to_excel(cw, sheet_name=bd, index=False)
                        self._axs(cw.sheets[bd], dd)
                        self._write_source_sheet(cw.book)
                    self._completed_dates_saved.add(bd)
                    self.log(f"📁 완결 파일 저장: 운행기록_{sd2}_완료.xlsx")
            self._saved_record_count = len(self.recorded_data)
        except PermissionError:
            self.log("⚠ 엑셀 파일이 열려 있어 저장을 건너뜁니다.")
        except Exception as e:
            self.log(f"❌ 저장 오류: {e}")

# ──────────────────────────────────────────────────────────
# 【6-40】 _write_source_sheet(wb)
#   엑셀 산출물(운행기록/완료 파일 공통)에 "출처 및 수집방법" 시트를
#   맨 앞(인덱스 0)에 추가. 공공누리 제1유형 출처표시 + 수집 프로그램/
#   소스코드 링크 + 실시간 데이터 유의사항을 담는다. _core_excel_save()에서
#   ExcelWriter의 with 블록 안, 날짜별 시트 작성 직후에 호출된다.
#   데이터셋 이름 아래 행에는 공공데이터포털 상세 페이지 URL을 하이퍼링크로 둔다.
# ──────────────────────────────────────────────────────────
    def _write_source_sheet(self, wb):
        ws = wb.create_sheet("출처 및 수집방법", 0)
        rows = [
            ("■ 출처 표시", ""),
            ("출처", "서울특별시(미래첨단교통과)"),
            ("제공처", "공공데이터포털 https://www.data.go.kr"),
            ("이용허락범위", "공공누리 제1유형(출처표시) / 제3자 권리 포함 : 저작권 표시"),
            ("이용 데이터셋", "서울특별시_버스위치정보조회 서비스 (15000332)"),
            ("", "https://www.data.go.kr/data/15000332/openapi.do"),
            ("", "서울특별시_버스도착정보조회 서비스 (15000314)"),
            ("", "https://www.data.go.kr/data/15000314/openapi.do"),
            ("", "서울특별시_정류소정보조회 서비스 (15000303)"),
            ("", "https://www.data.go.kr/data/15000303/openapi.do"),
            ("", "서울특별시_노선정보조회 서비스 (15000193)"),
            ("", "https://www.data.go.kr/data/15000193/openapi.do"),
            ("", ""),
            ("■ 수집 방법", ""),
            ("수집 프로그램", f"Seoul_Bus_Drive_Recorder v{APP_VERSION}"),
            ("소스 코드", "https://github.com/metapbl/BusR"),
            ("", ""),
            ("■ 유의사항", ""),
            ("", "실시간 API 응답을 기록한 것으로 통신 상태에 따라 누락·오차가 있을 수 있습니다."),
            ("", "본 자료를 외부에 제공·공표할 때에는 위 출처를 함께 표기하시기 바랍니다."),
        ]
        for r, (a, b) in enumerate(rows, start=1):
            ws.cell(row=r, column=1, value=a)
            cb = ws.cell(row=r, column=2, value=b)
            if a.startswith("■"):
                ws.cell(row=r, column=1).font = XlFont(bold=True)
            if isinstance(b, str) and b.startswith("http"):
                cb.hyperlink = b
                cb.font = XlFont(color="0563C1", underline="single")
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 78

# ──────────────────────────────────────────────────────────
# 【6-41】 _axs(ws, df)
#   엑셀 워크시트에 스타일 적용.
#   헤더 행(1행): 연파랑 배경(#DDEBF7) + 굵은 글자(11pt) + 가운데 정렬
#   컬럼 너비: max(데이터 최대 글자 수, 헤더 글자 수) + 컬럼별 가산값
#     데이터시각(+1), 운행시작/종료(+7), 정류소이름(번호)(+15), 노선(+8), 기타(+5)
# ──────────────────────────────────────────────────────────
    def _axs(self, ws, df):
        hf = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        hft = XlFont(bold=True, size=11)
        ha = Alignment(horizontal='center', vertical='center')
        cw = {
            "데이터시각": lambda b: b + 1,
            "운행시작/종료": lambda b: b + 7,
            "정류소이름(번호)": lambda b: b + 15,
            "노선": lambda b: b + 8
        }
        for ci, col in enumerate(df.columns):
            c = ws.cell(row=1, column=ci + 1)
            c.fill = hf
            c.font = hft
            c.alignment = ha
            ml = df[col].astype(str).map(len).max()
            bw = max(ml, len(str(col)))
            ws.column_dimensions[c.column_letter].width = cw.get(col, lambda b: b + 5)(bw)
            

# ──────────────────────────────────────────────────────────
# 【6-42】 _ask_interval()
#   갱신 주기(10초 이상)를 QLineEdit으로 입력 받는 대화상자.
# ──────────────────────────────────────────────────────────
    def _ask_interval(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("갱신주기 설정")
        dlg.setFixedSize(300, 140)
        lo = QVBoxLayout(dlg)
        lo.addWidget(QLabel("갱신주기(초)를 입력하세요 (10초 이상):"))
        ed = QLineEdit(str(self.refresh_interval))
        ed.setAlignment(Qt.AlignCenter)
        lo.addWidget(ed)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        lo.addWidget(bb)
        def ok():
            try:
                v = int(ed.text())
            except:
                QMessageBox.warning(dlg, "알림", "숫자를 입력해주세요.")
                return
            if v < 10:
                QMessageBox.warning(dlg, "알림", "갱신주기는 10초 이상이어야 합니다.")
                return
            self.refresh_interval = v
            self.log(f"갱신주기가 {v}초로 변경되었습니다.")
            dlg.accept()
        bb.accepted.connect(ok)
        bb.rejected.connect(dlg.reject)
        ed.returnPressed.connect(ok)
        dlg.exec()

# ──────────────────────────────────────────────────────────
# 【6-43】 _show_api_status()
#   메뉴 [API 현황] 클릭 시 호출. API별 오늘/어제 호출 횟수를 표로 표시.
#   setMinimumSectionSize(60)으로 컬럼이 너무 좁아지지 않도록 보호.
# ──────────────────────────────────────────────────────────
    def _show_api_status(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("API 현황")
        lo = QVBoxLayout(dlg)
        t = QTableWidget(6, 4)
        t.setHorizontalHeaderLabels(["구분", "API URL", "오늘", "어제"])
        hdr = t.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.Fixed)
        hdr.setSectionResizeMode(3, QHeaderView.Fixed)
        t.setColumnWidth(0, 145)
        t.setColumnWidth(2, 75)
        t.setColumnWidth(3, 75)
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.verticalHeader().setVisible(False)
        data = [
            ("운행 출발 판정 API", "POS1", URL_POS1),
            ("종점 도착 판정 API", "POS2", URL_POS2),
            ("노선 정류소 목록 API", "SLST", URL_SLST),
            ("버스 노선 정보 API", "RINF", URL_RINF),
            ("버스 노선 검색 API", "SRCH", URL_SRCH),
            ("기타", "기타", "기타"),
        ]
        for row, (lb, ky, url) in enumerate(data):
            t.setItem(row, 0, QTableWidgetItem(lb))
            t.setItem(row, 1, QTableWidgetItem(url))
            t.setItem(row, 2, QTableWidgetItem(f"{self.api_stats_today.get(ky, 0):,}회"))
            t.setItem(row, 3, QTableWidgetItem(f"{self.api_stats_yesterday.get(ky, 0):,}회"))
        t.horizontalHeader().setMinimumSectionSize(60)
        lo.addWidget(t)
        row_h = t.verticalHeader().defaultSectionSize()
        header_h = t.horizontalHeader().height()
        total_h = header_h + row_h * 6 + 10
        t.setFixedHeight(total_h)
        dlg.resize(750, total_h + 50)
        dlg.exec()

# ──────────────────────────────────────────────────────────
# 【6-44】 _show_program_info()
#   메뉴 [프로그램 정보] 클릭 시 호출.
#   ① 저작권/오픈소스(MIT, Qt·PySide6 LGPL v3) 고지
#   ② 공공데이터 4개 데이터셋 출처표시(공공누리 제1유형) + 공공누리 마크
#   ③ 이용 안내(행정심판 2022-21122 재결, 목적/금지사항, 인증키 암호화 한계 등)
#   ④ [라이선스 전문 폴더 열기] 버튼 → _open_licenses_dir()로 licenses 폴더 오픈
# ──────────────────────────────────────────────────────────
    def _show_program_info(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("프로그램 정보")
        dlg.resize(680, 840)

        # ① 제목 · 저작권 헤더
        header = (f"<h3 style='margin-bottom:2px'>서울시내버스 노선 운행기록 수집 프로그램 "
                  f"v{APP_VERSION}</h3>"
                  f"<p>Copyright &copy; 2026 박국환 "
                  f"(<a href='mailto:ggoyong2@naver.com'>ggoyong2@naver.com</a>)<br>"
                  f"저장소 : <a href='https://github.com/metapbl/BusR'>https://github.com/metapbl/BusR</a><br>"
                  f"이 프로그램은 <b>MIT 라이선스</b>로 배포됩니다.</p>")

        # ② 이용 안내
        notice = ("<p><b>■ 이용 안내</b></p>"
                  "<ul style='margin-top:0'>"
                  "<li>수집 정보는 서울특별시가 공개한 버스 운행 정보이며, 중앙행정심판위원회 "
                  "2022-21122(2023-05-16) 재결은 특정 버스의 위치정보를 특정 개인의 위치정보로 "
                  "볼 수 없다고 판단하였습니다.</li>"
                  "<li>이 프로그램은 법정 휴게시간 등 근로조건 점검을 목적으로 하며, 개별 운전자에 "
                  "대한 감시 목적의 사용을 금지합니다.</li>"
                  "<li>실시간 API 응답을 기록한 것으로 통신 상태에 따라 누락·오차가 있을 수 있습니다.</li>"
                  "<li>수집 결과를 외부에 제공·공표할 때에는 아래 출처표시를 함께 표기하시기 바랍니다.</li>"
                  "<li>입력하신 인증키는 설정 파일에 평문으로 저장되지 않고 암호화되어 보관됩니다. "
                  "다만 복호화에 필요한 정보가 프로그램에 포함되어 있으므로, 소스 코드를 분석할 수 "
                  "있는 사람에 대한 방어 수단은 아닙니다. 설정 파일을 타인과 공유하지 마십시오.</li>"
                  "<li>인증키 관리 및 API 이용약관 준수 책임은 이용자에게 있습니다.</li>"
                  "</ul>")

        # ③ 오픈소스 고지
        oss = ("<p><b>■ 오픈소스 고지</b></p>"
               "<ul style='margin-top:0'>"
               "<li>이 프로그램은 Qt 및 PySide6를 <b>GNU LGPL v3</b>에 따라 사용합니다.</li>"
               "<li>Qt 소스 : <a href='https://download.qt.io'>https://download.qt.io</a></li>"
               "<li>PySide6 소스 : "
               "<a href='https://code.qt.io/cgit/pyside/pyside-setup.git'>https://code.qt.io/cgit/pyside/pyside-setup.git</a></li>"
               "<li>그 밖에 requests, openpyxl 등을 사용하며, 전체 목록과 라이선스 전문은 "
               "아래 [라이선스 전문 폴더 열기] 버튼으로 확인할 수 있습니다.</li>"
               "<li>아이콘 : 박국환 직접 제작</li>"
               "</ul>")

        # ④ 공공데이터 출처표시
        opendata = ("<p><b>■ 공공데이터 출처표시</b></p>"
                    "<ul style='margin-top:0'>"
                    "<li>본 프로그램은 공공누리 제1유형(출처표시)에 따라 서울특별시(미래첨단교통과)가 "
                    "개방한 다음 공공저작물을 이용하였습니다."
                    "<br>&nbsp;&nbsp;· <a href='https://www.data.go.kr/data/15000332/openapi.do'>"
                    "서울특별시_버스위치정보조회 서비스</a>"
                    "<br>&nbsp;&nbsp;· <a href='https://www.data.go.kr/data/15000314/openapi.do'>"
                    "서울특별시_버스도착정보조회 서비스</a>"
                    "<br>&nbsp;&nbsp;· <a href='https://www.data.go.kr/data/15000303/openapi.do'>"
                    "서울특별시_정류소정보조회 서비스</a>"
                    "<br>&nbsp;&nbsp;· <a href='https://www.data.go.kr/data/15000193/openapi.do'>"
                    "서울특별시_노선정보조회 서비스</a></li>"
                    "<li>해당 데이터는 공공데이터포털 "
                    "(<a href='https://www.data.go.kr'>https://www.data.go.kr</a>)에서 "
                    "무료로 이용하실 수 있습니다.</li>"
                    "<li>위 데이터에는 제3자 권리가 포함되어 있어 저작권 표시가 필요합니다. "
                    "(<a href='https://ccl.cckorea.org/about/'>제3자 권리 포함 : 저작권 표시</a>)</li>"
                    "</ul>")

        inner = QWidget()
        v = QVBoxLayout(inner)

        def _add_label(text):
            lb = QLabel(text)
            lb.setWordWrap(True)
            lb.setOpenExternalLinks(True)
            lb.setTextFormat(Qt.RichText)
            v.addWidget(lb)
            return lb

        _add_label(header)
        v.addSpacing(18)
        _add_label(notice)
        v.addSpacing(18)
        _add_label(oss)
        v.addSpacing(18)
        _add_label(opendata)

        # 공공누리 제1유형 마크 — 내장 base64, 원본 크기 그대로(변형·scaled() 금지)
        mark = QLabel()
        pm = load_pixmap_from_b64(GG_IMG_B64)
        if not pm.isNull():
            mark.setPixmap(pm)
        mark.setAlignment(Qt.AlignHCenter)
        v.addWidget(mark)

        v.addStretch(1)

        sa = QScrollArea()
        sa.setWidgetResizable(True)
        sa.setWidget(inner)

        btn_lic = QPushButton("라이선스 전문 폴더 열기")
        btn_lic.clicked.connect(lambda: _open_licenses_dir(dlg))
        btn_close = QPushButton("닫기")
        btn_close.clicked.connect(dlg.accept)

        h = QHBoxLayout()
        h.addWidget(btn_lic)
        h.addStretch(1)
        h.addWidget(btn_close)

        lay = QVBoxLayout(dlg)
        lay.addWidget(sa)
        lay.addLayout(h)
        dlg.exec()

# ──────────────────────────────────────────────────────────
# 【6-45】 closeEvent(event)
#   창 닫기(X 버튼 또는 종료 메뉴) 시 Qt가 자동 호출.
#   ① 종료 확인 메시지 → 아니오이면 event.ignore()로 취소
#   ② is_monitoring=False 설정 (갱신 루프 중지)
#   ③ 미저장 기록 있으면 _perform_auto_save()로 마지막 저장
#   ④ event.accept()로 창 닫기 허용
# ──────────────────────────────────────────────────────────
    def closeEvent(self, event):
        if QMessageBox.question(self, "종료 확인", "프로그램을 종료하시겠습니까?") != QMessageBox.Yes:
            event.ignore()
            return
        self.is_monitoring = False
        if self.recorded_data and self.auto_save_path and self.can_auto_save and len(self.recorded_data) > self._saved_record_count:
            self._perform_auto_save()
        event.accept()

# ──────────────────────────────────────────────────────────
# 【6-46】 _load_favorites() → list
#   설정 파일(Seoul_Bus_Config.ini)의 [favorites] 섹션에서
#   즐겨찾기 노선 목록을 JSON으로 읽어 리스트로 반환.
#   파싱 실패 시 빈 리스트 반환.
#
# 【6-47】 _save_favorites(favs)
#   즐겨찾기 리스트를 JSON으로 직렬화하여 설정 파일에 저장.
#   동시에 현재 API 키도 갱신 (favorites 저장 시 keys도 함께 기록).
# ──────────────────────────────────────────────────────────
    def _load_favorites(self):
        cfg = configparser.ConfigParser()
        if os.path.exists(self._CFG_FILE):
            cfg.read(self._CFG_FILE, encoding="utf-8")
        raw = cfg.get("favorites", "data", fallback="[]")
        try:
            return json.loads(raw)
        except:
            return []

    def _save_favorites(self, favs):
        cfg = configparser.ConfigParser()
        if os.path.exists(self._CFG_FILE):
            cfg.read(self._CFG_FILE, encoding="utf-8")
        if "keys" not in cfg:
            cfg["keys"] = {}
        cfg["keys"]["main_key"] = self._enc_key(self.api_key_main)
        cfg["keys"]["back_key"] = self._enc_key(self.api_key_back)
        if "favorites" not in cfg:
            cfg["favorites"] = {}
        cfg["favorites"]["data"] = json.dumps(favs, ensure_ascii=False)
        try:
            with open(self._CFG_FILE, "w", encoding="utf-8") as f:
                cfg.write(f)
        except:
            pass

# ──────────────────────────────────────────────────────────
# 【6-48】 _cleanup_search_signal(handler)
#   sig_search_done 시그널에서 특정 핸들러(슬롯)의 연결을 안전하게 해제.
#   대화상자가 닫힌 후 더 이상 필요 없는 슬롯이 계속 연결되어
#   의도치 않게 호출되는 것을 방지.
# ──────────────────────────────────────────────────────────
    def _cleanup_search_signal(self, handler):
        try:
            self.sig_search_done.disconnect(handler)
        except:
            pass

# ══════════════════════════════════════════════════════════
# 【7】 프로그램 진입점
#   파이썬 파일을 직접 실행했을 때만 이 블록이 실행됨.
#   다른 파일에서 import할 때는 실행되지 않음.
# ══════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    mode = "dark" if detect_os_dark_mode() else "light"
    app.setPalette(_make_palette(mode))
    window = SeoulBusRecorder()
    window.show()
    sys.exit(app.exec())