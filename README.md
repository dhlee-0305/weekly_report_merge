# 파이썬 설치
윈도우용 Python 3.11.7 버전 설치

# 작업 폴더 및 설정
프로그램은 실행 시 `weekly_merge.ini`의 경로를 기준으로 작업 폴더를 자동 생성합니다.
기본 설정값은 아래와 같습니다.

| 설명                | 경로                               |
| ------------------- | ---------------------------------- |
| 작업 폴더            | C:\Downloads\weekly_report         |
| 취합 파일 생성 폴더   | C:\Downloads\weekly_report\output\ |
| 로그 폴더            | C:\Downloads\weekly_report\log\\   |



# 설정 파일 및 템플릿 준비

\weekly_report\output\ 폴더에 아래 2개 파일 복사합니다.

* weekly_merge.ini : 설정 파일
* 경영전략회의_template.docx : 템플릿 파일



# 라이브러리 설치

파이썬 설치 후 아래 명령을 실행합니다.
\> pip install python-docx
\> pip install pyinstaller



# 실행 방법

\> python weekly_merge.py



# 실행 파일 생성

단일 실행 파일 생성 시 아래 명령을 사용합니다.
\> pyinstaller --onefile weekly_merge.py



# 결과 파일

주간보고 취합 결과는 `weekly_merge.ini`의 `OUTPUT_PATH` 경로에 저장됩니다.
로그 파일은 `LOG_FILE_PATH` 경로에 날짜별로 생성됩니다.



# 주요 설정 항목

`REFERENCE_DAYOFWEEK`: 보고 기준 요일
`LOG_LEVEL`: 로그 레벨
`MERGE_SECTION`: 섹션별 취합 여부
