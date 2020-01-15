# Schema-Exporter
Schema-Exporter는 MariaD의 Information Schema를 기준으로 테이블 정보를 추출 하여 테이블 명세서를 작성해주는 간단한 툴입니다. 

## Import Module
+ tkinter
+ mysql.connector
+ openpyxl

## Issue
+ DB 접속정보 입력시 Password가 평문으로 입력됨
+ 스키마 추출 시 추출이 완료되기 까지 응답없음 발생 
+ 스키마 리스트 박스의 스크롤 미적용

