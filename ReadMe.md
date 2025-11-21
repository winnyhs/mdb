MDB → JSON → MDB 변환 도구

1. MDB → JSON 변환
   python main.py read

   output_json/
     Table1.json
     Table2.json
     ...

2. JSON → MDB 역삽입
   python main.py write

주의:
- Python 32bit 필수
- pywin32 설치 필수
- DAO 3.5/3.6/4.0 자동 감지
- Decimal, Datetime 안전 변환 포함

- pip freeze > requirements.txt
  $ cat requirements
   comtypes==1.4.12
   pillow==10.4.0
   psutil==5.9.5
   pywin32==311
   pywinauto==0.6.9
   six==1.17.0
