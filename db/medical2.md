Tables

| table name | columns | row 개수 | 설명 |
|------------|---------|----------|------|
| BODY | ID, fDAY, tDAY, fTIME, tTIME, TYPE, MEMO | 24 | 시간별로 8BODY와 숫자가 기록됨 | 
| HISTORY | ID, NAME, SEX, BODY, ADD, TEL, 4BODY, 8BODY, TIME, MEMO | 18723 | 사용자 정보 |
| M_INI      | ITEM, VALUE | 1 | 프로그램 경로, C:\PROGRAM FILES\MICROSOFT OFFICE\OFFICE\ |
| M_USER | USERID, PSWD, CMD0, ... | 3 | 사용자와 권한 map | 
| M_USER1 | USERID, PSWD, CMD0, ... | 3 | 사용자와 권한 map | 
| MAN | MAN, DESP | 4 | 소양인, 소음인, 태양인, 태음인 | 
| POINT | ID, P1, P2, P3, P4 |  24 | 모든 값이 0 | 
| POSITION | ID, CASE, ITEM, fTIME, tTIME, POST1, POST2, POST3, POST4, NAME, MEMO | 23 | 혈자리 시침정보? |
| QA | ID, T, Q1, P | 설문지 | 
| TYPE | TYPE, FILE,EMO | 비어 있음 |
| Y2E | ID, YANG, EUM, RUN |  양력 -> 음력 변환이 들어있다 |