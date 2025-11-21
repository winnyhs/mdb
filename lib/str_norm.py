import unicodedata
import re

def norm(s, len):
    """Normalize Korean/English text for matching while preserving () [] {}."""
    if not s:
        return ""

    # 1) Unicode Normalize
    s = unicodedata.normalize("NFKC", s)

    # 2) Lowercase (영문만 영향)
    s = s.lower()

    # 3) Strip leading/trailing spaces
    s = s.strip()

    # 4) Remove ALL white spaces (space/tabs/newlines)
    s = re.sub(r"[ \t\r\n]+", "", s)

    # 5) Remove punctuation EXCEPT (), [], {}, and ,
    #    we remove: . , ; : ! ? ' " - _ / \ |
    s = re.sub(r"[.;:!?'\"/_\\\-|]+", "", s)

    return s[:len]

# TODO: 그냥 matching 하는 것과 비교해 보기. 맨 앞 20글자만 matching 했을 때와도 비교하기
# 개수를 봐서 차이가 있는지 없는지 확인하기 
# TODO: table a, b, c, M_LIST, M_DATA 상호 비교해서 차이 정리하기 

if __name__ == "__main__": 
    string = 'abc  (def)    ghi\njkl\t \r (ttt)   [xxx]  '
    print(f"norm({string}) ===> {norm(string)}")