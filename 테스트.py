import re

text1 = "ㅁㄴㄹㄷ ㅁㄴㄹㄷㄷ ㅁㄴㄷㄹ"
text2 = "ㅂㄷㄱㅈ ㅈㄷㄷㄷ ㅂㅌㅇㄷ"

# 정규표현식 패턴
pattern = '[ㄱ-ㅎㅏ-ㅣ가-힣]+'

# 각 문장에서 패턴에 매칭되는 문자열 추출
matches1 = re.findall(pattern, text1)
matches2 = re.findall(pattern, text2)

# 두 문자열에서 공통으로 포함된 문자열 찾기
common_matches = list(set(matches1) & set(matches2))

# 결과 출력

for match in common_matches:
    print(match.encode('utf-8').decode('cp949'))