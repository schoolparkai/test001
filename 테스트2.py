def levenshtein_distance(s1, s2):
    # 두 문자열의 길이 계산
    m, n = len(s1), len(s2)
    # 거리를 저장할 행렬 초기화
    d = [[0] * (n+1) for _ in range(m+1)]
    # 초기값 설정
    for i in range(1, m+1):
        d[i][0] = i
    for j in range(1, n+1):
        d[0][j] = j
    # 행렬 채우기
    for j in range(1, n+1):
        for i in range(1, m+1):
            if s1[i-1] == s2[j-1]:
                d[i][j] = d[i-1][j-1]
            else:
                d[i][j] = min(d[i-1][j], d[i][j-1], d[i-1][j-1]) + 1
    # 결과 반환
    return d[m][n]

# 두 문자열 정의
text1 = "워정워정하며 세우러이 거의로다"
text2 = "흐롱흐롱하며 일운 일이 무사일고"

# 두 문자열을 음절 단위로 쪼갬
chars1 = list(text1)
chars2 = list(text2)

# 루벤슈타인 거리 계산
distance = levenshtein_distance(chars1, chars2)

# 두 문자열의 길이 계산
length1 = len(chars1)
length2 = len(chars2)

# 유사도 계산
similarity = (1 - distance / max(length1, length2)) * 100

# 결과 출력
print(f"두 문자열의 음절 유사도: {similarity:.2f}%")
