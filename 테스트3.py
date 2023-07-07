import difflib

text1 = "워정워정하며 세우러이 거의로다"
text2 = "흐롱흐롱하며 일운 일이 무사일고"
text3 = "아희로 전세계를 코딩하자!"

dist1 = difflib.SequenceMatcher(None, text1, text2).ratio()
dist2 = difflib.SequenceMatcher(None, text1, text3).ratio()
dist3 = difflib.SequenceMatcher(None, text2, text3).ratio()

print(f"Text1과 Text2 간의 유사도: {dist1}")
print(f"Text1과 Text3 간의 유사도: {dist2}")
print(f"Text2와 Text3 간의 유사도: {dist3}")
