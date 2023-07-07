#!/usr/bin/python 
# -*- coding: utf-8 -*-
# 검색 결과를 엑셀로 변환하여 저장
## 작품번호를 list로 받아, 딕셔너리를 불러 필요한 정보를 가져옴 
## 작품번호, 제목, 작자, 창작연대, 검색된 행 등을 Excel로 저장하여 제공토록 함


import os, re
import glob, time
import sys, codecs
import 말뭉치_함수RE
import Gasa_Dic
from openpyxl import Workbook

global fff, words, i, count_item, grab, startend, howmany, grabb, each_line1

# ㅇㅡㅣ ㄴ.{1,2}ㅣㄷㅏㄹㅡㄴㅣ" # 성공 "ㄴᆞㄴ\s\S.{1,15}ㅇㅏ\s"    /ㅇㅔㅅㅕ.{1,9}ㄹ\s?ㅅㅗㄴㅑ/
words = "ㄱㅓㄹ[ㅐ|ㆎ]ㅎ[ㅏ|ᆞ]ㅇ[ㅑ|ㅕ]"
words_han ="_by_거래하야"
# 정규표현식 문법 ?(앞 문자가 0개 또는 1개)/ .(최소 한개의 문자/공백포함)/ {n,m}(n번 이상 m번 이하 반복: dot와 조합시켜야)/ \s(스페이스) / \S(스페이스 아닌 거) / (그룹){반복횟수}
# 한자로 된 작품은 한음절 크기가 1임에 주의할 것. {범위} 입력 시 1부터 입력할 것.
# \s -> 그냥 띄어쓰기도 됨  #아래아 + ㅣ = 한글자(ᆡ) / .{1,2} 이런 형태로 검색해야 함.
# 띄어쓰기 유무에 주의 -> 선택적으로 검색 : \s?
### 의 다르니  ㅇㅡㅣ ㄴᆡㄷㅏㄹㅡㄴㅣ
# "(^ㅇㅣ.{1,15}ㄴ\s.{2,40}\sㅈㅕ.{1,15}ㄴ\s)|(\sㅇㅣ.{1,15}ㄴ\s.{2,40}\sㅈㅕ.{1,15}ㄴ\s)" # 성공(그룹으로 선택)
# "\S.{4,11}ㅇㅣ\s\S.{4,11}ㄱㅗ\S.{4,11}ㅇㅣ\s\S.{4,11}ᆡ" # 성공
# "ㄴᆞㄴ\s\S{0,1}ㅈㅓ\s\S.{1,8}ㅇㅏ\s" # 성공
#"ㄴᆞㄴ\s\S{0,1}[ㅈㅓ |ㄷㅕ |ㅈㅕ ]\s\S.{1,8}[ㅇㅏ |ㅇㅑ ]\s" # 성공
# "冬+.{0,6}ㅅ+.{2,4}[ㄷ|ㅼ]+.{1,1}ㄹ"     |ㄷㅕ |ㅈㅕ  |ㅇㅣㅇㅕ |ㅇㅑ 
# "ㄴᆞㄴ\s\S{1,5}ㄹ\s\S.{1,15}ㄴᆞㄴ\s\S{1,5}ㄹ\s" # 는 ㄹ 는 ㄹ #성공

mal="gasa_work"  
#"_data_history"  
#fff=""
fff=[]
sub_fff=[]
startend=[]
grab=1  # 홀수이면 선행 행의 값을 반영하지 않음 
howmany=0
grabb=0
each_line1=""

#print(glob.glob("*")) -> glob의 문제가 아님

path="gasa_work_paired" #아래처럼 절대경로로 바꾸어주니 작동함
path="C:\가사작업중20200208\Gasa_Divide\gasa_work_paired"
data_files = glob.glob(f"{path}/*") #

count_item=0



try:
    for e_f in data_files:
        i=0
        
        with codecs.open(e_f,"r", encoding="utf-16") as f:
            name=f.readline(0)
            # 작품번호, 작품제목 추출
            work_num, work_name = name.split(".")
            work_num.pop(' ')
            #print(name)
            #time.sleep(1)
            
            try:
                for each_line in f:
                    pattern_search = re.compile(words)
                    c = pattern_search.findall(each_line)  # 검색 건수를 아래에서 howmany로 저장할 것.

                    if  (grab%2) == 0:
                        each_line1=each_line
                    
                   
                        
                    
                    if len(c)==0 and (grab%2) != 0  : #검색 값이 없을 때, 선행 행의 값을 가져오지 않음  
                        pass
                    elif len(c)>0 and (grab%2) != 0 : #자모열에서 검색 값이 있을 때, 선행 행의 값을 가져오지 않은 상태,  AAAAAA로 넘겨
                        grab+=1 #선행 행의 값을 다음 행에 연결하도록 짝수로 만듦
                                           
                        
                        for m in re.finditer(words,each_line): #검색 구간을 검색하여 비율값을 리스트로 저장
                            start=int(m.start())
                            end=int(m.end())
                            startend=[]
                            
                            length=len(each_line) # 문장의 길이

                            pc_start=(start/length) #시작점의 비율
                            pc_end=(end/length) #끝점의 비율

                            
                            startend.append(pc_start)
                            startend.append(pc_end)

                            howmany=len(c)
                            
                            
                     
                    elif  (grab%2) == 0: #AAAAAAA 자모열의 검색 값을 넘겨 받아서 문자열의 검색 값을 저장 및 출력  grab은 짝수상태(검색유)
                        
                        grab+=1  # 다시 홀수로 만들어 닫음. 

                        

                        #print("AAA111111111111111111111111111111111111111AAAA")
                        #print(each_line1)
                        
                        
                        
                    
                        if i==0: # 문서의 처음/Gasa_Dic과 연동하여 '작가명/창작연대' 변수로 가져올 것
                            #딕셔너리와 연동 -> 작가/창작연대 추출 -> 아래에 이 변수를 반복 활용할 것.
                            #새 화일의 첫줄에서 자동 갱신됨.
                            author = Gasa_Dic.gasa_dic[work_num]['작가'].pop('\n')
                            era = Gasa_Dic.gasa_dic[work_num]['창작연대']
                            
                            sub_fff=[]
                            #fff=fff+"\n"+"\n"+"\n"+"\n"+"\n"+name+"\n"
                            i=1
                            lenlen=int(len(startend)/2)
                            print('len(each_line1)                        :' +str(len(each_line1))) #확인용>>>>>>>>>>>>>>>>>>

                            le=len(each_line1) #문장 길이
                            lc=howmany     #조회된 갯수   #len(each_line[new_start:new_end])
                            lw=(len(words)/5)

                            if 0<lc<2:
                                ll=int((le-lw)/(lc+1)) #조회할 앞 뒤 단어의 길이 계산
                                if ll>30:
                                    ll=32

                            elif 1<lc:
                                ll=int((le-lw)/lc) #조회할 앞 뒤 단어의 길이 계산
                                    
                                if ll>30:
                                    ll=32
                            else:
                                pass

                                
                            for m in range(lenlen):
                                
                                ppc_start=startend.pop(0)
                                ppc_end=startend.pop(0)
                                new_start=int(len(each_line1)*ppc_start)
                                new_end=int(len(each_line1)*ppc_end)
                                
                                
                                if len(each_line1)<50: 
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[:]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[:])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    #print('1111')
                                    '''print('le:' + str(le))
                                    print('len(each_line):' +str(len(each_line1)))'''
                                    
                                elif ll>new_start and (le-new_end)>ll: 
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[:new_end+ll]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[:new_end+ll])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    
                                    #print('1')
                            
                                elif ll<new_start and (le-new_end)>ll:
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[new_start-ll:new_end+ll]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)    
                                    sub_fff.append(each_line1[new_start-ll:new_end+ll])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    
                                    #print('2')
                                    
                                          
                            
                                elif ll<new_start and (le-new_end)<ll:
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[new_start-ll:]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[new_start-ll:])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    
                                    '''print('3')
                                    print('ll:' + str(ll))
                                    print('new_start:' + str(new_start)) 
                                    print('new_end:' + str(new_end))
                                    print('le:' + str(le))
                                    print('len(each_line1):' +str(len(each_line1)))
                                    print('len(ppc_start):' +str(ppc_start))
                                    print('len(ppc_end):' +str(ppc_end))'''
                            
                                elif ll>new_start and (le-new_end)<ll:
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[:]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[:])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    
                                    #print('4')
                                else:
                                    count_item+=1
                                    #fff=fff+str(count_item)+")"+ each_line1[new_start:new_end]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[new_start:new_end])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    #print('5')
                                    
                        
                            
                            
                            
                        elif i>0:
                            
                            lenlen=int(len(startend)/2)
                            
                            for m in range(lenlen):
                                ppc_start=startend.pop(0)
                                ppc_end=startend.pop(0)
                                new_start=int(len(each_line1)*ppc_start)
                                new_end=int(len(each_line1)*ppc_end)
                                
                                le=len(each_line1) #문장 길이
                                lc=len(each_line1[new_start:new_end])  #조회된 갯수
                                lw=(len(words)/5)

                                if 0<lc<2:
                                    ll=int((le-lw)/(lc+1)) #조회할 앞 뒤 단어의 길이 계산

                                elif 1<lc:
                                    ll=int((le-lw)/lc) #조회할 앞 뒤 단어의 길이 계산
                                    
                                    if ll>30:
                                        ll=22
                                else:
                                    pass


                                if len(each_line)<50: 
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[:]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[:])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    #print('2222')
                                    #print(each_line1)
                                    
                                elif ll>new_start and (le-new_end)>ll:
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[:new_end+ll]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[:new_end+ll])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    #print('6')
                            
                                elif ll<new_start and (le-new_end)>ll:
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[new_start-ll:new_end+ll]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[new_start-ll:new_end+ll])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    
                                    '''print('7')
                                    print('ll:' + str(ll))
                                    print('new_start:' + str(new_start)) 
                                    print('new_end:' + str(new_end))
                                    print('le:' + str(le))
                                    print('len(each_line):' +str(len(each_line)))
                                    print('len(ppc_start):' +str(ppc_start))
                                    print('len(ppc_end):' +str(ppc_end))'''
                            
                                elif ll<new_start and (le-new_end)<ll:
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[new_start-ll:]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[new_start-ll:])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    #print('8')
                            
                                elif ll>new_start and (le-new_end)<ll:
                                    count_item+=1
                                    #fff=fff+ str(count_item)+")"+each_line1[:]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[:])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    #print('9')

                                    
                                else:
                                    count_item+=1
                                    #fff=fff+str(count_item)+")"+ each_line1[new_start:new_end]+"\n"+"\n"
                                    sub_fff=[]
                                    sub_fff[0].append(count_item)
                                    sub_fff.append(work_num)
                                    sub_fff.append(work_name)
                                    sub_fff.append(author)
                                    sub_fff.append(era)
                                    sub_fff.append(each_line1[new_start:new_end])
                                    fff.append(sub_fff)
                                    sub_fff=[]
                                    #print('10')
                        
                                
                             
                    elif each_line==f.readline(-1):
                        f.close()
                        print("f.close()니다")
                        
                    else:
                        
                        pass
                    
                    
            except IOError as ioerr:
                print('File error (outer try): ' + str(ioerr))
                
    if not len(fff)==0 :
        out=open('C:\\가사작업중20200208\\Gasa_Divide\\result\\result_'+'('+str(count_item)+'건)'+words_han+"_from_"+mal+'.txt', mode= 'w', encoding='utf-16')
        print(fff, file=out)
        #ffff="error checking"   # 저장 경로 확인해 보니 c:\사용자\park9 
        #with open('resul9999t.txt', mode= 'w', encoding= 'utf-16') as f:
        #    f.write(fff)
                  
        out.close()
        print("검색어는 '" +str(words)+"' 입니다")
        
        print("검색 건수는 모두 '" +str(count_item)+"'건 입니다")
        
        
    else:
        print("해당되는 자료가 없습니다")
        
                    
except IOError as ioerr:
    print('File error (outer try): ' + str(ioerr))
print(fff)
    


    
    
    
