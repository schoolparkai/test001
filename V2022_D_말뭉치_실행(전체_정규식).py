#!/usr/bin/python 
# -*- coding: utf-8 -*-

import os, re
import glob
import sys, codecs
import 말뭉치_함수RE
global fff, words, i, count_item, grab, startend, howmany, grabb, each_line1

words ="ㄱㅓㄹㆎㅎᆞㅇㅑ" # 성공 "ㄴᆞㄴ\s\S.{1,15}ㅇㅏ\s"
# 정규표현식 문법 ?(앞 문자가 0개 또는 1개)/ .(최소 한개의 문자)/ {n,m}(n번 이상 m번 이하 반복: dot와 조합시켜야)/ \s(스페이스) / \S(스페이스 아닌 거) / (그룹){반복횟수}
# 한자로 된 작품은 한음절 크기가 1임에 주의할 것. {범위} 입력 시 1부터 입력할 것.

# "\S.{4,11}ㅇㅣ\s\S.{4,11}ㄱㅗ\S.{4,11}ㅇㅣ\s\S.{4,11}ᆡ" # 성공
# "ㄴᆞㄴ\s\S{0,1}ㅈㅓ\s\S.{1,8}ㅇㅏ\s" # 성공
#"ㄴᆞㄴ\s\S{0,1}[ㅈㅓ |ㄷㅕ |ㅈㅕ ]\s\S.{1,8}[ㅇㅏ |ㅇㅑ ]\s" # 성공
# "冬+.{0,6}ㅅ+.{2,4}[ㄷ|ㅼ]+.{1,1}ㄹ"     |ㄷㅕ |ㅈㅕ  |ㅇㅣㅇㅕ |ㅇㅑ 
# "ㄴᆞㄴ\s\S{1,5}ㄹ\s\S.{1,15}ㄴᆞㄴ\s\S{1,5}ㄹ\s" # 는 ㄹ 는 ㄹ #성공

mal="gasa_work"  
#"_data_history"  
fff=""
startend=[]
grab=1  # 홀수이면 선행 행의 값을 반영하지 않음 
howmany=0
grabb=0
each_line1=""


data_files = glob.glob("gasa_work_paired/*.txt") #

count_item=0




try:
    for e_f in data_files:
        i=0
        
        with codecs.open(e_f,"r", encoding="utf-16") as f:
            name=f.readline(0)
            print(name)
            
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
                        
                        
                        
                    
                        if i==0: # 문서의 처음이면 파일 이름을 앞에 쓰자구...
                            fff=fff+"\n"+"\n"+"\n"+"\n"+"\n"+name+"\n"
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
                                    fff=fff+ str(count_item)+")"+each_line1[:]+"\n"+"\n"
                                    #print('1111')
                                    '''print('le:' + str(le))
                                    print('len(each_line):' +str(len(each_line1)))'''
                                    
                                elif ll>new_start and (le-new_end)>ll: 
                                    count_item+=1
                                    fff=fff+ str(count_item)+")"+each_line1[:new_end+ll]+"\n"+"\n"
                                    #print('1')
                            
                                elif ll<new_start and (le-new_end)>ll:
                                    count_item+=1
                                    fff=fff+ str(count_item)+")"+each_line1[new_start-ll:new_end+ll]+"\n"+"\n"
                                    #print('2')
                                    
                                          
                            
                                elif ll<new_start and (le-new_end)<ll:
                                    count_item+=1
                                    fff=fff+ str(count_item)+")"+each_line1[new_start-ll:]+"\n"+"\n"
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
                                    fff=fff+ str(count_item)+")"+each_line1[:]+"\n"+"\n"
                                    #print('4')
                                else:
                                    count_item+=1
                                    fff=fff+str(count_item)+")"+ each_line1[new_start:new_end]+"\n"+"\n"
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
                                    fff=fff+ str(count_item)+")"+each_line1[:]+"\n"+"\n"
                                    #print('2222')
                                    #print(each_line1)
                                    
                                elif ll>new_start and (le-new_end)>ll:
                                    count_item+=1
                                    fff=fff+ str(count_item)+")"+each_line1[:new_end+ll]+"\n"+"\n"
                                    #print('6')
                            
                                elif ll<new_start and (le-new_end)>ll:
                                    count_item+=1
                                    fff=fff+ str(count_item)+")"+each_line1[new_start-ll:new_end+ll]+"\n"+"\n"
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
                                    fff=fff+ str(count_item)+")"+each_line1[new_start-ll:]+"\n"+"\n"
                                    #print('8')
                            
                                elif ll>new_start and (le-new_end)<ll:
                                    count_item+=1
                                    fff=fff+ str(count_item)+")"+each_line1[:]+"\n"+"\n"
                                    #print('9')

                                    
                                else:
                                    count_item+=1
                                    fff=fff+str(count_item)+")"+ each_line1[new_start:new_end]+"\n"+"\n"
                                    #print('10')
                        
                                
                             
                    elif each_line==f.readline(-1):
                        f.close()
                        print("f.close()니다")
                        
                    else:
                        
                        pass
                    
                    
            except IOError as ioerr:
                print('File error (outer try): ' + str(ioerr))
                
    if not fff=="" :
        out=open("result_"+"("+str(count_item)+"건)"+mal+".txt", "w", encoding="utf-16")
        print(fff, file=out)
        out.close()
        print("검색어는 '" +str(words)+"' 입니다")
        
        print("검색 건수는 모두 '" +str(count_item)+"'건 입니다")
        
        
    else:
        print("해당되는 자료가 없습니다")
        
                    
except IOError as ioerr:
    print('File error (outer try): ' + str(ioerr))
print(fff)
    


    
    
    
