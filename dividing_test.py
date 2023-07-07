#dividing_test.py
#coding:utf-8

import codecs
import glob
import codecs
import os

first_parts = ("ㄱ", "ㄲ", "ㄴ", "ㄷ", "ㄸ", "ㄹ", "ㅁ", "ㅂ", "ㅃ", "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅉ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ")
second_parts =("ㅏ", "ㅐ", "ㅑ", "ㅒ", "ㅓ", "ㅔ", "ㅕ", "ㅖ", "ㅗ", "ㅗㅏ", "ㅗㅐ", "ㅗㅣ", "ㅛ", "ㅜ", "ㅜㅓ", "ㅜㅔ", "ㅜㅣ", "ㅠ", "ㅡ", "ㅡㅣ", "ㅣ")
third_parts = ("", "ㄱ", "ㄲ", "ㄳ", "ㄴ", "ㄵ", "ㄶ", "ㄷ", "ㄹ", "ㄺ", "ㄻ", "ㄼ", "ㄽ", "ㄾ", "ㄿ", "ㅀ", "ㅁ", "ㅂ", "ㅄ", "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ")
def divide_korean(temp_string):
    temp_string_value = ord(temp_string)
    part_1 = (temp_string_value - 44032) // 588
    part_2 = (temp_string_value - 44032 - part_1 * 588) // 28
    part_3 = (temp_string_value - 44032 ) % 28
    return first_parts[part_1] + second_parts[part_2] + third_parts[part_3]



old_korean_dictionary = {}
read_file = codecs.open("old_korean_dictionary.txt", 'r', encoding="utf-8")
for each_line in read_file:
    old_korean, dividing_parts = each_line.split()
    old_korean_dictionary[old_korean] = dividing_parts


# wriet string to txt file
def write_string_to_file(temp_str, file_name):
    #the encoding must be same with the str
    file_object = open("./sizo_work_paired01/" +
                       file_name, 'w', encoding="utf-16")
    file_object.write(temp_str)   
    file_object.close()


data_files = glob.glob(os.getcwd() + "/sizo_work01/*.txt")
print ("the result files save in the " + os.getcwd())
for each_file in data_files:
    print (each_file + "-"*5 + ">dealing")   #begin to deal file
    with codecs.open(each_file, 'r', encoding="utf-16") as read_file:
        temp_file_string = ""
        for each_line in read_file:
            if each_line.strip() == "":
                continue
            temp_line = ""
            for i in range(0, len(each_line)):
                if each_line[i] in old_korean_dictionary:
                    temp_line = temp_line + old_korean_dictionary.get(each_line[i])
                elif each_line[i] >= u'\uAC00' and each_line[i] <= u'\uD7AF':
                    temp_line = temp_line + divide_korean(each_line[i])
                elif each_line[i] == "\n":                   ################
                    continue                                      ################
                else :
                    temp_line = temp_line + each_line[i]
            if len(temp_line)<2:
                temp_line = temp_line + each_line.strip('\n')
            else:
                temp_line = temp_line.strip('\n') + "\n" + each_line.strip('\n')     ############
            temp_file_string = temp_file_string.strip('\n') + "\n" +  temp_line.strip('\n') # + "\n"
        print ( each_file + "-"*5 + ">finished")   #finish
        #the name of new file to save the result, the new file is in the current dir
        new_filename = os.path.splitext(os.path.basename(each_file))[0] + "_result.txt"
        #write to the file
        write_string_to_file(temp_file_string, new_filename)



    

