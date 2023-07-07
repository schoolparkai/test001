#dividing_test.py
#coding:utf-8

import codecs
import glob
import codecs
import os




# wriet string to txt file
def write_string_to_file(temp_str, file_name):
    #the encoding must be same with the str
    file_object = open("./sizo_work_paired01/" +
                       file_name, 'w', encoding="utf-16")
    file_object.write(temp_str)   
    file_object.close()


data_files = glob.glob(os.getcwd() + "/sizo_work_paired/*.txt")
print ("the result files save in the " + os.getcwd())
for each_file in data_files:
    print (each_file + "-"*5 + ">dealing")   #begin to deal file
    with codecs.open(each_file, 'r', encoding="utf-16") as read_file:
        temp_file_string = ""
        temp_line = ""
        for each_line in read_file:
            if len(each_line.strip()) <2:
                continue
            else:
                temp_line  = temp_line + each_line.strip('\n')

            temp_file_string =  temp_line.strip('\n') # + "\n"
        print ( each_file + "-"*5 + ">finished")   #finish
        #the name of new file to save the result, the new file is in the current dir
        new_filename = os.path.splitext(os.path.basename(each_file))[0] + ".txt"
        #write to the file
        write_string_to_file(temp_file_string, new_filename)



    

