#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "suravita_roy"
import sys
import argparse
import os
import zipfile


def search_files(directory, text_to_search):
    searched_files = 0
    matched_files = 0
    files_in_dir = os.listdir(directory)
    #print(files_in_dir)
    for file in files_in_dir:
        #print(file)
        if  file.endswith('.dotm'):
            searched_files += 1
            #print(file)
            filepath = os.path.join(directory,file)
            if(zipfile.is_zipfile(filepath)):
                with zipfile.ZipFile(filepath) as myzip:
                    if 'word/document.xml' in myzip.namelist():
                        with myzip.open('word/document.xml') as docfile:
                            flag = False
                            for line in docfile:
                                # print(line)
                                indexline = line.find(text_to_search)
                            
                                if indexline >= 0:
                                    #print(indexline)
                                    flag = True
                                    print('Match found in {}\n...{}...'.format(filepath,line[indexline-40:indexline+40]))
                            if flag == True:
                                matched_files += 1
              


        else:
            #print('hello')
            continue

    print('Total dotm files searched: {}'.format(searched_files))   
    print('Total dotm files matched: {}'.format(matched_files))   

def main():
    #raise NotImplementedError("Your awesome code begins here!")
    parser = argparse.ArgumentParser()
    parser.add_argument('text_to_search', help="Provide the text you want to search inside files",type=str)
    parser.add_argument('--dir',help="Provide the directory to search for files",default=os.getcwd())
    args = parser.parse_args()
    # print(args.dir)
    # print(args.text_to_search)
    search_files(args.dir,args.text_to_search)

if __name__ == '__main__':
    main()
