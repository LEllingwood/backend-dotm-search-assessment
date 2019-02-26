#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""

import argparse
import os
import zipfile

# __author__ = LEllingwood

def create_parser():
    parser = argparse.ArgumentParser(description='Searches .dotm files for txt')
    parser.add_argument("--dir", help='directory to search', default='.')
    parser.add_argument("text", help='text to search for')
    return parser


# def search_zip(zipname, text_to_search):
#     """Opens a single zip file and searches for text"""
#     with zipfile.ZipFile(zipname, 'r') as myzip:
#         toc = myzip.namelist()
#         # get the table of contents
#         if 'word/document.xml' in toc:
#             with myzip.open('word/document.xml') as document:
#                 for line in document:
#                     idx = line.find(text_to_search)
#                     if idx >= 0:
#                         print(line[idx - 40: idx + 40])
#                         # I found a match!
#                         return True
#     # Nothing was found here.
#     return False


def main(directory_to_search, text_to_search):
    print ('Searching directory {} for text "{}" ...'.format(directory_to_search, text_to_search))
    
    # get a list of files to open
    file_list = os.listdir(directory_to_search)
    match_counter = 0
    search_counter = 0
    for file in file_list:
        search_counter += 1
        if not file.endswith('.dotm'):
            print("Ignoring file: {}".format(file))
            continue

        full_path = os.path.join(directory_to_search, file)
        # if search_zip(full_path):
        #     match_counter += 1

    # summarize findings

        with zipfile.ZipFile(full_path, 'r') as myzip:
            toc = myzip.namelist()
            # get the table of contents
            if 'word/document.xml' in toc:
                with myzip.open('word/document.xml') as document:
                    for line in document:
                        idx = line.find(text_to_search)
                        if idx >= 0:
                            match_counter += 1
                            print('...{}...'.format(line[idx - 40: idx + 40]))

    print('files matched: {}'.format(match_counter))
    print('files searched: {}'.format(search_counter)) 
        # index_
        # https://docs.python.org/2/library/zipfile.html
            

if __name__ == '__main__':
    parser = create_parser()
    namespace = parser.parse_args()
    # if namespace.dir is not None:
        # the "dir" name above will always match what you've called it in the add_argument above. 
        # Could the line above also be written like this: if namespace.dir
    # namespace.text 

    main(namespace.dir, namespace.text)
