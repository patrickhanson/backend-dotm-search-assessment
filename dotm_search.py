#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

import os
import argparse
import zipfile

DOC_FILENAME = 'word/document.xml'

def find_dotm(path_to_search, text_to_search):
    """Iterate through dotm files, looking for text"""
    file_list = os.listdir(path_to_search)

    files_searched = 0
    files_matched = 0

    for file in file_list:
        if not file.endswith('.dotm'):
            continue
        full_path = os.path.join(path_to_search, file)
        if not zipfile.is_zipfile(full_path):
            continue
        files_searched += 1

        with zipfile.ZipFile(full_path) as z:
            if DOC_FILENAME in z.namelist():
                with z.open(DOC_FILENAME) as doc:
                    for line in doc:
                        text_location = line.find(text_to_search)
                        if text_location >= 0:
                            files_matched += 1
                            left_location = max(0, text_location-40)
                            right_location = min(text_location + 41, len(line))
                            print('Match found in file {}'.format(full_path))
                            print('...' + line[left_location:right_location] + '...')
    print('Files searched: {}'.format(files_searched))
    print('Files matched: {}'.format(files_matched))

def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('text', help='Text to search for within dotm file')
    parser.add_argument('-d', '--dir', help='The directory to search within', default='.')
    return parser

def main():
    parser = create_parser()
    my_args = parser.parse_args()

    search_text = my_args.text
    search_path = my_args.dir
    
    if not my_args:
        parser.print_usage()
        exit(1)
    
    assert search_path is not None

    print('Searching directory {} for text {} ...'.format(search_path, search_text))
    
    find_dotm(my_args.dir, my_args.text)

if __name__ == '__main__':
    main()
