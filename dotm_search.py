#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import argparse
import zipfile

parser = argparse.ArgumentParser()
parser.add_argument('-d', '--directory', help = 'Specify a directory', required=True)
parser.add_argument('-t', '--text', help = 'Specify search text', required=True)
args = parser.parse_args()
"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Jordan"

def read_dotm(root, dotm_file):
    dotm = zipfile.ZipFile(os.path.join(root, dotm_file))
    content = dotm.read('word/document.xml').decode('utf-8')
    return content
    

def search_files(text, Dir):
    total = 0
    found = 0 
    for root, dirs, files in os.walk(Dir):
        print('Searching directory' + root + ' for ' + text + '...')
        for _ in files:
            if _.endswith('.dotm'):
                found += 1
                dotm_file = read_dotm(root, _)
                if text in dotm_file:
                    index = dotm_file.index(text)
                    print('match found in' + os.path.join(root, _))
                    print('...' + dotm_file[index-40:index+40] + '...')
            total += 1
    print('total matches: ' +  str(found))
    print('files searches: ' + str(total))

def main():
    search_files(args.text, args.directory)


if __name__ == '__main__':
    main()

