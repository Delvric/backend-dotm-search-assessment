#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = " Delvric Tezeno "

import zipfile
import argparse
import os


def search_all(str, sub):
    start = 0
    while True:
        start = str.find(sub, start)
        if start == -1:
            return
        yield start
        start += len(sub)


def find_file(file_name, search_text, files_with_string):
    with zipfile.ZipFile(file_name, 'r') as z:
        print(file_name)
        with z.open("word/document.xml") as contents:
            read_contents = contents.read()
            found_string = 0
            for is_string in search_all(read_contents, search_text):
                found_string += 1
                files_with_string.append(file_name)
                print(read_contents[is_string-40:is_string+40])


def loop_dir(args):
    list_of_files = os.listdir(args.dir)
    files_with_string = []
    file_count = 0
    for file in list_of_files:
        is_path = os.path.join(args.dir, file)
        if is_path.endswith('.dotm'):
            find_file(is_path, args.text, files_with_string)
            file_count += 1
    files_set = set(files_with_string)
    print("\n")
    print("Files with Matches: \n" + "\n".join(list(files_set)))
    print("Totals of Files with Matches: " + str(len(files_set)))
    print("Total of Files Searched: " + str(file_count))


def create_parser():
    parser = argparse.ArgumentParser(
        description="Searches for string in specified directory")
    parser.add_argument(
        "--dir", help="Please enter the directory you want to search", default=".")
    parser.add_argument("text", help="text to search for")
    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()
    loop_dir(args)


if __name__ == '__main__':
    main()
