#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Jonathan Jones"

import os
# import sys
import argparse
import zipfile


def creating_parse():
    parser = argparse.ArgumentParser(
        description="Finds a string in a specified file")
    parser.add_argument(
        "--dir", help="Input directory you wish to search", default=".")
    parser.add_argument(
        "text", help="Input text you wish to search for")
    return parser


def main():
    parser = creating_parse()
    files_searched = 0
    files_matched = 0
    args = parser.parse_args()
    dir_path = os.getcwd() + "/" + args.dir
    # print(dir_path)
    file_list = os.listdir(dir_path)
    print("Searching directory " + args.dir + " for text " + args.text)
    for file in file_list:
        if not file.endswith(".dotm"):
            print("File is not a .dotm: " + file)
            continue
        else:
            # print(file)
            files_searched += 1
            file_path = os.path.join(dir_path, file)
            with zipfile.ZipFile(file_path) as document:
                # print(document)
                with document.open("word/document.xml", "r") as words:
                    for word in words:
                        if args.text in word.decode("UTF-8"):
                            files_matched += 1
                            print("Found " + args.text + " in " + args.dir + "/" + file)
                            index = word.decode("UTF-8").find(args.text)
                            print("..." + word[index-40:index + 40].decode("UTF-8"))

    print("Total dotm files searched: " + str(files_searched))
    print("Total dotm files matched: " + str(files_matched))


if __name__ == '__main__':
    main()
