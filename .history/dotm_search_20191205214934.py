#!/usr/bin/env python
# -*- coding: utf-8 -*-

# """
# Given a directory path, search all files in the path for a given text string
# within the 'word/document.xml' section of a MSWord .dotm file.
# """
__author__ = "Raja"
import os
import zipfile
import argparse


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dir', help='which directory to search in', 
    default='.')
    parser.add_argument('text', help='text query')
    return parser


def main(directory, query):
    file_list = os.listdir(directory)
    files_searched = 0
    files_matched = 0
    parser = create_parser()
    args = parser.parse_args()
    print(args)
    print("Searching directory {} for text '{}' ...".format(directory, query))
    for file in file_list:
        files_searched += 1
        full_path = os.path.join(directory, file)
        if not full_path.endswith(".dotm"):
            print("  ...this isn't a dotm file: {}".format(full_path))
            continue
        if not zipfile.is_zipfile(full_path):
            print("  ...this isn't a zip file: {}".format(full_path))
            continue
        with zipfile.ZipFile(full_path, "r") as zipped:
            stuff = zipped.namelist()
            if "word/document.xml" in stuff:
                with zipped.open("word/document.xml", "r") as doc:
                    for thing in doc:
                        x = thing.find(query)
                        if x >= 0:
                            files_matched += 1
                            print("  ...{}...".format(thing[x - 40:x + 40]))
    print("Files matched: {}".format(files_matched))
    print("Files searched: {}".format(files_searched))



if __name__ == '__main__':
    main("dotm_files", "$")

