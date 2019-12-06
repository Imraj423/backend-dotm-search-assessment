#!/usr/bin/env python
# -*- coding: utf-8 -*-

# """
# Given a directory path, search all files in the path for a given text string
# within the 'word/document.xml' section of a MSWord .dotm file.
# """
__author__ = "Raja"
import os
import zipfile


def main(directory, query):
    file_list = os.listdir(directory)
    for file in file_list:
        full_path = os.path.join(directory, file)
        if not full_path.endswith(".dotm"):
            print("this isn't a dotm {}".format(full_path))
            continue
        if not zipfile.is_zipfile(full_path):
            print("this isn't a zip file {}".format(full_path))
            continue
        with zipfile.ZipFile(full_path, "r") as zipped:
            stuff = zipped.namelist()
            if "word/document.xml" in stuff:
                with zipped.open("word/document.xml", "r") as doc:
                    for things in doc:
                        x = things.find(query)
                        print(x)


if __name__ == '__main__':
    main("dotm_files","$")

