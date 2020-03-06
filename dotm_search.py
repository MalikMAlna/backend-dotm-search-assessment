#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "watched demo by madarp"

import os
import sys
import zipfile
import argparse

DOC_FILENAME = "word/document.xml"


def scan_zipfile(z, search_text, full_path):
    """Returns True if search_text was found"""
    with z.open(DOC_FILENAME) as doc:
        xml_text = doc.read()
    xml_text = xml_text.decode("utf8")
    text_location = xml_text.find(search_text)
    if text_location >= 0:
        # match_count += 1
        print(f"Match found in file {full_path}")
        print("    ..." +
              xml_text[text_location-40:text_location+40]
              + "...")
        return True
    return False


def create_parser():
    """Creates parser for dotm searches"""
    parser = argparse.ArgumentParser(
        description="Searches for text within all dotm files")
    parser.add_argument(
        '--dir', help="directory to search for dotm files", default=".")
    parser.add_argument('text', help="text to search within each dotm file")
    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()

    search_text = args.text
    search_path = args.dir

    if not search_text:
        parser.print_usage()
        sys.exit(1)

    print(f"Searching directory {search_path}")
    print(f"for dotm files with text {search_text}")

    file_list = os.listdir(search_path)
    match_count = 0
    search_count = 0

    for file in file_list:
        if not file.endswith(".dotm"):
            print(f"Disregarding file: {file}")
            continue
        else:
            search_count += 1

        full_path = os.path.join(search_path, file)

        if zipfile.is_zipfile(full_path):
            with zipfile.ZipFile(full_path) as z:
                z_names = z.namelist()
                if DOC_FILENAME in z_names:
                    if scan_zipfile(z, search_text, full_path):
                        match_count += 1
        else:
            print(f"Not a zipfile: {full_path}")
    print(f"Total dotm files searched: {search_count}")
    print(f"Total dotm files matched: {match_count}")


if __name__ == '__main__':
    main()
