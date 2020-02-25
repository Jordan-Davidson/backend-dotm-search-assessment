#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import argparse
parser = argparse.ArgumentParser()
parser.add_argument('-d', '--directory' help = 'Specify a directory' required=True)
parser.add_argument('-t', '--text', help = 'Specify search text' required=True)
args = parser.parse_args()
"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Jordan"


def main():
    raise NotImplementedError("Your awesome code begins here!")


if __name__ == '__main__':
    main()

