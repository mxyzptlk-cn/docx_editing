#!/usr/bin/env python3
# -*- coding:utf-8 -*- 
# Author: Mxyzptlk
# Date: 2019-11-18

import tempfile
import win32api
import win32print

import os


def print_file():
    path_prefix = os.getcwd()
    filename = path_prefix + '\\套表模板' + '\\' + 'WN-QR-0-3-A软件及升级包杀毒记录-1.5.docx'
    
    win32api.ShellExecute(
        0,
        "print",
        filename,
        #
        # If this is None, the default printer will
        # be used anyway.
        #
        '/d:"%s"' % win32print.GetDefaultPrinter(),
        ".",
        0
    )


printer_name = win32print.GetDefaultPrinter()
print(printer_name)