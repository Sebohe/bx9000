#!/bin/bash
rm -r __pycache__
rm -r build
rm -r dist
pyinstaller --noconsole --onefile --specpath=dist/ bx9000.py
mv dist/bx9000.exe .
