# README.md
# 03/03/2025
#
# BBS2 - main.py
#
# PURPOSE:
# To automate the merging of Timecards and Check Copies for filing.
#
# FUNCTIONALITY:
# Uses relative directories "input/Check_Copy" and "input/TC_Report", to scan crewmember's names by cropping sections of PDFs dropped inside those directories.
# Saves in "output", PDFs with each crewmember's Timecard on top and Check Copies behind them according to the detected name.
# Saves in "output" also two XLSX files with errors and info about timecards and check copies that were unable to be matched with eachother / anything.
