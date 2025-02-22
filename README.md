# rri

A simple script that calculates team rankings based on RRI formula


## Overview

This script was created for a local ultimate organization to calculate the weekly rankings for its league. 

The script has no formal input but rather uses the current directory to find an existing spreadsheet GameData.xlsx. This spreadsheet contains all game results for all teams in the league. The script then uses USAU's RRI algorithm (excluding the date weight) to calculate the team rankings. It places the final ranking data into a file Rankings.xlsx. If the output file does not exist, it will be created in the current directory. Otherwise the existing file will be modified.

The formula for RRI comes from [USAU's algorithm for calculating RRI.](https://play.usaultimate.org/teams/events/rankings/) It is typically recommended that teams have at least 10 games played when using RRI as a ranking algorithm. 


## Dependencies

This script was created using Python 3.9. It requires the openpyxl and pandas packages be installed.


## Contents

1. GameData.xlsx
    - An example spreadsheet format that can be used as an input to RRI.py
2. Rankings.xlsx
    - An example output spreadsheet created when RRI.py is ran
3. RRI.py
    - The main script to be ran


## How to Use

1. Create a spreadsheet called GameData.xlsx in the same directory as RRI.py
2. Add the team and game results to the spreadsheet using the required column headers. The column headers can be found in the heading of RRI.py or you can reference GameData.xlsx
3. Run RRI.py via python in a windows terminal
4. The script will create Rankings.xlsx in the RRI.py directory. If Rankings.xlsx already exists, the existing file will be modified

## Notes

The script frequently has path issues when ran on a MAC OS.