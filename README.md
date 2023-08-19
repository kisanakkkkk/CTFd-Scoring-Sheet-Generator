# CTFd-Scoring-Sheet-Generator
[![Python 3.8](https://img.shields.io/badge/Python-3.8-blue.svg)](https://www.python.org/downloads/)

Automatic CTFd Scoring Sheet Generator Written in Python 

## Installation
  - #### **Install Python 3.8**
    - [Python 3.8](https://www.python.org/downloads/)
  - #### **Install Python Dependency Packages**
    - [openpyxl](https://pypi.org/project/openpyxl/3.1.2/)
  - #### **Clone the Source Code**

```
git clone https://github.com/kisanakkkkk/CTFd-Scoring-Sheet-Generator.git
cd ./CTFd-Scoring-Sheet-Generator
pip3 install -r requirements.txt
python3 generator.py -d {user_data} -s {scoreboard_data} -c {chall_data}
```
## Basic Usage
1. Export the following files from your CTFd (https://docs.ctfd.io/docs/exports/ctfd-exports):
   * `{CTFNAME}-users+fields.csv` (for individual-based CTF)
    
        OR
    
        `{CTFNAME}-teams+members+fields.csv` (for team-based CTF)
   * `{CTFNAME}-scoreboard.csv`
   * `{CTFNAME}-challenges.csv`
2. Open file `generator.py`, modify the values of the variable 'CATEGORIES' according to the names of your CTF categories

    > `CATEGORIES = ["Web Exploitation","Binary Exploitation","Reverse Engineering","Cryptography","Forensic","Misc"]`

3. Run the code

    `python3 generator.py -d {CTFNAME}-users+fields.csv -s {CTFNAME}-scoreboard.csv -c {CTFNAME}-challenges.csv` (for individual-based CTF)
    
    OR

    `python3 generator.py -d {CTFNAME}-teams+members+fields.csv -s {CTFNAME}-scoreboard.csv -c {CTFNAME}-challenges.csv -t` (for team-based CTF)
4. Upload the the xlsx output file to google spreadsheet (DISCLAIMER: somehow opening it with excel will make it corrupted... so yeah)
5. Start Scoring
6. ???
7. Profit

## How to Score?

## Command Line Flags
See `--help` for the complete list
```text
usage: generator.py [-h] -d DATA -s SCORE -c CHALL [-o OUTPUT] [-t]

optional arguments:
  -h, --help            show this help message and exit
  -d DATA, --data DATA  specify user/team data CSV file
  -s SCORE, --score SCORE
                        specify scoreboard CSV file
  -c CHALL, --chall CHALL
                        specify challenges CSV file
  -o OUTPUT, --output OUTPUT
                        set output file (default: output.xlsx)
  -t, --team            indicates that the CTF is team-based (default: individual-based)
```

## Info
- The output file are INTENDED for spreadsheet use, opening with excel might corrupt the file
- Program will create a new folder named `sanitized` containing copies of your file
- Admin data will not be displayed on data sheets
- Team/User/Member's password will not be displayed on data sheets
- Output file will contain :
  - Accumulation sheet (to calculate the total score)
  - Sheets for each of category (scoring will be done there)
  - Sanitized copy of user/team data, scoreboard, and challenges