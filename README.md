# CTFd-Scoring-Sheet-Generator
[![Python 3.8](https://img.shields.io/badge/Python-3.8-blue.svg)](https://www.python.org/downloads/)

Automatic CTFd Scoring Sheet Generator Written in Python 

<img width="892" alt="Screenshot 2023-08-19 234625" src="https://github.com/kisanakkkkk/CTFd-Scoring-Sheet-Generator/assets/70153248/abb87192-5834-4f78-a16b-f44d4f1618d2">

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
<img width="683" alt="Screenshot 2023-08-19 234424" src="https://github.com/kisanakkkkk/CTFd-Scoring-Sheet-Generator/assets/70153248/593127d7-6b0e-42b2-ad91-9d59a6bc8e35">


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
- Program will create copies of your file, located at `sanitized` folder
- Admin data will not be displayed on data sheets
- Team/User/Member's password will not be displayed on data sheets
- Ignore the width exceeding
