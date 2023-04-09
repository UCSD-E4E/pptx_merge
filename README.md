# e4e_pptx_merge

## Setup Instructions
1. On Windows, ensure that Office 365 is downloaded (i.e. PowerPoint should be located in your Start Menu).
2. Create a virtual environment: `python -m venv .venv`
3. Activate the virtual environment: `.venv\Scripts\activate`
4. Install this repository: `python -m pip install .`
5. Test that the install is sane: `python -m pytest tests`
6. Add `\\e4e-nas.ucsd.edu\programmatics` to your computer
    1. Open Windows Explorer
    2. Go to "This PC"
    3. In the ribbon, click on "Add a network location"
    4. Click "Next"
    5. Select "Choose a custom network location", then click "Next"
    7. Enter `\\e4e-nas.ucsd.edu\programmatics`, then click "Next"
    8. Click "Next"
    9. Click "Finish"

## Usage Instructions
1. Activate the virtual environment: `.venv\Scripts\activate`
2. Use `e4e_meeting_tool`:
```
usage: e4e_meeting_tool [-h] --job_file JOB_FILE --week WEEK [--output OUTPUT] [--order {sequence,random}]

options:
  -h, --help            show this help message and exit
  --job_file JOB_FILE, -j JOB_FILE
  --week WEEK, -w WEEK
  --output OUTPUT, -o OUTPUT
  --order {sequence,random}
```
For example, for generating Week 2 in `\\e4e-nas.ucsd.edu\programmatics\2023_Q4` as `\\e4e-nas.ucsd.edu\programmatics\2023_Q4\Week 2 final.pptx`:
`> e4e_meeting_tool --job_file job.yml --week 2 --output \\e4e-nas.ucsd.edu\programmatics\2023_Q4\ --order random`

Note that you can specify the order of presentations by modifying the schedule in `job.yml`

## Installing for developers
Developers should do the following:
```
# Create a virtual environment
python -m venv .venv

# Activate the virtual environment
.venv/Scripts/activate.ps1   # for Windows PowerShell
.venv/Scripts/activate.bat   # for Windows Command Prompt
source .venv/bin/activate    # for bash

# Install in developer mode
python -m pip install -e .[dev]
```
