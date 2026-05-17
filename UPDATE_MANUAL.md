Pycharm
Terminal
Claude
Setup virtual environment

pls regenerate the data C:\Users\carol\PycharmProjects\EA26_AnaesthCareEurope\venv\Scripts\python.exe generate_data.py from this path            
  "C:\Users\carol\PycharmProjects\ICISA information" on this site https://github.com/ACCESSCD/ICISA_Dashboard    

# ICISA Dashboard — Update Manual

## Your two source files

| File | What it does |
|------|-------------|
| `Faculty list follow up.xlsx` | Speaker names, confirmation status, bio/photo received |
| `PGM ICISA 0305.xlsx` | Conference program — which sessions each speaker is in |

Both files must be saved in:
`C:\Users\carol\PycharmProjects\ICISA information\`

---

## How to update the dashboard

### Step 1 — Update your Excel files
Make your changes in **Faculty list follow up.xlsx** and/or **PGM ICISA 0305.xlsx** as usual. Save and close them before continuing.

### Step 2 — Open a terminal
In PyCharm, open the Terminal (bottom of the screen). Or open PowerShell and type:
```
cd "C:\Users\carol\PycharmProjects\ICISA_Dashboard"
```

### Step 3 — Regenerate the data
```
C:\Users\carol\PycharmProjects\EA26_AnaesthCareEurope\venv\Scripts\python.exe generate_data.py
```
You will see a summary printed, e.g.:
```
Speakers (excl. Declined): 66
  Confirmed : 51
  Pending   : 15
  No tasks  : 17
  4+ tasks  : 0
Written -> speakers.json
```

### Step 4 — Push to GitHub
```
git add speakers.json
git commit -m "Update speaker data"
git push
```

### Step 5 — Done
Wait about 60 seconds, then refresh **https://accesscd.github.io/ICISA_Dashboard/**
The live site will show the updated data.

---

## What the dashboard tracks

- **No sessions assigned** — speakers confirmed or pending who do not yet appear in the program file
- **4+ sessions** — speakers assigned to 4 or more sessions (potential overload)
- **Tasks column** in the table — number of times each speaker appears in the program
- **Bio / Photo** — whether the speaker has submitted their bio and photo

---

## Rules for the data

### Faculty list (`Faculty list follow up.xlsx`)
- Sheet used: **International Speakers**
- Speaker names: **column F** (first name), **column G** (last name)
- Confirmation status: **column U** (Participation confirmed)
- Declined speakers are automatically excluded from the dashboard
- If a name contains the word "moderator" it is ignored (moderators are not counted as tasks)
- Entries marked **DO NOT INVITE** are excluded

### Program file (`PGM ICISA 0305.xlsx`)
- Sheet used: **Sheet1**
- Speaker names are read from the session cells
- Each time a speaker's name appears in a session cell = 1 task
- Moderator lines are not counted as tasks

---

## If you rename the Excel files

Open `generate_data.py` in PyCharm and update these two lines near the top:

```python
FACULTY_PATH = Path(r'C:\Users\carol\PycharmProjects\ICISA information\Faculty list follow up.xlsx')
PGM_PATH     = Path(r'C:\Users\carol\PycharmProjects\ICISA information\PGM ICISA 0305.xlsx')
```

Change the filenames to match your new file names, then save.

---

## Quick reference — full update in one copy-paste

Open PowerShell and paste all four lines at once:

```
cd "C:\Users\carol\PycharmProjects\ICISA_Dashboard"
C:\Users\carol\PycharmProjects\EA26_AnaesthCareEurope\venv\Scripts\python.exe generate_data.py
git add speakers.json
git commit -m "Update speaker data" && git push
```

