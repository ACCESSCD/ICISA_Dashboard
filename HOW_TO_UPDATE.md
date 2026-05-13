# How to Update the ICISA Speaker Dashboard

## Each time you update the Excel files:

1. Update `Faculty list follow up.xlsx` and/or `PGM ICISA 0305.xlsx`
   in `C:\Users\carol\PycharmProjects\ICISA information`
   — save and close them

2. Open **Claude Code** (terminal)

3. Say:
   > **"regenerate the speakers data and push to GitHub"**

That's it. Claude will run the script, update the site, and push to GitHub.
The live site refreshes within a minute.

---

## You do NOT need to:
- Manually activate the virtual environment
- Type any commands yourself
- Copy any files

---

## If Claude Code isn't open yet:
Open a terminal and type `claude`, wait for it to load, then say the phrase above.

---

## What the script does (for reference):
- Reads the two Excel files from `ICISA information`
- Filters out Declined speakers and moderators
- Counts session assignments per speaker from the programme
- Writes `speakers.json` to the `ICISA_Dashboard` folder
- Commits and pushes to GitHub Pages
