# How to Use — Conductor Ampacity Calculator

This guide is for users with no coding experience. Follow each step in order.

---

## What You Need

- A Windows PC
- An internet connection (for the one-time setup in Steps 1 and 2)

---

## Step 1 — Install Python

1. Go to **https://www.python.org/downloads/**
2. Click the yellow **Download Python 3.x.x** button
3. Run the downloaded installer
4. **Important:** On the first screen, tick the box that says **"Add Python to PATH"** before clicking anything else

   ![Add Python to PATH checkbox](https://www.python.org/static/img/python-logo.png)

5. Click **Install Now** and wait for it to finish

**To verify it worked:** open the Start menu, search for `cmd`, open **Command Prompt**, type `python --version`, and press Enter. You should see something like `Python 3.11.4`.

---

## Step 2 — Install Required Libraries (one time only)

The tool needs two small Python libraries. You only need to do this once.

1. Open **File Explorer** and navigate to the project folder
2. Click on the address bar at the top of File Explorer, type `cmd`, and press Enter — a Command Prompt window will open directly in that folder
3. Type the following and press Enter:

   ```
   pip install -r requirements.txt
   ```

4. Wait for it to finish. You will see a success message when done.

---

## Step 3 — Find Your Conductor Name

1. Open the file `data/conductor_data.xlsx` in Excel
2. Look at the **Conductor Name** column and find the conductor you want
3. Note down the exact name (e.g., `Grosbeak`, `Drake`, `Hawk`)
4. Close the file

> The name must match the spelling in the Excel file. Capitalisation does not matter —
> `grosbeak`, `Grosbeak`, and `GROSBEAK` all work.

---

## Step 4 — Edit the Configuration File

1. Find `config.ini` in the project folder
2. Right-click it and choose **Open with → Notepad**
3. Find the line that reads:

   ```
   name = Grosbeak
   ```

4. Replace `Grosbeak` with your conductor name
5. Save the file (**Ctrl + S**)

> **You only need to change the conductor name.** All other values in `config.ini`
> have working default values based on IEEE 738 standard conditions. You can
> explore and adjust them later once you are comfortable with the tool.

---

## Step 5 — Run the Tool

Double-click **`run_ampacity.bat`** in the project folder.

A Command Prompt window will open, run the calculation, and close automatically.
This normally takes only a few seconds.

---

## Step 6 — View Your Results

1. Open the `output/` folder inside the project directory
2. You will find a new Excel file named `{ConductorName}_{date_time}.xlsx`
3. Open it — it contains two sheets:
   - **Summary** — all inputs and calculated results in a formatted table
   - **Transient Profile** — conductor temperature vs. time chart and data

---

## Troubleshooting

| Problem | What to do |
|---|---|
| `.bat` file flashes and closes immediately | Right-click `run_ampacity.bat` → **Open with → Command Prompt**. Read the error message. |
| `'python' is not recognized as a command` | Python was not added to PATH. Re-install Python and tick the **"Add Python to PATH"** box on the first screen. |
| `Conductor '...' not found in the database` | Check that the name in `config.ini` matches the spelling in `data/conductor_data.xlsx` exactly. |
| `No module named 'pandas'` or `'openpyxl'` | Repeat Step 2 — the libraries were not installed successfully. |
| The `output/` folder is empty after running | Open Command Prompt in the project folder, run `python conductor_ampacity.py`, and read the error message. |
