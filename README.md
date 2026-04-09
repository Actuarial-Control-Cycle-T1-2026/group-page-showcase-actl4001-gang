# SOA 2026 Student Research Case Study

# Note: this was written and used internally within the team as a "quick start guide" since we were unfamiliar with github. it contains step-by-step instructions on how github works, how to link it to RStudio (which all of us bar one used), and a quick equitte guide on how to commit/push changes onto the GitHub repo :)

## Project Structure
```
soa-2026-case-study/
├── data/
│   ├── raw/          ← Add SOA data files here locally (not synced to GitHub)
│   └── cleaned/      ← Output from 01_cleaning.R
├── R/                ← All R scripts, run in numbered order
├── report/           ← Report drafts
├── outputs/
│   └── figures/      ← Charts exported from R
├── .gitignore
└── README.md
```

## Script Run Order
Run R scripts in this order: `00` → `01` → `02` → `03` → `04` → `05`

---

# 👥 Collaboration Guide — Start Here

This section walks you through everything you need to go from zero to fully set up and collaborating through GitHub. Follow the steps in order.

---

## Stage 1 — Install Git on Your Computer

Git is the software that tracks changes to files. GitHub is the website that hosts the project online. You need Git installed locally before anything else works.

### Windows
1. Go to **https://git-scm.com/download/win** and download the installer
2. Run the installer — accept all default settings throughout
3. Once installed, you'll have a program called **Git Bash** on your computer — this is your terminal for running Git commands

### Mac
1. Open **Terminal** (search for it in Spotlight)
2. Type the following and press Enter:
```bash
git --version
```
3. If Git is not installed, macOS will automatically prompt you to install it — follow the prompts

---

## Stage 2 — Configure Git With Your Identity

This tells Git who you are so your changes are tagged with your name. You only do this once.

Open **Terminal** (Mac) or **Git Bash** (Windows) and run these two commands, replacing the name and email with your own:

```bash
git config --global user.name "Your Name"
git config --global user.email "your@email.com"
```
```
*note: shortcut to paste in the git terminal is middle mouse, otherwwise you can just right click and select paste (control + v does not work)
``` 
Use the same email address you signed up to GitHub with.

---

## Stage 3 — Accept the Repository Invitation (if u are looking at this on github you're done with this, go to next step)

1. Check your email for an invitation from GitHub to join the repository
2. Click **View invitation** → **Accept invitation**
3. You now have access to the repo at `https://github.com/ethan-boey/actl4001-soa-2026-case-study`

---

## Stage 4 — Clone the Repository to Your Computer

Cloning creates a local copy of the project on your laptop that stays connected to GitHub.

Open Terminal (Mac) or Git Bash (Windows) and run:

```bash
cd Documents
git clone https://github.com/ethan-boey/actl4001-soa-2026-case-study.git
cd actl4001-soa-2026-case-study
```

Replace `Documents` with whatever folder you want the project to live in.

After this you'll have a folder called `actl4001-soa-2026-case-study` on your computer containing the full project.

---

## Stage 5 — Add the Raw Data Files Locally

The SOA data files are intentionally excluded from GitHub (they're too large and don't need version tracking). You need to add them manually to your local copy.

Place all of the following files inside the `data/raw/` folder on your computer (create a new "raw" folder if not already there, the SOA files can be downloaded if you havent already from the SOA website):

- `srcsc-2026-claims-business-interruption.xlsx`
- `srcsc-2026-claims-cargo.xlsx`
- `srcsc-2026-claims-equipment-failure.xlsx`
- `srcsc-2026-claims-workers-comp.xlsx`
- `srcsc-2026-cosmic-quarry-inventory.xlsx`
- `srcsc-2026-cosmic-quarry-personnel.xlsx`
- `srcsc-2026-interest-and-inflation.xlsx`

These files will never be uploaded to GitHub — every team member manages their own local copy.

---

## Stage 6 — Connect RStudio to Git

### 6a — Enable Git in RStudio settings
1. Open RStudio
2. Go to **Tools → Global Options → Git/SVN**
3. Tick **"Enable version control interface for RStudio projects"**
4. Under **Git executable**, check the path is filled in automatically
   - If not, click **Browse** and find it:
     - **Windows:** `C:/Program Files/Git/bin/git.exe`
     - **Mac:** open Terminal, type `which git`, and paste the result
5. Click **OK** and restart RStudio

### 6b — Open the project in RStudio
This is the step that activates the Git panel inside RStudio.

1. In RStudio, go to **File → New Project → Existing Directory**
2. Navigate to your cloned `actl4001-soa-2026-case-study` folder
3. Click **Create Project**

This is where we will be working on files locally, which then can be uploaded to the Github repo (see below).

---

## Stage 7 — Create Your Personal Branch

Everyone works in their own branch to avoid overwriting each other's work. Think of it as your own personal lane.

In the RStudio Git panel, find the **purple branch icon** in the top-right corner of the panel:

1. Click the **purple branch icon**
2. Type your branch name from the table at the top of this README (e.g. `wc-bi-models`)
3. Make sure **"Sync branch with remote"** is ticked
4. Click **Create**

You are now working in your own branch. All your commits stay in your lane until you're ready to merge.

---

## Stage 8 — Understand How It All Works

### The big picture

```
Your Laptop            GitHub (cloud)         Teammate's Laptop
-----------            --------------         -----------------
your copy   →  push →  shared copy  ← pull ←  their copy
            ← pull ←               → push →
```

Everyone has their own copy of the project locally. GitHub is the shared copy in the middle. You **push** to send your work up, and **pull** to receive your teammates' work down.

### The four commands you'll use every day

| Action | What it does | How to do it in RStudio |
|--------|-------------|------------------------|
| **Pull** | Download latest changes from teammates | Click the blue ↓ arrow in the Git panel |
| **Stage** | Mark files as ready to save | Tick the checkbox next to each file in the Git panel |
| **Commit** | Save a labelled snapshot of your changes | Click **Commit**, write a message, click **Commit** |
| **Push** | Upload your commits to GitHub | Click the green ↑ arrow in the Git panel |

### The status letters in the Git panel

When you save a file, it appears in the Git panel with a letter showing its status:

| Letter | Meaning |
|--------|---------|
| `M` | Modified — you changed a file that already existed |
| `?` | Untracked — a new file Git hasn't seen before |
| `A` | Added — staged and ready to commit |
| `D` | Deleted — a file has been removed |

---

## Stage 9 — Your Daily Workflow

Follow this sequence every time you sit down to work on the project.

### Step 1 — Pull before you do anything else
In the RStudio Git panel, click the **blue ↓ Pull arrow**.
This downloads your teammates' latest work. Always do this first — skipping it is how you end up with conflicts.

### Step 2 — Write your R code as normal
Open your assigned script, write code, save the file. As soon as you save, the file appears in the Git panel.

### Step 3 — Stage your files
Tick the **checkbox** next to each file you want to include in your next commit. If you've only changed one file, only tick that one.

### Step 4 — Commit your changes
1. Click the **Commit** button
2. A window opens — you'll see your staged files on the top left, and a colour-coded preview of your changes at the bottom (green = added lines, red = removed lines)
3. In the **Commit message** box (top right), write a short description of what you did — e.g. *"fitted lognormal distribution to cargo severity"*
4. Click **Commit** — a confirmation popup appears, close it

### Step 5 — Push to GitHub
Click the **green ↑ Push arrow**. Your work is now on GitHub and visible to the whole team.

---

## Stage 10 — Merging Your Work Into Main

When your section is complete and ready for the project lead to review:

1. Go to the GitHub repo page in your browser
2. GitHub will show a yellow banner: **"Your branch had recent pushes — Compare & pull request"** — click it
3. Write a short summary of what you completed
4. Assign the **Project Lead** as the reviewer
5. Click **Create pull request**

The project lead reviews your changes, leaves any comments, and clicks **Merge** when ready. Your work is now in `main` and available to everyone on their next pull.

---

## Common Issues & Fixes

**The Git panel isn't showing in RStudio**
Go to `Tools → Global Options → Git/SVN` and make sure the Git executable path is set correctly, then restart RStudio. Also make sure you opened the project using the `.Rproj` file and not just the folder.

**I get an error when pushing**
You probably need to pull first. Click the blue ↓ Pull arrow, then try pushing again.

**I accidentally edited a file and want to undo it**
In the Git panel, right-click the file → **Revert** → confirm. This restores it to the last committed version. Note: this permanently discards your uncommitted changes.

**Two people edited the same file (merge conflict)**
Git will flag the file in the panel. Open it in RStudio — you'll see sections marked with `<<<<<<<` and `>>>>>>>` showing both versions. Edit the file to keep the correct version, delete the markers, save, then stage and commit as normal. The best prevention is to always pull first and stick to your assigned scripts.

**I committed to the wrong branch**
Tell the project lead before pushing — it's fixable without pushing. If you've already pushed, still tell the project lead and they can sort it out.

---

## Quick Reference Card

```bash
# These are the terminal equivalents of the RStudio buttons, for reference

git pull origin main              # download latest from GitHub
git checkout -b your-branch-name  # create and switch to a new branch
git checkout main                 # switch back to main
git add .                         # stage all changed files
git commit -m "your message here" # save a snapshot
git push origin your-branch-name  # upload to GitHub
git status                        # see what's changed locally
git log --oneline                 # see commit history
```
