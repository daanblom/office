# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!# Office Automations (Python)

A small collection of Python scripts to speed up repetitive Microsoft Office and document tasks.

PPTXtoPOTX.py — convert all .pptx files in a folder tree into .potx templates next to the source (no Office required).

footerMover.py — lightweight GUI automation to step through slides and perform a repeated footer action (delete → paste → page-down → click). Handy when doing bulk footer adjustments by hand but with robot speed.

footerRemover.py — companion GUI automation intended to remove footers across slides (see usage below).

Repo: https://github.com/daanblom/office (main branch)

Quick start
# 1) Create a virtual env (optional but recommended)
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# 2) Install requirements for GUI scripts
pip install -r requirements.txt


PPTXtoPOTX.py uses only Python’s standard library and works without any extra packages.
The GUI scripts (footerMover.py, footerRemover.py) require pyautogui and keyboard.

Scripts
1) PPTXtoPOTX.py

Convert every *.pptx under a chosen root into a *.potx in the same folder (same filename, new extension).
This is done by editing the Open XML ZIP package ([Content_Types].xml) — no PowerPoint/COM involved.

Usage

# Windows PowerShell
python PPTXtoPOTX.py "C:\Path\To\Root"

# macOS/Linux
python PPTXtoPOTX.py /path/to/root


Options

--overwrite — overwrite existing .potx files (default is to skip).

Why this approach?

Cross-platform (Windows/macOS/Linux)

No Office install needed

Avoids Protected View, template redirection, and COM/automation quirks

2) footerMover.py

A tiny “human-in-the-loop” helper that repeats a key sequence while you keep PowerPoint (or another app) focused.

What it does (per slide)

Delete

Ctrl+V

PageDown

Left mouse click

…and repeats until you press q to quit. There’s a 10-second countdown so you can focus the right window.

Usage

python footerMover.py


Prepare your first slide (place cursor/focus where needed).

When the terminal says “Starting in 10 seconds…”, switch to PowerPoint.

Press q any time to stop.

Pro tips

Disable PowerPoint animations/transitions to keep timing stable.

If some actions miss, adjust the time.sleep(0.12) delays in the script (slower machines may need 0.2–0.3).

3) footerRemover.py

A sibling utility intended to help remove footers at speed via GUI automation.
Current pattern is similar to footerMover.py: it simulates keystrokes/clicks so you can power through a deck.

Usage

python footerRemover.py


Follow on-screen prompt (focus PowerPoint during the countdown).

Press q to exit.

Tip: If you need a different sequence (e.g., select footer placeholder → Delete → PageDown), tweak the pyautogui calls to match your workflow.

Requirements

PPTXtoPOTX.py

Python 3.8+

No external dependencies

footerMover.py / footerRemover.py

Python 3.8+

pyautogui (GUI automation)

keyboard (hotkey detection on Windows; requires admin privileges on some systems)

Install with:

# requirements.txt
pyautogui>=0.9
keyboard>=0.13
# PPTXtoPOTX.py uses only stdlib; python>=3.8 recommended
python>=3.8


macOS users: keyboard doesn’t support macOS; use pynput as an alternative (you can adapt the scripts by replacing the key detection loop). pyautogui on macOS also needs Accessibility permissions (System Settings → Privacy & Security → Accessibility).

Safety & permissions

Windows (UAC): running the GUI scripts from an elevated terminal may be required for keyboard to detect global keys in other apps.

macOS: grant Accessibility permissions to your terminal/IDE for pyautogui to control the mouse/keyboard.

Display scaling: if clicks land slightly off, check your OS/UI scaling; pyautogui assumes the current screen coordinate system.

Common issues & fixes

Nothing happens when I run a GUI script
Ensure the PowerPoint window has focus after the countdown. On macOS, verify Accessibility permissions for your terminal/IDE.

The loop is too fast/slow
Increase/decrease the time.sleep(...) delays between actions.

Accidentally spamming inputs
Press q in the terminal to exit; if needed, add an emergency sleep or a longer starting delay.

Contributing

PRs welcome! If you have other Office/document chores you’d like to automate (batch find/replace in DOCX, auto-export to PDF, etc.), open an issue or drop a script in a PR.

License

MIT — do what you like, but no warranties. Enjoy!
