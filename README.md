# Office Automation Assistant

An AI-powered desktop assistant that lets you control **Excel**, **Word**, and **PowerPoint** using natural language — via voice, keyboard, or clipboard input.

---

## 📌 What It Does

You give a command like:
> *"create a table with 4 columns and 5 rows"*
> *"bold A1:E1 and sum B2:B10 in C10"*
> *"insert a heading Introduction in Word"*

And the assistant automatically executes it on the correct file — no clicking, no manual work.

---

## 🗂️ Project Structure

```
project/
│
├── server.py                    # Flask API server — entry point
│
├── parser/
│   └── command_parser.py        # Parses natural language → action dicts
│
├── executor/
│   ├── __init__.py              # Exports all executors
│   ├── excel_executor.py        # Executes Excel actions via openpyxl
│   ├── word_executor.py         # Executes Word actions via python-docx
│   └── ppt_executor.py          # Executes PowerPoint actions via python-pptx
│
├── commands/
│   ├── excel_commands.json      # Excel keyword → action mapping
│   ├── word_commands.json       # Word keyword → action mapping
│   └── powerpoint_commands.json # PowerPoint keyword → action mapping
│
├── modules/
│   ├── ui.py                    # Displays results to user
│   ├── system_core.py           # Launch/close apps, OS-level tasks
│   └── store_apps.py            # Tracks known installed apps
│
├── ai/
│   └── openai_handler.py        # Handles AI/GPT responses
│
├── listener/
│   ├── voice_listener.py        # Captures voice input
│   ├── keyboard_listener.py     # Captures keyboard input
│   └── clipboard_listener.py    # Captures clipboard input
│
├── requirements.txt             # All dependencies
└── README.md                    # This file
```

---

## ⚙️ How It Works

```
User Input (voice / keyboard / clipboard)
            │
            ▼
        server.py
     (Flask API — POST /command)
            │
            ▼
   command_parser.py
   → Detects app (excel / word / ppt)
   → Loads the matching JSON command file
   → Matches keywords to action names
   → Extracts parameters (range, rows, cols, text...)
   → Returns list of action dicts
            │
            ▼
   executor/ (correct executor is selected)
   → ExcelExecutor   → openpyxl
   → WordExecutor    → python-docx
   → PowerPointExec  → python-pptx
            │
            ▼
   Action executed on the file
            │
            ▼
   Response returned → ui.py displays result
```

---

## 🚀 Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/your-username/office-automation-assistant.git
cd office-automation-assistant
```

### 2. Create and activate virtual environment
```bash
python -m venv .venv

# Windows
.venv\\Scripts\\activate

# Mac/Linux
source .venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the server
```bash
python server.py
```

---

## 📦 Dependencies

```
flask
openpyxl
python-docx
python-pptx
openai
pillow
pytesseract
pyperclip
pynput
```

Install all at once:
```bash
pip install flask openpyxl python-docx python-pptx openai pillow pytesseract pyperclip pynput
```

---

## 🧠 Command Examples

| App         | Command                                      | Action                        |
|-------------|----------------------------------------------|-------------------------------|
| Excel       | `bold A1:E1`                                 | Bold the cell range           |
| Excel       | `sum B2:B10 in C10`                          | Write SUM formula in C10      |
| Excel       | `create a table with 4 columns and 5 rows`   | Create a bordered table       |
| Word        | `insert heading Introduction`                | Add a heading                 |
| Word        | `create table with 3 columns and 4 rows`     | Insert a Word table           |
| PowerPoint  | `add slide with title Overview`              | Add a new slide               |

---

## 🔧 Adding a New Command

1. Open the relevant JSON file (e.g. `excel_commands.json`)
2. Add a new entry:
```json
{
  "action": "your_action_name",
  "keywords": ["your keyword", "alternative phrase"],
  "parameters": {
    "param1": "\\\\b(\\\\d+)\\\\b"
  },
  "description": "What this command does"
}
```
3. Add the handler in the matching executor file:
```python
def _do_your_action_name(self, params):
    # your logic here
```
4. Wire it in the `run()` method:
```python
if action == "your_action_name":
    self._do_your_action_name(action_dict)
```

---

## 🤝 Contributing

Pull requests are welcome. For major changes, open an issue first to discuss what you would like to change.

---

## 📄 License

MIT License — feel free to use and modify.
"""

with open("README.md", "w", encoding="utf-8") as f:
    f.write(readme)

print("README.md created successfully")
