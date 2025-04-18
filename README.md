
# 📊 Reporter with AI Insights

Generate stunning PowerPoint 📽️ and Excel 📊 status updates from your Azure DevOps (ADO) queries — all with a single command!

This tool:
- Reads your ADO query results (using your Personal Access Token)
- Generates a PowerPoint presentation using a custom template
- Updates an Excel sheet in a readable format
- Uses AI 🤖 to generate a quick summary insight slide

---

## 🔧 Setup Instructions

### 🐍 Prerequisites

Make sure you have **Python 3.7+** installed.

### 📦 Install Dependencies

Run this command to install required Python packages:

```bash
pip install pandas openpyxl python-pptx requests openai
```

---

## 📂 Folder Structure

Place the following files in the same directory as `reporter.py`:

- ✅ `pat.txt` – Your **Personal Access Token** from ADO (one-liner)
- ✅ OpenAIKey.txt  – Your **Open AI Key**  (one-liner), if you like to leverge AI to interperate 
- 📄 `template.pptx` – Your PowerPoint template file
- 📄 `template.xlsx` – Your Excel template file

---

## 🚀 How to Run

```bash
python reporter.py [--ai true/false] [--isTree true/false]
```
- --ai true/false: Enables or disables AI functionality. Default is false.
- --isTree true/false: Fetches work items as a tree structure (true) or a flat list (false). Default is false.

That’s it! The script:
- Reads your ADO query (pre-configured)
- Uses the templates to generate clean output
- Auto-generates one slide with **AI-generated insights**
- Saves output files in the current directory

---

## 🤖 What’s Happening Under the Hood?

- Authenticates using the `pat.txt` token
- Fetches work items from a predefined ADO query
- Parses fields like title, state, and status
- Updates an Excel template with fresh data
- Populates a PowerPoint template:
  - Includes work item metrics
  - Adds a slide summarizing progress using **OpenAI**-powered insights (optional/configurable)
- Saves files like `output.pptx` and `output.xlsx`

---

## 🧠 AI Magic

The script calls an AI model to:
- Read the work items
- Generate a concise status summary
- Auto-fill a designated slide in the PowerPoint deck

---

## 🙌 Contributing

Pull requests welcome! Add new templates, suggest better formatting, or extend AI use cases!

---

## 📜 License

MIT License – use it freely and improve it boldly.
