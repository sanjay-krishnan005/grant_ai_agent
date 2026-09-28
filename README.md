# GrantAgent — Full Auto-Fill Mode

The agent opens the grant website, reads every form field, generates answers using Llama 3 (free, local), and types them in automatically. You review at the end and click Submit yourself.

**Zero copy-paste. Zero API costs.**

---

## One-time setup (Windows)

### 1. Install Node.js
Download from https://nodejs.org → choose "LTS" → install

### 2. Install Ollama
Download from https://ollama.com → install → then open Command Prompt and run:
```
ollama pull llama3
```
(This downloads ~4GB once)

### 3. Install project packages
Open Command Prompt in the grant-agent folder:
```
npm install
npx playwright install chromium
```

---

## Every time you use it

**Step 1 — Start Ollama** (open Command Prompt):
```
ollama serve
```
Leave this window open.

**Step 2 — Start GrantAgent** (open another Command Prompt in the grant-agent folder):
```
node server.js
```

**Step 3 — Open browser:**
Go to http://localhost:3000

---

## How to use

1. Paste the grant application URL
2. Enter the grant name
3. Choose AI model (llama3 is best default)
4. Paste your company profile (be detailed!)
5. Click **"Auto-fill form"**

The agent will:
- Open Chrome and navigate to the grant URL
- Scan all form fields (text, dropdowns, checkboxes, radio buttons)
- Generate answers for every field using your company profile
- Type everything in automatically
- Show you a summary of what was filled
- **You review and click Submit yourself**

---

## Field types supported

| Type | Handled |
|---|---|
| Text inputs | ✓ Auto-typed |
| Long essay / textarea | ✓ Auto-typed |
| Dropdowns (select) | ✓ Best matching option selected |
| Radio buttons | ✓ Best matching option clicked |
| Checkboxes | ✓ Checked/unchecked based on answer |
| File uploads | ✗ Skipped — upload manually |

---

## Tips

- **Login walls**: Log in manually in the Chrome window that opens, then click "Auto-fill form" again
- **CAPTCHAs**: Solve manually — the agent waits patiently
- **Bad fills**: Some fields may not match perfectly — review the summary and fix manually before submitting
- **Better answers**: Use llama3:70b model if your machine has 32GB+ RAM

---

## Switching models

Edit the model in the UI dropdown, or set the env variable:
```
set OLLAMA_MODEL=mistral
node server.js
```
