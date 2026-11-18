# 🧾 Automated POS & Bank Reconciliation App (نظام مطابقة نقاط البيع الآلي)

An open-source, template-driven **Streamlit** web application designed for retail and multi-branch businesses to reconcile daily **POS Movement Workbooks (`ملف الحركة`)** against **Bank / Terminal Receipts (`ملف البنك`)**, verify cashier variances (`عجز` / `زيادة`), and automatically generate ERP-ready double-entry journal workbooks.

Built with a native **Arabic Right-to-Left (RTL)** interface, an interactive **multi-page PDF/image hover magnifier** (zero OCR dependencies), smart receipt filename auto-fill, and a built-in **First-Time Setup & Settings UI (`إعدادات الشركة والفروع`)** so any company can configure its own branches, logos, and chart of accounts in minutes.

---

## ✨ Key Features

1. **Native Arabic RTL Interface & Standard Accounting Terminology:**
   - Full RTL layout (`direction: rtl; text-align: right;`) with top toolbar clearance (`padding-top: 2.5rem;`).
   - Uses standard Arabic accounting labels throughout: **ملف الحركة**, **مبلغ الحركة**, **ملف البنك**, **مبلغ بنكي**, **مبلغ الكاش / الوردية -**, **الوردية**, **مبلغ الفرق**, **المحسوب**, **اسم طريقة الدفع**, and **طريقة الدفع**.

2. **Split-Screen Workspace & Multi-Page Hover Magnifier (No OCR):**
   - Dual-column workspace (`st.columns([1.2, 1])`) keeping the receipt viewer inside a fixed-height scrollable container (`height=780`) alongside the cashier entry form.
   - Renders multi-page PDFs (`PyMuPDF`) and images as a **continuous white paper stream** inside an interactive HTML/CSS/JS viewer with a circular magnifying lens (`142px`) and floating zoom/toggle pill (`🔍 مكبر: تشغيل/إيقاف` + `2x / 2.5x / 3.5x`).
   - Dynamically tracks `event.target.src` and `event.target.getBoundingClientRect()` across every page so hovering over Page 2, Page 3, etc. magnifies accurately.

3. **Smart Filename Auto-Fill & Streamlined Card Inputs:**
   - Automatically extracts **Cashier Name (`اسم الكاشير`)**, **Cash or `-Shift` (`مبلغ الكاش / الوردية -`)**, and **Signed Variance (`مبلغ الفرق`: `عجز` / `زيادة`)** directly from the uploaded receipt's filename.
   - Primary card networks ordered for rapid data entry: **Mastercard**, **Mada**, then **Visa** (accepts space-separated terminal totals, e.g., `1992 1203`).
   - Compact signed variance input (`-` for shortage `عجز`, `+` for extra `زيادة`) synced with a `- عجز` / `+ زيادة` pill toggle.
   - Secondary **`خصم` (Discount)** and **Rare Networks (`0` Variants — GCCNET, Amex, Discover, UnionPay, Maestro)** neatly tucked inside a collapsed expander (`خيارات إضافية: الخصم والشبكات النادرة (0)`).

4. **Dual-Purpose Lookup (`مبلغ الكاش / الوردية -`) & Shift Disambiguation:**
   - **Lookup by Shift (`-Shift`, e.g., `-920`):** Start the input with `-` followed by the unique shift digits after the colon `:` (e.g., `920` in `03047309:920`) to match the **`الوردية`** column in `ملف الحركة` and read that row's cash amount from Column G (`مبلغ بنكي`)—even when the cash amount is `0.00`.
   - **Lookup by Cash Amount (e.g., `2780`):** Searches Column G (`مبلغ بنكي`) in `ملف الحركة`. If multiple shifts share the same cash amount, prompts for **الوردية** and disambiguates using the digits after `:`.

5. **Automated Double-Entry Journal Rules & 3-File Export:**
   - Generates 10-column ERP journal rows with Debits assigned to `البنك` and Credits assigned to `دفتر الأستاذ` (without hyphens) using your customizable deposit description template.
   - **Missing Network Single-Line Rule:** If a cashier has a bank terminal amount (`مبلغ بنكي` > 0) for a card network that does **not** exist (or is `0`) in `ملف الحركة`, the app generates **one single double-entry line** directly between `البنك` and the configured Shortage (`عجز`) account for that full amount while preserving that payment network's description.
   - Exports **3 files** (with multi-branch workbooks automatically bundled into a `.zip` archive):
     1. **Clean Journal Entry File (`ملف القيد اليومي`)** — ERP-ready 10-column journal entries.
     2. **Cashier Summary Report (`تقرير الكاشير`)** — Cash, shortages (`عجز`), extras (`زيادة`), discounts (`خصم`), network differences, calculated vs. actual variance, and match status.
     3. **Formatted App File (`ملف الحركة المنسق`)** — Preserves original workbook formulas while highlighting user-entered cells in **Yellow**, matched `مبلغ الفرق` cells in **Gray**, and mismatched variances in **Red**.

6. **Visual First-Time Setup & Settings (`⚙️ إعدادات الشركة والفروع`):**
   - Configure your company name, optional logo, theme colors, branch groups, branch bank/deposit accounts, clearing accounts, and shortage GL account directly in the browser UI and save to `config.json`.

---

## 🚀 Quick Start Installation

### 1. Clone the Repository
```bash
git clone https://github.com/mohammade9/pos-reconciliation-arabic.git
cd pos-reconciliation-arabic
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the Application
```bash
streamlit run app.py
```
*(On macOS, you can also double-click `Run_App.command` after running `chmod +x Run_App.command`.)*

On first launch, if `config.json` does not exist yet, the app opens the **Company & Branch Settings (`⚙️ إعدادات الشركة والفروع`)** screen pre-populated from `config.example.json`. Review or customize your branches and click **💾 حفظ الإعدادات في `config.json` والانتقال للمطابقة** to begin reconciling.

---

## 🤖 Set Up With AI (Ready-to-Copy Prompt)

Want an AI coding assistant (such as Cursor, Claude Code, GitHub Copilot, or ChatGPT) to generate your `config.json` and tailor the app to your company's chart of accounts automatically?

Copy the prompt below, paste it into your AI assistant along with your branch list and account numbers, and let it configure the project for you:

```markdown
I have cloned the open-source `pos-reconciliation-arabic` Streamlit app. Please inspect `config.example.json` and `app.py`, and generate a valid `config.json` file in the project root using my company's details below:

1. **Company Branding:**
   - Arabic Name (`company_name_ar`): [ENTER ARABIC COMPANY NAME]
   - English Name (`company_name_en`): [ENTER ENGLISH COMPANY NAME]
   - Primary Theme Color (`accent_color`): [e.g. #1D54B5]
   - Secondary Theme Color (`accent2_color`): [e.g. #5BC2E7]
   - Currency (`currency`): [e.g. SAR]

2. **Chart of Accounts (Clearing & Shortage GL Codes):**
   - Shortage / Over-Short GL Account (`shortage_account`): [ENTER SHORTAGE GL CODE]
   - Card Clearing Accounts (`clearing_accounts`):
     - Mastercard (`master`): [ENTER GL CODE]
     - Mada (`mada`): [ENTER GL CODE]
     - Visa (`visa`): [ENTER GL CODE]
     - GCCNET (`gccnet`): [ENTER GL CODE]
     - Amex (`amex`): [ENTER GL CODE]
     - Discover (`discover`): [ENTER GL CODE]
     - UnionPay (`unionpay`): [ENTER GL CODE]
     - Maestro (`maestro`): [ENTER GL CODE]

3. **Deposit Description Format (`description_template`):**
   - Use template: `"إيداع بنكي: {deposit_acc}-{network}, {currency}. 000001"` (or customize as needed).

4. **Branch Groups & Branches (`branch_groups`):**
   - Group 1:
     - ID: [e.g. MAIN]
     - Name: [e.g. Main Branches]
     - Company Code in Journal: [e.g. MAIN]
     - Multi-file upload mode (`multi_file`): [true / false]
     - Append branch code suffix `-CODE` to Clearing & Shortage GL accounts (`append_branch_code_to_gl`): [true / false]
     - Branches List (Name, Branch Code, Bank Account, Default Deposit Account, and any per-network Deposit Accounts):
       - [PASTE YOUR BRANCH TABLE OR LIST HERE]

Please create `config.json` with this exact structure and verify that `app.py` loads all branches and accounts cleanly.
```

---

## 📁 Project Structure

- `app.py` — Main Streamlit application (RTL UI, multi-page hover magnifier, accounting engine, and visual settings editor).
- `config.example.json` — Clean template configuration with sample branch groups and placeholder GL codes.
- `config.json` — Your local company configuration (auto-generated via the UI or from `config.example.json`; ignored by Git for privacy).
- `requirements.txt` — Python package dependencies.
- `Run_App.command` — Double-clickable macOS launcher script.
- `.gitignore` — Ensures `config.json`, spreadsheets (`.xlsx`), receipts (`.pdf`, `.png`, `.jpg`), and archives (`.zip`) are never committed to Git.

---

## 📄 License

Released into the public domain under [The Unlicense](LICENSE).
