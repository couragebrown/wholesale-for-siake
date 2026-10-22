# Siaka Wholesale Flow - PC Edition

A high-performance desktop management system tailored for wholesale distributors, dealer debt tracking, product stock management, supplier procurement, and multi-channel payment reconciliation.

---

## ⚡ 1-Click Installation (Zero Data Loss Architecture)
The application is designed to run directly from **Drive D (`D:\excel`)**, keeping all business records and databases safely insulated from Windows OS resets, formatting, and Drive C malware attacks.

To install or reinstall on your PC:
1. Double-click **`Install_Siaka_Wholesale_Flow.bat`**.
2. The installer automatically:
   - Verifies and creates all required local storage folders (`app\data\`, `app\data\backups\`, `app\data\app_profile\`).
   - Initializes a clean, blank production database (`wholesale_data.json`).
   - Creates a **Desktop Shortcut** with the high-resolution logo.
   - Creates a **Start Menu Shortcut** for instant search access.

---

## 🚀 Launching the App
After installation, you can launch Siaka Wholesale Flow at any time:
- **From Desktop**: Double-click the **Siaka Wholesale Flow** shortcut.
- **From Start Menu**: Press the Windows Key and search for **Siaka Wholesale Flow**.
- **From Folder**: Double-click **`Launch_Wholesale_App.vbs`** (quiet desktop mode) or **`Launch_Wholesale_App.bat`**.

---

## 🛡️ Zero Data Loss Safeguards
Your business data is guarded by a 3-layer protection system:
1. **Atomic File Writes**: The server writes data to `wholesale_data.json.tmp` and replaces the file atomically. This eliminates the risk of 0-byte file truncation or file corruption caused by abrupt power outages or PC restarts.
2. **Automated Rolling Backups**: Before any change is written to disk, a timestamped snapshot is archived to `app\data\backups\backup_YYYYMMDD_HHMMSS.json`. The system retains the last 100 historical snapshots.
3. **In-App Data Safety & Recovery Hub**: Click the **`🛡️ Data & Backups`** button in the top navigation bar to:
   - Download an instant complete JSON database backup to your flash drive or external disk.
   - Trigger instant server-side snapshot backups on demand.
   - Restore your database from any past snapshot or uploaded backup file with a single click.
4. **Auto-Recovery on Startup**: If `wholesale_data.json` is ever missing or damaged, the server automatically recovers from the latest backup snapshot.
5. **Drive D Isolation**: All application files, databases, and logs live on Drive D (`D:\excel`), ensuring your accounting records survive even if Drive C is wiped or repaired.

---

## 📦 Core Modules
- **📊 Real-time Dashboard**: Live financial overview with customizable date filters (Today, This Week, This Month, Last 30 Days, Custom Range).
- **📝 Wholesale Orders**: Fast multi-item billing with live inventory checks, unit conversion, and immediate payment recording.
- **👥 Dealer & Debt Management**: Track customer contact info, territory, credit terms, lifetime purchases, and balance due.
- **🏭 Suppliers & Payables**: Manage supplier procurement, inventory restocking, payment disbursements, and outstanding payables.
- **💵 Payment Ledger**: Complete chronological record of payments with payment method tracking (Cash, MoMo, Bank, Cheque), reference numbers, and notes.
- **📄 Official Statement Printing**: High-resolution, multi-page ready A4 statement generator with official branding, balance summaries, and acknowledgment signature blocks.
- **📗 Live Excel Export (`.xlsx`)**: One-click native Excel generation formatting all records across dedicated sheets (Dashboard, Products, Dealers, Orders, Payment History, Suppliers).
- **🗂️ Archived Demo Data**: Your previous demo test records have been safely preserved in `app\data\demo_backup_wholesale_data.json` for reference.

---

## 🛠️ System Requirements
- Windows 10 or Windows 11 (64-bit or 32-bit)
- Microsoft Edge or Google Chrome (used for the borderless app window)
- Microsoft Excel (optional, for native `.xlsx` generation)

