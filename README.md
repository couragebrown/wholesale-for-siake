# Siaka Wholesale Flow

A fast, responsive desktop management system designed for wholesale distributors, dealer accounts, product catalog management, supplier restocks, and multi-channel payment tracking.

## 🚀 Quick Start
To launch the application on Windows:
- Double-click **`Launch_Wholesale_App.bat`** (or **`Launch_Wholesale_App.vbs`** for quiet launch).
- The local backend service starts automatically and opens the desktop application window.

## 📦 Features
- **📊 Real-time Dashboard**: Live financial overview showing Total Sales Billed, Cash Collected, Outstanding Dealer Debt, and Total Orders.
- **📝 Wholesale Orders**: Fast multi-item order creation with real-time stock balance checking and instant payment logging.
- **👥 Dealer & Debt Management**: Track customer contact info, territory, credit terms, lifetime purchases, and balance due.
- **🏭 Suppliers & Payables**: Manage supplier procurement, inventory restocking, payment disbursements, and outstanding payables.
- **💵 Payment Ledger**: Complete chronological record of payments with payment method tracking (Cash, MoMo, Bank, Cheque), reference numbers, and notes.
- **📄 Official Statement Printing**: High-resolution, multi-page ready A4 statement generator with official branding, balance summaries, and acknowledgment signature blocks.
- **📊 Live Excel Export (`.xlsx`)**: One-click native Excel generation formatting all records across dedicated sheets (Dashboard, Products, Dealers, Orders, Payment History, Suppliers).
- **💾 Local Data Security**: Data is automatically persisted to local JSON on your PC and Excel workbook.

## 🛠️ Tech Stack
- Frontend: Vanilla HTML5, CSS3 (Modern Glassmorphism & Plus Jakarta Sans typography), ES6 JavaScript
- Backend: Lightweight PowerShell REST API Server (`server.ps1`)
- Storage: Local JSON database (`app/data/wholesale_data.json`) and Excel COM Integration
