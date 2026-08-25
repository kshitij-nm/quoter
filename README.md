# Quotation Pro

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![PySide6](https://img.shields.io/badge/PySide6-Qt-green.svg)](https://wiki.qt.io/Qt_for_Python)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Quotation Pro** is a powerful, modern desktop application designed to streamline the creation, management, and export of product and service quotations. Built with a highly responsive **PySide6** graphical interface, it acts as a localized CRM and quotation engine, seamlessly bridging a live UI with underlying Excel database files.

---

## ✨ Key Features

* **Smart Cart System:** Instantly add products or services to a customer's cart via single-click checkboxes, complete with dynamic quantity and discount-percentage inputs.
* **Intelligent Draft Management:** Never lose your work. Carts are silently serialized into local `.json` files in real-time. Typing a returning customer's name instantly restores their active session and re-checks their items in the database.
* **Excel-Backed Database:** The application reads from and auto-saves to `catalog.xlsx` in the background, making external database edits completely compatible.
* **Live Duplicate Detection:** Utilizes hidden Unique Identifiers (UIDs) to track rows. If a user accidentally creates an identical twin of an existing product, the system instantly flags older iterations with a soft red highlight.
* **Universal Search & Filtering:** A highly optimized search engine that scans across all columns simultaneously, combined with dropdown Category and Supplier filters.
* **Persistent UI States:** Employs `QSettings` to memorize user preferences between sessions, including Light/Dark mode, splitter geometry, sidebar states, and exact table column widths.
* **One-Click Export:** Finalized carts are automatically processed and exported as heavily formatted `.xlsx` quotation files ready to be sent to clients.

---

## 🛠️ Technology Stack

| Component | Technology | Description |
| :--- | :--- | :--- |
| **Language** | Python 3.x | Core application logic and data processing. |
| **GUI Framework** | PySide6 | Modern, dark-mode-compatible user interface. |
| **Data Handling** | Pandas | High-performance manipulation of catalog datasets. |
| **File I/O** | OpenPyXL & JSON | Generation of quotes and serialization of cart drafts. |

---

## 🚀 Installation & Setup

### 1. Prerequisites
Ensure you have Python 3.8 or higher installed on your system. 

### 2. Clone the Repository
```bash
git clone [https://github.com/yourusername/quotation-pro.git](https://github.com/yourusername/quotation-pro.git)
cd quotation-pro
```