# 🤖 RPA Automation Library

## Overview
A collection of robotic process automation (RPA) workflows built with **Microsoft Power Automate Desktop**. These "digital workers" handle repetitive file management and data entry tasks.

---

## 📂 Featured Workflows

### 1. Smart File Router (Matrix Logic)
**File:** [SmartFileRouter.txt](./SmartFileRouter.txt)

> **The Challenge:**
> The procurement team receives unorganized PDFs (Invoices, POs) that need to be filed into specific directories based on Supplier Name and Document Type. Manual sorting was slow and error-prone.
>
> **The Solution:**
> * **Matrix-Based Rules:** The bot reads an external Excel file (`Rules.xlsx`) to determine where files should go. This allows users to update sorting rules without modifying the bot's code.
> * **Multi-Criteria Matching:** Checks that the filename contains the *Supplier Name* AND starts with the correct *Prefix* before moving.
> * **Audit Logging:** Generates a text report of every file moved during the session.

**Logic Visualization:**
![Flow Diagram](./flow-diagram.png)
*(Make sure you uploaded an image named flow-diagram.png for this to show)*

---

## 🛠️ Tech Stack
* **Tool:** Microsoft Power Automate Desktop (PAD)
* **Integrations:** Excel, File System, Outlook
* **Concepts:** Loops, Conditionals, Error Handling, UI Automation
