# 🤖 RPA & Cloud Automation Library

## Overview
A collection of robotic process automation (RPA) workflows and Cloud Scripts built with **Microsoft Power Automate** and **TypeScript (Office Scripts)**. These "digital workers" handle repetitive file management, data entry, and report generation tasks.

---

## 📂 Featured Workflows

### 1. Smart File Router (Desktop RPA)
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

### 2. Auto-Refresh & KPI Snapshot (Cloud Script)
**File:** [Refresh_and_Snapshot_KPIs.ts](./Refresh_and_Snapshot_KPIs.ts)

> **The Challenge:**
> Management required a "frozen" snapshot of daily KPIs immediately after the data refresh. Doing this manually led to timing errors (copying data before the refresh finished) or formatting inconsistencies.
>
> **The Solution:**
> * **Async Orchestration:** Utilized TypeScript's `async/await` pattern to ensure `workbook.refreshAllDataConnections()` completes before any data manipulation occurs.
> * **Automated History:** Inserts a new column to "push" old data to the right, creating an infinite rolling history of daily snapshots without manual intervention.
> * **Cloud Compatible:** Unlike VBA, this script runs in Excel Online and can be triggered via Power Automate Cloud Flows.

---

## 🛠️ Technical Implementation & Workflow

* **Power Automate Desktop (PAD):** Used for tasks requiring access to the local file system (e.g., moving PDFs, interacting with legacy desktop apps).
* **Office Scripts (TypeScript):** Used for Excel-specific logic that needs to run in the cloud (SharePoint/OneDrive) without a user logged in.
* **Separation of Concerns:** Hard-coded logic is avoided. Scripts rely on external "Config Tables" (Excel) so business users can update rules (e.g., new folder paths or supplier names) without touching the code.

---
## ⚙️ Engineering Philosophy & AI-Augmented Workflow

This repository demonstrates a **Modern "Hybrid" Development Strategy**.
In an era where syntax is cheap but logic is expensive, my focus is on **Architectural Design** and **Business Value**. I utilize Large Language Models (LLMs) as a "force multiplier" to accelerate development across my tech stack.

### 🧠 The Division of Labor
* **Human Architect (Me):**
    * **Business Logic:** Defining the complex rules (e.g., FIFO Valuation, Line 0 Protection, File Routing Matrix).
    * **System Architecture:** Designing how disparate systems (Excel, Outlook, File System) interact securely.
    * **Validation:** Reviewing, debugging, and stress-testing all code for accuracy and edge cases.
* **AI Assistant (Tooling):**
    * **Syntax Generation:** Rapidly translating logic into specific syntax for **VBA**, **M Code (Power Query)**, and **DAX**.
    * **Pattern Optimization:** Refactoring nested loops into $O(N)$ Hash Maps (`Scripting.Dictionary`) for performance.
    * **Boilerplate:** Generating standard error-handling blocks and UI elements.

### 🛠️ Tech Stack & Methodology
This workflow allows me to maintain high standards of code quality across multiple domains:
* **VBA:** used for Event-Driven Automation and Object Model manipulation.
* **Power Automate (RPA):** used for OS-level orchestration and "Low-Code" integration.
* **TypeScript (Office Scripts):** used for cloud-native Excel automation and API-level data handling.
* **Power Query (M Code):** used for robust ETL data transformation and cleaning steps.
