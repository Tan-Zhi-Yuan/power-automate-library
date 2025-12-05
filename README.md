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

## 🛠️ Technical Implementation & Workflow

### Development Philosophy
This repository demonstrates a **"Hybrid Engineering"** approach. My focus is on designing robust algorithms and business logic, while utilizing modern AI tools to accelerate syntax generation and optimization.

* **Architectural Design (Human-Led):**
    * Defined the **O(N)** strategy for the logistics synchronization to solve the runtime freeze.
    * Created the **"Line 0 Protection"** business rule to ensure data integrity during updates.
    * Designed the **Matrix Logic** for the RPA bot to ensure non-technical users could update rules without touching code.

* **AI-Assisted Optimization (Tool-Assisted):**
    * Leveraged LLMs to refactor legacy nested loops into **Scripting.Dictionary** (Hash Maps) for performance.
    * Generated boilerplate error-handling blocks (`On Error GoTo`) to ensure production stability.
    * Used AI for rapid syntax translation between Excel Formulas and VBA Logic.

### Key Technical Concepts Demonstrated
| Concept | Application in this Repo |
| :--- | :--- |
| **Big O Optimization** | Reduced execution time from exponential ($O(N^2)$) to linear ($O(N)$) using Dictionaries. |
| **ETL Processes** | Automated the Extraction, Transformation, and Loading of data between unstructured reports and master tables. |
| **Defensive Coding** | Implemented `Option Explicit`, type checking, and object validation to prevent runtime crashes. |
| **Low-Code Integration** | Bridged the gap between Excel (VBA) and the OS (Power Automate) to handle file system operations. |
