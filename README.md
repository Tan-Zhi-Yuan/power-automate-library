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

## 🛠️ Technical Implementation & Workflow

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
* **Power Query (M Code):** used for robust ETL data transformation and cleaning steps.
