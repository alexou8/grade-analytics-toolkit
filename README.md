# Grade Analytics & Reporting Toolkit (Excel + Access)

A lightweight analytics and reporting tool built in Microsoft Excel (VBA) that connects to an Access database to import grade data, generate summary statistics, visualize distributions, and export a formatted Word report.

This project focuses on building a repeatable workflow for data selection → validation → calculation → visualization → reporting, using familiar desktop tools.

---

## Key Features

- **Database import (Access .mdb)**: load student, course, and assessment data from a selected database file
- **Flexible filtering**: choose a **course** and **assessment type** to analyze
- **Automatic metrics**:
  - min, max, average
  - median, mode
  - standard deviation
- **Histogram visualization**: generate a distribution chart for the selected grades
- **Word report export**: create a clean, shareable report containing metrics, chart output, and student info
- **One-click workflow** via a custom Excel ribbon tab

---

## Tech Stack

- **Excel VBA** (UI + logic + automation)
- **Microsoft Access (.mdb)** as the data source
- **Office automation** (Excel charts + Word export)

---

## Repository Contents

- `Grade Analytics & Reporting Toolkit.xlsm` — Excel workbook (macros + UI)
- `Registrar.mdb` — sample Access database (student + grades data)
- `/docs` — example exported report and overview documents

> Note: This repo includes sample data for demonstration. Replace with your own database if desired.

---

## Getting Started

### Requirements
- Windows desktop version of **Microsoft Excel** (macros enabled)
- Microsoft Access Database Engine available on your machine (typical with Office installs)

### Setup
1. Download / clone the repo
2. Open `Grade Analytics & Reporting Toolkit.xlsm`
3. Click **Enable Editing** and **Enable Content (Macros)** when prompted
4. In Excel, use the custom ribbon tab: **Grade Analytics / Student Marking Application**
5. Click **Import Student Grades** and select the provided `Registrar.mdb` (or your own)

---

## How to Use

1. **Import Data**
   - Click **Import Student Grades**
   - Select an Access `.mdb` file containing student/course/assessment/grade data

2. **Select Analysis Target**
   - Choose a **Course** and **Assessment Type**
   - The workbook populates:
     - a grade table
     - a student info table
     - computed statistics

3. **Create a Histogram**
   - Click **Graph Data**
   - A histogram chart is generated for the selected grades

4. **Export a Report**
   - Click **Export Data**
   - A Word document is generated containing:
     - selected course + assessment
     - summary metrics
     - histogram visualization
     - student listing

5. **Reset**
   - Click **Clear Data** to clear imported/calculated sheets

---

## What I Practiced / Built

- Designing a **UI-driven workflow** (forms + custom ribbon controls)
- Building a **data pipeline** from a relational source into structured worksheets
- Implementing **descriptive statistics** reliably in Excel
- Creating **repeatable reporting outputs** (export to Word with consistent formatting)
- Handling edge cases like missing selections and clearing state between runs
