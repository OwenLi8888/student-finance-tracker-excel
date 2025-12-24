# Excel-Based Student Finance Management Tool

## Overview
This project is an Excel-based financial management and decision-support system designed to help students track income, expenses, and savings goals while receiving automated financial insights. The tool reduces manual bookkeeping and cognitive load by automating data collection, analysis, visualization, and feedback through Excel and VBA.

## Problem
Many students track money in scattered notes, banking apps, or mental math. This creates uncertainty, missed goals, and financial stress. This tool centralizes all financial activity in one system and converts raw transactions into structured insight, projections, and actionable advice.

## User Workflow

### 1. Data Input
Users enter income and expenses on the `Expenses & Income` sheet using a VBA-powered `Add Item` form. A minimum of three months of historical data is required so the system can identify trends and generate forward-looking estimates. Input validation and reset functionality are handled through VBA to prevent data corruption.

### 2. Goal Management
Users define financial goals such as savings targets on the `Goals` sheet using a dedicated form. Progress toward each goal is recorded over time and automatically summarized for the selected month.

### 3. Automated Analysis and Visualization
When the `Generate Financial Data` button is clicked, VBA aggregates all transactions for the selected month and updates:
- Income vs expense comparisons
- Category-level spending breakdowns
- A short-term forecast for next-month income and expenses based on recent trends

All dashboards and charts update automatically without manual filtering or recalculation.

### 4. Goal Progress Tracking
For the selected month, the system computes progress toward each financial goal and displays:
- Percentage completion
- Visual progress bars on the Home Page

This provides immediate feedback on whether the user is on track.

### 5. Financial Advice Engine
The tool compares actual spending against an ideal allocation model:
- Savings: 25%
- Expenses: 45%
- Investments: 20%
- Emergency Fund: 10%

Based on deviations from these targets, a VBA-based rules engine generates tailored financial advice and motivational feedback.

## Technical Implementation

- **Excel formulas** for aggregation, validation, and summary metrics  
- **VBA modules** for:
  - Structured data input through UserForms
  - Automated sheet clearing and workflow control
  - Goal progress calculations
  - Financial advice logic  
- **Dynamic charts and tables** for visualizing spending, trends, and goal progress  
- **Modular VBA design** separating UI logic, analytics, and decision rules for maintainability

All VBA logic is exported into the `/vba` folder to make the system readable, versioned, and auditable.

## Impact
By automating bookkeeping, analysis, and feedback, the tool removes guesswork and decision fatigue from financial management. Users no longer need to manually interpret spreadsheets to understand their financial position, helping reduce stress and enabling more confident financial planning.
