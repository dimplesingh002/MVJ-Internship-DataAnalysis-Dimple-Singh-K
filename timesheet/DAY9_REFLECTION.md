# Day 9: Weekly Review & Reflection Note

## Phase 1 Review: SQL & Database Foundations
* **Key Technical Skills Mastered:**
  * Database design and SQLite setup (`online_retail.db`).
  * Core SQL aggregation: `SUM()`, `AVG()`, `COUNT(DISTINCT)`.
  * Advanced SQL analytics: Common Table Expressions (CTEs), `LAG()` for MoM revenue tracking, and `RANK() OVER()` window functions.
* **Project Completed:** 
  * **Project 1 (KPI Analysis Pack):** Successfully extracted Total Revenue, Average Order Value (AOV), top products, and repeat customer distributions, documenting all findings in `project1/README.md` and `project1/queries.sql`.

## Challenges & Solutions
* **Challenge:** Handling missing descriptions, canceled orders, and multi-line invoice aggregations in raw SQLite data.
* **Solution:** Used `WHERE Description IS NOT NULL` clauses and structured CTEs to isolate clean, actionable sales metrics.

## Goal for Phase 2
* Transition from SQL-based queries to Python/pandas (`KG-PD`) to replicate KPI calculations, perform automated data cleaning, and generate visualization plots.
