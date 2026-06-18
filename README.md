# Osgoode JD Budget Planner

An interactive Excel budgeting tool built for Osgoode Hall Law School students to estimate their income, expenses, and debt across the academic year, and to project how those numbers change over their 1L, 2L, and 3L years.

## What It Does

Students fill in a short set of dropdowns and inputs, and the tool automatically calculates a personalized budget:

- **Background inputs** — study year (1L/2L/3L), domestic/out-of-province/international status, and housing situation (residence, off-campus, or at home)
- **Income & resources** — employment income, savings, investments, family contributions, government aid, scholarships, and other income
- **Expenses & debt** — tuition, rent, food, transportation, health/dental, and other living costs, pre-populated with Osgoode's own cost estimates as a benchmark
- **Budget summary** — automatic surplus/shortfall calculation, with a chart comparing income vs. expenses
- **3-year projections** — expense and income categories are projected forward through 2L and 3L using average annual increase rates, so students can see their full multi-year financial picture, not just the current year
- **Bursary estimation** — estimates likely bursary funding based on calculated shortfall and debt level, using tiered thresholds
- **Helpful links hub** — one-click access to tuition deadlines, government funding, lines of credit, bursaries, scholarships, housing, and other financial aid resources

## How It Was Built

The workbook is built in Excel and automated with VBA, including:

- **Conditional logic** that pre-fills expense estimates based on a student's selected year, residency status, and housing situation, pulling from a built-in reference table of Osgoode's published costs
- **Input validation** that flags ineligible combinations (e.g., selecting "at home" housing while marked as an out-of-province or international student)
- **Dynamic hyperlinking**, where certain links (like health insurance) change destination based on a student's inputs
- **Custom navigation buttons** that move users between sheets (Welcome → Background → Income → Expenses → Summary → Links) without needing to find tabs manually
- **Formula-driven projections** that scale each budget line item forward across all three years of the JD program using category-specific growth rates

All calculations are handled through native Excel formulas and named ranges, so the budget updates live as inputs change, with no manual recalculation needed.

## Why I Built It

As a Student Financial Services Assistant at Osgoode, I saw students repeatedly asking the same questions about what their finances would look like across the full JD program, not just one year. I built this tool to give them a self-serve way to model their own budget, see if they're heading for a shortfall, and get a rough estimate of bursary support, before that conversation needs to happen with our office. Additionally, students were always curious on how to see bursary funding levels before applying, and this calculator helps show students the logic behind bursary funding to help give them a better understanding of how bursary funding is determined. 

## Tech Used

Excel, VBA

## Note

This file is macro-enabled (.xlsm). Excel will prompt you to "Enable Macros" on opening, this is required for the navigation buttons and validation checks to work. The estimates provided are for planning purposes only and are not official financial figures.
