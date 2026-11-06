# Data Job Market Analysis in Excel: Salary Calculator Dashboard & Skills vs Salary Analysis

## Introduction

What does the data job market actually pay, and which skills are worth learning? This repository answers that with two Excel projects built on 32,672 real data-job postings from 2023:

1. **Project 1: Data Science Salary Calculator.** An interactive dashboard where you pick a *job title*, *country* and *employment type* and instantly see the median salary, job count, top job platform and salary charts.
2. **Project 2: Skills vs Salary Analysis.** A pivot-table analysis of the 10 most requested skills: which are most in demand, which pay the most, how pay differs between the US and the rest of the world, and whether jobs that ask for more skills pay more.

**Contents**

- [Background](#background)
- [Tools I Used](#tools-i-used)
- [Project 1: Salary Calculator Dashboard](#project-1-data-science-salary-calculator-dashboard)
- [Project 2: Skills vs Salary Analysis](#project-2-skills-vs-salary-analysis)
- [What I Learned](#what-i-learned)
- [Conclusions](#conclusions)

---

## Background

The goal was to build something useful for anyone choosing a data career path: a tool to benchmark salaries, plus an evidence-based view of which skills the market rewards.

**Questions I set out to answer**

1. What is the median salary for a given data role, in a given country, for a given employment type?
2. Which skills are most in demand?
3. Which skills pay the best?
4. How much does pay differ between the US and the rest of the world?
5. Do roles that require more skills pay more?

**The dataset**

| Item | Detail |
|---|---|
| Postings | 32,672 job postings |
| Period | January to December 2023 |
| Job titles | 10 (Data Analyst, Senior Data Analyst, Business Analyst, Data Scientist, Senior Data Scientist, Data Engineer, Senior Data Engineer, Machine Learning Engineer, Software Engineer, Cloud Engineer) |
| Geography | 111 country values; the United States makes up about 77% of postings (25,141) |
| Salary coverage | 22,036 postings (67%) have an advertised yearly salary, which is what all medians in this repo use |
| Key columns | job title, location, country, schedule type, platform, posted date, yearly salary, skills list |

> Dataset source: **[add link / credit here]**

---

## Tools I Used

- **Microsoft Excel** (Excel 2021 / Microsoft 365)
  - Excel Tables and named ranges
  - Dynamic array formulas: `SORT`, `FILTER`, `XLOOKUP`
  - Array formulas: `MEDIAN(IF(...))`
  - Data validation drop-downs
  - PivotTables, PivotCharts and Slicers
  - Region Map, scatter, and combo (column + line) charts
- **GitHub** for version control and publishing the analysis


## Project 1: Data Science Salary Calculator (Dashboard)
---<img width="1151" height="486" alt="dashboard" src="https://github.com/user-attachments/assets/661a07d9-6f37-4aac-8c17-a8458b6f419f" />

*The Salary Calculator dashboard in its default view: Data Analyst, United States, Full-time.*

### What it does

Choose three inputs and the whole dashboard recalculates:

| Input | Options |
|---|---|
| **Job Title** | 10 data roles |
| **Country** | Any country in the dataset |
| **Type** | Full-time, Contractor, Part-time, Internship, Temp work |

| Output | What it shows |
|---|---|
| **Median Salary** (KPI card) | Median yearly salary for the selected title, country and type |
| **Job Count** (KPI card) | Number of matching postings |
| **Top Job Platform** (KPI card) | The job board with the most matching postings |
| **Median salary by job title** (bar chart) | All 10 roles side by side, with the selected role highlighted |
| **Median salary by job type** (bar chart) | Compares employment types for the selected role and country, with the selected type highlighted |
| **Median salary by country** (map) | Salary map for the selected role and type |

For the default view (Data Analyst, United States, Full-time) the dashboard shows **6,480 postings**, **Indeed** as the top platform, and a **median salary of $90,000**.

### How it works

- The raw data lives in an Excel Table called `jobs` (32,672 rows) on a hidden sheet.
- The three drop-downs are stored as named ranges (`title`, `country`, `type`).
- Helper sheets (hidden) calculate a median for every title, type and country using array formulas, for example:

```excel
=MEDIAN(IF((jobs[job_title_short]=A2) *
           (jobs[salary_year_avg]<>0) *
           (jobs[job_country]=country) *
           (ISNUMBER(SEARCH(type, jobs[job_schedule_type]))),
           jobs[salary_year_avg]))
```

- `SORT` + `FILTER` drop empty results and rank the bars, and `XLOOKUP` pulls out the selected value.
- The bar charts highlight the selection using two series: one returns the value when the bar matches the selection, the other returns `#N/A` (so it draws nothing), and vice versa.
- The KPI cards are text boxes linked to named cells (`median_salary`, `job`, `platform`).

### Insights from the dashboard

**Median yearly salary by job title (United States, Full-time view)**

| Job title | Median salary |
|---|---:|
| Senior Data Scientist | $155,000 |
| Senior Data Engineer | $150,000 |
| Machine Learning Engineer | $150,000 |
| Data Scientist | $130,000 |
| Data Engineer | $125,595 |
| Software Engineer | $125,000 |
| Cloud Engineer | $115,000 |
| Senior Data Analyst | $110,000 |
| Data Analyst | $90,000 |
| Business Analyst | $90,000 |

**Median yearly salary by employment type (Data Analyst, United States)**

| Type | Median salary | Postings with a yearly salary |
|---|---:|---:|
| Full-time | $90,000 | 4,219 |
| Part-time | $82,500 | 75 |
| Contractor | $81,477 | 138 |
| Temp work | $75,000 | 21 |
| Internship | $65,000 | 14 |

Key takeaways:

- **Seniority pays.** A Senior Data Analyst earns about **22% more** than a Data Analyst ($110,000 vs $90,000).
- **Analyst roles sit at the bottom of the data pay scale.** Data Analyst and Business Analyst both have a $90,000 median, about $65,000 below the top-paid Senior Data Scientist.
- **Machine Learning Engineers match senior engineering pay.** At $150,000 they earn the same as Senior Data Engineers.
- **Full-time roles give the best and most reliable benchmark.** The non-full-time medians rest on small samples (14 to 138 postings), so treat them as indicative only.
- **The country map needs context.** The highest values on the map come from countries with very few postings (for example, Belarus's $400,000 comes from a single posting).

<img width="1154" height="501" alt="dashboard 2" src="https://github.com/user-attachments/assets/87983252-9c57-435b-8404-226ba62727ca" />


---

## Project 2: Skills vs Salary Analysis

This analysis uses PivotTables, PivotCharts and Slicers (Country and Job Title) on the same dataset. The figures below are for the **United States**, where the sample is largest.

### 2.1 What are the most in-demand skills?

<img width="1060" height="311" alt="analysis" src="https://github.com/user-attachments/assets/f15677aa-e83f-4ae5-86bf-aa6144e9df69" />

*Count of the top 10 skills mentioned in US job postings.*

| Rank | Skill | Mentions in postings |
|---:|---|---:|
| 1 | SQL | 14,444 |
| 2 | Python | 13,314 |
| 3 | Tableau | 5,799 |
| 4 | R | 5,743 |
| 5 | Excel | 5,342 |
| 6 | AWS | 4,885 |
| 7 | SAS* | 4,114 |
| 8 | Spark | 3,611 |
| 9 | Azure | 3,339 |
| 10 | Power BI | 3,044 |



- **SQL and Python dominate.** SQL appears in roughly **57%** of US postings and Python in roughly **53%**, more than double the next skill.
- **BI tools come next.** Tableau is the most requested visualisation tool, with Power BI behind it.
- **Excel is still in the top 5.** It remains a core requirement even in a market focused on code.

### 2.2 Which skills pay the most?

<img width="1178" height="305" alt="analysis 2" src="https://github.com/user-attachments/assets/1916c240-40eb-432b-94b8-58508700c88e" />


| Skill | Median salary (USD) | Skill Count |
|---|---:|---:|
| Spark | $142,500 | 3,611 |
| AWS | $136,980 | 4,885 |
| Azure | $128,409 | 3,339 |
| Python | $128,050 | 13,314 |
| R | $121,500 | 5,743 |
| SQL | $120,000 | 14,444 |
| Tableau | $113,775 | 5,799 |
| SAS | $109,159 | 4,114 |
| Power BI | $101,000 | 3,044 |
| Excel | $92,500 | 5,342 |

*Overall median across these skills: $120,000.*

- **Big-data and cloud skills pay the most.** Spark, AWS and Azure take the top three spots. Spark's median is about **54% higher** than Excel's.
- **Python is the best mix of pay and demand.** At $128,050 it pays about $8,000 more than SQL and is requested almost as often.
- **SQL is the safest foundation.** It is the most requested skill and sits at the market median ($120,000).
- **Excel and Power BI pay the least of the ten.** They are still widely required, but on their own they are associated with lower-paid roles.

### 2.3 How does pay differ between the US and the rest of the world?

*The table below is from the `Salary analysis` sheet
:*

<img width="822" height="303" alt="analysis 3" src="https://github.com/user-attachments/assets/0c965656-144b-42fc-9572-3140bce3deb4" />

| Job title | US median | Non-US median | Difference |
|---|---:|---:|---:|
| Senior Data Scientist | $155,000 | $155,000 | 0% |
| Machine Learning Engineer | $150,000 | $101,029 | -33% |
| Senior Data Engineer | $150,000 | $147,500 | -2% |
| Data Scientist | $130,000 | $119,550 | -8% |
| Data Engineer | $125,000 | $123,500 | -1% |
| Software Engineer | $125,000 | $89,100 | -29% |
| Cloud Engineer | $115,000 | $89,100 | -23% |
| Senior Data Analyst | $110,000 | $111,175 | +1% |
| Data Analyst | $90,000 | $90,000 | 0% |
| Business Analyst | $90,000 | $75,000 | -17% |
| **All roles** | **$118,940** | **$111,175** | **-7%** |

- **Overall, the US premium is modest (about 7%).**
- **The gap depends heavily on the role.** Machine Learning Engineer, Software Engineer and Cloud Engineer show the biggest US premiums (23% to 33%).
- **Core data roles are priced almost globally.** Data Analyst, Data Engineer and the senior data roles are within a few percent of the US median outside the US.
- **Sample-size warning:** Cloud Engineer has only 65 salary postings in total (18 in the US), so treat that row with caution.

### 2.4 Do jobs that require more skills pay more?

<img width="1259" height="355" alt="Analysis 4" src="https://github.com/user-attachments/assets/425452e5-6f4d-4e44-a202-565835cbe5ff" />
*Each point is a job title: average number of skills listed per posting (x) against median salary (y).*

| Job title | Median salary | Skills per posting |
|---|---:|---:|
| Senior Data Scientist | $155,000 | 5.2 |
| Senior Data Engineer | $150,000 | 8.3 |
| Machine Learning Engineer | $150,000 | 4.7 |
| Data Scientist | $130,000 | 5.0 |
| Software Engineer | $125,000 | 5.1 |
| Data Engineer | $125,000 | 6.9 |
| Cloud Engineer | $115,000 | 4.8 |
| Senior Data Analyst | $110,000 | 4.4 |
| Data Analyst | $90,000 | 3.6 |
| Business Analyst | $90,000 | 3.3 |
| **Average** | **$118,940** | **5.0** |

- **There is a moderate positive relationship** (correlation about **0.65**, with only 10 job titles). More skills tend to go with higher pay, but it is far from a rule.
- **Role matters more than skill count.** Senior Data Engineers list the most skills (8.3) but are not the best paid, while Machine Learning Engineers earn $150,000 with fewer than average (4.7).
- **Analysts list the fewest skills and earn the least.** The step from Data Analyst to Senior Data Analyst comes with roughly one extra skill per posting (3.6 to 4.4) and about 22% higher pay.
- This is an association between job titles, not proof that adding a skill causes a raise.

---

## What I Learned

- **Building an interactive dashboard in Excel.** Drop-downs, named ranges and linked text boxes turn a static workbook into a calculator.
- **Dynamic and array formulas.** `MEDIAN(IF(...))` with multiple conditions, `SORT`, `FILTER`, `Power Pivot`, `Power Query`, `DAX`, `ETL`, `Measures`, `Data Modeling` and `XLOOKUP`, to build a clean, ranked calculation layer.
- **Conditional chart highlighting.** Using a value/`#N/A` pair of series to highlight the selected bar.
- **PivotTables and Slicers.** Slicing the same dataset by country and job title, and building combo and scatter charts on top.
- **Data cleaning matters as much as analysis.** Checking the numbers against the raw data exposed duplicate skill tags and a mislabelled country (details below).
- **Medians over averages.** Salary data is skewed by outliers, so the median is the honest summary.

---
---

## Conclusions

1. **Learn SQL and Python first.** They are by far the most requested skills, and Python also pays above the market median.
2. **Cloud and big-data skills carry the biggest pay premium.** Spark, AWS and Azure lead the salary ranking.
3. **Excel alone is not enough.** It remains widely required but has the lowest median of the top 10 skills.
4. **Seniority and specialisation drive pay.** Senior titles and Machine Learning / Engineering roles clearly out-earn analyst roles.
5. **Location matters, but only for some roles.** The US premium is large for ML, Software and Cloud Engineers, and small for core data roles.
6. **More skills help, but they are not the whole story.** Pay rises with skill breadth on average, yet role and seniority matter more.


