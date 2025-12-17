# 🦁 **CSK IPL Analysis Dashboard (2008–2025)**
### Python | JavaScript | Excel | Power BI

<img src="dashboard/imgs/csk_logo.png" width="180"/> 

This project presents a fully interactive and dynamic analysis of the Chennai Super Kings (CSK) across all IPL seasons from **2008 to 2025** (except 2016–17 due to the team’s suspension).

It covers **batting, bowling, fielding, captaincy stats, match results, partnerships, and overall team performance**, offering deep insights into player contributions and seasonal trends.


## 📌 1. Project Overview
This data-driven analytics system was built to analyze every major performance metric of CSK using:

- Automated **web scraping** (JavaScript)
- Data cleaning, merging & transformation (Python + Excel)
- Data modeling using a **Star Schema**
- KPI creation through **DAX measures**
- Multi-page interactive dashboard in **Power BI**

The final output is a professional sports analytics dashboard designed for fans, analysts, and recruiters evaluating data analytics skills.



## 💡 2. Business Problem & Objective

IPL generates massive volumes of cricket statistics across seasons, but raw scorecards
and tables do not easily reveal patterns related to consistency, player impact,
team evolution, and long-term performance trends.

This project addresses that gap by building a centralized analytics system that:
- Consolidates CSK’s cricket data across seasons
- Generates insights at team, player, and match levels
- Enables season-wise and overall comparisons   
- Supports decision-making with an interactive dashboard
- Presents clean, accurate, and visually compelling analytics


## 🧰 3. Tech Stack & Tools Used

### **🟨 JavaScript – Puppeteer + Cheerio**
Used for scraping CSK's:
- Season-wise batting records  
- Bowling records  
- Fielding records  
- Player profiles  
- Match results  
- Partnership data  

### **🐍 Python**
Used for:
- Merging scraped files  
- Cleaning inconsistent data  
- Removing duplicates & null values  
- Splitting columns  
- Combining season-wise files into a unified structure  
- Exporting clean datasets for Power BI  

### **📊 Excel**
Used for:
- Manual review of scraped data  
- Final adjustments  
- Sheet structuring (seasonal sheets and combined sheets)  

### **📉 Power BI**
- Data modeling (Star Schema)
- DAX measures (KPIs for batting, bowling, fielding, and team performance)
- Power Query Editor for shaping data
- Multi-page dashboard design
- Dynamic filtering (Season, Player, Opposition, Ground)


## 🏗️ 4. Data Pipeline (End-to-End Workflow)

```text
Web Scraping (ESPN Cricinfo Statsguru)
        ↓
Raw Json/CSV Files
        ↓
Python Cleaning, Merging & Transformation
        ↓
Final Clean Excel Files (Batting, Bowling, Fielding, Match Results, Players, Partnerships)
        ↓
Power BI Data Modeling (Dimensions + Facts)
        ↓
DAX Measures Creation
        ↓
Interactive Dashboard with 8 Pages

```

**Step 1 — Web Scraping**

Automated scraping was performed using Puppeteer and Cheerio to extract structured tables for all available CSK IPL seasons (excluding 2016–17) from ESPN Cricinfo Statsguru.

- Automated Puppeteer scripts visited Statsguru for each season.
- Cheerio extracted structured tables:
- Batting by season
- Bowling by season
- Fielding by season
- Career records
- Partnerships
- Match results

All scraped data was exported as raw JSON/CSV files.

[View Player Scraper Code](src/scraper/scrape_players_seasonwise.js) | [View Team Performance Scraper Code](src/scraper/scrape_team_performance.js)


**Step 2 — Python Transformation**

Using Python (Pandas), all season-wise data was cleaned, standardized, and merged:

Standardization & Cleaning:
- Processed all season-wise batting, bowling, and fielding Excel files using openpyxl and pandas.
- Standardized column names, removed unused/system-generated fields, and inserted key identifiers such as Player_ID, Full_Name, and Season.
- Applied sport-specific transformations:
    - Batting — calculated HS_Numeric, extracted Not_out_status, renamed 100/50 columns, reordered metrics.
    - Bowling — cleaned BBI, renamed 4-for/5-for fields, ensured text formatting, removed irrelevant columns.
    - Fielding — standardized Ct_Wk, Ct_Fi, MD, and eliminated empty columns.
- Saved cleaned structured files to /data/final_excel/.
- Standardized player names and merged all cleaned batting, bowling, and fielding datasets into a unified player master table.

Aggregation & Career Stats:
- Aggregated season-wise stats into career totals using grouping logic for batting, bowling, and fielding performance.
- Computed each player's career span (first–last season) and total seasons played.
- Ensured consistent Player_ID assignment across all merged tables.
- Generated a comprehensive Excel output containing four sheets — Batting, Bowling, Fielding, Combined.
- Applied formatting safeguards (BBI, MD as text) for accurate downstream loading into Power BI.
- **Captaincy stats** were **manually compiled** in Excel due to data complexity.

[View Python Cleaning & Merging Code](src/etl/clean_excel.py) | [View Overall Records Code](src/etl/overall_records.py)

**Step 3 — Excel Validation**

Before importing to Power BI, each sheet underwent manual validation:

- Checking duplicates
- Verifying player IDs
- Checking missing match results
- Cleaning categorical values
- Finalizing sheets for Power BI import

**Step 4 — Import to Power BI**

Loaded all 5 processed Excel files.

Used Power Query to:

- Remove extra columns
- Ensure data types
- Append sheets
- Normalize text formats
- Create relationships


**Step 5 — Build Data Model**

- Fact & Dimension tables created
- Relationships made one-directional
- Measure table created for all DAX KPIs and calculations


**Step 6 — Visualization & Dashboard Building**

- Built 8 pages based on cricket analysis logic
- Added slicers, drill-through, tooltips
- Applied CSK theme


## 🗂️ 5. Data Model (Star Schema)

Dimensions
```
Dim_Players
Dim_Seasons
Dim_Opposition
Dim_Ground
```

Fact Tables
```
Fact_Batting
Fact_Bowling
Fact_Fielding
Fact_Partnerships
Fact_MatchResults
Fact_TeamPerformance
Fact_PointsTable
Fact_Captaincy
```
All relationships flow from Dimensions → Facts ensuring clean filtering.

![Data Model](dashboard/imgs/datamodel.png)


## 📐 6. DAX Measures (Key KPIs)

**🏏 Batting KPIs**

```DAX
Total_Runs = SUM(Fact_Batting[Runs])
Avg_Batting= AVERAGE(Fact_Batting[Ave])
SR_Bat = AVERAGE(Fact_Batting[SR])
Total_Fours = SUM(Fact_Batting[Fours])
Total_Sixes = SUM(Fact_Batting[Sixes])
Fifties = SUM(Fact_Batting[Fifties])
Hundreds = SUM(Fact_Batting[Hundreds])
```

**🎯 Bowling KPIs**
```DAX
Total_Wickets = SUM(Fact_Bowling[Wkts])
Avg_Bowling = AVERAGE(Fact_Bowling[Ave])
SR_Bowl= AVERAGE(Fact_Bowling[SR])
Eco_Rate = AVERAGE(Fact_Bowling[Econ])
4_Wickets = SUM(Fact_Bowling[4_Wkts])
5_Wickets = SUM(Fact_Bowling[5_Wkts])
```
**🧤 Fielding KPIs**
```DAX
Total_Dismissals = SUM(Fact_Fielding[Dis])
Total_Catches = SUM(Fact_Fielding[Ct])
Total_CtFi = SUM(Fact_Fielding[Ct_Fi])
Total_CtWk = SUM(Fact_Fielding[Ct_Wk])
Total_Stumpings = SUM(Fact_Fielding[St])
```
**🤝 Partnership KPIs**
```DAX
Partnership_Runs = SUM(Fact_Partnerships[Runs])
Partnership_HR = 
CALCULATE(
    MAX(Fact_Partnerships[HR]),
    Fact_Partnerships[HR_numeric] = MAX(Fact_Partnerships[HR_numeric])
)
Partnership_Avg = DIVIDE([Partnership_Runs],[Total_Partnership_Inns],0)
Partnership_Fifties = SUM(Fact_Partnerships[Fifties])
Partnership_Hundreds = SUM(Fact_Partnerships[Hundreds])
```

**🏆 Team KPIs**
```d 
Total_Matches = COUNTROWS(Fact_MatchResults)

Total_Wins = CALCULATE(COUNTROWS(Fact_MatchResults), Fact_MatchResults[Result] = "Won")

Win % = DIVIDE([Total_Wins], [Total_Matches])

Total_Losses = CALCULATE(COUNTROWS(Fact_MatchResults), Fact_MatchResults[Result] = "Lost")

W/L Ratio = DIVIDE([Total_Wins], [Total_Losses], 0)

Total_NR = CALCULATE(COUNTROWS(Fact_MatchResults), Fact_MatchResults[Result] = "n/r")

Total_Points = CALCULATE(SUM('Fact_PointsTable'[Points]), 'Fact_PointsTable'[Team] = "CSK")

Total_Playoffs = IF(
    -- 1. Check if only ONE season is selected (e.g., on a slicer)
    HASONEFILTER('Dim_Seasons'[Season_ID]),

    -- If TRUE (Only one season selected): Return the status text for that year
    -- Uses MAX/MIN/VALUES because there is only one possible value
    MAX('Dim_Seasons'[CSK_Status]),

    -- If FALSE (Multiple seasons selected or no season selected): Return the TOTAL Runner-up Count
    CALCULATE(
        COUNTROWS('Dim_Seasons'),
        NOT ('Dim_Seasons'[CSK_Status] = "League Stage")
    )
)

Titles_Won = 
IF(
    -- 1. Check if only ONE season is selected (e.g., on a slicer)
    HASONEFILTER('Dim_Seasons'[Season_ID]),

    -- If TRUE (Only one season selected): Return the status text for that year
    -- Uses MAX/MIN/VALUES because there is only one possible value
    MAX('Dim_Seasons'[CSK_Status]),

    -- If FALSE (Multiple seasons selected or no season selected): Return the TOTAL Runner-up Count
    CALCULATE(
        COUNTROWS('Dim_Seasons'),
        'Dim_Seasons'[CSK_Status] = "Champions"
    )
)

Runner-Up = 
IF(
    -- 1. Check if only ONE season is selected (e.g., on a slicer)
    HASONEFILTER('Dim_Seasons'[Season_ID]),

    -- If TRUE (Only one season selected): Return the status text for that year
    -- Uses MAX/MIN/VALUES because there is only one possible value
    MAX('Dim_Seasons'[CSK_Status]),

    -- If FALSE (Multiple seasons selected or no season selected): Return the TOTAL Runner-up Count
    CALCULATE(
        COUNTROWS('Dim_Seasons'),
        'Dim_Seasons'[CSK_Status] = "Runner-up"
    )
)
```


## 📊 7. Dashboard Pages

The Power BI report consists of 8 fully interactive pages:

- **Page 1 – Overview: How I Introduced the Soul of CSK**

    I designed the Overview page to set the emotional and analytical tone for my entire dashboard.

    Here, I bring together **the core identity of CSK** across 16 IPL seasons — 254 matches, 142 wins, 5 titles, 5 runner-ups, nearly 40,000 runs, and 1,440 wickets.

    Every KPI card here is intentional.

    Instead of simply displaying numbers, I wanted these cards to feel like franchise milestones:

    - Dhoni leading with 4,865 runs
    - Jadeja dominating with 143 wickets
    - The highest score ever — 246
    - The lowest score — 79
    - Leadership eras captured visually in the Captaincy Stats table

    As soon as someone sees this page, they understand:

    **This is not just a dashboard. This is the legacy of CSK.**

    This page is my foundation — the **“Who Are We?”** chapter — before the viewer travels deeper into the analysis.  

    ![Overview](dashboard/imgs/1_overview.jpg)

- **Page 2 – Player Insights: How I Capture Individual Journeys**

    This page is where my dashboard becomes personal.

    I built it to give users the ability to explore any CSK player’s evolution over the years.

    For example, the screenshot shows Ravindra Jadeja’s complete journey — batting, bowling, fielding, all woven together through KPIs and trend charts.

    My favorite visual here is the Batting vs Bowling Average over Seasons.
    It doesn’t just show numbers — it shows a story arc:

    -  early inconsistencies
    -  mid-career stabilization
    -  maturity as a true all-rounder

    The intention of this page is simple:
    When someone views a player here, they shouldn’t just see statistics — they should understand the player’s role, evolution, and impact.

    ![Player Profile](dashboard/imgs/2_player.jpg) 

- **Page 3 – Batting Insights: Telling the Story of CSK’s Batting Dynasty**

    I built this page to celebrate how CSK has crafted one of the most consistent batting lineups in the IPL.

    I highlight:

    - Total Runs **39,045**.
    - Franchise batting legends
    - Leaderboards for Runs, Average, Strike Rate, Highest Run, Sixes, Fours, Fifties, Hundreds
    
    The Top 10 Run Scorers chart acts like a visual Hall of Fame:
    - Dhoni and Raina towering over eras
    - Faf and Gaikwad transitioning into new phases
    - Middle-order stabilizers like Rayudu and Hussey

    The **Runs per Season** line chart was added intentionally to show the **rhythm of CSK** — peaks, rebuild years, revival seasons, and dominance phases.

    This page tells everyone who views my dashboard:

    **CSK never depended on one batter. They built a batting culture.**

    ![Batting](dashboard/imgs/3_batting.jpg)

- **Page 4 – Bowling Insights: Revealing the Tactical Backbone**

    This is one of my most analytical sections.

    I wanted to break the misconception that CSK are purely a batting-heavy team.

    Here, I visualize:

    - **1440** team wickets
    - Best Strike Rates, Averages, Economies, 4 Wickets Hauls, 5 Wickets Hauls, Best Bowling Figures, Most Overs Bowled.
    - Top wicket takers.
    - Season-wise economy variations.

    The **Economy Rate over Seasons** line chart is the strategic heart of this page.
    It reveals how CSK’s bowling philosophies shifted over time — from spin-heavy control years to pace-transition phases.

    This page proves a key insight I wanted to highlight:

    **CSK’s bowling dominance doesn’t come from mystery spin or raw pace — it comes from discipline, control, and smart overs.**

    ![Bowling](dashboard/imgs/4_bowling.jpg)

- **Page 5 – Fielding Insights: Highlighting the Unseen Impact**

   Fielding doesn’t often get the spotlight, so I dedicated a whole page to it.

    This section quantifies:

    - **1080** total dismissals
    - **1038** catches (fielding + wicketkeeping)
    - **42** stumpings
    - Dhoni’s staggering **180** total dismissals
    - Other stats like Most Dismissals, Most Catches (Fielding), Most Catches (Wicketkeeping), Most Stumpings.

    The **Dismissals per Season** chart helps show the consistency of CSK’s fielding standards across eras. And Donut charts for **Catches (Fi) vs Catches (Wk) vs Stumpings** break down the contributions for the dismissals visually.


    This page tells the quieter, often ignored part of the story:

    **CSK wins not only with bat and ball — but with moments of brilliance in the field.**

    ![Fielding](dashboard/imgs/5_fielding.jpg)

- **Page 6 – Partnership Insights: Mapping Chemistry Between Players**

    This page is one of the most emotionally rich parts of my dashboard.

    I highlight:

    - **41,061** partnership runs
    - **234** fifty-run stands
    - **42** century partnerships
    - **182** the highest partnership in an innings
    - Most partnerships by Runs, Highest Run, Fifties, Hundreds.

    What I personally love is how the visuals bring out iconic duos:
    - Dhoni–Jadeja: 1708 runs
    - Dhoni–Raina: 1482 runs
    - Hussey–Vijay: Era-defining synergy
    - Conway–Gaikwad: Modern machine

    Through these charts, I wanted to prove a key narrative:

    **CSK is not a team of individual performers — they are built on partnerships, chemistry, and trust between pairs.**

    ![Partnership](dashboard/imgs/6_partnership.jpg)

- **Page 7 – Team Insights: Understanding Rivalries, Patterns & Trends**

    This section presents the competitive heartbeat of CSK’s IPL journey.

    With visuals such as:

    - Wins vs Losses by Season
    - Head-to-head records
    - Home vs Away performance
    - Toss impact
    - Result distributions

    I show long-term patterns:
    - MI is the most challenging rivalry
    - CSK outperforms most teams like SRH, KKR, RCB
    - Interestingly, CSK has **more away wins than home wins** — breaking a common myth

    **This page is designed to help users understand the strategic evolution of the team across seasons, opponents, venues, and conditions.**

    ![Team](dashboard/imgs/7_team_performance.jpg)

- **Page 8 – Match Explorer: A Complete Match Explorer**

    This final page is the most interactive and the most analyst-friendly section I built.

    Includes:
    - Match results table
    - Ground mapping
    - Win margins
    - Toss influence
    - Batting first vs second outcomes

    This page serves as the ultimate reference point for anyone wanting to analyze CSK’s match-by-match performance across IPL history.

    ![Match](dashboard/imgs/8_match_explorer.jpg)


**Each page includes:**
- Top Nav bar for navigation
- Slicers for dynamic filtering
- Detailed visuals (charts, tables, cards, image cards)
- DAX-driven dynamic values
- Custom tooltips 
- CSK-themed UI design

### ⭐ **Final Reflection — My Analytical Vision Behind This Dashboard**

When I built this dashboard, my goal wasn’t just visualization.

I wanted to create a **narrative-driven analytical experience** that blends:
- Sports analytics
- Design clarity
- Storytelling
- Interactivity

Each page contributes to a multi-layered story:
- Legacy
- Evolution
- Individual brilliance
- Tactical strengths
- Partnerships
- Rivalries
- Match-by-match insights

This is my attempt to show how data isn’t just numbers.
**It is memory. It is emotion. It is strategy. It is the heartbeat of CSK’s IPL journey.**


### 🎨 **Design Theme – CSK Colors**

| Color Name      | Hex Code   | Preview |
|-----------------|------------|---------|
| Primary Yellow  | `#FFC30D`  | <div style="width:20px;height:20px;background:#FFC30D;border:1px solid #ddd;display:inline-block"></div> |
| Royal Blue      | `#004BA0`  | <div style="width:20px;height:20px;background:#004BA0;border:1px solid #ddd;display:inline-block"></div> |
| Black    | `#000000`  | <div style="width:20px;height:20px;background:#000000;border:1px solid #ddd;display:inline-block"></div> |
| White           | `#FFFFFF`  | <div style="width:20px;height:20px;background:#FFFFFF;border:1px solid #ddd;display:inline-block"></div> |
| Dark Grey    | `#1C1C1C`  | <div style="width:20px;height:20px;background:#1C1C1C;border:1px solid #ddd;display:inline-block"></div> |


This analysis surfaces clear performance patterns across seasons, players, partnerships, venues, and match situations, using interactive filters (Season, Player, Opposition, Ground).

### Executive Insight Summary (Concise)

- CSK’s peak seasons **(2010, 2013, 2018, 2021, 2023)** are driven by **balanced partnerships, controlled bowling economy, and stable leadership** not individual brilliance.

- **MS Dhoni and Suresh Raina** anchor CSK’s batting legacy, while **Ravindra Jadeja** stands out as the franchise’s most impactful all-rounder.

- **Bowling economy** are stronger predictors of match success than **total wickets**.

- **Partnership-led batting** (Dhoni–Raina for consistency, Conway–Gaikwad for impact) plays a critical role in converting starts into wins.

- CSK shows **high adaptability**, with **more away wins** overall, while maintaining a **strong home advantage at Chennai (~65%).**

- **Post-2018 fielding improvements** significantly improved win conversion in close matches.

### Key Analytical Insights

- **Team & Seasons:** High win-percentage seasons consistently align with lower bowling economy and strong partnerships rather than aggressive scoring.

- **Players & Roles:** Dhoni remains the central influence across batting, keeping, and partnerships; Jadeja delivers the highest combined all-round value; Bravo and Ashwin anchor wicket-taking across eras.

- **Batting & Partnerships:** CSK’s success correlates more with sustained partnerships than isolated high individual scores.

- **Bowling & Fielding:** Controlled economy and improved fielding (especially post-2018) convert pressure into match-winning outcomes.

- **Match Strategy:** CSK performs consistently both batting first and second, indicating tactical flexibility beyond toss dependency.

### Business-Level Outcomes

- Identified **peak performance eras** to evaluate leadership and squad stability.
- Highlighted **sustained player impact** beyond traditional averages.
- Demonstrated that CSK’s success model prioritizes **partnerships, bowling control, and fielding efficiency**.
- Enabled **interactive, decision-oriented** exploration at season and match level.

---

## 📂 9. Repository Structure
```text
CSK_Ipl_Analytics/
|
├── dashboard/
│   ├── imgs/
│   │     ├── logos.png                   # Logos for the dashboard
│   │     └── dashboard_screenshots.png   # Images/screenshots of the final dashboard
│   ├── CSK_IPL_Analysis_Dashboard.pbix   # Power BI Desktop file for the final dashboard
│   └── CSK_IPL_Analysis_Dashboard.pdf    # Exported PDF version of the dashboard
│
├── data/
│   ├── final_excel/                      # Cleaned, transformed, and merged data files 
│   │    ├── All_Players_record.xlsx
│   │    ├── batting_csk_records_cleaned.xlsx
│   │    ├── bowling_csk_records_cleaned.xlsx
│   │    ├── fielding_csk_records_cleaned.xlsx
│   │    └── match_results.xlsx
|   |
│   └── raw_temp/                         # Original, raw and temporary scraped data files
│        ├── batting_records_csk.xlsx
│        ├── bowling_records_csk.xlsx
│        ├── csk_match_results.xlsx
│        ├── fielding_records_csk.xlsx
│        ├── partnerships_records.xlsx
│        ├── pointstable.xlsx
│        ├── team_results.xlsx
|        └── other temporary raw data json files
│
├── src/                                 # Source code for data acquisition and cleaning
│   ├── etl/                             # Python scripts for Extract, Transform, and Load (ETL)
│   │     ├── clean_excel.py             # Script to clean and transform raw data files
│   │     └── overall_records.py         # Script to generate overall aggregated records
│   └── scraper/                         # JavaScript/Python scripts for web scraping
│         ├── merge_csk_player_data.py   # Python script to merge scraped data into single files
│         ├── scrape_players_seasonwise.js # JavaScript scraper for season-wise player stats
│         └── scrape_team_performance.js   # JavaScript scraper for team performance data
│
├── CSK Performance Dashboard.docx       # Detailed documentation and analysis report
├── package-lock.json                    # Lock file for JavaScript
|
├── package.json                         # Configuration for JavaScript (for scraper)
|
└── README.md                            
```
## ▶️ 10. How to Use This Project

1. Clone or download the repository
2. Open `CSK_IPL_Analysis_Dashboard.pbix` using Power BI Desktop
3. Ensure all Excel files inside `/data/final_excel/` are present
4. Refresh the dataset to reload all visuals
5. Use slicers to explore seasons, players, opponents, and venues


## 📘 11. Learnings

This project significantly strengthened both my technical and analytical skills:

- Learned how to automate large-scale sports data collection using **Puppeteer and Cheerio**
- Gained hands-on experience in cleaning highly inconsistent real-world datasets
- Developed strong data transformation logic using **Python (pandas, openpyxl)**
- Designed and implemented a complete **Star Schema data model** in Power BI
- Built complex **DAX measures** involving conditional logic, context handling, and aggregation
- Understood the importance of performance optimization in **Power BI (relationships, slicers, measures)**
- **Improved data storytelling** by aligning visuals with analytical narratives
- Learned to balance **domain knowledge (cricket)** with analytical design
- Strengthened **end-to-end project thinking** — from raw data to business insights
- Improved **documentation and project communication skills** through detailed README writing

## 🚀 12. Future Enhancements

The project can be further extended in the following ways:

- Add predictive analytics to forecast player form and team performance
- Build a custom MVP / Player Impact Index combining batting, bowling, and fielding metrics
- Introduce venue-based geospatial analysis using map visuals
- Implement season comparison bookmarks for side-by-side analysis
- Add advanced tooltip pages for deeper drill-down insights
- Integrate match phase analysis (Powerplay vs Middle vs Death Overs) by importing ball-by-ball data 
- Automate data refresh pipelines for future IPL seasons
- Convert the dashboard into a public Power BI Service embed for wider accessibility

## 👤 Author
**Uthayanithi U**  
*Aspiring Data Analyst / Power BI Developer / SQL Enthusiast*  
📧 [LinkedIn Profile](https://www.linkedin.com/in/uthaya7)  
📁 [Portfolio Website](https://github.com/uthaya7/)