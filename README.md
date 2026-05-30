# ai-augmented-sales-incentive-simulator
An advanced corporate operations model built in Excel utilizing multi-dimensional lookups (INDEX/MATCH, 2-Way matrix matching) and bracket evaluation, featuring an AI-engineered synthetic data pipeline and an ETL data-cleaning log.

## Business Scenario
In global sales organizations, commission structures are rarely flat. They shift based on performance tiers (Bronze, Silver, Gold) and geographical regions (NA, EMEA, APAC). Manual tracking is highly prone to error and scalability issues.

This project simulates a corporate operations data pipeline. Utilizing a multi-tab relational structure inside Microsoft Excel, the model automates transaction-level revenue generation, tracks individual sales representative performance brackets, and dynamically calculates exact commission payouts using multi-dimensional matrix lookups.

---

## Technical Skills Demonstrated
* **Advanced Lookups:** Multi-dimensional vector arrays (`INDEX` + Double `MATCH`), structural bypass left-lookups (`INDEX/MATCH`), and range/bracket evaluation (`VLOOKUP` with Approximate Match `TRUE`).
* **Data Ingestion & ETL:** Transitioning raw CSV data into relational structures, isolating text-to-columns delimiters, and stripping rich-text formatting corruptions.
* **AI-Augmented Engineering:** Orchestrating targeted prompts for complex synthetic data generation and reverse-auditing LLM formula structures.

---

## Relational Data Model
The workbook consists of five interconnected datasets:
1. **`Raw_Sales_Ledger`**: Granular transaction logs containing `Order_ID`, `Sales_Rep_ID`, `SKU`, and `Units_Sold`.
2. **`Product_Lookup`**: Master catalog structured with constraints (SKU is *not* the first column, rendering classic `VLOOKUP` useless for product mapping).
3. **`Sales_Rep_Lookup`**: Employee roster mapping IDs to Names and regional footprints.
4. **`Incentive_Matrix`**: A 2D grid intersecting performance tiers (Rows) with global regions (Columns).
5. **`Tier_Thresholds`**: Corporate revenue minimums and maximums defining tier eligibility.

---

## Core Formula Architecture

### 1. The Leftward Structural Bypass (`INDEX/MATCH`)
In the product master file, the unique identifier (`SKU`) sits to the *right* of the `Product_Name`. To map the product name back to the raw sales ledger without structurally altering the master database layout, I bypassed `VLOOKUP` restrictions:
```excel
=IFERROR(INDEX(Product_Lookup!A:A, MATCH(D2, Product_Lookup!B:B, 0)), "Unmapped Product")
