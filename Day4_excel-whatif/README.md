# Day 4 - Excel 2: What-if and Forecasting
Date: Tue 20 Oct | Topic: Goal Seek, Data Tables, Scenario Manager, Forecast Sheet

## What-if Model: What price and footfall hit a revenue target?
Assumptions:
- Current Price = 250
- Current Footfall / Day = 120
- Conversion = 35%
- Revenue / Day = Price * Footfall * Conversion

## Tasks Done
1. Goal Seek: Find Price needed for Revenue/Day = 15000, keeping Footfall=120
2. Data Table: Price 200 to 300 (step 10) vs Revenue
3. Scenario Manager: 
   - Base: Price 250, Footfall 120
   - Optimistic: Price 280, Footfall 150
   - Pessimistic: Price 220, Footfall 90
4. Forecast Sheet: Forecast next 3 months revenue from last 12 months

## Result
- Goal Seek Price for 15000 target = ~357
- Best lever is Footfall, not Price alone.
File: whatif_model.csv - open in Excel and apply formulas above.