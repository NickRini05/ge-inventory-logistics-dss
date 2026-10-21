# GE Inventory & Logistics Decision Support System

An Excel/VBA-based decision support system designed to support demand forecasting, inventory management, and logistics planning for three GE Appliances product lines: refrigerators, dishwashers, and microwaves.

## Project Overview

Developed as a four-person academic team project at Virginia Tech, this system combines forecasting techniques, optimization models, and VBA-driven navigation into an interactive Excel workbook.

The objective was to help users evaluate product demand, allocate inventory, plan warehouse shipments, and assess associated costs.

## Features

### Demand Forecasting
- Forecasts product sales using linear, exponential, and polynomial models.
- Compares forecast accuracy using average absolute error.
- Identifies the best-performing forecasting approach for each product line.
- Includes charts illustrating product demand and forecast trends.

### Transportation and Shipping
- Models shipments from four warehouses to six store locations.
- Accounts for warehouse supply, store demand, and transportation costs.
- Supports cost-efficient shipment allocation through an Excel optimization model.

### Delivery Vehicle Assignment
- Includes separate assignment models for refrigerators, dishwashers, and microwaves.
- Evaluates semi-trucks, step vans, and cutaway vans.
- Incorporates vehicle capacities, assignment costs, and delivery times.

### Inventory and Financial Analysis
- Tracks beginning inventory, orders placed, and remaining inventory.
- Summarizes product sales and allocations across store locations.
- Calculates revenue, shipping costs, and profitability.

### VBA Interface
- Provides a welcome screen and login interface.
- Uses VBA to support navigation and interaction with the workbook's analytical components.

## Technologies and Tools

- **Microsoft Excel:** Spreadsheet modeling, formulas, and visualization
- **Visual Basic for Applications (VBA):** Interface functionality and automation
- **Excel Solver:** Transportation and assignment optimization
- **Forecasting Methods:** Linear, exponential, and polynomial forecasting

## How to Use

1. Download `GE_Inventory_Logistics_DSS.xlsm` from this repository.
2. Open the workbook in the desktop version of Microsoft Excel.
3. Enable VBA macros if you trust the downloaded workbook.
4. Use the Welcome screen to navigate through the system.
5. Explore the forecasting, inventory, shipping, and assignment models.

Microsoft Excel with VBA support is required. The Excel Solver add-in may also be needed to rerun optimization models.

## My Contributions

My primary responsibilities included developing transportation and transshipment modeling components and implementing workbook navigation functionality using VBA.

I collaborated with three teammates to integrate the system's analytical features into a unified decision support application. Our team presented the completed project to a Deloitte Senior Consultant for feedback and evaluation.

## Project Context

**Course:** BIT 3424, Introduction to Business Analytics Modeling, Virginia Tech  
**Completed:** Spring 2026  
**Project Type:** Four-person academic team project

This project was completed for educational purposes using a GE Appliances business scenario and is not an official GE Appliances application.
