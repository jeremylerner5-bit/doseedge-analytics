# DoseEdge Analytics Dashboard
## Demo Instructions

---

## Quick Start (5 Steps)

### Prerequisites
- **Node.js** installed on your computer ([Download here](https://nodejs.org/))
- A web browser (Chrome, Firefox, Edge)

### Step 1: Open a Terminal/Command Prompt
- **Windows**: Press `Win + R`, type `cmd`, press Enter
- **Mac**: Open Terminal from Applications > Utilities

### Step 2: Navigate to the Project Folder
```bash
cd "C:\Users\jerem\Desktop\DoseEdge Analytics"
```
Or navigate to wherever you've saved this folder.

### Step 3: Install Dependencies
```bash
npm install
```
This installs the required packages. Only needed the first time.

### Step 4: Start the Server
```bash
npm start
```
You should see: `Server running on http://localhost:3000`

### Step 5: Open in Browser
Open your web browser and go to:
```
http://localhost:3000
```

---

## What You'll See

The dashboard has multiple tabs:

### Dashboard Tab
- **Summary Cards**: Total doses, average turnaround time, bypass rate
- **Time Period Filters**: Last 7 days, 30 days, 90 days, or custom range
- **Quick Stats**: Key performance indicators at a glance

### Analytics Tab
- **Production Charts**: Daily dose production trends
- **Turnaround Analysis**: Preparation time metrics
- **Location Breakdown**: Performance by facility/area

### Trends Tab
- **Historical Data**: Long-term trend visualization
- **Comparison Views**: Week-over-week, month-over-month

### Inventory Tab
- **Stock Doses**: Current inventory levels
- **Product Usage**: Which products are used most frequently
- **Wastage Tracking**: Detailed waste analysis by product

---

## Try It Out

### Explore Pre-Loaded Data
The dashboard comes with real sample data. Try:

1. **Change Time Periods**: Click "7 Days", "30 Days", or "90 Days" to see different views
2. **View Production Trends**: See how daily dose counts vary over time
3. **Check Wastage**: Navigate to Inventory and explore the wastage breakdown
4. **Product Analysis**: See which products have the highest usage/wastage

### Upload New Data
1. Click the "Upload" button
2. Select the data type (Production, Turnaround, Bypass, etc.)
3. Upload an Excel file exported from DoseEdge
4. Watch the dashboard update with new data

### Export Data
1. Click "Export" to download current data as CSV
2. Use for further analysis in Excel or other tools

---

## Key Metrics Explained

| Metric | What It Measures | Why It Matters |
|--------|------------------|----------------|
| **Total Doses** | Number of IV preparations completed | Workload and capacity planning |
| **Turnaround Time** | Order to completion time | Efficiency and patient care timeliness |
| **Bypass Rate** | % of scans bypassed | Medication safety compliance |
| **Wastage** | Unused/expired preparations | Cost control and waste reduction |
| **Stock Doses** | Pre-made inventory levels | Inventory optimization |

---

## Features to Explore

1. **Interactive Charts**: Hover over data points for details
2. **Filtering**: Narrow down by date, location, or product
3. **Drill-Down**: Click on summary cards for detailed views
4. **Data Grouping**: View by day, week, or month

---

## Stopping the Server

Press `Ctrl + C` in the terminal window to stop the server.

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "npm not found" | Install Node.js from nodejs.org |
| "Port 3000 in use" | Close other apps or edit port in server.js |
| Charts not loading | Refresh the page, check console for errors |
| No data showing | Check the data/ folder has JSON files |

---

## About This Project

This analytics dashboard was built to transform raw DoseEdge data exports into actionable operational insights. Instead of manually compiling reports from spreadsheets, pharmacy operations teams can:

- Monitor production trends in real-time
- Identify turnaround time bottlenecks
- Track wastage patterns and reduce costs
- Optimize inventory levels for stock doses

**Technologies Used**: Node.js, Express.js, Chart.js, XLSX parsing

**Data Files**: The `data/` folder contains JSON files with sample data for production, turnaround times, bypass rates, product usage, wastage, and stock doses.

---

*Questions? Contact Jeremy for a live demonstration.*
