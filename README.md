# Data Cleaning and Analytics System

BAT403 - Foundations of Enterprise Data Management

This is a deployable React/Vite system for the workflow:

Upload -> Profile -> Clean -> Analyze -> Visualize

## Project Description

This project involves the design and development of a Data Cleaning and Analytics System that allows users to upload raw datasets, perform data cleaning operations, and generate meaningful insights and visualizations. The system supports different tabular datasets, including CSV and Excel files, and simulates a real-world data analysis workflow.

## Project Objectives

- Develop a system for data preprocessing and analysis.
- Apply data cleaning techniques.
- Generate insights from datasets.
- Create interactive dashboards.
- Understand data-driven decision making.

## Core System Workflow

Upload -> Profile -> Clean -> Analyze -> Visualize

## Features

- Upload CSV or Excel files.
- Preview original and cleaned datasets in table format.
- Profile rows, columns, data types, missing values, and basic statistics.
- Clean data by filling or removing missing values, removing duplicates, converting data types, standardizing text formats, and filtering invalid numeric ranges.
- Compare original vs cleaned data with changed cells highlighted.
- Generate summary statistics, frequent values, trend notes, and simple interpretations.
- Visualize data with bar, line, pie, and gauge-style dashboard cards.
- Choose diagram types in a Weka-inspired Diagram Builder: bar, line, pie, donut, or scatter.
- Customize dashboard colors, title text, subtitle text, and add note text blocks like a lightweight Weka-style workspace.
- Save the current workspace in the browser so refreshes do not reset uploaded or cleaned data.
- Open the app offline after it has been loaded once in the browser.

## Requirement Coverage

### Data Input

- Users can upload CSV, XLS, or XLSX files.
- The uploaded dataset is displayed in a table format.
- The app keeps a copy of the original dataset for comparison.

### Data Profiling

- The app shows the number of rows and columns.
- Data types are detected for each column.
- Missing values are identified per column and as a total.
- Basic statistics are generated, including unique count, most frequent value, mean, median, minimum, and maximum where applicable.

### Data Cleaning Features

- Handle missing values using mean, median, mode, zero, blank, custom value, or row removal.
- Remove duplicate records by comparing complete row values.
- Convert data types to number, text, date, or boolean.
- Standardize text formats by trimming spaces, uppercase, lowercase, or title case.
- Filter invalid numeric data using minimum and maximum limits.

### Data Comparison

- The app stores the uploaded dataset as the original dataset.
- Cleaning actions update a separate cleaned dataset.
- The Original vs Cleaned table lets users switch between versions.
- Changed cells are highlighted in the cleaned view after cleaning operations.

### Insights Generation

- Summary statistics are generated for numeric columns.
- Missing value totals, duplicate counts, unique values, means, medians, minimums, and maximums are shown.
- Most frequent values are identified per column.
- Simple interpretations are displayed in the Insights panel.
- Trends and patterns are explored through line charts, bar charts, pie charts, donut charts, and scatter plots.

### Dashboard Visualization

- Bar chart for comparing values across categories.
- Line chart for trend or ordered-variable analysis.
- Pie and donut charts for proportional category summaries.
- Ratio gauge for comparing two numeric variables.
- User-selectable variables are available through dropdown controls.

### System Functionality

- The system follows the workflow: Upload -> Profile -> Clean -> Analyze -> Visualize.
- Data processing happens in the browser, so no backend is required.
- The interface is responsive and works on desktop and smaller screens.
- Uploaded and cleaned data are saved in IndexedDB so the workspace remains after refresh.
- A service worker caches the app shell for offline use after the first visit.
- The project is configured for Vercel deployment.

## Cleaning Methods

- Missing values can be filled by mean, median, mode, zero, blank text, or a custom value. Rows with missing values in a chosen column can also be removed.
- Duplicate records are detected by full-row equality and removed while keeping the first record.
- Type conversion supports text, number, date, and boolean formats.
- Standardization supports trimming spaces, uppercase, lowercase, and title case.
- Invalid numeric values can be filtered using optional minimum and maximum limits.

## Weka-Inspired Workflow

The interface follows the same general idea as Weka Explorer: load data, inspect attributes, apply preprocessing, then visualize relationships by choosing variables. The Diagram Builder mirrors Weka's visualizer pattern by letting users choose axes, a value field, and a color/class field before viewing the plot.

## Screenshots to Include

Add screenshots of these screens to your final document:

1. Upload screen before importing data.
2. Data profile after uploading a dataset.
3. Cleaning Workbench after applying a cleaning method.
4. Original vs Cleaned comparison table with highlighted changes.
5. Insights panel.
6. Dashboard visualizations with selected variables.

## Defense Guide

Explain the logic in this order:

1. Upload: The user imports a CSV or Excel file, and the system reads the first sheet or CSV table.
2. Profile: The system detects rows, columns, data types, missing values, duplicates, and statistics.
3. Clean: The user chooses a column and cleaning method based on the detected data issue.
4. Compare: The app keeps the original data unchanged and shows the cleaned version separately.
5. Analyze: The app summarizes numeric and categorical patterns.
6. Visualize: The user selects chart variables to show trends, category comparisons, proportions, or numeric relationships.

Example defense explanation:

> I applied missing-value handling because empty cells can affect statistics and charts. I removed duplicate rows because repeated records can overstate totals and frequencies. I used type conversion so numeric and date fields can be analyzed correctly. I used standardization to make values consistent, such as fixing uppercase/lowercase differences. The insights panel then explains the cleaned data through statistics, frequent values, and chart patterns.

## Run Locally

```bash
npm install
npm run dev
```

Open the local URL printed by Vite.

## Build

```bash
npm run build
```

## Deploy to Vercel

1. Push this project to GitHub.
2. Import the repository in Vercel.
3. Vercel will use `npm run build` and publish the `dist` folder.

The included `vercel.json` is already configured for Vite.

## Sample Dataset

The app does not auto-load sample data. For testing or documentation, upload `public/sample-dataset.csv` manually from the app.

The sample dataset intentionally includes:

- Missing values in `Revenue`, `Compensation`, and `Engagement`.
- One duplicate row.
- One invalid negative `Revenue` value.
- Numeric, date, and categorical columns for testing charts and cleaning methods.
