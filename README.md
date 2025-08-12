# Data Analysis Excel Projects

This repository features a collection of Excel-based data analysis projects designed to improve analytical skills. It covers key techniques such as pivot tables, data visualisation, and advanced filtering for in-depth analysis. Additionally, the projects include creating dynamic dashboards to present insights effectively. Ideal for anyone looking to enhance their proficiency in Excel for data analysis and decision-making.

-------

## Project 1 - Amazon Sales Data Analysis

This project shows advanced PivotTables and PivotCharts. The interactive dashboard pulls together key retail metrics - total orders, revenue, quantity, and average rating, in case it makes them easy to explore with slicers for date, payment method, and city. 
   
Dashboard: ![amazon](https://github.com/valerijarudjko/data-analysis-xlsx-projects/blob/main/amazon_sales_analysis/Amazon%20Sales%20Interactive%20Dashboard.png)

-----
## Project 2 - Sales and Profit Data Analysis

This project demonstrates advanced PivotTables and PivotCharts in Excel. The interactive dashboard brings together key sales metrics - total sales, total profit, and customer counts, making easy to explore with slicers for order date and product category. It also visualises profit by year, sales by state, and highlights top‑performing sub‑categories and customers.


Dashboard: ![sales](https://github.com/valerijarudjko/data-analysis-xlsx-projects/blob/main/sales_profit_analysis/sales_profit_dashboard.png)

------
## Project 3 - Employee Workflow Data Analysis



______

## Refresh Data button to your Excel dashboard 

1. Turn on the Developer tab:

- Go to Excel > Preferences > Ribbon & Toolbar (on Mac) or File > Options > Customize Ribbon (on Windows).
- Enable the Developer tab.
  
2. Insert a button:
- Go to Developer → Insert → Form Controls → Button (Form Control).
- Draw the button on your dashboard where you want the icon.
  
3. Assign a macro:

When you release the mouse after drawing:
- Excel will ask you to assign a macro.
- Click New and in the VBA editor, paste this (vba):

```vba
Sub Refresh_All()
ThisWorkbook.RefreshAll
End Sub
```
- Save and close the VBA editor.
- Select the Refresh_All macro for your button.

4. (Optional) Format it:

- Right‑click the button → Edit Text to rename it to something like "🔄 or icon Refresh Data".
- Resize, style, or place an icon image behind it for a cleaner look.

After this function when you click Refresh Data shape or icon, it will run the macro and refresh all PivotTables, PivotCharts, and queries in the workbook.

-----
