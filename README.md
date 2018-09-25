# ExcelAddIn_MultipleCTPs

Introduction:

This project contains simple C# code for handling multiple Custom Task Panes in Excel 2013+.
It is an Excel Add-In (only tested on Excel 2016 so far).

I could not find any projects online that handled this elegantly, and/or were written in C#.


How it works:

Upon running, there will be a new Velocity Multi tab in the ribbon. Click the Log In button to open the custom task pane.
This will work across different workbooks, and if a workbook is inactive, it will hide the pane.


Credit:

Ideas based on work found in this article: https://www.codeproject.com/Tips/1063935/Implementing-CTPs-in-Excel-and-with-Excel
