README
Stock Analysis Script
Overview
This repository contains a VBA script designed to analyze stock data across multiple quarters. The script loops through all stocks for each quarter and outputs the following information:

Ticker symbol
Quarterly change from the opening price at the beginning of a quarter to the closing price at the end of that quarter
Percentage change from the opening price at the beginning of a quarter to the closing price at the end of that quarter
Total stock volume for each stock
Additionally, the script identifies the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

The script is designed to work on all worksheets within an Excel workbook, enabling a comprehensive analysis across different quarters.

Features
Data Retrieval:

The script reads and stores the following values from each row of stock data:
Ticker symbol
Volume of stock
Open price
Close price
Column Creation:

The script creates the necessary columns for:
Ticker symbol
Total stock volume
Quarterly change in dollars
Percentage change
Conditional Formatting:

Conditional formatting is applied to the "Quarterly Change" column to highlight positive changes in green and negative changes in red.
Similar formatting is applied to the "Percent Change" column.
Calculated Values:

The script calculates and displays the following:
Greatest % increase
Greatest % decrease
Greatest total volume
Worksheet Looping:

The script is capable of running on all sheets within an Excel workbook, ensuring consistent processing across different quarters.
Performance:

The script is optimized to run efficiently, even on larger datasets. Testing is recommended using the provided alphabetical_testing.xlsx file to ensure quick execution.
Usage
To use the script, follow these steps:

Load the Data:

Ensure that the Excel workbook containing stock data (e.g., alphabetical_testing.xlsx) is opened in Excel.
Run the Script:

Open the VBA editor (Alt + F11) in Excel.
Import the VBA script file into the editor.
Run the script to perform the analysis across all worksheets.
Review Results:

The results, including the ticker symbol, quarterly change, percentage change, and total stock volume, will be displayed on the same or a new worksheet within the Excel workbook.
Conditional formatting will visually differentiate positive and negative changes.
Requirements
Excel Workbook: The script requires an Excel workbook containing quarterly stock data.
VBA Script: The script is written in VBA and needs to be run from within the VBA editor in Excel.
Development and Testing
Test File: Use alphabetical_testing.xlsx for initial testing. This smaller dataset allows for faster execution and debugging.
Consistency: Ensure the script operates consistently across all worksheets by testing with different datasets.
Submission Requirements
For submission, ensure the following items are uploaded to GitHub/GitLab:

Screenshots of Results: Provide screenshots showing the output results in Excel.
VBA Script File: Upload the VBA script file separately for easy access.
README File: Include this README file to explain the script's functionality and usage.
Grading Criteria
The script will be evaluated based on the following criteria:

Data Retrieval (20 points)
Column Creation (10 points)
Conditional Formatting (20 points)
Calculated Values (15 points)
Looping Across Worksheets (20 points)
GitHub/GitLab Submission (15 points)
Additional Notes
Ensure the script is optimized for performance and runs efficiently on all worksheets.
The script should not modify the original data but should output results either on the same or a new worksheet.
Include comments in the VBA code to explain each step for easier understanding and maintenance.
Support and Resources
If you encounter any challenges or need support, the following resources are available:

Class Slack Channel: Reach out to peers or instructors for support.
AskBCS Learning Assistants: Utilize the class Slack application for immediate help.
Office Hours: Attend scheduled office hours for one-on-one assistance.
Tutoring: Schedule a session through Bootcamp Spot for personalized support.
If additional help is needed, contact your instructional team, Student Success Advisor, or submit a support ticket through the Student Support section of your BCS application.

Acknowledgements
If you used any external code sources, including Stack Overflow or received help from instructors or peers, please note this in your submission. Proper attribution ensures transparency and academic integrity.

By following the instructions and utilizing the resources provided, you can successfully complete the stock analysis script and meet all the requirements for this assignment. Happy coding!
