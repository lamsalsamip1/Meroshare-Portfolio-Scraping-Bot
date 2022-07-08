# Meroshare-Portfolio-Scraping-Bot
This is an automation done using Selenium in Python to login to a meroshare account and scrape the data from the portfolio page and update the data in an existing excel file.

A step by step description of it's main operation is as follows:

1) Go to meroshare website
2) Login by providing necessary keys
3) Click on My Portfolio Tab
4) Calculate number of rows and columns of portfolio table
5) Loop through the table with a nested loop
6) Format those values(removing commas), and write it to required location of excel file by indexing the sheet through loop variables.
