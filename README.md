# Automated basis generation with sales forecast


## Task
Build a tool to diminish the amount of time needed to update the basis for the monthly sales forecast review that a company does. The output must be an excel sheet with one product per line, some relevant information about it and the monthly sales forecast that was built last month. The new month must be automatically added to the sheet, with no values in it, since it will be later handled by the team.

## Summary of the tech task
This Notebook was built using Python 3.7.3. Since this is a data manipulation problem, I chose to use pandas as the main library for it. The input comes in form of 6 different sheets that must be

## Inputs
Use the following excel sheets as inputs, with the relevant fields listed below:
- Product log:
 	- Product code
 	- Product description
 	- Category
 	- Subcategory
 	- Status
 	- Brand
 	- Date of release
- Analyst log:
 	- Category
 	- Analyst responsible
- Stock on Store:
 	- Product code
 	- Product description
 	- Status
 	- Stock (in units)
- Stock on distribution center:
 	- Booked for
 	- Product code
 	- Stock (in units)
- Sales from last twelve months
 	- Product code
 	- Monthly sale (in units)
 	- Date
 	- Month
 	- Year
- Previous sales forecast basis
 	- Product code
 	- 14x Monthly sale (in units)

## Functionality
This notebook saves about 4 hours of work with data manipulation. It also provides all the necessary information for the review of the sales forecast. 

## Tech Stack
- Python
	- Pandas
