# rebalance-portfolio

Google Apps Script (GAS) code to calculate optimal purchases to re-balance a portfolio.

Demo (copy the Google Sheet): https://docs.google.com/spreadsheets/d/1ntywtnbHjnhOAXBWm3uxT4xhSpne-s9qYo59h99WP40/edit

Uses [LinearOptimizationService](https://developers.google.com/apps-script/reference/optimization/linear-optimization-service) for the calculation:

Constraints:
- 0 <= remaining cash <= 10
- Optimal purchase - maximum deviation* <= Delta <= optimal purchase + maximum deviation*

Where maximum deviation is between 0.01% and 0.1% from the target allocation.

Objective:
- Minimize remaining cash

The resulting calculation will provide a suggested purchase that closely matches the target allocation, without resulting in negative cash.

## rebalance.vba

This is a legacy Excel VBA script I wrote in 2016, to be used for reference only. This code uses the Excel Solver add-in to calculate the optimal purchase, which is slightly better than Google Sheets (Solver uses a non-linear algorithm which allows multiple objectives).

Note the VBA will not work without modifications, the functions defined on it are tied to Excel form controls.
