# Expenditure Projections

This project handles all quarterly expenditure projections. Monthly automated projections are in development.

## 1: Distribute Templates

This script takes the monthly expenditure file corresponding to the projection month (ex: September's file for Q1) and joins analyst-specified calculations from the previous quarter (ex: last FY's Q3 calculations would be applied for a Q1 projection).

Excel formulas, corresponding the calculations, are written as strings within the script. The openxlsx package is used to make them 'active' formulas within Excel upon export: the data frame is subset into agency-specific data frames, Excel formula classes applied, and then the dataframes are exported into distrubtion folders by analyst.

Analysts then make a copy of their template files elsewhere (typically their G drive analyst folders) and make changes by line item.

## 2: Compile Projections

Once analysts have completed their projections using the template files, they place a copy of the edited files into a centralized folder on the G drive.

This script then imports all Excel files following a specified naming pattern in that centralized folder and binds them together. This compiled data frame is then exported three times:

1. As an Excel file for the Assistant Budget Director. The full compiled dataset is saved for reference.
    - Changes can also be made in this file, and then re-imported to generate the PDF (#2). This is usually done only once analysts are done making changes and the Assistant Director wants to make some global adjustments without needing to go into individual files.
2. As a formatted PDF for the Director of Finance. Only the surplus/deficit is shown, by object.
3. As a 'calculations' CSV, which will be used to generate the next quarter's template. Only a few columns are saved.