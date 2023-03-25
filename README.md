# This module is part of data processing of an environmental chemical fate model (Gridded-SoilPlusVeg)

Results obtained in from the model are saved into a .csv file on an hourly basis

Each row belong to one cell in a grid of 10x10 including calculated fugacity value for upper-air, lower-air, five soil layers, vegetation, row number, and column number

Each 100 cells represent one hour and entire grid values

This module organizes values for each compartment and put them in their respecting sheet, where each hour is shown within a grid structure

At the end ColorScale will give color to each cell depending on their value ranging from red, yellow to blue for high to low values, respectively 

