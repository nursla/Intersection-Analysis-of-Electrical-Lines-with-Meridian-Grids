# Intersection-Analysis-of-Electrical-Lines-with-Meridian-Grids
This script performs a spatial intersection analysis between electrical line shapefiles and meridian grid shapefiles. It reads shapefiles from specified directories, processes each meridian individually, and outputs the intersection results to an Excel file. The script ensures that the coordinate reference systems (CRS) of the shapefiles are consistent before performing the intersection. The results are saved in separate sheets for each meridian in the Excel file.
#Input: Shapefiles of electrical lines and meridian grids.
Processing: Spatial intersection of electrical lines with meridian grids.
Output: An Excel file containing the intersection results, with separate sheets for each meridian.
Dependencies: geopandas, pandas, os
#1-Update the file paths in the script to point to your shapefiles.
2-Ensure that the required Python packages (geopandas, pandas) are installed.
3-Run the script to generate the Excel file with intersection results.
You can also include a brief example or link to the dataset if it’s publicly available, and provide instructions on how to set up the environment or any prerequisites.
