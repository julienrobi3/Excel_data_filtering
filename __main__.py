# This is an example of how someone could use this project (so far).

import pandas as pd
import matplotlib.pyplot as plt
from filtrer_donnees import ExcelData

if __name__ == '__main__':

    # Load the file
    data1 = ExcelData('DATA2019.xlsx', sheet='GENERAL')
    print(type(data1.source_df))

    # Plot source data for Chlorophylle column. You need to close the figure to continue with the script.
    data1.view_data(source=True, var='Chlorophylle')

    # Remove outliers for the Chlorophylle column
    data1.remove_outliers_moving_average('Chlorophylle', 0.5, 60)

    # Plot source data for Chlorophylle column
    data1.view_data(info=True)

    # Save the changes
    data1.save_changes()

    # Save dataframe to a new Excel file
    data1.save_to_new_file('Data_without_outliers.xlsx')

    #
    #
    # View data grouped in specific sections. It can be the date, the ID of the instrument used, site, etc.
    data1.view_data_by_sections('Sonde_ID', 'Temperature')





