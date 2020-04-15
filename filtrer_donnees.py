import pandas as pd
import matplotlib.pyplot as plt
import os

class ExcelData:
    def __init__(self, file, sheet=None):
        """
        Args:
            file (str): Excel file to be used
            sheet (str): Name of the Excel spreadsheet where the variable to be filtered is. The default spreadsheet is
                  the first one.
        """
        self.source_df = pd.read_excel(file, sheet_name=sheet)
        print(type(self.source_df))
        self.temp_df = None
        self.clean_df = self.source_df.copy(deep=True)  # Copy the source df to the clean one.

        # Store variables in attributes. But first, initialize in the __init__
        self.var = ''
        self.cut_mov = None
        self.mov_num = None

    def remove_outliers_moving_average(self, var, cut_mov, mov_num=2):
        """Uses pandas module to perform a simple moving average on a variable of an Excel spreadsheet. It will perform
        analysis by converting the file into a pandas dataframe.

        Args:
            var (str): Name of the variable in the file to be filtered. It must be a column of
            cut_mov (float): Cut off difference between the moving average and the data. Ex. If equals to 1, all the
                data that are different by more than 1 from the moving average will be filtered out (replaced by NaN).
            mov_num (int): Number of data used in the moving average. The default number is 2.
        Returns:
            df (dataframe): the filtered Excel spreadsheet as a dataframe.

        """
        pd.options.mode.chained_assignment = None

        # Save variables in attributes to reuse them in other methods if necessary
        self.var = var
        self.cut_mov = cut_mov
        self.mov_num = mov_num

        # Print the number of row (number of data) your spreadsheet contains.
        print(f"The number of row in the original file is {len(self.source_df)}")

        # MOVING AVERAGE - REMOVAL OF OUTLIERS

        self.temp_df = self.source_df.copy(deep=True) # Copy the source df to the temp one.
        # 1. Creates a new column containing the values of the moving average with the number of values specified
        # in mov_num.
        self.temp_df['r_mean'] = self.temp_df[self.var].rolling(mov_num, center=True, min_periods=round(mov_num /
                                                                                                        3)).mean()

        # 2. For every values of your variable, it will compare it to the moving average value and will replace
        # it by NaN if the difference is greater than cut_mov.
        self.temp_df[self.var].where((abs(self.temp_df[self.var] - self.temp_df['r_mean']) < self.cut_mov),
                                     inplace=True)

        # 3. Delete the column r_mean used for the moving average
        # self.temp_df.drop('r_mean', axis=1, inplace=True)

        # Print the number of data after the filtering performed.
        print(f"The number of data left after filtering is {self.temp_df[self.var].count()}")
        print(f"\nNote that your changes for column {self.var} have not been saved yet.\n"
              f"To do so, please use the method save_changes.")

    @staticmethod
    def plot_data(df, var):
        """
        Static method used in view_data method in order to make a simple scatter plot
        Args:
            df (dataframe): Dataframe of your data
            var (str): variable to plot

        Returns:
            fig, a figure object
        """
        # Making sure there is an index to plot with
        if 'level_0' not in df:
            df.reset_index(inplace=True)

        # Create a figure
        fig = plt.figure()
        ax = fig.add_axes([0.05, 0.05, 0.90, 0.90])

        # Create a scatter plot
        ax.scatter(df.index, df[var], s=1)

        # set title
        ax.set_title(f"{var} data")

        return fig

    def view_data(self, source=False, var=None, info=False):
        """
        Plot either the source or the new filtered data. The default option is the filtered data.
        Args:
            source (Bool): To plot the original (source) data, needs to be True. To plot the changes after removing
                outliers,needs to be False (default option)
            var (str): If "source" is True, "var" needs to be specified. It represents the data to plot. If "source" is
                False, the data to plot is the one that has been filtered (you do not need to specify it).
            info (Bool): If True, add information about the removal of outliers to the figure: Number of data filtered
                out, mov_num and cut_mov.

        """

        if source:
            figure = self.plot_data(self.source_df, var)
            plt.show()

        else:
            figure = self.plot_data(self.temp_df, self.var)
            # # set title
            # ax.set_title(f"Data after removing outliers on column {self.var}")

            if info:
                # Add text to figure if it is specified by the user.
                figure.text(0, 0.85,
                            f"Number of data left after outliers removal is:"
                            f" {self.temp_df[self.var].count()} / {len(self.source_df)} \n"
                            f"Number of data for moving average: {self.mov_num}\nMax difference between value and "
                            f"moving average: {self.cut_mov}", transform=plt.gca().transAxes)
            plt.show()

    def save_changes(self):
        if self.temp_df is not None:
            self.clean_df[self.var] = self.temp_df[self.var]
            print(f"Changes from column name {self.var} have been correctly saved.\n"
                  f"To export the changes in a new Excel file, please use the save_to_new_file method")
        else:
            print('No new changes have been saved')
        self.temp_df = None

    def save_to_new_file(self, name_file, sheet_name='clean'):
        self.clean_df.to_excel(name_file, sheet_name)

    def save_to_new_sheet(self):
        pass


# Only for testing:
if __name__ == '__main__':
    data1 = ExcelData('C:\\Users\\client\\OneDrive\\certificat_info\\Tahiti_donnees\\ARUTUA_2019.xlsx', 'Raw')
    data1.remove_outliers_moving_average('Chlorophylle', 0.5, 60)
    data1.view_data(source=True,var='Chlorophylle')
    # data1.view_data(info=True)
