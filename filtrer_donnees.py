import pandas as pd
import matplotlib.pyplot as plt


def clean_data_excel(file, var, cut_rol, sheet=None, cut_neighboor=None, rol_num=2):
    """Uses pandas module to perform a simple rolling mean on a variable of an Excel spreadsheet. It will perform
    analysis by converting the file into a pandas dataframe.

    Args:
        file (str): Excel file to be used
        var (str): Name of the variable in the file to be filtered. It must be a column of
        cut_rol (float): Cut off difference between the rolling mean and the data. Ex. If equals to 1, all the data
                that are different by more than 1 from the rolling mean will be filtered out (replaced by NaN).
        sheet (str): Name of the Excel spreadsheet where the variable to be filtered is. The default spreadsheet is
              the first one.
        rol_num (int): Number of data used in the rolling mean. The default number is 2.
        cut_neighboor (float): Cut off difference between the neighboor data (-1 and +1 indices) and the data itself.
                Ex. If equals to 0.5, all the data that are different by more than 0.5 from both neighboors
                will be filtered out (replaced by NaN). This filtring method is optional. It is not performed
                by default.
    Returns:
        df (dataframe): the filtered Excel spreadsheet as a dataframe.

    """
    pd.options.mode.chained_assignment = None

    # Load the excel file
    df = pd.read_excel(file, sheet_name=sheet)

    # Print the number of row (number of data) your spreadsheet contains.
    print(f"The number of row in the original file is {len(df)}")

    # Will perform the neighboor filtration only if user specified a value to cut_neighboor.
    if cut_neighboor is not None:
        df[var].where(((abs(df[var] - df[var].shift(-1)) < cut_neighboor) | (abs(df[var] - df[var].shift(1) <
                                                                                 cut_neighboor))), inplace=True)

    # ROLLING MEAN FILTRATION

    # 1. Creates a new column containing the values of the rolling mean with the number of values specified in rol_num.
    df['r_mean'] = df[var].rolling(rol_num, center=True, min_periods=round(rol_num / 3)).mean()

    # 2. For every values of your variable, it will compare it to the rolling mean value and will replace it by NaN if
    # the difference is greater than cut_rol.
    df[var].where((abs(df[var] - df['r_mean']) < cut_rol), inplace=True)

    # 3. Delete the column r_mean used for the rolling mean
    df.drop('r_mean', axis=1, inplace=True)

    # Print the number of data after the filtering performed.
    print(f"The number of data left after filtering is {df[var].count()}")

    # Return the dataframe filitered.
    return df


def generate_graph(df, var, cut_rol, rol_num=2, cut_neighboor=None):
    '''
    Pas le temps de commenter! À suivre...
    Args:
        df:
        var:
        cut_rol:
        rol_num:
        cut_neighboor:

    Returns:

    '''
    df.reset_index(inplace=True)
    plt.figure()

    # Create a scatter plot
    plt.scatter(df.index, df[var], s=1)

    # Add text to it.
    plt.text(0, 1, f"Nombre de données pour rolling mean: {rol_num}\nDifférence max entre valeur et rolling mean:"
                   f" {cut_rol}", transform=plt.gca().transAxes)
    if cut_neighboor is not None:
        plt.text(0, 0.90,
                 f"Données avec différence > {cut_neighboor}\navec les valeurs voisines ont été "
                 f"retirées", transform=plt.gca().transAxes)
    else:
        plt.text(0, 0.9, f"La filtration par valeur\nvoisine n'a pas été effectuée", transform=plt.gca().transAxes)
    plt.show()



# Only for testing:
if __name__ == '__main__':
    clean_df = clean_data_excel('ARUTUA_2019.xlsx', 'Chlorophylle', sheet='Raw', rol_num=20, cut_rol=1,
                                cut_neighboor=0.3)
    clean_df.reset_index(inplace=True)
    plt.scatter(clean_df.index, clean_df.Chlorophylle, s=1)
    plt.show()
