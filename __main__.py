import sys
print(sys.path)

from filtrer_donnees import clean_data_excel, generate_graph

# Perform the filtering
var = 'Chlorophylle'
cut_rol = 0.5
rol_num = 10
cut_neighboor = 0.3
clean_df = clean_data_excel('C:\\Users\\client\\OneDrive\\certificat_info\\Tahiti_donnees\\ARUTUA_2019.xlsx',
                            var, cut_rol, sheet='Raw', rol_num=rol_num,
                            cut_neighboor=cut_neighboor)

# Show the result
generate_graph(clean_df, var, cut_rol, rol_num,cut_neighboor, info=True)


# Convert filtered data into excel file
#clean_df.to_excel('clean_Chloro_data.xlsx', sheet_name='Clean')


