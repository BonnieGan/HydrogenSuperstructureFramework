import os
#project name
project_name = 'case study'
#database name
db_name = 'ecoinvent 3.5 cutoff'
#LCA result file name
result_name = 'LCA results'
#brightway2 compatible inventory table file name
input_file_name = "input"
#calculated LCIA result file name
env_result = 'scenario results'

#brightway2 compatible inventory table file path
input_file_path = r"C:\Users\chenl\OneDrive - UBC\Desktop\bw2\hydrogen superstructure\files\CaseStudy\input_43.xlsx"
#calculated LCIA result file path
LCA_result_path = r"C:\Users\chenl\OneDrive - UBC\Desktop\bw2\hydrogen superstructure\files\CaseStudy\LCA results_7.xlsx"

#ecoinvent database file path
ei35_path = r"C:\Users\chenl\OneDrive - UBC\Desktop\ecoinvent 3.5_cutoff_ecoSpold02\ecoinvent 3.5_cutoff_ecoSpold02\datasets"
#raw inventory table file path
raw_data_path = r"C:\Users\chenl\OneDrive - UBC\Desktop\bw2\hydrogen superstructure\files\CaseStudy\inventory_CaseStudy.xlsx"
#ecoinvent locations file path
ei_35_loc_path = r"C:\Users\chenl\OneDrive - UBC\Desktop\bw2\hydrogen superstructure\files\ei35 locations.csv"
#all exported file folder
output_path = r"C:\Users\chenl\OneDrive - UBC\Desktop\bw2\hydrogen superstructure\files\CaseStudy"
#scenario data file path
supp_path = r"C:\Users\chenl\OneDrive - UBC\Desktop\bw2\hydrogen superstructure\files\CaseStudy\scenario_CaseStudy.xlsx"


#avoid overwritting
def get_excel_path(file_nm, file_dir):
    file_type = '.xlsx'
    aname = file_nm + file_type
    anum = 0
    apath = os.path.join(file_dir, aname)
    while os.path.isfile(apath) == True:
        anum += 1
        aname = file_nm + '_' + str(anum) + file_type
        apath = os.path.join(file_dir, aname)
    return apath

