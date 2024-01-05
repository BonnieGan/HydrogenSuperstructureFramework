import os

'''edit before running the code'''
#project name
project_name = 'demo1'
#database name
db_name = 'ecoinvent 3.5 cutoff'
#LCA result file name
result_name = 'LCA results_demo'
#brightway2 compatible inventory table file name
input_file_name = 'input_demo'
#calculated optimization result file name
env_result = 'solutions_demo'
#ecoinvent database file path
ei35_path = '/Users/Bonnie/Desktop/ecoinvent 3.5_cutoff_ecoSpold02/datasets'
#ecoinvent locations file path
ei_35_loc_path = '/Users/Bonnie/Desktop/HydrogenSuperstructureFramework/ei35 locations.csv'
#raw inventory table file path
raw_data_path = '/Users/Bonnie/Desktop/demo/inventory_demo.xlsx'
#scenario data file path
supp_path = '/Users/Bonnie/Desktop/demo/scenario_demo.xlsx'
#all exported file folder
output_path = '/Users/Bonnie/Desktop/demo/results'

'''Monte Carlo Simulatio part'''
#Monte Carlo Simulation project name
mc_project = 'mc'
#Monte Carlo Simulation iterations
iterations = 1000
#Monte Carlo Simulation file path
mc_path = '/Users/Bonnie/Desktop/demo/mc_demo.xlsx'

'''edit after running the code'''
#brightway2 compatible inventory table file path
input_file_path = '/Users/Bonnie/Desktop/demo/results/input_demo.xlsx'
#calculated LCIA result file path
LCA_result_path = '/Users/Bonnie/Desktop/demo/results/LCA results_demo.xlsx'


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

