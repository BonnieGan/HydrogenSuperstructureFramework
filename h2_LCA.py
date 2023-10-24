from brightway2 import *
import numpy as np
import pandas as pd
import utils_config
import os

def lca_calc(md, fu):
    impact_categories = []
    impact_results = []
    for method in md:
        lca = LCA(fu, method)
        lca.lci()
        lca.lcia()
        impact_categories.append(method[2])
        impact_results.append(lca.score)
    ser = pd.Series(impact_results, index=impact_categories)
    return ser

projects.set_current(utils_config.project_name)

imp = ExcelImporter(utils_config.input_file_path)

# find if the database already exist
if imp.data[0]['database'] in databases:
    del databases[imp.data[0]['database']]

# link exchanges
imp.apply_strategies()
imp.match_database(fields=('name', 'unit', 'location'))
imp.match_database("ecoinvent 3.5 cutoff", fields=('name', 'unit', 'location', 'reference product'))
imp.match_database("biosphere3", fields=('name', 'unit', 'location', 'categories', 'reference product'))
imp.statistics()

# write the database
imp.write_database()

# get the methods
my_method = pd.read_excel(utils_config.input_file_path, sheet_name='method')
my_method_list = []
md_unit = []
for i in range(my_method.shape[0]):
    my_method_list.append(tuple(my_method.iloc[i, 1:]))
    md_unit.append(my_method.iloc[i, 0])

result_dict = {'unit': md_unit}
for i in imp.data:
    functional_unit = {(i['database'], i['code']): 1}
    en_imp = lca_calc(my_method_list, functional_unit)
    result_dict.update({i['name']: en_imp})

dis = pd.read_excel(utils_config.input_file_path, sheet_name='distribution')
ei35 = Database(utils_config.db_name)

for i in range(dis.shape[0]):
    if i >0:
        b = ei35.search(dis.loc[i]['skip'], limit = 10000)
        for k in b:
            if k['unit'] == dis.iloc[i][3] and k['reference product'] == dis.iloc[i][7] and k['location'] == dis.iloc[i][1]:
                functional_unit = {k: dis.iloc[i][2]}
                en_imp = lca_calc(my_method_list, functional_unit)
                a = 'distribution by ' + dis.iloc[i][8]
                result_dict.update({a: en_imp})

result_df = pd.DataFrame(result_dict)
result_df_transposed = result_df.T

# avoid overwrite
file_path = utils_config.get_excel_path(utils_config.result_name, utils_config.output_path)

# write excel
result_df_transposed.to_excel(file_path)

