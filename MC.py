from brightway2 import *
import numpy as np
import pandas as pd
import stats_arrays
import matplotlib.pyplot as plt
import matplotlib as mpl
import time
import utils_config
import os

def multiImpactMC(functional_unit, methods, activity, path):
    all_scores = []
    colour1 = '#8f5362' #red
    colour2 = '#6888a5' #bar colour
    colour3 = '#706d94' #mean
    colour4 = '#79b4a0' #median
    mpl.rcParams['font.family'] = 'Times New Roman'
    for i in methods:
        data = []
        mc = MonteCarloLCA(functional_unit, i)
        # Split the string from the right at the last space
        parts = activity.rsplit(' ', 1)
        # The last element is the one after the last space
        if len(parts) > 1:
            last_element = parts[-1]
            sheet_name = [last_element + '_' + i[1]]  # Output: bg
        else:
            sheet_name = parts
        scores = [next(mc) for _ in range (utils_config.iterations)] # run 1000 simulations
        # Replace : or / with _
        name = sheet_name[0]
        if ':' in name:
            name = name.replace(":", "_")
        if '/' in name:
            name = name.replace('/', '_')
        data.append([name])
        data.append(scores)
        all_scores.append(data)
        # draw plot
        # Calculate statistics
        percentile_5 = np.percentile(scores, 5)
        percentile_95 = np.percentile(scores, 95)
        mean_value = np.mean(scores)
        median_value = np.median(scores)
        sd_value = np.std(scores)
        plt.clf()
        plt.figure(figsize=(8, 5))
        plt.subplots_adjust(left=0.1, right=0.75)
        plt.hist(scores, bins=50, color=colour2, alpha=0.7)
        # Add vertical lines for the percentiles
        plt.axvline(percentile_5, color=colour1, linestyle='--', label='5th Percentile', linewidth=1)
        plt.axvline(percentile_95, color=colour1, linestyle='--', label='95th Percentile', linewidth=1)
        # Add a vertical line for the mean
        plt.axvline(mean_value, color=colour3, linestyle='--', label='Mean', linewidth=1)
        # Add a vertical line for the median
        plt.axvline(median_value, color=colour4, linestyle='--', label='Median', linewidth=1)
        plt.xlabel(i[1])
        plt.ylabel('Frequency')
        plt.title(name)
        # Add notes with the statistics
        def format_number(number):
            if abs(number) < 1e-2:  # Set your threshold for small numbers here
                return f'{number:.2e}'  # Format as scientific notation
            else:
                return f'{number:.2f}'  # Format with two decimal places

        plt.text(plt.xlim()[1] * 1.01, plt.ylim()[1], f'Mean: {format_number(mean_value)}', color=colour3, fontsize=10, ha='left')
        plt.text(plt.xlim()[1] * 1.01, plt.ylim()[1] * 0.95, f'Median: {format_number(median_value)}', color=colour4, fontsize=10, ha='left')
        plt.text(plt.xlim()[1] * 1.01, plt.ylim()[1] * 0.9, f'5th Percentile: {format_number(percentile_5)}', color=colour1, fontsize=10, ha='left')
        plt.text(plt.xlim()[1] * 1.01, plt.ylim()[1] * 0.85, f'95th Percentile: {format_number(percentile_95)}', color=colour1, fontsize=10, ha='left')
        plt.text(plt.xlim()[1] * 1.01, plt.ylim()[1] * 0.8, f'Standard Deviation: {format_number(sd_value)}', color='black', fontsize=10, ha='left')
        fig_path = os.path.join(path, name)
        plt.savefig(fig_path)
        plt.close()  
    return all_scores

projects.set_current(utils_config.mc_project)

# setup
if not 'biosphere3' in databases:
    bw2setup()
# import Ecoinvent3.5
if not utils_config.db_name in databases:
    fpei = utils_config.ei35_path
    ei = SingleOutputEcospold2Importer(fpei, utils_config.db_name)
    ei.apply_strategies()
    ei.write_database()

imp = ExcelImporter(utils_config.mc_path)

# find if the database already exist
if imp.data[0]['database'] in databases:
    del databases[imp.data[0]['database']]

# link exchanges
imp.apply_strategies()
imp.match_database(fields=('name', 'unit', 'location'))
imp.match_database("ecoinvent 3.5 cutoff", fields=('name', 'unit', 'location','reference product'))
imp.match_database("biosphere3", fields=('name', 'unit', 'location', 'categories', 'reference product'))
imp.statistics()

# write the database
imp.write_database()

# get the methods
imp_cat = pd.read_excel(utils_config.mc_path, sheet_name='impact_category')
imp_cat_list = []
for i in range(imp_cat.shape[0]):
    imp_cat_list.append(tuple(imp_cat.iloc[i, :]))

def create_unique_folder(base_path, base_name):
    counter = 1
    folder_path = os.path.join(base_path, base_name)

    # Keep trying new folder names until an unused one is found
    while os.path.exists(folder_path):
        folder_path = os.path.join(base_path, f"{base_name}{counter}")
        counter += 1

    # Create the folder with the unique name
    os.makedirs(folder_path)
    return folder_path

# Usage
base_path = utils_config.output_path # Replace with your desired path
base_folder_name = "mc results"
folder_path = create_unique_folder(base_path, base_folder_name)

mc_results = []
for i in imp.data:
    fu = {(i['database'], i['code']): 1}
    mc_result = multiImpactMC(fu, imp_cat_list, i['name'], folder_path)
    mc_results += mc_result

# avoid overwrite
file_path = utils_config.get_excel_path('mcResult', utils_config.output_path)
    
# write the workbook
with pd.ExcelWriter(file_path) as writer:
    for i in mc_results:
        result_df = pd.DataFrame(i[1])
        result_df.to_excel(writer, sheet_name=i[0][0][:30], header=False, index=False)
      