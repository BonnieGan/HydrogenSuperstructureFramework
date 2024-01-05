from brightway2 import *
import numpy as np
import pandas as pd
import utils_config
import os

# extract database data
def extract_data(path, sheet):
    database_data = pd.read_excel(path, sheet_name=sheet)

    # extract exchange data
    other_list, bio_list = [], []
    for i in range(database_data.shape[0]):
        exc = {
            'name': database_data.loc[i, 'Input'],
            'amount': database_data.loc[i, 'Input amount'],
            'unit': database_data.loc[i, 'Input unit']
            }
        if database_data.loc[i,'Biosphere (y/n)'] == 'y':
            bio_list.append(exc)
        else:
            other_list.append(exc)
    return other_list,bio_list

# find RoW and RER data
def find_GLO_location(filtered_ei_data):
    output_list = []
    for i in filtered_ei_data:
        if i['location'] == 'RoW' or i['location'] == 'RER' or i['location'] == 'GLO':
            output_list.append(i)
    return output_list

# find data with exactly the same geography
def find_exa_location(filtered_ei_data, ts):
    output_list = []
    for i in filtered_ei_data:
        if ts in i['location']: 
            output_list.append(i)
    return output_list

# find data with geography in the contained and overlapping lists
def find_contained_location(contained_geo_list, filtered_ei_data, ts):
    output_list = []
    for i in filtered_ei_data:
        if i['location'] in contained_geo_list:
            output_list.append(i)
    return output_list

# find RoW data
def find_RoW_location(filtered_ei_data):
    output_list = []
    for i in filtered_ei_data:
        if i['location'] == 'RoW':
            output_list.append(i)
    return output_list

# find exchange locations
def find_exc_location(test, df, lt, unit):
    # test = project location
    # df = ecoinvent location dataframe
    # lt = filtered exchange search result
    # unit = unit of the exchange (e.g., kilowatt hour for electricity)

    contained_test_loc = []
    exc_list_GLO = []
    exc_w_unit = []
    exc_loc_list = []
    output_list = []

    for i in range(df.shape[0]):
        if test in df.loc[i, 'Contained and Overlapping Geographies']:
            contained_test_loc.append(df.loc[i, 'Shortname'])  

    # filter the exchange list with correct unit
    for i in lt:
        if i['unit'] == unit:
            exc_w_unit.append(i)

    # exchange list with GLO geography
    for i in exc_w_unit: 
        if i['location'] == 'GLO':
            exc_list_GLO.append(i)

    # find all locations in exc_list
    for i in exc_w_unit:
        exc_loc_list.append(i['location'])

    # if the geography is GLO but we don't have GLO data in evoinvent, return data with location of RoW and RER            
    if test == 'GLO' and len(exc_list_GLO) == 0:
        output_list = find_GLO_location(exc_w_unit)

    # find if we have exactly the same geography
    if test in exc_loc_list:
        output_list = find_exa_location(exc_w_unit, test)
    
    # this applied to situations where the exactly same geography is not found, but we add additional geographies that are close to the case. e.g., 'CA-BC' is not found, we add geographies that are within Canada context
    if len(output_list) == 0:
        if '-' in test:
            state = test.split('-')[0]
            output_list = find_exa_location(exc_w_unit, state)
            output_list1 = find_contained_location(contained_test_loc, exc_w_unit, test)
            output_list2 = find_RoW_location(exc_w_unit)
            output_list3 = find_GLO_location(exc_w_unit)
            output_list += output_list1
            output_list += output_list2
            output_list += output_list3

    # if not, find if the geography is in the contained and overlapping lists
    if len(output_list) == 0:
        output_list = find_contained_location(contained_test_loc, exc_w_unit, test)

    # if not, we assume it's in the rest of the world, since we checked Europe in the above scripts
    if len(output_list) == 0:
        output_list = find_RoW_location(exc_w_unit)
    
    # if not, we assume it's GLO
    if len(output_list) == 0:
        output_list = find_GLO_location(exc_w_unit)
    
    output_list0 = list(set(output_list))

    return output_list0

# search flows in the database
def search_flow(db,flow):
    exc_list1= db.search(flow, limit=10000)
    a = flow
    b = a.lower()
    exc_list2 = [x for x in exc_list1 if b in x['name']] # filter the results with only identification factor appear on name
    if len(exc_list2) != 0: # if no elements in the filtered result, we adopt the initial result
        exc_list = exc_list2
    else:
        exc_list = exc_list1
    return exc_list

# get all exchanges
def get_all_exc(db, input_list, location, db_loc, act, a_db, a_sheet):
    output_list = [a_sheet, a_db, act]
    for i in input_list:
        a_list = []
        a_list.append([i['name'], i['amount']])
        exc_list = search_flow(db, i['name'])
        a_list.append(find_exc_location(location, db_loc, exc_list, i['unit']))
        output_list = output_list + a_list
    return output_list

# get biosphere exchanges
def get_bio_exc(db, input_list):
    output_list = []
    for i in input_list:
        a_list = []
        a_list.append([i['name'], i['amount']])
        bio_list = db.search(i['name'], limit=10000)
        a_list.append(bio_list)
        output_list = output_list + a_list
    return output_list

# get distribution exchange sheet
# a_list = row header
# a_input = distribution exchange list
def get_dis_exc(a_list, a_input):
    b_exchange = [['skip'], a_list]
    for j in range(len(a_input)):
        for k in range(len(a_input[j])):
            if (k % 2)!= 0:
                for i in range(len(a_input[j][k])):
                    a_exchange = [
                        a_input[j][k][i]['name'],
                        a_input[j][k][i]['location'],
                        a_input[j][k-1][1],
                        a_input[j][k][i]['unit'],
                        '',
                        '',
                        '',
                        a_input[j][k][i]['reference product'],
                        a_input[j][k-1][0]
                        ]
                    b_exchange.append(a_exchange)
                b_exchange.append([])
    output_df = pd.DataFrame(b_exchange)
    return output_df

# get exchange sheets
def get_exc_sheet(a_input, a_list, db_name):
    b_exchange = [['skip'], a_list]
    for j in range(len(a_input)):
        if j >= 4:
            if (j % 2) == 0:
                b_exchange.append([a_input[j-1][0]])
                for k in range(len(a_input[j])):
                    a_exchange = [
                        a_input[j][k]['name'],
                        a_input[j][k]['location'],
                        a_input[j - 1][1],
                        a_input[j][k]['unit'],
                        'technosphere',
                        'None',
                        db_name,
                        a_input[j][k]['reference product'],
                        a_input[j][k]['comment']
                    ]
                    b_exchange.append(a_exchange)
                b_exchange.append([])
    return b_exchange

# add biosphere exchenges to exchange sheet
def add_bio_exc(a_input, a_list):
    b_exchange = a_list
    if len(a_input) != 0:
        for j in range(len(a_input)):
            if (j % 2) != 0:
                b_exchange.append([a_input[j-1][0]])
                for k in range(len(a_input[j])):
                    a_exchange = [
                        a_input[j][k]['name'],
                        'None',
                        a_input[j-1][1],
                        a_input[j][k]['unit'],
                        'biosphere']
                    cat = a_input[j][k]['categories'][0]
                    if len(a_input[j][k]['categories']) > 1:
                        for i in range(len(a_input[j][k]['categories'])):
                            if i > 0:
                                cat += '::' + a_input[j][k]['categories'][i]
                    a_exchange.append(cat)
                    a_exchange.append(a_input[j][k]['database'])
                    b_exchange.append(a_exchange)
                b_exchange.append([])
    output_df  = pd.DataFrame(b_exchange)
    return output_df         

# get input sheets
def get_input_sheet(a_input, aloc, a_list):
    # input data sheet
    in_data = [
        ['cutoff', 8],
        [],
        ['Activity', a_input[2]],
        ['location', aloc],
        ['production amount', 1],
        ['type', 'process'],
        ['unit', 'kilogram'],
        [],
        ['Exchanges'],
        a_list,
        [a_input[2], aloc, 1, 'kilogram', 'production','None', 'hydrogen energy system']
        ]
    output_df = pd.DataFrame(in_data)
    return output_df

projects.set_current(utils_config.project_name)

# setup
if not 'biosphere3' in databases:
    bw2setup()
# import Ecoinvent3.5
if not utils_config.db_name in databases:
    fpei = utils_config.ei35_path
    ei = SingleOutputEcospold2Importer(fpei, utils_config.db_name)
    ei.apply_strategies()
    ei.write_database()

ei35 = Database(utils_config.db_name)
bio3 = Database('biosphere3')

# extract information from excel sheet
raw_data_df = pd.read_excel(utils_config.raw_data_path, header=None, index_col=0)

# extract location
raw_loc = raw_data_df.loc['location', 1]

# extract all the locations
ei_35_loc_df = pd.read_csv(utils_config.ei_35_loc_path)

# convert contained location columns to list
for j in range(ei_35_loc_df.shape[0]):
    if type(ei_35_loc_df.loc[j, 'Contained and Overlapping Geographies']) == str:
        contained_loc = ei_35_loc_df.loc[j, 'Contained and Overlapping Geographies'].split(';')
        ei_35_loc_df.at[j, 'Contained and Overlapping Geographies'] = contained_loc
    else:
        ei_35_loc_df.at[j, 'Contained and Overlapping Geographies'] = str(ei_35_loc_df.loc[j, 'Contained and Overlapping Geographies'])

# find the location
for j in range(ei_35_loc_df.shape[0]):
    if ei_35_loc_df.loc[j, 'Name'] == raw_loc:
        loc = ei_35_loc_df.loc[j, 'Shortname']
        break
    else:
        loc = 'GLO'

# extract summary sheet
database_data = pd.read_excel(utils_config.raw_data_path, sheet_name='summary')
database_names = database_data.iloc[0].to_dict()

# replace nan value with 0
database_data = database_data.fillna(0)

# extract activity data
activity_names = []
all_exc_list = []
all_bio_list = []
for x, y in database_names.items():
    for j in range(database_data.shape[0]):
        if j >= 1:
            sheet = database_data.loc[j, x]
            if not sheet == 0:
                sheet_nm = x + sheet
                activity_name = y + ' by ' + sheet
                activity_names.append(activity_name) # extract activity name
                input_exc, bio_exc = extract_data(utils_config.raw_data_path, sheet)
                in_exc_list = get_all_exc(ei35, input_exc, loc, ei_35_loc_df, activity_name, y, sheet_nm)
                bio_exc_list = get_bio_exc(bio3, bio_exc)
                all_exc_list.append(in_exc_list)
                all_bio_list.append(bio_exc_list)

# extract distribution data
dis_data = pd.read_excel(utils_config.raw_data_path, sheet_name='distribution')

# initialize distribution list 
# [
#  [[method1, amount1], 
#   [exchange1, exchange2, ...]],  
#  [[method2, amount2], 
#   [exchange1, exchange2, ...]],
# ]
dis_list = [] 

for i in range(dis_data.shape[0]): # loop row
    sub_dis = []
    sub_dis.append([dis_data.loc[i, 'Method'], dis_data.loc[i, 'Input amount']]) # append [method, amount]
    dis_flow = search_flow(ei35, dis_data.loc[i, 'Input'])
    sub_dis.append(find_exc_location(loc, ei_35_loc_df, dis_flow, 'ton kilometer')) # append [exchange1, exchange2, ...]
    dis_list.append(sub_dis)

# extract lcia method
lcia_method = raw_data_df.loc['LCIA method', 1]

# get impact category list
imp_cat = []
all_methods = list(methods)
my_method = [['skip']] 

if pd.isna(raw_data_df.loc['impact category', 1]) == True:
    for md, key in methods.items():
        if md[0] == lcia_method:
            my_method.append([key['unit']]+list(md))
else:
    for j in range(raw_data_df.shape[1]):
        imp_cat.append(raw_data_df.loc['impact category', j+1])
    # find the method
    for md, key in methods.items():
        if md[0] == lcia_method:
            for j in imp_cat:
                if md[1] == j:
                    my_method.append([key['unit']]+list(md))

# method sheet
my_method_df = pd.DataFrame(my_method)

# database sheet
my_db = [['Database', 'hydrogen energy system'], ['format', 'Excel spreadsheet']]
my_db_df = pd.DataFrame(my_db)

exc_cols = ['name', ' location', ' amount', ' unit', ' type', 'categories', ' database', 'reference product']
dis_cols = exc_cols + ['method']
dis_cols0 = [['skip'], dis_cols]

dis_cols_df = pd.DataFrame(dis_cols0)
dis_ex_df = get_dis_exc(dis_cols, dis_list)

# avoid overwrite
file_path = utils_config.get_excel_path(utils_config.input_file_name, utils_config.output_path)
    
# write the workbook
with pd.ExcelWriter(file_path) as writer:
    my_method_df.to_excel(writer, sheet_name='method', header = False, index = False)
    my_db_df.to_excel(writer, sheet_name='database', header = False, index = False)
    dis_ex_df.to_excel(writer, sheet_name='ex_distribution', header = False, index = False)
    dis_cols_df.to_excel(writer, sheet_name='distribution', header = False, index = False)
    for j in range(len(all_exc_list)):
        all_in_exc_list = get_exc_sheet(all_exc_list[j], exc_cols, utils_config.db_name)
        all_in_exc_list_df = add_bio_exc(all_bio_list[j], all_in_exc_list)
        in_data_df = get_input_sheet(all_exc_list[j], loc, exc_cols)
        exc_sheet_nm = 'ex_' + all_exc_list[j][0]
        in_sheet_nm = 'input_' + all_exc_list[j][0]
        all_in_exc_list_df.to_excel(writer, sheet_name=exc_sheet_nm, header = False, index = False)
        in_data_df.to_excel(writer, sheet_name=in_sheet_nm, header = False, index = False)
