import networkx as nx
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import utils_config
from scipy.optimize import linprog

# add edges
def add_edges_AtoB(alist, blist):
    output_list = []
    for j in range(len(alist)):
        for k in range(len(blist)):
            output_list.append((alist[j], blist[k]))
    return output_list         

# node(n) = sum(node(n+1)/eff)*['portion']
def get_por_cons(aeq, beq, G, po_list, limit, op, position):
    # blist -> coefficient list
    # clist -> portion list
    # anum -> counts # of production node that has assigned portion
    # po_list -> position of nodes that amount > capacity
    # limit -> capacity list
    alist, blist, clist, polist, anum = [], [], [], [], 0
    if op == 1:
        A = 0
        B = position[0]
    
    if op == 2:
        A = position[0]
        B = position[1]

    for i in po_list:
        if A <= i < B:
            polist.append(i)
    
    # find how many node has assigned portion
    for n1, n2 in G.nodes(data=True):
        if n2['layer'] == op:
            if n2['portion'] == 0:
                blist.append(0)
                clist.append(0)
            else:
                blist.append(1)
                clist.append(n2['portion'])
                anum += 1
        else: 
            blist.append(0) 
            clist.append(0) 

    if anum != 0:
        for i in range(len(blist)):        
            if blist[i] == 1:
                alist = list(blist)
                for j in range(len(blist)):
                    if j != i:
                        alist[j] = 0
                if i in polist:
                    aeq.append(alist)
                    beq.append(limit[i])
                else:
                    for n1, n2 in G.nodes(data=True):
                        if n2['layer'] == op + 1:
                            portion = clist[i]
                            alist[B] = (-portion/n2['efficiency'])
                            B += 1
                    aeq.append(alist)
                    beq.append(0)
    return aeq, beq

# usex = (use1 + use2 + use3)*['portion']
# 0*prodx + 0*storx - use1*['portion'] - use2*['portion'] - use3*['portion'] + usex*(1-['portion']) = 0
def get_use_cons(aeq, beq, G, po_list, limit, position):
    # blist -> coefficient list
    # clist -> portion list
    # anum -> counts # of production node that has assigned portion
    # po_list -> position of nodes that amount > capacity
    # limit -> capacity list
    alist, blist, clist, polist, anum = [], [], [], [], 0

    A = position[1]
    B = position[2]

    for i in po_list:
        if A <= i < B:
            polist.append(i)

    # find how many node has assigned portion
    for n1, n2 in G.nodes(data=True):
        if n2['layer'] == 3:
            if n2['portion'] == 0:
                blist.append(0)
                clist.append(0)
            else:
                blist.append(1)
                clist.append(n2['portion'])
                anum += 1
        else:
            blist.append(0)
            clist.append(0)
    
    if anum != 0:
        for i in range(len(clist)):
            if clist[i] != 0:
                alist = list(blist)
                if i in polist:
                    for j in range(len(blist)):
                        if j != i:
                            alist[j] = 0
                    aeq.append(alist)
                    beq.append(limit[i])
                else:
                    for k in range(A, B):
                        if k == i:
                            alist[k] = 1 - clist[i]/100
                        else:
                            alist[k] = -clist[i]
                    aeq.append(alist)
                    beq.append(0)
    return aeq, beq

# get objective function
# min(environmental impacts*nodes) [prodx + storx + usex]
def get_obj_fun(env_imp, G):
    output_list = []
    for n1, n2 in G.nodes(data=True):
        output_list.append(n2[env_imp])
    return output_list   

# sum(nodeA) = sum(nodeB/eff)
def A_to_B(G, layer):
    output_list = []
    for n1, n2 in G.nodes(data=True):
        if n2['layer'] == layer:
            output_list.append(1)
        elif n2['layer'] == layer + 1:
            output_list.append(-100/n2['efficiency'])
        else:
            output_list.append(0)
    return output_list

# nodex = 'amount'
# nodex + 0*prodx + 0*storx + 0*usex = 'amount'
def assi_nodex(G, mode, aeq, beq):
    alist, xlist, ylist, amount, anum = [], [], [], [], 0

    # get default list [0, 0, 0, ... 0] (xlist) and the amount list [0, 0, ... 'amount', ... 0] (ylist)
    for n1, n2 in G.nodes(data=True):
        xlist.append(0)
        amount.append(n2['amount'])
    
    # get list of nodes with assigned amount
    for n1, n2 in G.nodes(data=True):
        if n2['layer'] in mode:
            ylist.append(anum)
        anum += 1
    
    # change the coefficent of the node with assigned amount to 1
    for i in ylist:
        alist = list(xlist)
        alist[i] = 1
        aeq.append(alist)
        beq.append(amount[i])
    return aeq, beq

# sum(nodex) = tot_node
def sum_node(G, mode):
    output_list = []
    for n1, n2 in G.nodes(data=True):
        if n2['layer'] in mode:
            output_list.append(1)
        else:
            output_list.append(0)
    return output_list

# get constraints
def get_all_cons(G, mode, tot, po_list, limit, position):
    aeq, beq = [], []

    # prod1 + prod2 + prod3 = stor1/eff + stor2/eff + stor3/eff
    # prod1 + prod2 + prod3 - stor1/eff - stor2/eff - stor3/eff + 0*use1 + 0*use2 + 0*use3 = 0
    aeq.append(A_to_B(G, 1))
    beq.append(0)

    # stor1 + stor2 + stor3 = use1/eff + use2/eff + use3/eff
    # 0*prod1 + 0*prod2 + 0*prod3 + stor1 + stor2 + stor3 - use1/eff - use2/eff - use3/eff = 0
    aeq.append(A_to_B(G, 2))
    beq.append(0)

    if mode[0] == 'specific':
        # nodex = ['amount']
        aeq, beq = assi_nodex(G, mode, aeq, beq)
    else:
        # sum(nodex) = tot_node
        aeq.append(sum_node(G, mode))
        beq.append(tot)

    if 'p_p' in mode:
        # prodx = (stor1/eff + stor2/eff + stor3/eff +...)*['portion'] 
        # prodx - (stor1/eff + stor2/eff + stor3/eff +...)*['portion'] + 0*usex = 0
        aeq, beq = get_por_cons(aeq, beq, G, po_list, limit, 1, position)
    
    if 's_p' in mode:
        # storx = tot_stor*['portion'] or capacity
        # 0*prodx + storx + 0*usex = tot_stor*['portion'] or capacity
        aeq, beq = get_por_cons(aeq, beq, G, po_list, limit, 2, position)
    
    if 'u_p' in mode:
        # usex = (use1 + use2 + use3)*['portion']
        # 0*prodx + 0*storx + use1*['portion'] + use2*['portion'] + use3*['portion'] + usex*(1-['portion']) = 0
        aeq, beq = get_use_cons(aeq, beq, G, po_list, limit, position)

    return aeq, beq

# node amount less thant capacity
def get_all_bound(G, min, max):
    output_list = []
    for n1, n2 in G.nodes(data=True):
        if n2[max] == 'None' or n2[max] == 0:
            output_list.append((n2[min], None))
        else:
            output_list.append((n2[min], n2[max]))
    return output_list 

# infeasible solution
def err_reg(G, aeq, beq, C, Bound, mode, position, tot):
    # run linear programming with unlimited capacity
    Bound0 = get_all_bound(G, 'min_amount', None)
    res0 = linprog(C, A_eq = aeq, b_eq= beq, bounds=Bound0)
    num_list, cap = [], []
    # num_list -> position of nodes that amount > capacity
    # cap -> capacity list

    for i in range(0, position[2]):
        attr = list(G.nodes(data=True))[i][1]
        cap.append(attr['capacity'])
        if res0.x[i] > attr['capacity']:
            num_list.append(i)
        
    aeq1, beq1 = get_all_cons(G, mode, tot, num_list, cap, position)
    res = linprog(C, A_eq = aeq1, b_eq= beq1, bounds=Bound) 
    return res

def run_linprog(env_imp, env_imps, G, mode, tot, position, dict):
    solutions, results = [], []

    # get objective function
    C = get_obj_fun(env_imp, G)

    # get constraints
    aeq, beq = get_all_cons(G, mode, tot, [], [], position)

    # get bounds
    Bound = get_all_bound(G, 'min_amount', 'capacity')

    # linear programming
    res = linprog(C, A_eq = aeq, b_eq= beq, bounds=Bound)

    if 'The problem is infeasible.' in res.message:
        if len(mode) < 2:
            solutions.append(['error', 'missing infomation'])
            results.append(['error', 'no solutions'])
        if len(mode) == 2:
            solutions.append(['error', 'assigned amount out of capacity'])
            results.append(['error', 'no solutions'])
        else:
            res = err_reg(G, aeq, beq, C, Bound, mode, position, tot)
            solution = ['solution']
            for k in res.x:
                solution.append(k)
            solutions.append(solution)
            solutions.append(['note', 'The amount exceeds the capacity in some processes. The corresponding amount has been adjusted to match the capacity.'])
            results.append([env_imp, res.fun])
            for i in env_imps:
                if i != env_imp:
                    A = get_obj_fun(i, G)
                    result = sum(x * y for x, y in zip(res.x, A))
                    results.append([i, result])
    else:
        solution = ['solution']
        for k in res.x:
            solution.append(k)
        solutions.append(solution)
        for i in env_imps:
            re = [i]
            for k in range(len(solution)):
                if k > 0:
                    a = G.nodes[list(G.nodes())[k-1]][i]*solution[k]
                    re.append(a)
            A = get_obj_fun(i, G)
            re.append(sum(x * y for x, y in zip(res.x, A)))
            re.append(dict[i])
            results.append(re)
        
    return solutions, results

# set up graph
G = nx.DiGraph()

# import excel file
df_scenario = pd.read_excel(utils_config.supp_path, sheet_name='scenario', index_col=0, header=None)
df_prod = pd.read_excel(utils_config.supp_path, sheet_name='H2 production', index_col=0)
df_stor = pd.read_excel(utils_config.supp_path, sheet_name='H2 storage', index_col=0)
df_use = pd.read_excel(utils_config.supp_path, sheet_name='H2 use', index_col=0)

# replace NaN with 0
df_scenario = df_scenario.fillna(0)
df_prod = df_prod.fillna(0)
df_stor = df_stor.fillna(0)
df_use = df_use.fillna(0)

tot_prod = df_scenario.loc['H2 production', 1]
tot_stor = df_scenario.loc['H2 storage', 1]
tot_use = df_scenario.loc['H2 demand', 1]
d_dis = df_scenario.loc['Distribution', 1]
op_imp = df_scenario.loc['Optimization parameter', 1]

position = [
    df_prod.shape[1]- 1, 
    df_prod.shape[1] + df_stor.shape[1] - 2, 
    df_prod.shape[1] + df_stor.shape[1] + df_use.shape[1] - 3
    ]

# add nodes
# G.add_node('raw material', layer = 1)
prod, stor, use = [], [], []

# layer: 1 -> production; 2 -> storage; 3 -> use
# add H2 production nodes
for i in range(df_prod.shape[1]):
    if i >= 1:
        name = 'H2 production by ' + df_prod.columns[i]
        prod.append(name)
        G.add_node(
            name, 
            layer = 1, 
            capacity = df_prod.loc['Total capacity', df_prod.columns[i]], 
            portion = df_prod.loc['Anticipated portion', df_prod.columns[i]],
            amount = df_prod.loc['Anticipated amount', df_prod.columns[i]],
            min_amount = df_prod.loc['Minimum amount', df_prod.columns[i]]
            )

# add H2 storage nodes
for i in range(df_stor.shape[1]):
    if i >= 1:
        name = 'H2 storage by ' + df_stor.columns[i]
        stor.append(name)
        G.add_node(
            name, 
            layer = 2, 
            capacity = df_stor.loc['Total capacity', df_stor.columns[i]], 
            portion = df_stor.loc['Anticipated portion', df_stor.columns[i]], 
            efficiency = df_stor.loc['Efficiency', df_stor.columns[i]],
            amount = df_stor.loc['Anticipated amount', df_stor.columns[i]],
            min_amount = df_stor.loc['Minimum amount', df_stor.columns[i]]
            )    

# add H2 use nodes
for i in range(df_use.shape[1]):
    if i >= 1:
        name = 'H2 use by ' + df_use.columns[i]
        use.append(name)
        G.add_node(
            name, 
            layer = 3, 
            amount = df_use.loc['Anticipated amount', df_use.columns[i]], 
            efficiency = df_use.loc['Efficiency', df_use.columns[i]],
            portion = df_use.loc['Anticipated portion', df_use.columns[i]], 
            min_amount = df_use.loc['Minimum amount', df_use.columns[i]],
            capacity = df_use.loc['Capacity', df_use.columns[i]]
            )    
        
# add environmental impact attributes
env_imp = pd.read_excel(utils_config.LCA_result_path)
env_imp = env_imp.T.dropna()

env_imps = []
attrs = {}
env_unit = {}

for j in range(env_imp.shape[1]): # loop column
    attr = {}
    if 'production' in env_imp.iloc[0][j] or 'use' in env_imp.iloc[0][j]: # for production and application node, add environmental impact directly
        for k in range(env_imp.shape[0]): # loop row
            if k > 0:
                attr.update({env_imp.index[k]: env_imp.iloc[k, j]}) #{environmental impact category: environmental impact}
    if 'storage' in env_imp.iloc[0][j]:
        pro = env_imp.iloc[0][j].split(' by ')[1] # get storage process name (e.g., compression)
        for i in range(env_imp.shape[1]): # loop column
            if 'distribution' in env_imp.iloc[0][i] and pro in env_imp.iloc[0][i]: # check in header
                m = i              
        for k in range(env_imp.shape[0]): # loop row
            if k > 0:
                new_imp = env_imp.iloc[k, j] + d_dis*env_imp.iloc[k, m] # storage impact + distance * distribution impact
                attr.update({env_imp.index[k]: new_imp})
    attrs.update({env_imp.iloc[0][j]: attr})
nx.set_node_attributes(G, attrs)

for i in range(env_imp.shape[0]):
    if i > 0:
        env_imps.append(env_imp.index[i])
        env_unit.update({env_imp.index[i]: env_imp[0][i]})

# add edges
prod_to_stor = add_edges_AtoB(prod, stor)
stor_to_use = add_edges_AtoB(stor, use)

G.add_edges_from(prod_to_stor)
G.add_edges_from(stor_to_use)

# show the plot
fig, ax = plt.subplots()
pos = nx.multipartite_layout(G, subset_key='layer')
nx.draw(G, ax=ax, pos = pos, with_labels=True, font_size=6)

cond_prod, cond_stor, cond_use = [], [], []

for n1, n2 in G.nodes(data=True):
    if n2['layer'] == 3:
        if n2['amount'] != 0:
            cond_use.append('amount')
        elif n2['min_amount'] != 0:
            cond_use.append('min amount')
        elif n2['portion'] !=0:
            cond_use.append('portion')
    if n2['layer'] == 2:
        if n2['amount'] != 0:
            cond_stor.append('amount')
        elif n2['min_amount'] != 0:
            cond_stor.append('min amount')
        elif n2['portion'] != 0 :
            cond_stor.append('portion')
    if n2['layer'] == 1:
        if n2['amount'] != 0:
            cond_prod.append('amount')
        elif n2['min_amount'] != 0:
            cond_prod.append('min amount')
        elif n2['portion'] != 0 :
            cond_prod.append('portion')   

cond_prod = set(cond_prod)    
cond_stor = set(cond_stor)
cond_use = set(cond_use)

# general cases
mode, tot = ['default'], 0

if tot_prod != 0: 
    mode.append(1)
    tot = tot_prod
elif tot_stor != 0:
    mode.append(2)
    tot = tot_stor
elif tot_use != 0:
    mode.append(3)
    tot = tot_use

if tot == 0:
    mode[0] = 'specific'
else:
    mode[0] = 'total'

if 'amount' in cond_prod:
    mode.append(1)
elif 'amount' in cond_stor:
    mode.append(2)
elif 'amount' in cond_use:
    mode.append(3)

if 'portion' in cond_prod:
    mode.append('p_p')
elif 'portion' in cond_stor:
    mode.append('s_p')
elif 'portion' in cond_use:
    mode.append('u_p')

solutions, results = run_linprog(op_imp, env_imps, G, mode[:2], tot, position, env_unit) 
if 'error' not in solutions[0]:
    solutions, results  = run_linprog(op_imp, env_imps, G, mode, tot, position, env_unit) 

header = ['process']
header = header + list(G.nodes()) + ['total impact', 'unit']
output = [header] + solutions + results
output_df = pd.DataFrame(output)

# avoid overwrite
file_path = utils_config.get_excel_path(utils_config.env_result, utils_config.output_path)

# write excel
output_df.to_excel(file_path, header=False, index=False)
