import pandas as pd
import numpy as np

import sailing_classes as sc

# Load input data
df_camper = pd.read_excel('input_data.xlsx', sheet_name='Camper', index_col=0)
df_leader = pd.read_excel('input_data.xlsx', sheet_name='Leader', index_col=0)
df_boat = pd.read_excel('input_data.xlsx', sheet_name='Boat', index_col=0)
df_schedule = pd.read_excel('input_data.xlsx', sheet_name='Schedule', index_col=0)

# Initiate camp class
camp_plan = sc.Camp(df_camper, df_leader, df_boat, df_schedule)
# print(camp_plan.balance_log.head(5))
# print('Session 1 beach:', camp_plan.allocations[0].sessions[0].crew[-1].campers)
# print('Session 2 beach:', camp_plan.allocations[0].sessions[1].crew[-1].campers)
# print('Session 3 beach:', camp_plan.allocations[0].sessions[2].crew[-1].campers)
camp_plan.excel_export()