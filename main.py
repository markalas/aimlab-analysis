import pandas as pd
import os
import numpy as np
import pathlib as Path
import pyodbc
import openpyxl

# Directories
current_dir = os.curdir
aim_lab_data_dir = os.path.join(current_dir, "aimlab_data",)
aim_lab_output_dir = os.path.join(current_dir, "aimlab_output_dir")

# Import aimlab csv data
grid_shot_data = os.path.realpath(aim_lab_data_dir + "\Gridshot.csv")

# Read in and Transform Aimlab Data
def movecol(df, cols_to_move=[], ref_col='', place='After'): # movecol function to reorder dataframe columns
    
    cols = df.columns.tolist()
    if place == 'After':
        seg1 = cols[:list(cols).index(ref_col) + 1]
        seg2 = cols_to_move
    if place == 'Before':
        seg1 = cols[:list(cols).index(ref_col)]
        seg2 = cols_to_move + [ref_col]
    
    seg1 = [i for i in seg1 if i not in seg2]
    seg3 = [i for i in cols if i not in seg1 + seg2]
    
    return(df[seg1 + seg2 + seg3])

grid_shot_raw = pd.read_csv(grid_shot_data) # read in grid_shot_data
grid_shot_df = grid_shot_raw.copy() # create a copy of grid_shot_df
grid_shot_df.drop(columns = ["weaponName", "map" , "version"], inplace = True) # drop unecessary columns
grid_shot_df[["createDate", "createTime"]] = grid_shot_df["createDate"].str.split(" ", expand=True) # split createDate column to new columns "createDate" and "createTime"
grid_shot_df = movecol(grid_shot_df, cols_to_move=['createTime'], ref_col='createDate', place='After') # reorder columns with movecol

grid_shot_df.to_excel(aim_lab_output_dir + "\Gridshot_clean.xlsx", index=False) # output dataframe to_excel

print(grid_shot_df)

