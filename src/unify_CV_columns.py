#!/usr/bin/env python
# coding: utf-8

import sys
import numpy as np
import pandas as pd
import regex

from statistics import mean


def unify_columns(data, text_in_column, new_column_name, expected_items=1, 
                  expected_types=(float, int), delete=True):

    grup_str = "Grup"

    keys = data.keys()

    miss_sess = []

    colums_names = [c for c in keys if text_in_column.lower() in c.lower()]

    cols = data[keys[0]]
    new_values = np.zeros_like(cols)
    grup_col = np.zeros_like(cols) if grup_str not in keys else None

    for idx, alumne in data.iterrows():
        vals = []
        sess = []
        missing = []
        grup = ""
        for column in colums_names:
            value = alumne[column]
            if type(value) in (float, int):
                vals.append(value)
                if grup_col is not None and not grup:
                    match = regex.search("M1A|M1B|M2A|M2B|T1A|T1B", column)
                    grup = match.group() if match else ""
                matchS = regex.search(r"Sessió\s+(\d+)\s+", column)
                sess.append(int(matchS.group(1)) if matchS else "")

        if grup_col is not None:
            grup_col[idx] = grup

        num_vals = len(vals)
        if num_vals == expected_items:  # average of several marks or just that mark.
            new_values[idx] = round(mean(vals), 3) if expected_items > 1 else vals[0]
        elif num_vals == 0:
            new_values[idx] = -9  # error code -9 for no values found
        else:
            if len(set([round(v, 2) for v in vals])) == 1:  # If several but equals
                new_values[idx] = vals[0]  # take the first
            else:
                new_values[idx] = -1*abs(num_vals-expected_items)  # set an error code
                if expected_items > 1 :
                    missing = [str(s) for s in range(1, expected_items+1) if s not in sess]
        
        miss_sess.append(','.join(missing))


    if grup_col is not None:
        data[grup_str] = grup_col
        
    data[new_column_name] = new_values

    if delete:
        data = data.drop(columns=colums_names)

    if expected_items > 1:
        data["Missing sessions"] = miss_sess

    return data


def unify_informes(data):

    text_in_column = "entrega informes"
    expected_items = 7
    new_column_name = "Mitjana informes"

    return unify_columns(data, text_in_column, new_column_name, expected_items)


def unify_oral(data):

    text_in_column = "prova oral"
    expected_items = 1
    new_column_name = "Prova oral unif."

    return unify_columns(data, text_in_column, new_column_name, expected_items)

def reformating(data):
    """ Long names are not usefull
    """
    keys = data.keys()
    renaming_olds = [c for c in keys if "continuada" in c.lower()]
    renaming_news = [c.replace("Avaluació Continuada",
                               "AC").replace("Tasca:", 
                                             "").replace("(Real)",
                                                         "").replace("Exercici", 
                                                                     "Exerc.") 
                     for c in renaming_olds]

    renaming = {k: v for k, v in zip(renaming_olds, renaming_news)}

    data = data.rename(columns=renaming)

    data = data.drop(columns=["Darrera descàrrega des d'aquest curs",
                              "Adreça electrònica"])


    return data


def export(data, filename):
    data.to_excel(filename.replace('.xlsx', "_unified.xlsx"))


if __name__ == "__main__":
    
    args = sys.argv

    filename = args[1] if len(args)>1 else "NotesFromCV.xlsx"
    
    sheet = 0  # sheet name or sheet number or list of sheet numbers and names

    data = pd.read_excel(filename, sheet_name=sheet)

    data = unify_informes(data)
    data = unify_oral(data)
    
    data = reformating(data)

    print("Resulting file:\n")
    print(data)

    export(data, filename)