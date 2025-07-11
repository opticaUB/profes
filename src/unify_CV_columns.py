#!/usr/bin/env python
# coding: utf-8

import sys, os
import glob
import numpy as np
import pandas as pd
import regex

from statistics import mean

UNIFIED_SUFFIX = "_unified"  # Final filename's suffix
GRUP_LAB = "gLab."  # Column name for the Laboraty group
GRUP_TEO = "gTeo."  # Column name for the Theory group

ORAL_NEW = "Oral"  # New column name for Oral Exam
ORAL_PATTERN = "prova oral"  # substring to be contained in expected oral columns
ORAL_EXPECTED = 1  # Final expected marks

INFORMES_NEW = "Informes"  # New column name for Averaged reports
INFORMES_PATTERN = "entrega informes"  # substring to be contained in expected reports columns
INFORMES_EXPECTED = 7  # Final expected marks to be able to make the average

MISSING_SESSIONS = "Missing sessions"  # column name for the missing sessions

POSSIBLE_GRUPS_LAB = "M1A|M1B|M2A|M2B|T1A|T1B"  # regex pattern for possible teo groups
GRUPS_TEO = ['Q2_M1', 'Q2_M2', 'Q2_T1', 'Q1_T1']  # possibles grups teorics
Q1T1_LABEL = 'ReAv.'  # Nom del grup de tardor
SESSION_PATTERN = r"Sessió\s+(\d+)\s+"  # regex pattern to look for session number
GRUPS_LABEL_CV = 'Grups'  # Column name for grups in participants Excel
NIUB = 'Número ID'  # ID column name

FINAL_KEYS = ['Nom', 'Cognoms', NIUB,  # Desired columns in the final file
              ORAL_NEW, INFORMES_NEW, MISSING_SESSIONS, 
              GRUP_TEO, GRUP_LAB]  

def load_main(pattern):
    fn_not_exists = True

    dir_list = [fn for fn in glob.glob(pattern) if UNIFIED_SUFFIX not in fn]
    excel_filename = dir_list[0] if len(dir_list) == 1 else None
    fn_not_exists = not os.path.isfile(excel_filename or '')

    while fn_not_exists:
        excel_filename = input("Nom del fitxer Excel:")
        excel_filename = (excel_filename if excel_filename.endswith('.xlsx') 
                          else excel_filename+'.xlsx')
        fn_not_exists = not os.path.isfile(excel_filename)
        print(f"\n >>> {excel_filename} does NOT found. Try again," if fn_not_exists 
              else "Found!")

    return pd.read_excel(excel_filename, sheet_name=0), excel_filename

def load_part(pattern):

    dir_list2 = [fn for fn in glob.glob(pattern)]
    excel_fn_part = dir_list2[0] if len(dir_list2) == 1 else None
    part_exists = os.path.isfile(excel_fn_part or '')

    while not part_exists:
        print("  (Deixa en blanc si vols ignorar-ho)")
        excel_fn_part = input("Nom del fitxer Excel (participants):") 
        if excel_fn_part == '': break
        excel_fn_part = (excel_fn_part if excel_fn_part.endswith('.xlsx') 
                         else excel_fn_part+'.xlsx')
        part_exists = os.path.isfile(excel_fn_part)
        print(f"\n >>> {excel_fn_part} does NOT found. Try again," 
              if excel_fn_part else "Found!")

    return pd.read_excel(excel_fn_part, sheet_name=0) if part_exists else None

def unify_columns(data, text_in_column, new_column_name, expected_items=1, 
                  expected_types=(float, int), delete=True):

    keys = data.keys()

    miss_sess = []

    colums_names = [c for c in keys if text_in_column.lower() in c.lower()]

    cols = data[keys[0]]
    new_values = np.zeros_like(cols)
    grup_col = np.zeros_like(cols) if GRUP_LAB not in keys else None

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
                    match = regex.search(POSSIBLE_GRUPS_LAB, column)
                    grup = match.group() if match else ""
                matchS = regex.search(SESSION_PATTERN, column)
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
                    missing = ['s'+str(s) for s in range(1, expected_items+1) if s not in sess]
        
        miss_sess.append(','.join(missing))


    if grup_col is not None:
        data[GRUP_LAB] = grup_col
        
    data[new_column_name] = new_values

    if delete:
        data = data.drop(columns=colums_names)

    if expected_items > 1:
        data[MISSING_SESSIONS] = miss_sess

    return data


def unify(data):

    data = unify_columns(data, ORAL_PATTERN, ORAL_NEW, ORAL_EXPECTED)

    return unify_columns(data, INFORMES_PATTERN, INFORMES_NEW, INFORMES_EXPECTED)


def find_grup_teo(data, data_part):

    if data_part is None:
        return data

    grup_teo_list = []
    for idx, alumne in data_part.iterrows():
        grups_str = alumne[GRUPS_LABEL_CV].split(',')[0]
        for cand in GRUPS_TEO:
            if cand in grups_str:
                gt = cand[3:] if cand.startswith('Q2') else 'ReAv.'
                break
            else:
                gt = '-'
        grup_teo_list.append(gt)

    data_part[GRUP_TEO] = grup_teo_list
    return data.merge(data_part[[NIUB, GRUP_TEO]], on=NIUB, how='left')


def reformating(data, keep_AC=False):
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

    desired_columns = FINAL_KEYS + renaming_news if keep_AC else FINAL_KEYS

    drop_keys = [k for k in data.keys() if k not in desired_columns]
    data = data.drop(columns=drop_keys)

    return data

def print_summary(data):

    print("Resulting file:\n")
    print(data[:5])
    print("  (això és un resum de la taula final)")
    print("---------------\n")

    print("No values found: (repetidors?)")
    repes = data[(data[INFORMES_NEW]==-9) & (data[ORAL_NEW]==-9)]
    print('\n'.join([rn+' '+rc for rn, rc in zip(repes['Nom'], repes['Cognoms'])])
          if len(repes)<10 else f"{len(repes)} estudiants no tenen cap nota. ")
    print("---------------\n")

    print("Missing grades: (falta alguna nota?)")
    missi = data[((data[INFORMES_NEW]>-9) &
                  (data[INFORMES_NEW]<0) )    | 
                 ((data[ORAL_NEW]>-9) &
                  (data[ORAL_NEW]<0))]

    missing_dict = {}
    for idx, m_st in missi.iterrows():
        grup = m_st[GRUP_LAB]
        num1, num2 = missing_dict.get(grup, (0, 0))
        missing_dict[grup] = (num1+1, num2+m_st[INFORMES_NEW])

    for gr in missing_dict.keys():
        print(f"{gr}: {-missing_dict[gr][1]} (de {missing_dict[gr][0]} alumnes)")
    print(" ")    

    print('Sessions que falten:')
    print('\n'.join([f'{nota}: ({grup}) {rn} {rc} -> {sess}' 
                     for rn, rc, nota, grup, sess 
                     in zip(missi['Nom'], missi['Cognoms'], 
                            missi[INFORMES_NEW], missi[GRUP_LAB], 
                            missi[MISSING_SESSIONS])]))
    print("---------------\n")

def export_to_excel(data, filename, new_filename=None):
    new_filename = new_filename or filename.replace('.xlsx', UNIFIED_SUFFIX+".xlsx")
    try:
        if GRUP_TEO in data.keys():
            with pd.ExcelWriter(new_filename, engine='openpyxl') as writer:
                for nom_grup, df_grup in data.groupby(GRUP_TEO):
                    if nom_grup != Q1T1_LABEL:
                        df_grup.to_excel(writer, sheet_name=nom_grup, index=False)
        else:
            data.to_excel(new_filename)
    except PermissionError:
        print(" >>> No s'ha pogut desar, probablement perquè el document ja existeix "
              "i ESTÀ OBERT.\n     Pot existir, però ha d'estar tancat.")
    

def export(data, filename):
    data.to_excel(filename.replace('.xlsx', UNIFIED_SUFFIX+".xlsx"))


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