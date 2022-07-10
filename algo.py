import pandas as pd
import numpy as np
from tkinter import filedialog

def browseFiles():
    filepath = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Text files",
                                                      "*"),
                                                     ("all files",
                                                      "*.*")))
    # label_file_explorer.configure(text="File Opened: " + filename)
    print(filepath)
    return filepath


def createAndSave():
    filepath = filedialog.askopenfilename(initialdir='/',
                                          title='Select a saving directory',
                                          filetypes='*')
    listToMatrix(filepath)



def duplicates_check(input_list):
    """
    Checks for duplicates in the input excel file.
    :param input_list:
    :return bool:
    """
    print('Checking for duplicates...')
    if len(input_list) != len(set(input_list)):
        print('Found duplicates in the excel file!')
        return True
    else:
        print('No duplicates found! Great Job :)')
        return False


def add_empty_rows(df, n_empty, n_steps):
    """ adds 'n_empty' empty rows every 'n_steps' rows  to 'df'.
        Returns a new DataFrame. """

    # to make sure that the DataFrame index is a RangeIndex(start=0, stop=len(df))
    # and that the original df object is not mutated.
    df = df.reset_index(drop=True)

    # length of the new DataFrame containing the NaN rows
    len_new_index = len(df) + n_empty * (len(df) // n_steps)
    # index of the new DataFrame
    new_index = pd.RangeIndex(len_new_index)

    # add an offset (= number of NaN rows up to that row)
    # to the current df.index to align with new_index.
    df.index += n_empty * (df.index
                           .to_series()
                           .groupby(df.index // n_steps)
                           .ngroup())

    # reindex by aligning df.index with new_index.
    # Values of new_index not present in df.index are filled with NaN.
    new_df = df.reindex(new_index)

    return new_df


def list_split(input_list, chunk_size=104):
    """
     Takes a list and returns list of `n` splited list
    :param input_list:
    :param chunk_size:
    :return `list`:
    """
    for i in range(0, len(input_list), chunk_size):
        yield input_list[i:i + chunk_size]


def listToMatrix(raw_file,output_path='generated_matrix'):
    """
    Takes an one col excel file containing nuids and convert it into a bootcamp matrix excel file
    :param raw_file: one line excel file
    :param output_name: output file name
    :return: df # pandas DataFrame before exporting to .xlsx
    """
    df = pd.read_excel(raw_file, header=None)
    df.index += 1
    raw_list = [value for cell, value in df[0].iteritems()]
    duplicates_check(raw_list)
    spereted_list = list(list_split(raw_list))
    boards = list(map(pd.Series, spereted_list))
    reshaped_list = list()
    for board in boards:
        reshaped_list.append(board.values.reshape(8, 13, order='F'))
        for idx in range(len(boards)):
            idx += 1

    print(f"Attention! The length of the input file is {len(raw_list)}, \nCreated {idx} Boards!")

    df = pd.DataFrame(np.concatenate(reshaped_list))
    new_df = add_empty_rows(df, 1, 8)
    new_df.to_excel(output_path + r'/output.xlsx', header=None, index=False)
    print(f"Created .xlsx file in {output_path}")
    return new_df


