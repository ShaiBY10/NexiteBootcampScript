from shutil import move
from pandas import read_excel, RangeIndex, Series, DataFrame
import numpy as np
from tkinter import filedialog
import os
from datetime import datetime


def browseFiles():
    filepath = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Text files",
                                                      "*"),
                                                     ("all files",
                                                      "*.*")))
    return filepath


def duplicatesCheck(input_list):
    """
    Checks for duplicates in the input Excel file.
    :param: input_list:
    :return bool:
    """
    print('Checking for duplicates...')
    if len(input_list) != len(set(input_list)):
        print('Found duplicates in the excel file!')
        return True
    else:
        print('No duplicates found! Great Job :)')
        return False


def addEmptyRows(df, n_empty, n_steps):
    """ adds 'n_empty' empty rows every 'n_steps' rows  to 'df'.
        Returns a new DataFrame. """

    # to make sure that the DataFrame index is a RangeIndex(start=0, stop=len(df))
    # and that the original df object is not mutated.
    df = df.reset_index(drop=True)

    # length of the new DataFrame containing the NaN rows
    len_new_index = len(df) + n_empty * (len(df) // n_steps)
    # index of the new DataFrame
    new_index = RangeIndex(len_new_index)

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


def listSplit(input_list, chunk_size=104):
    """
     Takes a list and returns list of `n` splited list

    :param input_list:
    :param chunk_size:
    :return `list`:
    """
    for i in range(0, len(input_list), chunk_size):
        yield input_list[i:i + chunk_size]

def CreateAndMoveRawFile(raw_file):
    #Check if directory exists
    directory_content = os.listdir(os.path.dirname(raw_file))
    if "Raw file" in directory_content:
        dstn = move(raw_file, os.path.join(os.path.dirname(raw_file),"Raw file"))
        print('DSTNNNN', dstn)
    else:
        print('case where there is no rawfile folder')
        os.mkdir(os.path.join(os.path.dirname(raw_file),"Raw file"))
        dstn = move(raw_file,os.path.join(raw_file,os.path.dirname(raw_file),"Raw file"))
        # print('Created new folder named: "Raw file", The raw file is there now :)')
        return dstn

def listToMatrix(raw_file, output_name=''):
    """
    Takes a one col Excel file containing NUID's and convert it into a bootcamp matrix Excel file
    :param output_name: name of output file
    :param raw_file: one line Excel file
    :param output_name: output file name
    :return: df # pandas DataFrame before exporting to .xlsx
    """
    global idx
    df = read_excel(raw_file, header=None)
    print(raw_file)
    df.index += 1
    raw_list = [value for cell, value in df[0].iteritems()] # list of first col of excel file
    raw_list = [str(i)[-6:] for i in raw_list]  # Take only last 6 chars of each item in list
    duplicatesCheck(raw_list)
    boards_list = list(listSplit(raw_list)) # spereate each 104 nuids to one board
    boards = list(map(Series, boards_list)) # convert every board in the board list to series
    reshaped_list = list()
    for board in boards:
        reshaped_list.append(board.values.reshape((8,13),order='F')) # Create a matrix with a shape of (8,13) to each board in board list
    df = DataFrame(np.concatenate(reshaped_list)) # Stack all boards into one dataframe
    new_df = addEmptyRows(df, 1, 8)
    new_df.to_excel(output_name, header=None, index=False,engine='openpyxl') # Export to excel
    print(f"Attention! The length of the input file is {len(raw_list)}, \nCreated {len(boards_list)} Boards!")
    new_df = CreateAndMoveRawFile(raw_file)
    return new_df


def getTime():
    now = datetime.now()
    now = now.strftime('%d-%m_%H%M')
    return now
