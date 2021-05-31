# Extract keywords from multiple PDF files, create a dataframe, then export it to an .xlsx file.

import os                        # provides functions for creating and removing a directory (folder), fetching its contents, changing and identifying the current directory
import pandas as pd              # flexible open source data analysis/manipulation tool
import glob                      # generates lists of files matching given patterns
import pdfplumber                # extracts information from .pdf documents

"""
Obtain key words from repetitive documents, then extract as a dataframe to an .xlsx !
"""

# defining the functions used in main()
def get_keyword(start, end, text):
    """
    start: should be the word prior to the keyword.
    end: should be the word that comes after the keyword.
    text: represents the text from the page(s) you've just extracted.
    """
    for i in range(len(start)):
        try:
            field = ((text.split(start[i]))[1].split(end[i])[0])
            return field
        except:
            continue

def main():
    # create an empty dataframe, from which keywords from multiple .pdf files will be later appended by rows.
    my_dataframe = pd.DataFrame()

    for files in glob.glob("C:\\Users\IceBanshee\Desktop\python test docs\sampledocs\*.pdf"):
        with pdfplumber.open(files) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            text = " ".join(text.split())

            # use the function get_keyword as many times to get all the desired keywords from a pdf document.

            # obtain keyword #1
            start = ['Curabitur']
            end = ['ante']
            keyword1 = get_keyword(start, end, text)

            # obtain keyword #2
            start = ['Duis quis']
            end = ['nunc']
            keyword2 = get_keyword(start, end, text)

            # obtain keyword #3
            start = ['Aenean']
            end = ['mi']
            keyword3 = get_keyword(start, end, text)

            # obtain keyword #4
            start = ['Row 1 ']
            end = [' ']
            keyword4 = get_keyword(start, end, text)

            # obtain keyword #5
            start = ['Row 2 ']
            end = [' ']
            keyword5 = get_keyword(start, end, text)

            # obtain keyword #6
            start = ['Row 3 ']
            end = [' ']
            keyword6 = get_keyword(start, end, text)

            # create a list with the keywords extracted from current document.
            my_list = [keyword1, keyword2, keyword3, keyword4, keyword5, keyword6]

            # append my list as a row in the dataframe.
            my_list = pd.Series(my_list)

            # append the list of keywords as a row to my dataframe.
            my_dataframe = my_dataframe.append(my_list, ignore_index=True)

            print("Document's keywords have been extracted successfully!")

    # rename dataframe columns using dictionaries.
    my_dataframe = my_dataframe.rename(columns={0:'Keyword 1',
                                                    1:'Keyword 2',
                                                    2:'Keyword 3',
                                                    3:'Keyword 4',
                                                    4:'Keyword 5',
                                                    5:'Keyword 6'})

    # change my current working directory
    save_path = ('C:\\Users\IceBanshee\Desktop\python test docs\sampleexcel')
    os.chdir(save_path)

    # extract my dataframe to an .xlsx file!
    my_dataframe.to_excel('sample_excel.xlsx', sheet_name = 'my dataframe')
    print("")
    print(my_dataframe)

if __name__ == '__main__':
    main()