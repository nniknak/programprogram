import pandas as pd
import numpy as np
import argparse
from openpyxl import formatting, styles
from datetime import datetime

# another nice thing would be if the data frame kept track of the index ... reported the theater name with it ... or inserted / created a pretty final version by itself
def df_byletter(df, alphabetletter):
    return df.loc[df['Letter']==alphabetletter]

# this is the one we're using now
def get_letter(theatername):
        theater = str(theatername)
        if theater.startswith("The "):
            return theater[4]
        elif theater.startswith("A "):
            return theater[2]
        elif theater.startswith("!"):
            return theater[1]
        elif theater.startswith("(New) "):
            return theater[6]
        else:
            return theater[0]

if __name__ == "__main__":
    # allow for input of filename as argument
    parser = argparse.ArgumentParser(description='Get the spreadsheet filename.')
    parser.add_argument('filename',
                        help='the downloaded spreadsheet filename', 
                        nargs='?',
                        default='Theater Programs by Theater Name.xlsx')

    args = parser.parse_args()

    # clear report file -- could also generate a new text file with a new name (datetime generating the)
    open('report.txt', 'w').close()

    # how many letters where the difference is 0 between the original and the split
    goodcount = 0
    missedletters = []
    
    # get a dataframe with the full list and a column for the first letter
    full_list_df = pd.read_excel(args.filename, header=1, sheet_name="Full list, A-Z") # get rid of first merged cell
    full_list_df['Letter'] = full_list_df['Theater'].apply(get_letter)

    # get the totals for each because we are curious
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        # dataframe from the full list with just the letter we're working with
        splitletterdf = full_list_df[full_list_df["Letter"]==letter]
        # dataframe from the original split in the spreadsheet -- with just the letter we're working with
        originalletterdf = pd.read_excel(args.filename, sheet_name=letter)

        # find lengths
        lensplit = len(splitletterdf)
        lenoriginal = len(originalletterdf)

        # generate counts of number of times a play title comes up
        splitcounts = splitletterdf.value_counts(subset="Play title", 
                                                normalize=False, 
                                                sort=True, 
                                                ascending=False)
        
        originalcounts = originalletterdf.value_counts(subset="Play title",
                                                       normalize= False,
                                                       sort=True,
                                                       ascending=False)

        with open('report.txt', 'a') as f:
            f.write("\n" + letter + "\n")
            f.write("Split Length:\n")
            slensplit = str(lensplit)
            f.write(slensplit + "\n")
            f.write("Original Length:\n")
            slenoriginal = str(lenoriginal)
            f.write(slenoriginal + "\n")
            f.write("Comparison:\n")
            dif = lenoriginal-lensplit
            sdif = str(dif)
            f.write(sdif + "\n")

            if dif == 0:
                goodcount += 1
            else:
                missedletters += letter
                try:
                    mergedcounts = pd.merge(splitcounts, originalcounts, how="outer", on=["Play title"]) # on=["Play title"]
                    mergedcounts["Difference"] = mergedcounts["count_x"].astype('int64', errors='ignore') - mergedcounts["count_y"].astype('int64', errors='ignore')
            
                    differentdf = mergedcounts[(mergedcounts["Difference"] != 0)]
                    sdifdf = differentdf.to_string(header=True, index=True)
                    f.write(sdifdf + "\n-----------------------\n")
                except ValueError as err:
                    print("Here's a problem!")
                    print(letter)

    with open('report.txt', 'a') as f:
        f.write("SUMMARY of Matching Letters: " + str(goodcount) + "/26")
        f.write(" Missing Letters: " + ", ".join(missedletters))

# pandas is a visualization tool too...

    