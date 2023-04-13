import pandas as pd
import itertools
import re
import csv

class Plugin:
    def __init__( self, name):
        self.name = name
        self.data = None
    def process( self):
        #procedss the data
        pass
    def load ( self ):
        # load data
        pass
    def get_sublist(self, list_of_strings, start_str, end_str, start_offset, end_offset):
    # initialize an empty sublist
        sublist = []
        # loop through the list of strings
        for s in list_of_strings:
            # if the string contains 'Case Number', start adding it to the sublist
            if start_str in s:
                sublist.append(s)
            # if the sublist is not empty and the string is not '\t\tBack to Top\t', keep adding it to the sublist
            elif sublist and s != end_str:
                sublist.append(s)
            # if the string is '\t\tBack to Top\t', stop adding to the sublist and break the loop
            elif re.search(end_str, s ):
                break
        # return the sublist without the last two elements
        return sublist[start_offset:end_offset]
    def create_df_from_list(self, list_of_strings, row_size):
    # create an iterator that yields sublists of 7 elements from the list
        iterator = (list(itertools.islice(list_of_strings,i ,i+row_size)) for i in range(0, len(list_of_strings), row_size))
    
        # convert the iterator to a list of lists
        data = list(iterator)
        #print(data)
        #for i in data:
        #    print(len(i))
        
        # create a pandas dataframe from the data
        df1 = pd.DataFrame(data[1:], columns=data[0])
        # return the dataf
        return df1
        


class EWS( Plugin):
    def __init__( self, name):
        super().__init__( name)
    def get_sublist(self, list_of_strings):
        return super().get_sublist(list_of_strings, 'Case Number', "\t\tBack To Top\t", 0, -2)
    def create_df_from_list(self, list_of_strings):
        # create an iterator that yields sublists of 7 elements from the list
        list_of_strings.remove("Case Criticality")
        return super().create_df_from_list( list_of_strings, 7)
        
    def iterate_strings(self, lst):
        # merge alerts together into one cell.
            new_lst = []
            i = 0
            while i < len(lst):
                if lst[i].find("(") > 0:
                    if lst[i].find("High") > 0 or lst[i].find("Low") > 0 or lst[i].find("Medium"):
                        temp = lst[i]
                        i += 1
                        temp2 = lst[i]
                        while i < len(lst) and lst[i].find("(") > 0:
                            temp2 = lst[i]
                            if lst[i].find("High") > 0 or lst[i].find("Low") > 0 or lst[i].find("Medium"):
                                temp += " " + lst[i]
                                i += 1
                            else:
                                break
                        new_lst.append(temp)
                    else:
                        new_lst.append(lst[i])
                        i += 1
                else:
                    new_lst.append(lst[i])
                    i += 1
            return new_lst
    def load( self, data):
        self.data = data

    def process( self):
        list_body = self.data.split("\r\n")
        print(list_body)
        list_body = self.iterate_strings(list_body)
        print(list_body)
        df1 =  self.create_df_from_list( self.get_sublist( list_body))
        # recreate this as a column since it is just formatting which we have trouble parsing out.
        df1['Case Critically'] = df1['Alert Summary'].apply(lambda x: 'High' if x.find('High') != -1 else ('Medium') if x.find('Medium') != -1 else 'Low')
        df1.to_csv("ews.csv", index=False, quoting=csv.QUOTE_NONNUMERIC)
