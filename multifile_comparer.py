# Object for comparing a list of Microsoft Excel Spreadsheets with extension .xlsx
# Derek Fujimoto
# May 2018

import openpyxl
import numpy as np
import os,glob

from comparer import comparer

# ========================================================================== #
class multifile_comparer(object):
    """
        Do a pairwise comparison of all files in a list, flag files which have 
        given similarities.
        
        Usage:
            construct: c = multifile_comparer(filelist)
                filelist: list of filenames, 
                            OR string of directory name to fetch all contents, 
                            OR wildcard string of filenme format.
            compare: c.compare()
            show results: c.print_table([filename])
                print to stdout if filename missing
        Derek Fujimoto
        May 2018 
    """

    # some colours
    colors={'HEADER':'\033[95m',
            'OKBLUE':'\033[94m',
            'OKGREEN':'\033[92m',
            'WARNING':'\033[93m',
            'FAIL':'\033[91m',
            'ENDC':'\033[0m',
            'BOLD':'\033[1m',
            'UNDERLINE':'\033[4m'}
            
    # thresholds for cell similarity (warning,fail)
    cell_sim_thresh = (0.5,0.8)

    # list of good extensions
    extensions = ('.xlsx','.xls')

    # ====================================================================== #
    def __init__(self,filelist):
        """
            filelist: list of filenames, OR string of directory to fetch all 
                      files, or wildcard string of file format. 
        """
        
        # save filelist
        if type(filelist) == str:
            self.set_filelist(filelist)
        else:
            self.filelist = filelist

        # build comparer objects
        nfiles = len(self.filelist)
        self.comparers = [comparer(self.filelist[i],self.filelist[j]) 
                            for i in range(nfiles-1) for j in range(i+1,nfiles)]

    # ====================================================================== #
    def set_filelist(self,string):
        """
            Get list of files based on directory structure. String is prototype 
            filename or directory name.
        """
        
        # check if string is directory: fetch all files there
        if os.path.isdir(string):
            if string[-1] != '/': string += '/'
            filelist = glob.glob(string+"*")
        
        # otherwise get files from wildcard
        else:
            filelist = glob.glob(string)
        
        # discard all files with bad extensions
        self.filelist = [f for f in filelist 
                           if os.path.splitext(f)[1] in self.extensions]  

    # ====================================================================== #
    def compare(self,options='meta,exact,string',do_print=False):
        """
            Run comparisons on the paired files
            
            Options: same as comparer.compare
        """
        
        for c in self.comparers:
            c.compare(options=options,do_print=do_print)
        
    # ====================================================================== #
    def print_table(self,filename=''):
        """
            Print a table of pairs, with results
            
            filename: if not '' then write table to file, else write to stdout.
        """
        
        # get column: file1 names
        file1 = [c.file1 + " " for c in self.comparers]
        file1_size = max(map(len,file1))
            
        # get column: file2 names
        file2 = [c.file2 + " " for c in self.comparers]
        file2_size = max(map(len,file2))    
    
        # get columns: keys
        keys_columns = {}
        for c in self.comparers:
            for k in c.results.keys():
                try: 
                    keys_columns[k].append(str(c.results[k]))
                except KeyError:
                    keys_columns[k] = [str(c.results[k])]
        keys_columns_size = {}
        for k in keys_columns.keys():
            keys_columns_size[k] = max(len(k),max(map(len,keys_columns[k])))+2
    
        # make print header
        s = "file1".ljust(file1_size) + "file2".ljust(file2_size)
        colkeys = list(keys_columns.keys())
        colkeys.sort()
        
        for k in colkeys:  
            # don't print ntotal or nsame
            if 'ntotal' in k or 'nsame' in k : continue  
            s += k.ljust(keys_columns_size[k])
            
        s += '\n'
        s += "-"*len(s)
        s += '\n'
        
        # make print columns
        for i in range(len(self.comparers)):
            s += file1[i].ljust(file1_size) + file2[i].ljust(file2_size)
            for k in colkeys:
                
                # don't print ntotal or nsame
                if 'ntotal' in k or 'nsame' in k : continue  
                    
                # get value 
                value = keys_columns[k][i]
                
                # set text color
                s1 = value.ljust(keys_columns_size[k])
                
                if filename == '':
                    if value == "True":
                        s1 = self.colors['FAIL']+s1+self.colors['ENDC']
                    
                    elif value == "False":
                        s1 = self.colors['OKGREEN']+s1+self.colors['ENDC']
                    
                    elif 'nexcess' in k: 
                        if float(value) == 0:
                            s1 = self.colors['WARNING']+s1+self.colors['ENDC']
                        else:
                            s1 = self.colors['OKGREEN']+s1+self.colors['ENDC']
                    
                    elif "sim" in k:
                        if float(value) > self.cell_sim_thresh[0]:
                            if float(value) > self.cell_sim_thresh[1]:
                                s1 = self.colors['FAIL']+s1+self.colors['ENDC']
                            else:
                                s1 = self.colors['WARNING']+s1+self.colors['ENDC']
                        else:
                            s1 = self.colors['OKGREEN']+s1+self.colors['ENDC']
                s += s1
            s += '\n'
            
        # write results
        if filename == '':
            print(s)
        else:
            with open(filename,'a+') as fid:
                fid.write(s)
