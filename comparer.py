# Object for comparing Microsoft Excel Spreadsheets with extension .xlsx
# Derek Fujimoto
# May 2018

import openpyxl
import numpy as np
import os,glob

# ========================================================================== #
class comparer(object):
    """
        Object for comparing Microsoft Excel Spreadsheets with extension .xlsx.
        
        Usage: 
        
            Construct object: c = comparer(file1,file2)
            Compare files: c.compare([options],[do_print])
                
                options: comma-seperated list of the following
                    meta: compare metadata
                    exact: compare cell values by coordinate
                do_print: if true, print results to stdout
        
        Saves results of comparison to dictionary "results", which is formatted 
        with pretty representation and dot operator access. 
        
        Derek Fujimoto
        May 2018 
    """

    # ====================================================================== #
    def __init__(self,file1,file2):
        """
            file1,file2: filename, with path, of spreadsheets to compare
        """
        
        self.file1 = file1
        self.file2 = file2
        self.book1 = openpyxl.load_workbook(file1)
        self.book2 = openpyxl.load_workbook(file2)
        
        # results
        self.results = result_dict()
        
    # ====================================================================== #
    def cmpr_exact_values(self,do_print=False):
        """
            Compare cell values for exact match. Not intelligent. 
            
            Output: 
                nsame: number of cells which are identical by coordinate. 
                ntotal: number of cells total which are in a shared range. 
                nexcess: number of non-empty cells which are not in a shared 
                        range. 
                        
            Sets to results: 
                nsame, ntotal, nexcess: as described above
                cell_similarity: nsame/ntotal
        """
        
        # track statistics
        nsame = 0
        ntotal = 0
        nexcess = 0
        
        # get sheet names 
        sheet_names1 = self.book1.sheetnames
        sheet_names2 = self.book2.sheetnames
        
        # iterate over sheets
        for sht1nm,sht2nm in zip(sheet_names1,sheet_names2):
            sht1 = self.book1[sht1nm]
            sht2 = self.book2[sht2nm]
            
            # iterate over cells
            sheet1_cells = [cell.value for row in sht1.rows for cell in row]
            sheet2_cells = [cell.value for row in sht2.rows for cell in row]

            # compare cells for which there is identical content
            size = min(len(sheet1_cells),len(sheet2_cells))
            cmpr = np.equal(sheet1_cells[:size],sheet2_cells[:size])
            ntotal += len(cmpr)
            nsame += np.sum(cmpr)
            
            # add content that is in excess from one sheet
            if len(sheet1_cells) > len(sheet2_cells):
                nexcess += np.sum(np.array(sheet1_cells[size:]) != None)
            else:
                nexcess += np.sum(np.array(sheet2_cells[size:]) != None)
            
        # check for excess sheets
        ndiff_sheets = abs(len(sheet_names1)-len(sheet_names2))
        if ndiff_sheets != 0:            
            
            # get exces sheets
            if len(sheet_names1) > len(sheet_names2):
                sheets = [self.book1[n] for n in sheet_names1[len(sheet_names2):]]
            else:
                sheets = [self.book2[n] for n in sheet_names2[len(sheet_names1):]]
            
            # count number of excess cells in excess sheets    
            for sht in sheets:
                cells = [cell.value for row in sht.rows for cell in row]
                nexcess += np.sum(np.array(cells) != None)
                
        # print results
        if do_print:
            print("Shared range cell content exact match: %d/%d (%.2f" % \
                        (nsame,ntotal,float(nsame)/ntotal*100) +\
                  "%) " + "with %d cells in excess." % nexcess)            
        
        # set to self
        self.results['nsame'] = nsame
        self.results['ntotal'] = ntotal
        self.results['nexcess'] = nexcess
        self.results['cell_similarity'] = np.around(float(nsame)/ntotal,4)
        
        return (nsame,ntotal,nexcess)
        
    # ====================================================================== #
    def cmpr_meta(self,do_print=False):
        """
            Compare meta data
            
            Output: 
                mod: boolean, are file last modified times the same?
                create: boolean, are file creation times the same?
            
            Sets to results: 
                same_modifiy_time, same_create_time: as described above
        """
        
        prop1 = self.book1.properties
        prop2 = self.book2.properties
        
        # compare mod time
        mod = prop1.modified == prop2.modified
        
        # compare create time
        create = prop1.created == prop2.created
        
        # print results
        if do_print:
            print('Sheet modification time is identical: %s' % str(mod))
            print('Sheet creation time is identical:     %s' % str(create))
        
        # set to self
        self.results['same_modify_time'] = mod
        self.results['same_create_time'] = create
        
        return (mod,create)

    # ====================================================================== #
    def compare(self,options='meta,exact',do_print=False):
        """
            Run comparisons on the two files
            
            Options: 
                meta: compare metadata
                exact: compare cell values by coordinate
        """
        
        # get options
        options = options.split(',')
        
        # run options
        if 'meta' in options:
            self.cmpr_meta(do_print=do_print)
        
        if 'exact' in options: 
            self.cmpr_exact_values(do_print=do_print)

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
    def compare(self,options='meta,exact',do_print=False):
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
            if k == 'ntotal' or k == 'nsame': continue  
            
            s += k.ljust(keys_columns_size[k])
        s += '\n'
        s += "-"*len(s)
        s += '\n'
        
        # make print columns
        for i in range(len(self.comparers)):
            s += file1[i].ljust(file1_size) + file2[i].ljust(file2_size)
            for k in colkeys:
                
                # don't print ntotal or nsame
                if k == 'ntotal' or k == 'nsame': continue
                    
                # get value 
                value = keys_columns[k][i]
                
                # set text color
                s1 = value.ljust(keys_columns_size[k])
                
                if filename == '':
                    if value == "True":
                        s1 = self.colors['FAIL']+s1+self.colors['ENDC']
                    
                    elif value == "False":
                        s1 = self.colors['OKGREEN']+s1+self.colors['ENDC']
                    
                    elif k == 'nexcess': 
                        if float(value) == 0:
                            s1 = self.colors['WARNING']+s1+self.colors['ENDC']
                        else:
                            s1 = self.colors['OKGREEN']+s1+self.colors['ENDC']
                    
                    elif k == 'cell_similarity':
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
                
# ========================================================================== #
class result_dict(dict):
    """
        Pretty formatting and nice retrieval of dictionary items.
    """

    # ====================================================================== #
    def __getattr__(self, name):
        """Allow element access via dot operator"""
        try:
            return self[name]
        except KeyError:
            raise AttributeError(name)

    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__
    
    # ====================================================================== #
    def __repr__(self):
        """Nice representation of results dictionary"""
        
        # get max length of keys
        keys = list(self.keys())
        if len(keys) == 0: return "No results found."
        
        max_key_len = max(map(len,keys))
        
        # get max length of key items
        items = map(str,[self[key] for key in keys])
        max_item_len = max(map(len,items))
        
        # sort keys
        keys.sort()
        
        # make a table
        s = self.__class__.__name__+': \n'
        for k in keys:
            s += "'" + k.ljust(max_key_len) + "': " + \
                 str(self[k]).ljust(max_item_len) + '\n'
        return s
        
    # ====================================================================== #
    def __dir__(self):
        return list(self.keys())
