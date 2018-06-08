# Compare Microsoft Spreadsheet Files

Look for instances of plagiarism. 

## Strategies

* Compare file meta data
    * File creation time stamp
    * File modification time stamp
* Cell values
    * Non-formula strings
    * Cell-by-cell values
* Cell layout
    * Check locations of filled/unfilled cells

## Restrictions

File types must be xlsx or xls. 

## Manual

```
usage: compsheet [-h] [-l LOGFILE] [-d PATH] [-o OPTIONS] [-p] [--explain]

optional arguments:
  -h, --help            show this help message and exit
  -l LOGFILE, --log LOGFILE
                        write printout to file
  -d PATH, --dir PATH   evaluate spreadsheets found on path
  -o OPTIONS, --options OPTIONS
                        comma-separated list of items to compare (default:
                        'meta,exact,string,geo')
  -p, --print           Print full summary of each comparison
  --explain             Print calculation methodology of table values
```

Run `compsheet` from command line in terminal. 

Ensure that `compsheet` is on the `PATH` and that `comparer.py` and `multifile_comparer.py` are on the `PYTHONPATH`. 
