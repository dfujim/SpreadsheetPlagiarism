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

File types must be xlsx. 

## Manual

```
usage: compsheet [-h] [-d] [--explain] [-l LOGFILE] [-o OPTIONS] [-p] [-q]
                 [-s SAVEFILE]
                 PATH

Run a pairwise comparison of all spreadsheets on target PATH. Look for pairs
with common features indicative of plagiarism.

positional arguments:
  PATH                  evaluate spreadsheets found on PATH

optional arguments:
  -h, --help            show this help message and exit
  -d, --dry             Dry run, don't write to speadsheet
  --explain             Print calculation methodology of table values
  -l LOGFILE, --log LOGFILE
                        write print out table to text file
  -o OPTIONS, --options OPTIONS
                        comma-separated list of items to compare (default:
                        'meta,exact,string,geo')
  -p, --print           Print full summary of each comparison
  -q, --quiet           No print output to stdout
  -s SAVEFILE, --save SAVEFILE
                        write printout to xlsx file
```

Run `compsheet` from command line in terminal. 

Ensure that `compsheet` is on the `PATH` and that `comparer.py` and `multifile_comparer.py` are on the `PYTHONPATH`. 
