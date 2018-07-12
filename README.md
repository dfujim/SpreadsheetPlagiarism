# Compare Microsoft Spreadsheet Files

Look for instances of plagiarism. 

* Compare file meta data
    * File creation time stamp
    * File modification time stamp
* Cell values
    * Non-formula strings
    * Cell-by-cell values
* Cell layout
    * Check locations of filled/unfilled cells

Spreadsheet files must be of type `.xlsx`

# Installation

## Automatic Setup (for unix beginners)

The setup script is a little inelegant, but it should do the job. If you are not logged in directly on a machine in Henn203, then this tutorial will walk you through how to send the file to the phas server and run the script remotely. If you are on a linux machine in Henn203, then skip steps 3 and 4. The tutorial will also show you how to run the script from linux or mac machines. 

1. Download the installation script

   Right-click on [this link](https://raw.githubusercontent.com/dfujim/SpreadsheetPlagiarism/master/install_compsheet.bash) and click "Save Link As". Save it somehwere on your machine. I'll assume it's in your Downloads folder. 

2. Open a terminal

3. Copy the files to the phas server. If your phas email address is smith@phas.ubc.ca then your username is `smith`. Type the following with your username, and press enter. 

   ```bash
   scp Downloads/install_compsheet.bash username@ssh.phas.ubc.ca:~
   ```
   
   This may prompt you with a long message about whether you should add an RSA key. You should. Type `yes` and press enter. You will be prompted for a password. This is the password you made when you first created your account. The keys you press when typing your password will not produce any visible effect. Don't worry, this is expected. Type the password and press enter. 

4. We will now remotely access our phas account. This is not needed if you are logged in directly on a machine in Henn203. To remotely run commands on the phas server type the following and press enter. 

   ```bash
   ssh username@ssh.phas.ubc.ca
   ```
   Now when you enter commands into the terminal, they are run on the phas server. 

5. Here are some basic commands you can use to move the installation script into the directory you want to install into: 

   ```bash
      ls                # "list":                     print all contents of current directory.
      cd path           # "change directory":         move into the location specified by "path"
      mv item1 item2    # "move":                     rename item1 to item2, or if a path is included, move item1 to the position item2. 
      mkdir dirname     # "make directory":           make a new directory with name "dirname"
      pwd               # "print working directory":  show where you are in the file system
   ```
   
   Check out [this tutorial](https://maker.pro/linux/tutorial/basic-linux-commands-for-beginners) for more commands, but for now here are some examples of the above: 
   
   ```bash
      # make a new directory
      mkdir test_dir
   
      # move into the new directory 
      cd test_dir
   
      # move up one level
      cd ..
      
      # rename directory
      mv test_dir test2
      
      # make a new directory inside the renamed one
      mkdir test2/test_sub
      
      # move right into that sub directory
      cd test2/test_sub
      
      # go up two levels
      cd ../..
      
      # let's say we have a file named file.txt inside the test2 directory.
      # we can get the file and move it to where we are (note the period)
      mv test2/file.txt .
      
      # we can then put it into the sub folder we made
      mv file.txt test2/test_sub
   ```

6. Run the installation script. Type the following and press enter: 

   ```bash
   bash install_compsheet.bash
   ```
   
   We then have to update the command list. Type the following and press enter:
   
   ```bash
   source ~/.bashrc
   ```
   
You can now use `compsheet` as a command from anywhere in the directory system. See the examples below on how to use the program. 
   

## Manual Setup (for users familiar with unix)

In truth, compsheet is just a set of python objects so there's no real installation necessary. There are a few things you can do to make your life easier though. 

1. Make sure you have the dependencies. Run the following: 

```bash
pip install numpy --user
pip install openpyxl --user
```

2. Make sure the python executable location is correct. Open the `compsheet` file with your favourite text editor and modify the first line to be either `#!/opt/anaconda3/bin/python3` for phas servers or `#!/usr/bin/python3` for general usage.

3. Make compsheet easily executable from anywhere on your system. Open `~/.bashrc` in an text editor and add an alias for compsheet: append the line `alias compsheet='/path_to_file/compsheet'` to the end of the file. 

4. Update your session with the new command. Run `source ~/.bashrc`.

You can now use `compsheet` as a command from anywhere in the directory system. See the examples below on how to use the program. 

# Manual

To use, run `compsheet` from command line in terminal, with the appropriate switches and inputs. 

## Some basic examples:

```bash
compsheet -h            # show help message
compsheet               # compare all files in current directory
compsheet ./dirname     # compare all files in directory 'dirname'
compsheet -d ./dirname  # do a dry run: write no files. 
compsheet --explain     # print description of table headers
```

## Help Message: 

```text
usage: compsheet [-h] [-d] [--explain] [-l LOGFILE] [-o OPTIONS] [-p] [-q] [-s SAVEFILE] [PATH]

Run a pairwise comparison of all spreadsheets on target PATH. Look for pairs with common features indicative of plagiarism.

positional arguments:
  PATH                           Evaluate spreadsheets found on PATH. Default is the current working directory.

optional arguments:
  -h, --help                     Show this help message and exit
  
  -d, --dry                      Dry run, don't write to speadsheet
  
  --explain                      Print calculation methodology of table values
  
  -l LOGFILE, --log LOGFILE      Write print out table to text file
  
  -o OPTIONS, --options OPTIONS  Comma-separated list of items to compare (default:'meta,exact,string,geo')
  
  -p, --print                    Print full summary of each comparison
  
  -q, --quiet                    No print output to stdout
  
  -s SAVEFILE, --save SAVEFILE   Write printout to xlsx file SAVEFILE
```
