#!/bin/bash
# Clone and install compsheet things
# Derek Fujimoto

# Clone files 
git clone https://github.com/dfujim/SpreadsheetPlagiarism.git

# Check for dependencies
pip install openpyxl --user
pip install numpy --user

# Set up alias: add to .bash_aliases
if [ -f $HOME/.bash_aliases ] 
then
    if [ `grep -F "alias compsheet" $HOME/.bash_aliases | wc -l` == 0 ] 
    then
        echo "Adding compsheet alias to .bash_aliases"
        echo 'alias compsheet=${PWD}/SpreadsheetPlagiarism/compsheet' >> $HOME/.bash_aliases
    else
        echo "compsheet alias found in .bash_aliases. Doing nothing."
    fi
else
    if [ `grep -F "alias compsheet" $HOME/.bashrc | wc -l` == 0 ] 
    then
        echo "Adding compsheet alias to .bashrc"
        echo 'alias compsheet=${PWD}/SpreadsheetPlagiarism/compsheet' >> $HOME/.bashrc
    else
        echo "compsheet alias found in .bashrc. Doing nothing."
    fi
fi

# final message
echo "Run 'source $HOME/.bashrc' to finish installation."
