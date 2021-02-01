# excel-print-fmt


## What does this do?
* formats and prints excel spread sheet cell data by row
* allows user to exclude certain columns from output

## Usage
```python3 main.py filename
python3 main.py filename --exclude 1 2 4
```
optional argument: -e/--exclude, column numbers to exclude from formatted print


### Example on test sheet in /test
```python3 main.py test.xlsx -e 1 2```

first 2 rows out output: 
```
excluding columns: ['last date modified', 'cost']

Vendor name: Vendor 1
pass or fail: fail
Requires login: yes
is imporant: yes
uses financial: yes
Notes: Failing module 3
####################
Vendor name: Vendor 2
pass or fail: pass
Requires login: no
is imporant: no
uses financial: yes
Notes: send this to so and so before completion
```
