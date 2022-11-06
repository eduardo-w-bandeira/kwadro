# QUADRO MODULE
 **Quadro** is a rudimentary ORM for xlsx files, that works with openpyxl package.

## First of all, install openpyxl
```console
pip install openpyxl
```

## EXAMPLES OF USAGE
### Import
```python
from quadro import BaseTable, Column, Board
```

### Define a table, i.e. a sheet-class
```python
class Clients(BaseTable):
    __title__ = 'Clients'
    name = Column(1)  
    phone = Column(2)
    address = Column('C') # or Column(3)
    country = Column(4)
```

### Define a file to use as a board
```python
board = Board('my-file.xlsx') # new file or existing one
```

### Create the sheet, if not created
```python
board.create_sheet(Clients)
```

### Or force a new one
```python
board.create_sheet(Clients, force_new=True)
```

### Add an entry at first empty row
```python
entry = Clients(
    name='Joshua King',
    phone=123,
    address='80 Bla St, Canberra',
    country='Australia')

board.add(entry)
```

### If one wants to get the row
```python
print(entry._row) # Outputs: 1
```

### Find the first entry that matches given args
```python
entry = board.find(
    Clients, name='Joshua King', address='80 Bla St, Canberra')

print(entry.phone) # Outputs: 123
```

### find_all( ) yields matches
```python
entries = board.find_all(Clients, country='Australia')
```

### If one wants to access openpyxl.Worksheet object
```python
worksheet = entry._worksheet
```

### Saving
```python
board.save('new-or-same-file.xlsx')
```