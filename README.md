## Quadro module
 *Quadro* is rudimentary ORM for xlsx files, woking with openpyxl package

## First of all, install openpyxl, once this layer is written on the top of it
pip install openpyxl

## Example of usage:

```
from quadro import BaseTable, Column, Board
```

## Define a table, i.e. a sheet-class
```
class Clients(BaseTable):
    __title__ = 'Clients'
    name = Column(1)  
    phone = Column(2)
    address = Column('C') ## or Column(3)
    country = Column(4)
```


## Define your file
```
board = Board('my-file.xlsx') ## new file or existing one
```

## Create the sheet, if not created
```
board.create_sheet(Clients)
```

## Or force a new one
```
board.create_sheet(Clients, force=True)
```

## Add an entry at first empty row
```
entry = Clients(
    name='Joshua King', phone=123,
    address='80 Bla St, Birmingham',
    country='Australia')

board.add(entry)
```

## If one wants to get the row
```
print(entry._row) # Outputs: 1
```

## Find an entry
```
entry = board.find(
    Clients, name='Joshua King',
    address='80 Bla St, Birmingham')

print(entry.phone) ## Outputs: 123
```

## find_all() yields matches
```
entries = board.find_all(Clients, country='Australia')
```

## If one wants to access openpyxl-Worksheet object
```
worksheet = entry._worksheet
```

## Saving
```
board.save('new-or-same-file.xlsx')
```