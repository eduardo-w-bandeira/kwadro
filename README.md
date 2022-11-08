# QUADRO MODULE
 **Quadro** is a rudimentary ORM for xlsx files, that works with openpyxl package.

## First of all, install openpyxl
```console
pip install openpyxl
```

## EXAMPLES OF USE
### Importing
```python
from quadro import BaseTable, Column, Board
```

### Load your file
```python
board = Board("my-file.xlsx")
```

### Or start a new file
```python
board = Board()
```

### Create a sheet, if not already there
#### Define a table, i.e., a sheet-class
```python
class Clients(BaseTable):
    __title__ = "clients"
    name = Column("A") # Or Column(1) 
    phone = Column("b") # It's case insensitive
    address = Column(3) # You can use numbers instead
    country = Column("D")
```
#### Then create it on the board
```python
board.create_sheet(Clients)
```

#### Or force a new sheet
*Warning*: The actual sheet will be deleted and a new one will be created.
```python
board.create_sheet(Clients, force_new=True)
```

### Add an entry in the first empty row
```python
client = Clients(
    name="John Doe",
    phone=7654321,
    address="80 Bla St, Canberra",
    country="Australia")

board.add(client)
```

### If you want to get the row
```python
print(client._row) # Outputs: 1
```

### Find the first entry that matches given args
```python
client = board.find(
    Clients, name="John Doe", address="80 Bla St, Canberra")

print(client.phone) # Outputs: 7654321
```

### Find all matches
```python
result = board.find_all(Clients, country="Australia")
```
If you want to retrive all rows, just use `find_all(Clients)` without other args.


### If you want to access openpyxl.Worksheet object
```python
worksheet = client._worksheet
```

### Saving
```python
board.save("new-or-same-file.xlsx")
```