# Overview
 **kwadro** is a simple Python ORM (object-relational mapper) for manipulating data in Excel files (xlsx) using object-oriented abstractions. It's built on top of the openpyxl package.


# Usage examples

### Import
```python
from kwadro import Board, Table, Column # or import *
```

### Load your file
```python
board = Board("my-file.xlsx")
```

### Or start a new file
```python
board = Board()
```

### Define a table-class for every sheet you want to work with
You'll need to derive your class from `Table`, like the example below:
```python
class Employees(Table):
    __title__ = "Employees"
    name = Column("A") # Or Column(1) 
    birth = Column("B")
    phone = Column("c") # It's case insensitive
    address = Column(4) # You can use numbers instead
    country = Column("E")
```

### If the sheet doesn't exist yet, you can create it on the board
```python
board.create_sheet(Employees)
```
Optionally you can choose the sheet index: `create_sheet(Employees, index=3)`.

### Or force a new sheet
```python
board.create_sheet(Employees, force_new=True)
```
>[!warning] Warning
> If you use `force_new=True`, when you save the file, the pre-existing sheet will be permanently deleted and a new one will be created.

### Add a new record in the first empty row
```python
import datetime

employee = Employees(
    name="John Doe",
    birth=datetime.date(1987, 3, 12),
    phone=7654321,
    address="80 Bla St, Vancouver",
    country="Canada"
)

board.add(employee)
```

### If you want to get the row number
```python
print(employee.get_row()) # Outputs: 1
```

### Find the first record that matches your filters
```python
employee = board.find(Employees, name="John Doe", address="80 Bla St, Vancouver")

print(employee.phone) # Outputs: 7654321
```

### Get the board-instance through the record
```python
board = employee.get_board()
```

### Change value in memory
```python
employee.name = "Tom Smith"
```

### Find all records that match your filters
```python
result = board.find_all(Employees, country="Canada")
```

### Retrieve all records
```python
result = board.find_all(Employees)
```

### Save changes
```python
board.save("new-or-same-file.xlsx")
```


# Accessing useful internal objects

### Access openpyxl-worbook instance
```python
workbook = board._workbook
```

### Acess openpyxl-worksheet instance
```python
worksheet = employee._worksheet
```