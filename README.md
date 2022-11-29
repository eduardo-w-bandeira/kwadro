# KWADRO MODULE
 **Kwadro** is a elementary Python ORM (object relational mapping) for Excel files (xlsx). It runs on top of openpyxl package.

## Examples of use
### Import
```python
from kwadro import Board, BaseTable, Column
```

### Load your file
```python
board = Board("my-file.xlsx")
```

### Or start a new file
```python
board = Board()
```

### Define a table for every sheet you want to work with
You'll need to derive your class from `BaseTable`, like the example below:
```python
class Employees(BaseTable):
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
*Warning*: If you use `force_new=True`, when you save the file, the pre-existing sheet will be permanently deleted and a new one will be created.

### Add a new entry in the first empty row
```python
import datetime

employee = Employees(
    name="John Doe",
    birth=datetime.date(1987, 3, 12),
    phone=7654321,
    address="80 Bla St, Canberra",
    country="Australia")

board.add(employee)
```

### If you want to get the row number
```python
print(employee._row) # Outputs: 1
```

### Find the first row that matches your filters
```python
employee = board.find(Employees, name="John Doe", address="80 Bla St, Canberra")

print(employee.phone) # Outputs: 7654321
```

### Find all rows that match your filters
```python
result = board.find_all(Employees, country="Australia")
```

### Retrieve all rows
```python
result = board.find_all(Employees)
```

### Save
```python
board.save("new-or-same-file.xlsx")
```

## Access openpyxl objects

### To access openpyxl-worbook instance, use
```python
workbook = board._workbook
```

### As for openpyxl-worksheet instance
```python
worksheet = employee._worksheet
```
