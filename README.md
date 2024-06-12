# aioexcel

A simple library to work with xlsx files in an asynchronous manner

Example: 
```python
from aioexcel import ExcelReader

reader = ExcelReader("example.xlsx")
print(await reader.read_cell("A", 3))
print(await reader.sheet_size())
```

This library was sponsored by our lord [timoniq](https://github.com/timoniq).