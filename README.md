# aioexcel

A simple library to work with xlsx files in an asynchronous manner

Example: 
```python
from aioexcel import ExcelReader

reader = ExcelReader("example.xlsx")
print(await reader.read_cell("A", 3))
print(await reader.sheet_size())
```

## Installation

With poetry: `poetry add git+https://github.com/prostomarkeloff/aioexcel.git`\
With pip: `pip install git+https://github.com/prostomarkeloff/aioexcel.git@master`


This library was sponsored by our lord [timoniq](https://github.com/timoniq).