# python-spreadsheet

![Travis (.org)](https://img.shields.io/travis/Data-Innovation-Research-Institute/python-spreadsheet.svg)

A wrapper around the openpyxl Workbook that uses a typewriter style write-and-advance metaphor. Cell values are set without needing to explicitly state column/row coordinates such as A1 and D6.

## Install

```bash
pip install git+https://github.com/Data-Innovation-Research-Institute/python-spreadsheet.git
```

## Use

Import the ```Spreadsheet``` class:

```python
from spreadsheet import Spreadsheet
```

Create a spreadsheet:

```python
spreadsheet = Spreadsheet()
```

Select the name of a sheet to make it active with the ```select_sheet``` method. If the named sheet does not exist, a sheet with that name will be created.

```python
spreadsheet.select_sheet('Stats')
```

### Setting Single Cell Values

The current cell of a new sheet is A1. Use the ```set_value``` method to set the value of the current cell:

```python
spreadsheet.set_value(1)
```

The current cell will advance to the next column on the same row. After setting the value of A1, the current cell will advance to B1.

To set the value of a cell and advance to the first cell on the next line, use the ```advance_row``` parameter with a value of ```True```:

```python
spreadsheet.set_value(1, advance_row=True)
```

To advance by more than one line, use the ```advance_by_rows``` parameter with the number of rows to advance:


```python
spreadsheet.set_value(1, advance_by_rows=3)

```

### Setting Multiple Cell Values

To set multiple column values at once, use the ```set_values``` method with a list of values. Each value will be used to set the value of a column on the current row, the column advancing by one after each value is set. For example, if the current cell is C3, the following call to ```set_values``` will set the values of cells C3, D3, and E3 to 1, 2 and 3, respectively:

```python
spreadsheet.set_values([1, 2, 3])
```

The default behaviour of the ```set_values``` method is to automatically advance to the first cell of the next line after the final cell value in the list is set, as this is a common use case. To prevent this behaviour, use the ```advance_row``` parameter with a value of ```False```: 

```python
spreadsheet.set_values([1, 2, 3], advance_row=False)

```

A shortcut for setting the value of a single cell and advancing to the first cell of the next line is to use the ```set_values``` method with a single-value list:

```python
spreadsheet.set_values([1])

```

### Saving Spreadsheets

The ```save``` method is provided as a convenience that wraps the ```save``` method of the underlying openpyxl Workbook:

```python
spreadsheet.save("Stats.xlsx")

```

### The openpyxl Workbook

The underlying openpyxl Workbook is available using the ```workbook``` property:

```python
spreadsheet.workbook

```
