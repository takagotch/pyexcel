### pyexcel
---
https://github.com/pyexcel/pyexcel

```
pip install pyexcel
python setup.py install

import pyexel as p
css_insight2 = p.Sheet()
css_insight2.name = "Worldwide Mobile Phone Shipments (Billions), 2017-2021"
css_insight2.ndjson = """
  {"year": ["2017", "2018", "2020", "2021"]}
""".strip()
css insight2

import pyexcel
a_list_of_dictionaries = []
pyexcel.save_as(records=a_list_of_dictionaries, dest_file_name="your_file.xls")

import pyexcel as p
records = p.iget_records(file_name="your_file.xls")
for record in records:
p.free_resources()

def increase_everyones_age(generator):
  for row in generator:
    row['Age'] += 1
    yield row
def duplicate_each_record(generator):
  for row in generator:
    yeild row
    yield row
records = p.iget_records(file_name="your_file.xls")
io=p.isave_as(records=duplicate_each_record(increase_everyones_age(records)),
  dest_file_type='csv', dest_lineterminator='\n')
print(io.get value())
```

```py
import pyexel as p
sheet = p.get_sheet(file_name='google.ods')
sheet.save_as('me.sortable.html', display_length=10)
from IPython.display import IFrame
IFrame("me.sortable.html", width=600, height=500)

```

```
```


