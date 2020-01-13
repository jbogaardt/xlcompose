import pandas as pd
import xlcompose as xlc

def test_simple_exhibit():
    df=pd.DataFrame({'Fruit': ['Apple', 'Pear'],
                     'Quantity': [1,2]})
    col = xlc.Column(xlc.DataFrame(df), xlc.CSpacer(), xlc.DataFrame(df))
    composite = xlc.Row(
        col, col,
        title=['This title spans both Column Objects']
    )
    x = xlc.Tabs(
       ('a_sheet', composite),
       ('another_sheet', composite)
    ).to_excel('workbook.xlsx')
