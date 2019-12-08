.. _examples:

.. currentmodule:: xlcompose


Writing to Excel
================

Example output of Exhibits
--------------------------
The following example highlights the general look of the outputs when using
xlcompose to output to Excel.

.. image:: ../_static/images/exhibits.PNG

Dataframe
---------

xlcompose comes with a :class:`DataFrame` class that has a supercharged version of
`pd.DataFrame.to_excel`.  This version allows for the export of any number of
DataFrames in any layout desired with formats of your choosing.  Simply wrap a
pandas dataframe or a 2D tringle representation in the Dataframe to get access
to this supercharged version.

**Example:**
   >>> import xlcompose as cl
   >>> raa = cl.load_dataset('raa')
   >>> xlc.DataFrame(raa).to_excel('workbook.xlsx')

This class is used exclusively for exporting to Excel and the normal
``pd.DataFrame`` should be used for any other purpose.  There are commonalities
between ``xlc.DataFrame.to_excel()`` and ``pd.DataFrame.to_excel()``.  For example,
both have arguments for ``header``, ``index``, ``index_label`` that behave
identically, however the argument placement happens at object initialization
for the :class:`DataFrame`:

**Example:**
   >>> # xlcompose
   >>> xlc.DataFrame(raa, header=False, index=True, index_label='Origin').to_excel('workbook.xlsx')
   >>> # vs
   >>> # pandas
   >>> raa.to_frame().to_excel('workbook.xlsx', header=False, index=True, index_label='Origin')

By placing the arguments at object initialization allows for the construction
of composite objects as we will see later. ``pd.DataFrame.to_excel`` provides
additional ``startrow`` and ``startcol`` arguments in the case where you want
the dataframe exported anywhere other than cell A1.  `xlc.DataFrame.to_excel`
replaces these with a single `margin` argument that behaves similarly to a
Cascading Stylesheet margin.  The ``margin`` option can be expressed as an
integer e.g. ``margin=1`` which will place empty cells around the dataframe.
The ``margin`` can also be expressed as a tuple. ``margin=(1,0)`` puts a top
and bottom set of cells, but not left and right. ``margin=(2,0,0,1)`` puts two
empty rows above the dataframe and one to the left. Example comparison to
``pd.DataFrame.to_excel``

**Example:**
   >>> # xlcompose
   >>> xlc.DataFrame(raa, margin=(3,0,0,1)).to_excel('workbook.xlsx')
   >>> # vs
   >>> # pandas
   >>> raa.to_frame().to_excel('workbook.xlsx', startrow=3, startcol=1)

**Formatting**

Formatting output is key to having a polished looking spreadsheet, but
unfortunately pandas does not help much.  xlcompose, uses xlsxwriter to apply
formats to the data cells in a dataframe.  Formats are expressed as
dictionaries.  You can specify a single set of formats for the entire dataframe:

**Example:**
   >>> formats={'num_format':'#,#', 'font_color':'red'}
   >>> xlc.DataFrame(raa, formats=formats).to_excel('workbook.xlsx')

Alternatively, you can specify formats for each column individually using a
nested dictionary.

**Example:**
   >>> formats={'Ultimate':{'num_format':'#,#', 'font_color':'red'},
   ...          'Latest':  {'num_format':'#,0.00', 'bold':True}}

Formatting options exist for the `index` and `header`.  Simply pass the desired
formats through using `index_formats` and `header_formats`.

**Example:**
   >>> formats={'italic':True, 'font_color':'red'}
   >>> xlc.DataFrame(raa, index_formats=formats).to_excel('workbook.xlsx')

   .. note::
      xlcompose already has default formats set up.  As you apply your own
      formats, the defaults will be applied first followed by your own.

For more information on available formats refer to
https://xlsxwriter.readthedocs.io/format.html

**Other Features**
`xlc.DataFrame` also allows for adding a title and column numbering. Titles are
expressed as a list:

**Example:**
   >>> title=['Sample Accident Year Triangle',
   ...        'Sourced from Mack',
   ...        'Evaluated as of Dec-1990']
   >>> xlc.DataFrame(raa, title=title).to_excel('workbook.xlsx')

Many actuarial exhibits include column numbering for easier reference.
This can be turned on using the `col_nums=True`.  As with everything else,
formats are adjustable through the `title_formats` argument.

**Example:**
   >>> title_formats = [{'font_color': 'red'},
   ...        {'font_color': 'green'},
   ...        {'font_color': 'blue'}]
   >>> xlc.DataFrame(raa, title=title, title_formats=title_formats).to_excel('workbook.xlsx')


Laying out composite objects
============================

While the addition of formats, titles and column numbering provide a little more
flexibility that can be obtained from ``pd.DataFrame.to_excel``, xlcompose
provides a lot more flexibility with its layout objects.  There are three
layout objects `Tabs`, `Row`, and `Column`.

.. note::
   The layout API borrows from the bokeh/holoviz API and should be familiar to
   the practitioner who uses those for visualization.

Rows and Columns
----------------

:class:`Column` takes multiple objects and displays them vertically.

**Example:**
   >>> col = xlc.Column(
   ...     xlc.DataFrame(raa, margin=(0,0,1,0)),
   ...     xlc.DataFrame(raa.link_ratio, formats={'italic': True})
   ... )
   >>> col.to_excel('workbook.xlsx')

:class:`Row` takes multiple objects and displays them horizontally.

**Example:**
   >>> xlc.Row(
   ...     xlc.DataFrame(raa, margin=(0,0,1,0)),
   ...     xlc.DataFrame(raa.link_ratio)
   ... ).to_excel('workbook.xlsx')

You can also nest ``Row`` and ``Column``  within rows and columns.  Nesting can
be a deep as you want allowing for a highly customized layout.
**Example:**
   >>> xlc.Row(col, col).to_excel('workbook.xlsx')

``Row`` and ``Column`` optionally take `title`, ``title_formats`` and `margin`
keywords that function the same as those in ``xlc.DataFrame``.

**Example:**
   >>> composite = xlc.Row(
   ...     col, col,
   ...     title=['This title spans both Column Objects'],
   ...     title_formats=[{'underline': True}]
   ...     margin=(0,1,0,0)
   ... )
   >>> composite.to_excel('workbook.xlsx')

Tabs
----

:class:`Tabs` are the sheet representation of these objects.  Tabs are different
from ``Row`` and ``Column`` in that each object passed to ``Tabs`` must be
expressed as a 2-tuple corresponding to ``('sheet_name', object)``.

**Example:**
   >>> xlc.Tabs(
   ...    ('a_sheet', composite),
   ...    ('another_sheet', composite)
   ... ).to_excel('workbook.xlsx')

Modifying defaults for all objects
----------------------------------
You may choose to override all defaults.  For example, by default, the font is
set to 'Calibri'.  `to_excel()` takes an additional parameter `default_formats`
to will apply to all nested objects you intend to export.

**Example:**
   >>> xlc.Tabs(
   ...    ('a_sheet', composite),
   ...    ('another_sheet', composite)
   ... ).to_excel('workbook.xlsx', default_formats={'font_name': 'Arial'})

If any nested object has a default override, the override will be honored over
this default.
