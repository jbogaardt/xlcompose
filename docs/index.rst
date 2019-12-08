xlcompose
=========
A declarative API for composing spreadsheets from python that is built on
`xlsxwriter` and `pandas` and inspired by `bokeh`.

Why use xlcompose?
----------------------
`xlcompose` provides a sweet spot between pandas `to_excel` and the `xlsxwriter`
API.  If you've ever needed to export multiple dataframes to a spreadsheet
or apply custom formatting, then you know pandas isn't up to the task.

On the other hand, `xlsxwriter` provides a tremendous amount of customization,
but it's imperative style can often lead to very verbose code.

With `xlcompose`, Several components are provided
through a set of classes to allow for highly composible spreadsheets.

There are several classes that allow for the export of data from Python:
  * `Series` - A class with formatting options for pandas Series
  * `DataFrame` - A class with formatting options for pandas DataFrames
  * `Image` - A class for exporting images
  * `Title` - A convenience class around a Series for titling


Layout components include:
  * `Row` - A class for laying out other components horizontally
  * `Column` - A class for laying our other components vertically
  * `CSpacer` - A class for adding spacing between components in a Column
  * `RSpacer` - A class for adding spacing between components in a Row
  * `Sheet` - A class for specifying sheet options
  * `Tabs` - A class for laying out other components across sheets

.. toctree::
   :maxdepth: 2

   modules
   examples




Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
