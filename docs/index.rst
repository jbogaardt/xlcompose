xlcompose
=========
A declarative API for composing spreadsheets from python that is built on
`xlsxwriter` and `pandas` and inspired by `bokeh`.

Why use xlcompose?
----------------------
**xlcompose** provides a sweet spot between **pandas** ``to_excel`` and the **xlsxwriter**
API.  If you've ever needed to export multiple dataframes to a spreadsheet
or apply custom formatting, then you know pandas isn't up to the task.

On the other hand, **xlsxwriter** provides a tremendous amount of customization,
but it's imperative style can often lead to very verbose code.

With **xlcompose**, several components are provided
through a set of classes to allow for highly composable spreadsheets.

There are several classes that allow for the export of data from Python:

  * :class:`~xlcompose.core.Series` - A class with formatting options for pandas Series
  * :class:`~xlcompose.core.DataFrame` - A class with formatting options for pandas DataFrames
  * :class:`~xlcompose.core.Image` - A class for exporting images
  * :class:`~xlcompose.core.Title` - A convenience class around a Series for titling


Layout components include:

  * :class:`~xlcompose.core.Row` - A class for laying out other components horizontally
  * :class:`~xlcompose.core.Column` - A class for laying our other components vertically
  * :class:`~xlcompose.core.CSpacer` - A class for adding spacing between components in a Column
  * :class:`~xlcompose.core.RSpacer` - A class for adding spacing between components in a Row
  * :class:`~xlcompose.core.Sheet` - A class for specifying sheet options
  * :class:`~xlcompose.core.Tabs` - A class for laying out other components across sheets

.. toctree::
   :maxdepth: 2

   examples
   modules




Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
