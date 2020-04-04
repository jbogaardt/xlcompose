[![Documentation Status](https://readthedocs.org/projects/xlcompose/badge/?version=latest)](https://xlcompose.readthedocs.io/en/latest/?badge=latest)

# xlcompose
A declarative API for composing spreadsheets from python that is built on
`xlsxwriter`, `pandas` and inspired by `bokeh` and `flask`.

### Why use xlcompose?
`xlcompose` provides a sweet spot between pandas `to_excel` and the `xlsxwriter`
API.  If you've ever needed to export multiple dataframes to a spreadsheet
or apply custom formatting, then you know pandas isn't up to the task.

On the other hand, `xlsxwriter` provides a tremendous amount of customization,
but it's imperative style can often lead to very verbose code.

With `xlcompose`, we take a compositional approach to spreadsheet design.

### Features
#### A rich set of container classes to manage layout of data
With `Row`, `Column`, `Tabs`, and `Sheet` containers, we can conceptualize the
placement of our data in an Excel spreadsheet.  Containers can be nested within
other containers allowing for highly customized layout within Excel.  Did I lose you?
Perhaps this conceptual image of how a layout might look will help.
![alt](https://raw.githubusercontent.com/jbogaardt/xlcompose/master/docs/_static/images/layout.PNG)


#### Build your own custom template library as YAML files
Borrowing inspiration from HTML templates of static web design, why not create
detailed Excel files from YAML templates? Like web frameworks, `xlcompose` templates
are fully compatible with the `jinja2` templating language allowing for context-aware
rendering of Excel files with its `load_yaml` function.  Simply pass a template and
data to `load_yaml` to create `xlcompose` objects.
![alt](https://raw.githubusercontent.com/jbogaardt/xlcompose/master/docs/_static/images/templating.PNG)

#### Formats, formats, everywhere!
Ultimately, `xlcompose` is just a wrapper around `xlsxwriter` which has near 100%
coverage of Excel formatting. Whether you want to change the color of a cell, set page breaks,
add headers and footers, `xlsxwriter` has got you covered.  We strive to provide
full access to `xlsxwriter` functionality, with just a more convenient API.  If you
see something missing, let me know!


## Documentation
Please visit the [Documentation](https://xlcompose.readthedocs.io/en/latest/) page for examples, how-tos, and source
code documentation.

## Installation
To install using pip:
`pip install xlcompose`

Alternatively, install directly from github:
`pip install git+https://github.com/jbogaardt/xlcompose/`

Note: This package requires Python 3.5 and later, xlsxwriter 1.1.8. and later,
pandas 0.23.0 and later.
