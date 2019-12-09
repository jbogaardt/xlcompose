[![Documentation Status](https://readthedocs.org/projects/xlcompose/badge/?version=latest)](https://xlcompose.readthedocs.io/en/latest/?badge=latest)

# xlcompose
A declarative API for composing spreadsheets from python that is built on
`xlsxwriter` and `pandas` and inspired by `bokeh`.

### Why use xlcompose?
`xlcompose` provides a sweet spot between pandas `to_excel` and the `xlsxwriter`
API.  If you've ever needed to export multiple dataframes to a spreadsheet
or apply custom formatting, then you know pandas isn't up to the task.

On the other hand, `xlsxwriter` provides a tremendous amount of customization,
but it's imperative style can often lead to very verbose code.

With `xlcompose`, Several components are provided
through a set of classes to allow for highly composible spreadsheets.

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
