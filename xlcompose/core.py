import pandas as pd
import numpy as np
import copy
import json


class _Workbook:
    """
    Excel Workbook level configurations.  This is not part of the end_user API
    """
    max_column_width = 30
    max_portrait_width = 120
    footer = '&CPage &P of &N\n&A'

    def __init__(self, workbook_path, exhibits, default_formats):
        self.formats = {}
        self.writer = pd.ExcelWriter(workbook_path)
        self.exhibits = exhibits
        self.workbook_path = workbook_path
        self.default_formats = {} if default_formats is None else default_formats

    def to_excel(self):
        """ Outputs object to Excel.

        Parameters:
        -----------
        workbook_path : str
            The target path and filename of the Excel document
        """
        self.exhibits = copy.deepcopy(self.exhibits)
        if self.exhibits.__class__.__name__ != 'Tabs':
            self.exhibits = Tabs(('sheet1', self.exhibits))
        for sheet in self.exhibits:
            self._write(sheet[1], sheet[0])
            sheet[1].kwargs.update(self.exhibits.kwargs)
            self._set_worksheet_properties(sheet[1], sheet[0])
        self.writer.save()
        self.writer.close()

    def _write(self, exhibit, sheet, start_row=0, start_col=0):
        klass = exhibit.__class__.__name__
        if getattr(exhibit, 'title', None) is not None:
            t = copy.deepcopy(exhibit.title)
            exhibit.title = None
            exhibit = Column(t, exhibit)
            self._write(exhibit, sheet, start_row, start_col)
        elif klass in ['Row', 'Column']:
            start_row = start_row
            start_col = start_col
            for item in exhibit.args:
                self._write(item, sheet, start_row, start_col)
                if klass == 'Column':
                    start_row = start_row + item.height
                if klass == 'Row':
                   start_col = start_col + item.width
        else:
            exhibit.start_row = start_row
            exhibit.start_col = start_col
            exhibit.sheet_name = sheet
            try:
                exhibit.worksheet = self.writer.sheets[exhibit.sheet_name]
            except:
                pd.DataFrame().to_excel(self.writer, sheet_name=exhibit.sheet_name)
                exhibit.worksheet = self.writer.sheets[exhibit.sheet_name]
            if klass in ['DataFrame', 'RSpacer', 'CSpacer']:
                if exhibit.header:
                    self._write_header(exhibit)
                if exhibit.index:
                    self._write_index(exhibit)
                self._register_formats(exhibit)
                self._write_data(exhibit)
            if klass in ['Title', 'Series']:
                self._write_title(exhibit)
            if klass == 'Image':
                self._write_image(exhibit)

    def _set_worksheet_properties(self, exhibit, sheet):
        ''' Format column widths, headers footers, etc.'''
        exhibit.worksheet = self.writer.sheets[sheet]
        widths = [min(self.max_column_width, item)
                  for item in exhibit.column_widths]
        for num, item in enumerate(widths):
            exhibit.worksheet.set_column(num, num, item)
        kwargs = exhibit.kwargs
        if kwargs.get('set_header', None) is not None:
            exhibit.worksheet.set_header(kwargs['set_header'])
        if kwargs.get('set_footer', None) is not None:
            exhibit.worksheet.set_footer(kwargs['set_footer'])
        if kwargs.get('repeat_rows', None) is not None:
            exhibit.worksheet.repeat_rows(*kwargs['repeat_rows'])
        if kwargs.get('fit_to_pages', None) is not None:
            exhibit.worksheet.fit_to_pages(*kwargs['fit_to_pages'])
        if kwargs.get('hide_gridlines', None) is not None:
            exhibit.worksheet.hide_gridlines()
        if sum(widths) > self.max_portrait_width:
            exhibit.worksheet.set_landscape()
        else:
            exhibit.worksheet.set_portrait()

    def _write_title(self, exhibit):
        start_row = exhibit.start_row
        start_col = exhibit.start_col
        end_row = start_row + exhibit.height
        end_col = start_col + exhibit.width - 1
        row_rng = range(start_row, end_row)
        title_format = []
        for item in exhibit.title_formats:
            v = self.default_formats.copy()
            v.update(item)
            title_format.append(self.writer.book.add_format(v))
        if exhibit.width > 1:
            for r in row_rng:
                exhibit.worksheet.merge_range(
                    r, start_col, r, end_col,
                    exhibit.data.iloc[r - start_row][0],
                    title_format[r - start_row])
        else:
            for r in row_rng:
                exhibit.worksheet.write(
                    r, start_col,
                    exhibit.data.iloc[r - start_row][0],
                    title_format[r - start_row])

    def _write_image(self, exhibit):
        exhibit.worksheet.insert_image(
            exhibit.start_row, exhibit.start_col, exhibit.data,
            options=exhibit.formats)
        exhibit.worksheet.merge_range(
            exhibit.start_row, exhibit.start_col,
            exhibit.start_row + exhibit.height - 1,
            exhibit.start_col + exhibit.width - 1, '')

    def _write_header(self, exhibit):
        ''' Adds column headers to data table '''
        if not exhibit.index:
            headers = exhibit.data.columns
        else:
            headers = [exhibit.index_label]+list(exhibit.data.columns)
        header_format = self.default_formats.copy()
        header_format.update(exhibit.header_formats)
        header_format = self.writer.book.add_format(header_format)
        for col_num, value in enumerate(headers):
            exhibit.worksheet.write(
                exhibit.start_row,
                col_num + exhibit.start_col,
                value, header_format)
            if exhibit.col_nums:
                exhibit.worksheet.write(
                    exhibit.start_row + 1, col_num, -col_num-1, header_format)

    def _write_index(self, exhibit):
        ''' Adds row index to data table '''
        index_format = self.default_formats.copy()
        index_format.update(exhibit.index_formats)
        index_format = self.writer.book.add_format(index_format)
        for row_num, value in enumerate(exhibit.data.index.astype(str)):
            exhibit.worksheet.write(
                row_num + exhibit.start_row + exhibit.header + \
                exhibit.col_nums,
                exhibit.start_col,
                value, index_format)
            exhibit.worksheet.set_column(
                first_col=exhibit.start_col, last_col=exhibit.start_col,
                width=exhibit.column_widths[0])

    def _register_formats(self, exhibit):
        """
        Registers all unique user-defined formats with the Workbook
        """
        for num, k in enumerate(exhibit.formats.keys()):
            # Add unique formats
            v = exhibit.formats[k]
            if type(v) is dict:
                col_formats = v
            elif type(v) is str:
                col_formats = {'num_format': v}
            else:
                raise ValueError('Cannot infer format ' + str(v))
            col_formats = self.default_formats.copy()
            col_formats.update(v)
            if self.formats.get(json.dumps(col_formats), None) is None:
                self.formats[json.dumps(col_formats)] = \
                    self.writer.book.add_format(col_formats)

        for k in exhibit.formats.keys():
            # Assign formats to columns
            v = self.default_formats.copy()
            v.update(exhibit.formats[k])
            exhibit.formats[k] = self.formats[json.dumps(v)]

    def _write_data(self, exhibit):
        start_row = exhibit.start_row + exhibit.col_nums + exhibit.header
        start_col = exhibit.start_col + exhibit.index
        end_row = start_row + exhibit.data.shape[0]
        end_col = start_col + exhibit.data.shape[1]
        row_rng = range(start_row, end_row)
        col_rng = range(start_col, end_col)
        d = exhibit.data.fillna('').values
        for c in col_rng:
            c_idx = c - exhibit.index - exhibit.start_col
            fmt = exhibit.formats[exhibit.data.columns[c_idx]]
            for r in row_rng:
                r_idx = r - exhibit.col_nums - exhibit.header - \
                        exhibit.start_row
                exhibit.worksheet.write(r, c, d[r_idx, c_idx], fmt)
                if r == start_row:
                    exhibit.worksheet.set_column(
                        first_col=c, last_col=c,
                        width=exhibit.column_widths[c_idx + exhibit.index])


class Title:
    """ Make cool looking titles

    Parameters
    ----------
    data : str or list of str
        Title and subtitle texts
    formats : list of dicts
        Formats to be applied to the title
    width :
        The width the title should span
    """
    def __init__(self, data, formats=[], width=None, column_widths=None, *args, **kwargs):
        if type(data) is str:
            data = [data]
        self.data = pd.DataFrame(data)
        self.title_formats = self._set_format(formats)
        self.width = width
        self.height = len(self.data)
        self.header = False
        self.index = False
        self.col_nums = False
        self.formats = {}
        if column_widths is not None and width is not None:
            self._column_widths = [column_widths]*self.width
        self.kwargs = kwargs

    @property
    def column_widths(self):
        if hasattr(self, '_column_widths'):
            return self._column_widths
        return [0]*self.width

    @column_widths.setter
    def column_widths(self, value):
        self._column_widths = value

    def __len__(self):
        return len(self.data)

    def _default_format(self):
        return [{'font_size': 20, 'align': 'center'},
                {'font_size': 16, 'align': 'center'},
                {'font_size': 16, 'align': 'center'}] + \
               [{'font_size': 13, 'align': 'center'}] * (len(self.data)-3)

    def _set_format(self, overlay):
        original = self._default_format()
        if overlay is not None:
            if type(overlay) is list:
                for num, item in enumerate(overlay):
                    original[num].update(item)
            else:
                for num, item in enumerate(original):
                    original[num].update(overlay)
        return original

    def to_excel(self, workbook_path, default_formats=None):
        """ Outputs object to Excel.

        Parameters:
        -----------
        workbook_path : str
            The target path and filename of the Excel document
        """
        _Workbook(workbook_path=workbook_path, exhibits=self,
                  default_formats=default_formats).to_excel()

class Series(Title):
    def __init__(self, data, formats=[], width=1, column_widths=1):
        data = pd.Series(data).to_frame()
        super().__init__(data, formats, width, column_widths)

    def _default_format(self):
        base_formats = {
            'float64': {'num_format': '#,0.00', 'align': 'center'},
            'float32': {'num_format': '#,0.00', 'align': 'center'},
            'int64': {'num_format': '#,0', 'align': 'center'},
            'int32': {'num_format': '#,0', 'align': 'center'},
            '<M8[ns]': {'num_format': 'yyyy-mm-dd hh:mm', 'align': 'center'},
            'datetime64[ns]': {'num_format': 'yyyy-mm-dd hh:mm', 'align': 'center'},
            'object': {'align': 'left'},
        }
        return [base_formats.get(
                    str(self.data[0].dtype), base_formats['object'])
                ] * len(self.data)

class Image:
    """ Image allows for the embedding of images into a spreadsheet

    Parameters
    ----------
    data : str
        path to the image file, e.g. sample.png or ./sample.jpg
    width : int
        the number of columns consumed by the image
    height : int
        the number of rows consumed by the image
    formats : dict
        xlsxwriter options for modifying the image
    """
    def __init__(self, data, width=1, height=1, formats={}, *args, **kwargs):
        self.data = data
        self.width = width
        self.height = height
        self.formats = formats
        self.kwargs = kwargs

    def to_excel(self, workbook_path, default_formats=None):
        """ Outputs object to Excel.

        Parameters:
        -----------
        workbook_path : str
            The target path and filename of the Excel document
        """
        _Workbook(workbook_path=workbook_path, exhibits=self,
                  default_formats=default_formats).to_excel()


class DataFrame:
    """
    Excel-ready DataFrame

    Parameters:
    -----------
    data : DataFrame or Triangle (2D)
        The data to be places in the exhibit
    formats : dict
        The formats to be applied to the data.  Options include
        'money', 'percent', 'decimal', 'date', 'int', and 'text'.  Each
        format can be overriden at the class level by overriding its:
        resepective format dict (e.g. DataFrame.money_format, ...)
    header : bool or list (len of data.columns)
        False uses no headers, True uses headers from data. Alternatively,
        a list of strings will override headers.
    header_formats : dict
        The formats to be applied to the header, if any
    col_nums : bool
        Set to True will insert column numbers into the exhibit.
    index : bool, default True
        Write row names (index).
    index_formats : dict
        The formats to be applied to the index, if any
    index_label : str or sequence, optional
        Column label for index column(s) if desired.
    title : list
        A list of strings up to length 4 (Title, subtitle1, subtitle2,
        subtitle3) to be placed above the data in the exhibit.

    """

    min_numeric_col_width = 12
    # Padding since bold characters are slightly larger than regular
    # and need a bit more width
    col_padding_multiplier = 1.1

    def __init__(self, data, formats=None,
                 header=True, header_formats=None, col_nums=False,
                 index=True, index_label='', index_formats=None,
                 title=None, column_widths=None, *args, **kwargs):
        self.index_formats = {
            'num_format': '0;(0)', 'text_wrap': True,
            'bold': True, 'valign': 'bottom', 'align': 'center'}
        self.header_formats = {
            'num_format': '0;(0)', 'text_wrap': True, 'bottom': 1,
            'bold': True, 'valign': 'bottom', 'align': 'center'}
        if type(data) is not pd.DataFrame:
            data = data.to_frame()
        self.data = data
        self.header = header
        self.index = index
        self.index_label = index_label
        self.col_nums = col_nums
        self.format_validation(formats)
        if column_widths is None:
            self.column_widths = self.get_column_widths()
        else:
            self.column_widths = column_widths
        self.height = data.shape[0] + self.col_nums + self.header
        self.width = data.shape[1] + self.index
        if type(title) is str:
            title = [title]
        if title is None or title == []:
            title = None
        else:
            self.height = self.height + len(title)
        if type(title) is list:
            title = Title(title)
        self.title = title
        if header_formats is not None:
            self.header_formats.update(header_formats)
        if index_formats is not None:
            self.index_formats.update(index_formats)
        self.kwargs = kwargs

    def to_excel(self, workbook_path, default_formats=None):
        """ Outputs object to Excel.

        Parameters:
        -----------
        workbook_path : str
            The target path and filename of the Excel document
        """
        _Workbook(workbook_path=workbook_path, exhibits=self,
                  default_formats=default_formats).to_excel()

    def get_column_widths(self):
        """ Default column widths """
        if self.index:
            row_w = [max(self.data.index.astype(str).str.len())]
            header_w = [max([len(token)
                             for token in str(self.index_label).split(' ')])]
        else:
            row_w = []
            header_w = []
        headers = list(self.data.columns)
        header_w = header_w + \
                   [max([len(token) for token in str(item).split(' ')])
                    for item in headers]
        numeric_cols = self.data.select_dtypes('number').columns
        row_w = row_w + \
                [(self.min_numeric_col_width if item in numeric_cols
                 else max(self.data[item].astype(str).str.len()))
                 for item in headers]
        return [max(item)* self.col_padding_multiplier
                for item in zip(header_w, row_w)]

    def format_validation(self, formats):
        ''' Creates an Excel format compatible dictionary '''
        base_formats = {
            'float64': {'num_format': '#,0.00', 'align': 'center'},
            'float32': {'num_format': '#,0.00', 'align': 'center'},
            'int64': {'num_format': '#,0', 'align': 'center'},
            'int32': {'num_format': '#,0', 'align': 'center'},
            '<M8[ns]': {'num_format': 'yyyy-mm-dd hh:mm', 'align': 'center'},
            'datetime64[ns]': {'num_format': 'yyyy-mm-dd hh:mm', 'align': 'center'},
            'object': {'align': 'left'},
        }
        if self.data.columns.name is not None:
            idx = self.data.index.to_frame().dtypes
            idx.index = [self.data.columns.name]
        else:
            idx = pd.Series()
        cols = self.data.dtypes.append(idx)
        self.formats = {
            k: base_formats.get(v, base_formats['object'])
            for k, v in dict(cols).items()
        }

        if type(formats) is list:
            self.formats.update(dict(zip(self.data.columns, formats)))
        elif type(formats) is str:
            self.formats.update(dict(zip(
                self.data.columns,
                [{'num_format': formats}] * len(self.data.columns))))
        elif type(formats) is dict and formats != {}:
            if list(formats.keys())[0] not in self.data.columns:
                self.formats.update(dict(zip(
                    self.data.columns,
                    [formats] * len(self.data.columns))))
            else:
                formats = {k: v if type(v) is dict else {'num_format': v}
                           for k, v in formats.items()}
                self.formats.update(formats)
        else:
            pass


class RSpacer(DataFrame):
    """ Convenience class to create a vertical spacer in a Row container"""
    def __init__(self, width=1, column_widths=2.25, *args, **kwargs):
        data = pd.DataFrame(dict(zip(list(range(width)), [' '] * width)),
                            index=[0])
        temp = DataFrame(data, index=False, header=False)
        for k, v in temp.__dict__.items():
            setattr(self, k, v)
        self.column_widths = [column_widths] * width
        self.kwargs = kwargs


class VSpacer(RSpacer):
    pass


class CSpacer(DataFrame):
    """ Convenience class to create a horizontal spacer in a Column container"""
    def __init__(self, height=1, column_widths=2.25, *args, **kwargs):
        data = pd.DataFrame({' ': [' '] * height})
        temp = DataFrame(data, index=False, header=False)
        for k, v in temp.__dict__.items():
            setattr(self, k, v)
        self.column_widths = [column_widths]
        self.kwargs = kwargs


class HSpacer(CSpacer):
    pass

class _Container():
    """ Base class for Row and Column
    """
    def __init__(self, *args, **kwargs):
        self.args = tuple([copy.deepcopy(item) for item in args])
        self._title_len = 0
        for item in self.args:
            if item.__class__.__name__ in ['Title']:
                self._title_len = len(item)
                item.width = self.width
        self.kwargs = kwargs

    def __getitem__(self, key):
        return self.args[key]

    def __len__(self):
        return len(self.args)

    def to_excel(self, workbook_path, default_formats=None):
        """ Outputs object to Excel.

        Parameters:
        -----------
        workbook_path : str
            The target path and filename of the Excel document
        """
        _Workbook(workbook_path=workbook_path, exhibits=self,
                  default_formats=default_formats).to_excel()


class Row(_Container):
    """
    Lay out child components in a single horizontal row.
    Children can be specified as positional arguments, as a single argument
    that is a sequence.

    Parameters
    ----------
    args:
        Children can be of the chainlader DataFrame, Row, and Column classes.
    title: optional (str or list)
        The title to be displayed across the top of the container.  Must be
        specified using the keyword `title=`

    Attributes
    ----------
    height : int
        Height of the container and is a function of the elements it contains
    width : int
        Width of the container and is a function of the elements it contains

    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for num, item in enumerate(self.args):
            if item.__class__.__name__ in ['Title']:
                self.args = Column(item, Row(*self.args[num + 1:])),

    @property
    def height(self):
        return max([item.height for item in self.args])

    @property
    def width(self):
        return sum([item.width for item in self.args
                    if item.__class__.__name__ not in ['Title']])

    @property
    def column_widths(self):
        if hasattr(self, '_column_widths'):
            return self._column_widths
        column_widths = []
        for item in [getattr(item, 'column_widths', []) for item in self.args]:
            column_widths = column_widths + item
        return column_widths

    @column_widths.setter
    def column_widths(self, value):
        self._column_widths = value


class Column(_Container):
    """
    Lay out child components in a single vertical column.
    Children can be specified as positional arguments, as a single argument
    that is a sequence.

    Parameters
    ----------
    args:
        Children can be of the chainlader DataFrame, Row, and Column classes.
    title: optional (str or list)
        The title to be displayed across the top of the container.  Must be
        specified using the keyword `title=`

    Attributes
    ----------
    height : int
        Height of the container and is a function of the elements it contains
    width : int
        Width of the container and is a function of the elements it contains

    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for item in self.args:
            if item.__class__.__name__ not in ['Title']:
                item.column_widths = self.column_widths

    @property
    def height(self):
        return sum([item.height for item in self.args])


    @property
    def width(self):
        return max([item.width for item in self.args
                    if item.__class__.__name__ not in ['Title']])

    @property
    def column_widths(self):
        if hasattr(self, '_column_widths'):
            return self._column_widths
        data = np.array(
            [item.column_widths for item in self.args
             if item.__class__.__name__ not in ['Title']])
        lens = np.array([len(i) for i in data])
        mask = np.arange(lens.max()) < lens[:,None]
        out = np.zeros(mask.shape, dtype=data.dtype)
        out[mask] = np.concatenate(data)
        return list(np.max(out, axis=0))

    @column_widths.setter
    def column_widths(self, value):
        self._column_widths = value

class Tabs:
    """
    Layout exhibits across worksheets.

    Parameters
    ----------
    args:
        Children must be a tuple with a sheet name and any of chainlader
        DataFrame, Row, and Column classes.  For example,
        ('sheet1', cl.DataFrame(data))
    """

    def __init__(self, *args, **kwargs):
        if len(args) != set([item[1] for item in args]):
            self.args = tuple([(item[0], copy.deepcopy(item[1]))
                               for item in args])
        else:
            self.args = args
        valid = ['Row', 'Column', 'Title', 'Series', 'DataFrame', 'Image']
        if len([item[1].__class__.__name__ for item in self.args
                if item[1].__class__.__name__ not in valid]) > 0:
             raise TypeError('Valid objects include '  + ', '.join(valid))
        self.kwargs = kwargs

    def __getitem__(self, key):
        return self.args[key]

    def __len__(self):
        return len(self.args)

    def to_excel(self, workbook_path, default_formats=None):
        """ Outputs object to Excel.

        Parameters:
        -----------
        workbook_path : str
            The target path and filename of the Excel document

        """
        _Workbook(workbook_path=workbook_path, exhibits=self,
                  default_formats=default_formats).to_excel()
