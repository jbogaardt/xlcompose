import pandas as pd
import numpy as np
import copy
import json
import os
from io import BytesIO
import xlsxwriter
import yaml

settings = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.yaml')
with open(settings, 'r') as f:
    settings = yaml.load(f.read(),Loader=yaml.SafeLoader)


class _Workbook:
    """
    Excel Workbook level configurations.  This is not part of the end_user API.
    This class  facilitates:
    1. Crawling the xlcompose object on which `to_excel` is called
    2. Writing the each nested object to an Excel file

    """

    def __init__(self, workbook_path, exhibits, default_formats):
        """ Initialize the writer object
        """
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
        if self.exhibits.__class__.__name__ == 'Sheet':
            self.exhibits = Tabs(self.exhibits)
        if self.exhibits.__class__.__name__ != 'Tabs':
            self.exhibits = Tabs(Sheet('sheet1', self.exhibits))
        for sheet in self.exhibits:
            self._write(sheet.layout, sheet.name)
            sheet.layout.kwargs.update(sheet.kwargs)
            self._set_worksheet_properties(sheet.layout, sheet.name)
        self.writer.save()


    def _write(self, exhibit, sheet, start_row=0, start_col=0):
        """
        Parameters
        ----------
        exhibit :
            An xlcompose object
        sheet : str
            The sheet name in Excel to write to
        start_row : int
            The starting row on which to write the exhibit
        start_col : int
            The starting column on which to write the exhibit
        """
        # Get xlcompose object for special handling
        klass = exhibit.__class__.__name__
        if getattr(exhibit, 'title', None) is not None:
            ## Special handling of title.  It must live in a Column
            #  if it doesn't already
            t = copy.deepcopy(exhibit.title)
            exhibit.title = None
            exhibit = Column(t, exhibit)
            self._write(exhibit, sheet, start_row, start_col)
        elif klass in ['Row', 'Column']:
            ## Need to render each object in Row and Column args keeping in
            #  mind the start_row and start_col of the container
            start_row = start_row
            start_col = start_col
            for item in exhibit.args:
                self._write(item, sheet, start_row, start_col)
                if klass == 'Column':
                    start_row = start_row + item.height
                if klass == 'Row':
                   start_col = start_col + item.width
        else:
            # General rendering for Non-container objects
            exhibit.start_row = start_row
            exhibit.start_col = start_col
            exhibit.sheet_name = sheet
            # Create sheet if it doesn't already exist
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
                self._write_series(exhibit)
            if klass == 'Image':
                self._write_image(exhibit)

    def _set_worksheet_properties(self, exhibit, sheet):
        """ Set worksheet level properties. Called once the entire sheet has
        been rendered. These set worksheet level settings and in many cases is
        a straight pass through to xlsxwriter.  Can we inspect the xlsxwriter
        properties to dynamically support its functionality?

        Parameters
        ----------
        exhibit :
            An xlcompose object
        sheet : str
            The sheet name in Excel to write to
        """
        exhibit.worksheet = self.writer.sheets[sheet]
        widths = [min(settings['max_column_width'], item)
                  for item in exhibit.column_widths]
        heights = [min(settings['max_row_height'], item) if item is not None else item
                   for item in exhibit.row_heights]
        for num, item in enumerate(widths):
            exhibit.worksheet.set_column(num, num, item)
        for num, item in enumerate(heights):
            if item is not None:
                exhibit.worksheet.set_row(num, item)
        kwargs = exhibit.kwargs
        exhibit.worksheet.fit_to_pages(*kwargs.get('fit_to_pages', (1,0)))
        bool_funcs = [
            'set_page_view', 'print_row_col_headers', 'hide_row_col_headers',
            'center_vertically', 'center_horizontally']
        for func in bool_funcs:
            if kwargs.get(func):
                getattr(exhibit.worksheet, func)()
        passthru_funcs = [
            'hide_gridlines', 'set_print_scale', 'set_start_page', 'set_paper'
            'set_h_pagebreaks', 'set_v_pagebreaks', 'print_across']
        for func in passthru_funcs:
            if kwargs.get(func):
                getattr(exhibit.worksheet, func)(kwargs[func])
        starg_funcs = [
            'freeze_panes', 'repeat_rows', 'repeat_columns', 'set_margins',
            'print_area']
        for func in starg_funcs:
            if kwargs.get(func):
                getattr(exhibit.worksheet, func)(*kwargs[func])
        if kwargs.get('set_header', None) is not None:
            if type(kwargs['set_header']) is list:
                exhibit.worksheet.set_header('\n'.join(kwargs['set_header']))
            else:
                exhibit.worksheet.set_header(kwargs['set_header'])
        if kwargs.get('set_footer', None) is not None:
            if type(kwargs['set_footer']) is list:
                exhibit.worksheet.set_footer('\n'.join(kwargs['set_footer']))
            else:
                exhibit.worksheet.set_footer(kwargs['set_footer'])
        if kwargs.get('set_landscape'):
            exhibit.worksheet.set_landscape()
        elif kwargs.get('set_portrait'):
            exhibit.worksheet.set_portrait()
        else:
            if sum(widths) > settings['max_portrait_width']:
                exhibit.worksheet.set_landscape()
            else:
                exhibit.worksheet.set_portrait()

    def _write_series(self, exhibit):
        """ Writes a Series or Title object.  Special considerations are
        that these objects can take a format list that applies to each element
        of the Series.  These also merge cells to span their designated `width`.
        """
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
        """ Writes an image object and merges the cells behind it based on the
        images designated `height` and `width`
        """
        exhibit.worksheet.insert_image(
            exhibit.start_row, exhibit.start_col, exhibit.data,
            options=exhibit.formats)
        exhibit.worksheet.merge_range(
            exhibit.start_row, exhibit.start_col,
            exhibit.start_row + exhibit.height - 1,
            exhibit.start_col + exhibit.width - 1, '')

    def _write_header(self, exhibit):
        ''' Adds column headers to data table '''
        if type(exhibit.data.columns) == pd.PeriodIndex:
            headers = exhibit.data.columns.astype(str)
        else:
            headers = exhibit.data.columns
        if exhibit.index:
            headers = [exhibit.index_label]+list(headers)
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


class _XLCBase:
    px_per_row = 15
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'styles.css'), 'r') as f:
        styles = '<style>' + f.read() + '</style>'

    def to_excel(self, workbook_path, default_formats=None):
        """ Outputs object to Excel.

        Parameters:
        -----------
        workbook_path : str
            The target path and filename of the Excel document
        """
        _Workbook(workbook_path=workbook_path, exhibits=self,
                  default_formats=default_formats).to_excel()

    def _repr_html_(self):
        return self.styles + self._get_html()


    def _get_html(self, my_height=0, my_width=100):
        width = 'width:' + str(my_width) + '%;' if my_width < 100 else 'width: auto;'
        name = self.__class__.__name__
        return '<div class="xlccontainer-' + name + \
                '" style="height:' + str(self.height*self.px_per_row) + 'px;' + \
                width + '"><div class="xlclabel-' + \
                name + '">' + name + \
                '</div></div>\n'


class Title(_XLCBase):
    """ Title objects are Series-like objects that has its own formatting style.

    Difference between a Series and a Title is that a Title `width` will be
    inferred from its container.  Additionally, Title objects have a different
    default format.

    Parameters
    ----------
    data : str or list of str
        Title and subtitle texts
    formats : list
        The formats to be applied to the Title. Each element in the Title
        will be assigned a format from this list with the corresponding index.
    width :
        The width the title should span.  If omitted, the title will take on
        the width of its container.
    column_widths : list
        list of floats representing the column widths of each column within the
        DataFrame.  If omitted, then widths are set by inspecting the data.
    row_heights : list
        list of floats representing the row heights of each row within the
        Series.  If omitted, then heights are set by inspecting the data.
    """

    def __init__(self, data, formats=[], width=None,
                 column_widths=None, row_heights=None, *args, **kwargs):
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
        if row_heights is not None:
            self._row_heights = [row_heights]*len(self.data)
        self.kwargs = kwargs

    @property
    def column_widths(self):
        if hasattr(self, '_column_widths'):
            return self._column_widths
        return [0]*self.width

    @column_widths.setter
    def column_widths(self, value):
        self._column_widths = value

    @property
    def row_heights(self):
        if hasattr(self, '_row_heights'):
            return self._row_heights
        return [None] * len(self.data)

    @row_heights.setter
    def row_heights(self, value):
        self._row_heights = value

    def __len__(self):
        return len(self.data)

    def _default_format(self):
        title_formats = copy.deepcopy(settings['title_formats'])
        return title_formats[:3] + title_formats[-1:] * (len(self.data)-3)

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


class Series(Title):
    """ Excel-ready Series object. This component does not render the index or
    series name.  To include those, see the `DataFrame` component.

    Parameters
    ----------
    data : list or Series
        The data to be placed into the series
    formats : list
        The formats to be applied to the Series. Each element in the Series
        will be assigned a format from this list with the corresponding index.
    widths: int
        The number of columns the Series should span. For values higher than 1,
        the columns spanned will be merged.
    column_widths : list
        list of floats representing the column widths of each column within the
        DataFrame.  If omitted, then widths are set by inspecting the data.
    row_heights : list
        list of floats representing the row heights of each row within the
        Series.  If omitted, then heights are set by inspecting the data.
    """

    def __init__(self, data, formats=None, width=1, column_widths=1,
                 row_heights=None, *args, **kwargs):
        formats = formats or []
        if type(data) == pd.Series:
            data = data.to_frame()
        else:
            data = pd.Series(data).to_frame()
        super().__init__(data, formats, width, column_widths, row_heights,
                         *args, **kwargs)

    def _default_format(self):
        base_formats = copy.deepcopy(settings['base_formats'])
        return [base_formats.get(
                    str(self.data.iloc[0].dtype), base_formats['object'])
                ] * len(self.data)



class Image(_XLCBase):
    """ Image allows for the embedding of images into a spreadsheet


    Parameters
    ----------
    data : str
        path to the image file, e.g. sample.png or ./sample.jpg or a matplotlib
        `AxesSubplot`
    width : int
        the number of columns consumed by the image
    height : int
        the number of rows consumed by the image
    formats : dict
        xlsxwriter options for modifying the image
    """

    def __init__(self, data, width=1, height=1, formats={}, *args, **kwargs):
        if data.__class__.__name__ in ['AxesSubplot', 'Figure']:
            #inch_to_row = 0.01431127
            #inch_to_col = 0.077469335
            #img_shape = data.get_figure().get_size_inches()
            imgdata = BytesIO()
            data.get_figure().savefig(imgdata, format="png")
            data = '_.png'
            formats.update({'image_data': imgdata})
        self.data = data
        self.width = width
        self.height = height
        self.formats = formats
        self.kwargs = kwargs
        if kwargs.get('column_widths'):
            self.column_widths = kwargs.get('column_widths')
        self.column_widths = [8.09]*width


class DataFrame(_XLCBase):
    """
    An Excel-ready DataFrame.

    Parameters:
    -----------
    data : DataFrame
        The data to be placed in the exhibit. Must be a pandas DataFrame or an
        object with the `to_frame()` method.
    formats : dict
        The formats to be applied to the data columns.  Dictionary keys can be
        either column names to do column specific formatting OR `xlsxwriter`
        format names to use consistent formating across the entire `DataFrame`.
    header : bool or list (len of data.columns)
        False uses no headers, True uses headers from data. Alternatively,
        a list of strings will override headers.
    header_formats : dict
        The formats to be applied to the header, if any
    col_nums : bool
        Set to True will insert column numbers into the exhibit.
    index : bool, default True
        Write row names (index).
    index_label : str or sequence, optional
        Column label for index column(s) if desired.
    index_formats : dict
        The formats to be applied to the index, if any
    column_widths : list
        list of floats representing the column widths of each column within the
        DataFrame.  If omitted, then widths are set by inspecting the data.
    row_heights : list
        list of floats representing the row heights of each row within the
        DataFrame.  If omitted, then heights are set by inspecting the data.
    """

    index_formats = copy.deepcopy(settings['index_formats'])
    header_formats = copy.deepcopy(settings['header_formats'])
    base_formats = copy.deepcopy(settings['base_formats'])

    def __init__(self, data, formats=None,
                 header=True, header_formats=None, col_nums=False,
                 index=True, index_label='', index_formats=None,
                 column_widths=None, row_heights=None, *args, **kwargs):

        if type(data) is not pd.DataFrame:
            data = data.to_frame()
        self.data = data
        self.header = header
        self.index = index
        self.index_label = index_label
        self.col_nums = col_nums
        self._format_validation(formats)
        if column_widths is None:
            self.column_widths = self._get_column_widths()
        else:
            self.column_widths = column_widths
        self.height = data.shape[0] + self.col_nums + self.header
        self.width = data.shape[1] + self.index
        if header_formats is not None:
            self.header_formats.update(header_formats)
        if index_formats is not None:
            self.index_formats.update(index_formats)
        if row_heights is not None:
            self._row_heights = row_heights
        self.kwargs = kwargs

    def _get_column_widths(self):
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
                [(settings['min_numeric_col_width'] if item in numeric_cols
                 else max(self.data[item].astype(str).str.len()))
                 for item in headers]
        return [max(item)* settings['col_padding_multiplier']
                for item in zip(header_w, row_w)]

    @property
    def row_heights(self):
        if hasattr(self, '_row_heights'):
            return self._row_heights
        return [None]*(len(self.data) + (1 - (not self.header)) + self.col_nums)

    @row_heights.setter
    def row_heights(self, value):
        self._row_heights = value


    def _format_validation(self, formats):
        ''' Creates an Excel format compatible dictionary.  Users can specify
        formats in a variety of shorthand ways and this conforms them to an
        xlsxwriter style.
        '''

        if self.data.columns.name is not None:
            idx = self.data.index.to_frame().dtypes
            idx.index = [self.data.columns.name]
        else:
            idx = pd.Series(dtype='object')
        cols = self.data.dtypes.append(idx)
        self.formats = {
            k: self.base_formats.get(v, self.base_formats['object'])
            for k, v in dict(zip(cols.index, cols.values)).items()}
        if type(formats) is list:
            self.formats.update(dict(zip(self.data.columns, formats)))
        elif type(formats) is str:
            self.formats.update(dict(zip(
                self.data.columns,
                [{'num_format': formats}] * len(self.data.columns))))
        elif type(formats) is dict and formats != {}:
            available_formats = [
                item[4:] for item in dir(xlsxwriter.format.Format)
                if item[:3]=='set']
            if len(set(self.data.columns).intersection(formats.keys()))==0:
                self.formats.update(dict(zip(
                    self.data.columns,
                    [formats] * len(self.data.columns))))
            elif len(set(available_formats).intersection(formats.keys()))==0:
                formats = {k: v if type(v) is dict else {'num_format': v}
                           for k, v in formats.items()}
                self.formats.update(formats)
            else:
                raise AttributeError(
                    'DataFrame.formats must be a dict with keys set to column names or xlsxwriter formats.')
        else:
            pass


class RSpacer(DataFrame):
    """ A blank vertical space in a Row container.

    Parameters
    ----------
    width: int
        Width (in number of columns) of the RSpacer
    column_widths: float
        The width to apply to each column of the RSpacer
    """

    def __init__(self, width=1, column_widths=2.25, *args, **kwargs):
        data = pd.DataFrame(dict(zip(list(range(width)), [' '] * width)),
                            index=[0])
        temp = DataFrame(data, index=False, header=False)
        for k, v in temp.__dict__.items():
            setattr(self, k, v)
        self.column_widths = [column_widths] * width
        self.row_heights = [None]
        self.kwargs = kwargs

    def _get_html(self, my_height=0, my_width=100):
        return '<div class="xlccontainer-Spacer" style="height:' + \
                str(my_height*self.px_per_row) + 'px;width:'+ \
                str(my_width)+'%;"><div class="xlclabel-Spacer">RSpacer</div></div>\n'


class CSpacer(DataFrame):
    """ A blank horizontal space in a Column container.

    Parameters
    ----------
    height: int
        Height (in number of rows) of the CSpacer
    row_heights: float
        The height to apply to each row of the CSpacer
    """

    def __init__(self, height=1, row_heights=None, *args, **kwargs):
        data = pd.DataFrame({' ': [' '] * height})
        temp = DataFrame(data, index=False, header=False)
        for k, v in temp.__dict__.items():
            setattr(self, k, v)
        self.column_widths = [2.25]
        self.row_heights = [row_heights] * height
        self.kwargs = kwargs

    def _get_html(self, my_height=0, my_width=100):
        return '<div class="xlccontainer-Spacer" style="width: auto;"><div class="xlclabel-Spacer">CSpacer</div></div>\n'


class _Container(_XLCBase):
    """ Base class for Row and Column """

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


class Row(_Container):
    """
    A container object respresenting a horizontal layout of other `xlcompose`
    objects.

    Parameters
    ----------
    args:
        Any xlcompose objects with the exception of `Sheet` and `Tabs`

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
                self.args = Column(item, Row(*self.args[num + 1:]))

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

    @property
    def row_heights(self):
        if hasattr(self, '_row_heights'):
            return self._row_heights
        data = np.array(
            [item.row_heights for item in self.args
             if item.__class__.__name__ not in ['Title', 'Image']])
        lens = np.array([len(i) for i in data])
        mask = np.arange(lens.max()) < lens[:,None]
        out = np.zeros(mask.shape, dtype=data.dtype)
        out[mask] = np.concatenate(data)
        out[out==None]=0
        return [item if item!=0 else None for item in np.max(out, axis=0)]

    @row_heights.setter
    def row_heights(self, value):
        self._row_heights = value

    def _get_html(self, my_height=100, my_width=100):
        direction = ''
        widths = [item.width/self.width*100 for item in self.args]
        heights = [item.height for item in self.args]
        contents = []
        for num in range(len(self.args)):
            if self.args[num].__class__.__name__ == 'RSpacer':
                contents.append(self.args[num]._get_html(max(heights), widths[num]))
            else:
                contents.append(self.args[num]._get_html(heights[num], widths[num]))
        contents = ''.join(contents)
        width = 'width:' + str(my_width) + '%;' if my_width < 100 else 'width: auto;'
        return '<div class="xlccontainer-Container" style="' + width + '"><div class="xlclabel-Container">Row</div>' + contents + '</div>\n'


class Column(_Container):
    """
    A container object respresenting a vertical layout of other `xlcompose`
    objects.

    Parameters
    ----------
    args:
        Any xlcompose objects with the exception of `Sheet` and `Tabs`

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

    @property
    def row_heights(self):
        if hasattr(self, '_row_heights'):
            return self._row_heights
        row_heights = []
        for item in [getattr(item, 'row_heights', []) for item in self.args]:
            row_heights = row_heights + item
        return row_heights

    @row_heights.setter
    def row_heights(self, value):
        self._row_heights = value

    def _get_html(self, my_height=100, my_width=100):
        direction = ''
        widths = [item.width/self.width*100 for item in self.args]
        heights = [100]*len(widths)
        contents = []
        for num in range(len(self.args)):
            contents.append(self.args[num]._get_html(heights[num], widths[num]))
        contents = ''.join(contents)
        width = 'width:' + str(my_width) + '%;' if my_width < 100 else 'width: auto;'
        return '<div class="xlccontainer-Container" style="flex-direction: column;' + \
                width + ';"><div class="xlclabel-Container">Column</div>' + \
                contents + '</div>\n'

class Tabs(_XLCBase):
    """
    A container object representing multiple worksheets.

    Parameters
    ----------
    args:
        A list of sheets or a tuple with a sheet name and any xlcompose object.
        For example, `('sheet1', xlc.DataFrame(data))`
    """
    _repr_html_ = None

    def __init__(self, *args, **kwargs):
        self.args = [
            Sheet(item[0], copy.deepcopy(item[1]))
            if type(item) is tuple
            else copy.deepcopy(item)
            for item in args]
        self.kwargs = kwargs

    def __getitem__(self, key):
        return self.args[key]

    def __len__(self):
        return len(self.args)


class Sheet(_XLCBase):
    """
    A container object representing an Excel worksheet.

    Parameters
    ----------
    name : str
        The name of the worksheet
    layout :
        An xlcompose object
    fit_to_pages:
        refer to `xlsxwriter` for `fit_to_pages` options
    freeze_panes:
        refer to `xlsxwriter` for `freeze_panes` options
    set_page_view:
        refer to `xlsxwriter` for `set_page_view` options
    print_row_col_headers:
        refer to `xlsxwriter` for `print_row_col_headers` options
    hide_row_col_headers:
        refer to `xlsxwriter` for `hide_row_col_headers` options
    center_vertically:
        refer to `xlsxwriter` for `center_vertically` options
    center_horizontally:
        refer to `xlsxwriter` for `center_horizontally` options
    set_header:
        refer to `xlsxwriter` for `set_header` options
    set_footer:
        refer to `xlsxwriter` for `set_footer` options
    repeat_rows:
        refer to `xlsxwriter` for `repeat_rows` options
    repeat_columns:
        refer to `xlsxwriter` for `repeat_columns` options
    set_margins:
        refer to `xlsxwriter` for `set_margins` options
    hide_gridlines:
        refer to `xlsxwriter` for `hide_gridlines` options
    set_print_scale:
        refer to `xlsxwriter` for `set_print_scale` options
    set_start_page:
        refer to `xlsxwriter` for `set_start_page` options
    set_h_pagebreaks:
        refer to `xlsxwriter` for `set_h_pagebreaks` options
    set_v_pagebreaks:
        refer to `xlsxwriter` for `set_v_pagebreaks` options
    print_across:
        refer to `xlsxwriter` for `print_across` options
    print_area:
        refer to `xlsxwriter` for `print_area` options
    set_paper:
        refer to `xlsxwriter` for `set_paper` options
    set_landscape:
        refer to `xlsxwriter` for `set_landscape` options
    set_portrait:
        refer to `xlsxwriter` for `set_portrait` options
    """
    def __init__(self, name, layout, **kwargs):
        self.name = name
        self.kwargs = kwargs
        if type(layout) is Image:
            self.layout = Column(layout)
        else:
            self.layout = layout
        self.kwargs = kwargs
        self.column_widths = self.layout.column_widths
        self.row_heights = self.layout.row_heights

    def _repr_html_(self):
        return self.layout._repr_html_()

class VSpacer(RSpacer):
    pass

class HSpacer(CSpacer):
    pass
