.. _examples:

.. currentmodule:: xlcompose


Templating
============

`xlcompose` objects have a nice dictionary/list like representation.  This allows
the structure of the object to be defined nicely using YAML notation.

**Example:**
   >>> xlc.Column(
   ...     xlc.Title(data=['This is a', 'sample title'], formats={'align':'left'})
   ... )

The python object above can be expressed in a YAML notation and read into
`xlcompose` as follows:

**Example:**
   >>> my_template = """
   ... - Column:
   ...     - Title:
   ...         data: ['This is a', 'sample title']
   ...         formats:
   ...             align: 'left'
   ... """
   >>> xlc.load_yaml(my_template)

The benefit of doing this is we can save templates as YAML files for reuse with
new data. Maybe you run a small business and have an Invoice template.  Perhaps
you're a financial analyst and you have P&L templates for you balance sheet and\
income statement.

Context aware templates with jinja2
-----------------------------------
Like most files, YAML files can be manipulated using the jinja templating language.
Jinja is used extensively in python web frameworks like `flask` and `django`.  By
using jinja, we can insert data into our template.

**Example:**
   >>> my_template = """
   ... - Column:
   ...     - Title:
   ...         data: ['This is a', {{title}}]
   ...         formats:
   ...             align: 'left'
   ... """
   >>> xlc.load_yaml(my_template, title='context-aware title')

Notice how `load_yaml` takes an arbitrary set of keyword arguments (kwargs), in
this case, just 'title' and `xlcompose` uses jinja to inject these values into
your template.  The jinja templating language has a ton of features that are
covered at great length elsewhere so we will only touch on the basics here.
That said, the full power of jinja can be used with your `xlcompose` templates.

.. note::
   Jinja templating is incredibly powerful and can lead to an impressive Excel
   template library as it does for impressive web sites. To learn more about it
   visit the Jinja Docs. (https://jinja.palletsprojects.com/en/master/templates/)

Python objects
---------------
YAML supports primitive data types - `str`, `int`, `float`, `dict`, `list`.
Likewise, `jinja2` was designed to inject text into templates.  This is great
for text based outputs like HTML, but `xlcompose` is trying to build complex
python objects like the `DataFrame`.  If you pass a complex object into a jinja
variable, jinja will grab the string representation of the object.

**This will not work:**
   >>> my_template = """
   ... - Column:
   ...     - Title:
   ...         data: ['This is a', 'sample title']
   ...     - DataFrame:
   ...         data: {{data}}
   ... """
   >>> xlc.load_yaml(my_template, data=data)

**It is equivalent to:**
   >>> xlc.Column(
   ...     xlc.Title(data=['This is a', 'sample title']),
   ...     xlc.DataFrame(data.__str__())
   ... )

We don't want the string representation of our `DataFrame`, we want the actual
`DataFrame` itself.  To accomplish this, `xlcompose` has a jinja2 directive
`{% eval %}` that evaulates the object as it would in python.

**This is correct:**
   >>> my_template = """
   ... - Column:
   ...     - Title:
   ...         data: ['This is a', 'sample title']
   ...     - DataFrame:
   ...         data: {% eval %}data{% endeval %}
   ... """
   >>> xlc.load_yaml(my_template, data=data)

.. warning::
  `{% eval %}` is called `eval` because it is evaluating text as python code.
  It is important to know and trust the templates you are using so that you do not
  unintentionally run malicious code.

`{% eval %}` can even run on methods of your context objects.  It is important to
note that `{% eval %}` runs as a final step in your template.  This means you can
pass other variables as context to your `eval` directive.

**Example:**
   >>> my_template = """
   ... - Column:
   ...     - Title:
   ...         data: ['This is a', 'sample title']
   ...     - DataFrame:
   ...         data: {% eval %}data.groupby('{{group}}').sum(){% endeval %}
   ... """
   >>> xlc.load_yaml(my_template, data=data, group='country')

In the above example, the variable `{{group}}` gets replaced with 'country'
before the `{% eval %}` directive evaluates the python expression.

Other jinja directives
----------------------
Again, there is a wealth of information on the web about how to use `jinja2` and
its best to learn about it from existing tutorials.  That said, here is one example
that shows loops through a hypothetical `DataFrame` - for each country we total
revenue and expense by product.

**Example:**
   >>> my_template = """
   ... - Column:
   ...     - Title:
   ...         data: ['Summary of Revenue and Expense', 'By Product and Country']
   ...     {% for country in data['country'].unique() %}
   ...     - Series:
   ...         data: 'Revenue and Expense for {{country}}'
   ...         formats:
   ...           bold: True
   ...     - DataFrame:
   ...         data: {% eval %}data.groupby({{country}})[['Revenue', 'Expense']].sum(){% endeval %}
   ...         formats:
   ...           align: 'center'
   ...           num_format: '$#,#'
   ...     {% endfor %}
   ... """
   >>> xlc.load_yaml(my_template, data=data)

We could have done this looping in python, but sometimes it is far more convenient
to have the context-aware template use the data it is passed to generate itself.
This approach leaves re-use of the template much simpler.  In the above,
we only need to pass it `data`, which is much simpler then recreating the loop in
python every time we want to re-use the template.

Setting a jinja environment
---------------------------
Jinja has a concept of an environment.  This is a folder where all templates
reside.  To make all the templates aware of each other (so that they can call
on each other), you would create a Jinja environment.  You do not need to
set an environment if you're working with only one template at a time.  However,
if you do have a jinja environment, you can pass it into `xlcompose`:

**Example:**
   >>> xlc.load_yaml(template='template.yaml', env=my_jinja_env, data=data, ...)
