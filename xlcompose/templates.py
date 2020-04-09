import yaml
import json
import pandas as pd
import re
import ast
import os
import xlcompose.core as core
from jinja2 import nodes, Template, TemplateSyntaxError, FileSystemLoader, Environment, BaseLoader
from jinja2.ext import Extension
from jinja2.nodes import Const


class EvalExtension(Extension):
    """ enabled the { % eval %}{% endeval %} jinja functionality """
    tags = {'eval'}

    def parse(self, parser):
        line_number = next(parser.stream).lineno
        eval_str = [Const('')]
        body = ''
        try:
            eval_str = [parser.parse_expression()]
        except TemplateSyntaxError:
            body = parser.parse_statements(['name:endeval'], drop_needle=True)
        return nodes.CallBlock(self.call_method('_eval', eval_str), [], [], body).set_lineno(line_number)

    def _eval(self, eval_str, caller):
        return '__eval__' + caller() + '__eval__'


def _kwarg_parse(formula):
    names = pd.Series(list(set([node.id for node in ast.walk(ast.parse(formula))
                                if isinstance(node, ast.Name)])))
    names = names[~names.isin(['str', 'list', 'dict', 'int', 'float'])]
    names = names.loc[names.str.len().sort_values(ascending=False).index].to_list()
    names = [(item, '__' + str(num) + '__',  'kwargs[\'' + item + '\']')
             for num, item in enumerate(names)]
    for item in names:
        formula = formula.replace(item[0],item[1])
    for item in names:
        formula = formula.replace(item[1],item[2])
    return formula

def _make_xlc(template, **kwargs):
        """ Recursively generate xlcompose object"""
        if type(template) is list:
            tabs = [_make_xlc(element, **kwargs) for element in template]
            try:
                return core.Tabs(*[(item.name, item) for item in tabs])
            except:
                return core.Tabs(*[('Sheet1', item) for item in tabs])
        key = list(template.keys())[0]
        if key in ['Row', 'Column']:
            return getattr(core, key)(*[_make_xlc(element, **kwargs)
                                       for element in template[key]])
        if key in ['Sheet']:
            sheet_kw = {k: v for k, v in template[key].items()
                        if k not in ['name', 'layout']}
            return core.Sheet(template[key]['name'],
                             _make_xlc(template[key]['layout'], **kwargs),
                             **sheet_kw)
        if key in ['DataFrame', 'Title', 'CSpacer', 'RSpacer', 'HSpacer',
                   'Series', 'Image']:
            for k, v in template[key].items():
                if type(v) is str:
                    if v[:8] == '__eval__':
                        v_adj = v.replace('__eval__', '')
                        template[key][k] =  eval(_kwarg_parse(v_adj))
            return getattr(core, key)(**template[key])

def load(template, env, kwargs):
    if env:
        env.add_extension(EvalExtension)
    else:
        try:
            path = os.path.dirname(os.path.abspath(template))
            template = os.path.split(os.path.abspath(template))[-1]
            env = Environment(loader=FileSystemLoader(path))
            env.add_extension(EvalExtension)
            template = env.get_template(template).render(kwargs)
        except:
            env = Environment(loader=BaseLoader())
            env.add_extension(EvalExtension)
            template = env.from_string(template).render(kwargs)
    replace = [item.strip() for item in re.findall('[ :]{{.+}}', template)]
    for item in replace:
        template = template.replace(item, '\'' + item + '\'')
    return template


def load_yaml(template, env=None, str_only=False, **kwargs):
    """ Loads a YAML template specifying the structure of the XLCompose Object.

    Paramters
    ---------
    template: str (path-like)
        A string representing the path of a YAML file or a YAML string.
    env: jinja2.Environment (optional)
        The jinja2 environment to be used. If omitted, one will be created at
        at the location of the template.
    str_only: bool
        Whether to load the string representation of the template only.  When set
        to True (default), the `xlcompose` object will be constructed.
    """
    template = load(template, env, kwargs)
    if str_only:
        return template
    else:
        return _make_xlc(yaml.load(template, Loader=yaml.SafeLoader), **kwargs)

def load_json(template, env=None, **kwargs):
    """ Loads a JSON template specifying the structure of the XLCompose Object.
    """
    template = load(template, env, kwargs)
    return _make_xlc(json.loads(template), **kwargs)
