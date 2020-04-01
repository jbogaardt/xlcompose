import yaml
import json
from jinja2 import Template
import re
import xlcompose.core as core

def make_xlc(template, **kwargs):
    """ Recursively generate xlcompose object"""
    if type(template) is list:
        tabs = [make_xlc(element, **kwargs) for element in template]
        return core.Tabs(*[(item.name, item) for item in tabs])
    key = list(template.keys())[0]
    if key in ['Row', 'Column']:
        return getattr(core, key)(*[make_xlc(element, **kwargs)
                                    for element in template[key]])
    if key in ['Sheet']:
        return core.Sheet(template[key]['name'],
                          make_xlc(template[key]['data'], **kwargs))
    if key in ['DataFrame', 'Title', 'CSpacer', 'RSpacer', 'HSpacer',
               'Series', 'Image']:
        for k, v in template[key].items():
            if type(v) is str:
                if v.replace('{{', '').replace('}}', '') in kwargs.keys():
                    template[key][k] = \
                        kwargs[v.replace('{{', '').replace('}}', '')]
        return getattr(core, key)(**template[key])

def load(template, primitives):
    with open(template) as f:
        template = Template(f.read()).render(primitives)
        replace = [item.strip() for item in re.findall('[ :]{{.+}}', template)]
        for item in replace:
            template = template.replace(item, '\'' + item + '\'')
        return template

def parse_kwargs(**kwargs):
    primitives = {k: v if type(v) in [str, int, float] else '{{' + k + '}}'
                  for k, v in kwargs.items()}
    objects = {k: v for k, v in kwargs.items()
               if type(v) not in [str, int, float]}
    return primitives, objects


def load_yaml(template, **kwargs):
    """ Loads a YAML template specifying the structure of the XLCompose Object.
    """
    primitives, objects = parse_kwargs(**kwargs)
    template = load(template, primitives)
    return make_xlc(yaml.load(template, Loader=yaml.SafeLoader), **objects)

def load_json(template, **kwargs):
    """ Loads a JSON template specifying the structure of the XLCompose Object.
    """
    primitives, objects = parse_kwargs(**kwargs)
    template = load(template, primitives)
    return make_xlc(json.loads(template), **objects)
