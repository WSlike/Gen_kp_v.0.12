from xml.dom import minidom


def create_tag(name: str = None, text: str = None, attributes: dict = None, *, cdata: bool = False):
    doc = minidom.Document()

    if name is None:
        return doc

    tag = doc.createElement(name)

    if text is not None:
        if cdata is True:
            tag.appendChild(doc.createCDATASection(text))
        else:
            tag.appendChild(doc.createTextNode(text))

    if attributes is not None:
        for k, v in attributes.items():
            tag.setAttribute(k, str(v))

    return tag


def create_var(root: minidom.Document, var_name: str = None, var_type: str = None, var_simple_value: str = None):
    tag_variable = create_tag('variable', attributes={'name': str(var_name)})
    tag_type = create_tag('type')
    tag_type_init = create_tag(str(var_type))
    tag_initial_value = create_tag('initialValue')
    tag_simple_value = create_tag('simpleValue', attributes={'value': str(var_simple_value)})
    root.appendChild(tag_variable)
    tag_variable.appendChild(tag_type)
    tag_type.appendChild(tag_type_init)

    if var_simple_value is not None:
        tag_variable.appendChild(tag_initial_value)
        tag_initial_value.appendChild(tag_simple_value)