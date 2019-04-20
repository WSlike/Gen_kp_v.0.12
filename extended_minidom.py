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