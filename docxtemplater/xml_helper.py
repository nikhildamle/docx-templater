from xml.dom.minidom import Node, Element


def get_text(node: Element):
    text = None
    for child in node.childNodes:
        if child.nodeType == Node.TEXT_NODE:
            if text is None: text = ''
            text += child.nodeValue
    return text


def set_text(node: Element, text):
    dom = node.ownerDocument

    for child in node.childNodes:
        if child.nodeType == Node.TEXT_NODE:
            node.removeChild(child)

    text_node = dom.createTextNode(text)
    text_node.nodeValue = text
    node.appendChild(text_node)


def remove_whitespace_only_nodes(node: Node):
    """Removes all of the whitespace-only text descendants of a DOM node.

    :param node: Node to cleanup

    If the specified node is a whitespace-only text node then it is left
    unmodified.
    """
    remove_list = []
    for child in node.childNodes:
        # Below if returns true for TEXT_NODE which contains no characters
        # other than spaces and old lines
        if child.nodeType == Node.TEXT_NODE and not child.data.strip():
            remove_list.append(child)
        elif child.hasChildNodes():
            remove_whitespace_only_nodes(child)

    for node in remove_list:
        node.parentNode.removeChild(node)
        node.unlink()  # Garbage collect unneeded Nodes
