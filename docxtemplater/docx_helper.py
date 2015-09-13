from xml.dom.minidom import Element

from docxtemplater.xml_helper import get_text, set_text

_namespaces = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}


def run_contains_text(run) -> bool:
    """Check if run element contains text element

    :param run: Run
    :return: True if run contains text element
    """
    text_element = run.getElementsByTagNameNS(_namespaces['w'], 't')
    if len(text_element) == 1:
        return True
    else:
        return False


def get_text_from_paragraph_runs(node: Element, list_of_runs_indexes: list = None) -> str:
    """Get all the text from the paragraph

    :param node: Paragraph to extract text from
    :param list_of_runs_indexes: List of runs indexes to extract text from
    :return: String containing all text inside paragraph

    Iterates through all runs inside paragraph appends text and returns it.
    """
    text = ''
    runs = node.getElementsByTagNameNS(_namespaces['w'], 'r')

    if list_of_runs_indexes is not None:
        runs = list(runs[i] for i in list_of_runs_indexes)

    for run in runs:
        if run_contains_text(run):
            text_element = run.getElementsByTagNameNS(_namespaces['w'], 't')
            # Each run will contains only one text element. So select first
            # (0th index) element
            text += text_element[0].firstChild.nodeValue
    return text


def get_run_text(run: Element):
    if run_contains_text(run):
        text_element = run.getElementsByTagNameNS(_namespaces['w'], 't')[0]
        return get_text(text_element)
    else:
        return None


def set_run_text(run: Element, text: str):
    dom = run.ownerDocument

    if run_contains_text(run):
        text_element = run.getElementsByTagNameNS(_namespaces['w'], 't')[0]
        set_text(text_element, text)
    else:
        text_element = dom.createElementNS(_namespaces['w'], 'w:t')
        set_text(text_element, text)
        run.appendChild(text_element)


def clear_run_text(run: Element):
    if run_contains_text(run):
        text_element = run.getElementsByTagNameNS(_namespaces['w'], 't')[0]
        set_text(text_element, '')
