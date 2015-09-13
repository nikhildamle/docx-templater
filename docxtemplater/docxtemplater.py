import os
import re
from zipfile import ZipFile
from xml.dom import minidom
from xml.dom.minidom import Document, Element

import pystache

from docxtemplater.xml_helper import remove_whitespace_only_nodes
from docxtemplater.docx_helper import run_contains_text, get_text_from_paragraph_runs
from docxtemplater.docx_helper import _namespaces, get_run_text, set_run_text


def write(template_docx_path: str, replacements: dict, output_docx_path: str):
    """Replace data into docx template file

    :param template_docx_path: Path to the docx template file
    :param replacements: Data to replace with placeholders
    :param output_docx_path: Path to the rendered docx file
    """
    dom = minidom.parseString(_read_docx(template_docx_path))
    remove_whitespace_only_nodes(dom)
    _prepare_mustache_template(dom)
    rendered_document_xml = pystache.render(dom.toxml(encoding='utf-8'), replacements)
    _write_docx(template_docx_path, rendered_document_xml, output_docx_path)

MUSTACHE_LOOP_REGEX = '\{\{\s*[#/]{1}\s*\w+\s*\}\}'
MUSTACHE_PLACEHOLDER_REGEX = '\{\{\s*[#/]{0,1}\s*\w+\s*\}\}'


def _read_docx(file_path: str) -> str:
    """Read word/document.xml from docx file

    :param file_path: Path to the template docx file
    """
    if os.path.exists(file_path):
        with ZipFile(file_path, mode='r') as docx:
            return docx.read('word/document.xml')
    else:
        raise FileNotFoundError('docx file not found')


def _write_docx(template_docx_path: str, rendered_document_xml: str, output_docx_path: str):
    """Write rendered output docx

    :param template_docx_path: Path to the template docx file
    :param rendered_document_xml: Rendered word/document.xml
    :param output_docx_path: Path of the location to save rendered docx
    """
    template_docx = ZipFile(template_docx_path, mode='r')
    output_docx   = ZipFile(output_docx_path,   mode='w')

    # Copy every other file except word/document.xml into output docx (zip) file
    # For word/document.xml replace it with rendered template.
    for item in template_docx.infolist():
        buffer = template_docx.read(item.filename)
        if item.filename != 'word/document.xml':
            output_docx.writestr(item, buffer)
        else:
            output_docx.writestr('word/document.xml', rendered_document_xml)

    template_docx.close()
    output_docx.close()


def _prepare_mustache_template(document: Document):
    """Prepares mustache compatible template from OpenXML template file

    :param document: Document node of document.xml file

    If the paragraph contains.
    """
    for paragraph in document.getElementsByTagNameNS(_namespaces['w'], 'p'):
        _merge_placeholder_broken_inside_runs_if_required(paragraph)

    for paragraph in document.getElementsByTagNameNS(_namespaces['w'], 'p'):
        text = get_text_from_paragraph_runs(paragraph)
        if re.search(MUSTACHE_LOOP_REGEX, text):
            _replace_loop_placeholder(paragraph, text)


def _replace_loop_placeholder(paragraph: Element, placeholder_text: str):
    """Replace paragraph containing mustache loop

    :param paragraph:
    :param placeholder_text:
    """
    parent_node = paragraph.parentNode
    if parent_node.tagName == 'w:body':
        # If placeholder is within body replace paragraph with xml comment
        # containing mustache loop
        comment = Document.createComment(parent_node, placeholder_text)
        parent_node.replaceChild(comment, paragraph)
    elif parent_node.tagName == 'w:tc':
        # If placeholder is within a table, replace row it resides with xml
        # comment containing mustache loop
        row = parent_node.parentNode
        table = row.parentNode
        comment = Document.createComment(table, placeholder_text)
        table.replaceChild(comment, row)


def _merge_placeholder_broken_inside_runs_if_required(paragraph: Element):
    """Merge broken runs containing mustache placeholders.

    :param paragraph: paragraph xml element containing broken runs

    docx document is made up of paragraph among other things. A run is a part
    of a paragraph with different formatting(color, bold...). But most times
    Microsoft word and libreoffice Writer splits up text with same formatting
    into different runs. If this text contains mustache placeholders, it will
    be missed by mustache renderer.
    This method merges runs into one if it contains mustache placeholders.
    """

    runs = paragraph.getElementsByTagNameNS(_namespaces['w'], 'r')

    def _merge(run: Element, text_to_replace='', open_brace_count=0, close_brace_count=0, runs_to_merge=None):
        """Merge placeholders broken into runs

        Microsoft Word and libreoffice most times split placeholders into
        multiple runs. For example
        <w:r>
            <w:rPr>
                <w:b w:val="false"/>
                <w:bCs w:val="false"/>
            </w:rPr>
            <w:t>{{PRODUCTS</w:t>
        </w:r>
        <w:r>
            <w:rPr/>
            <w:t>}}</w:t>
        </w:r>

        We need to merge this into one run while retaining the style
        """
        if runs_to_merge is None:
            runs_to_merge = []

        if run is None:
            return
        elif not run_contains_text(run):
            pass
        else:
            text = get_run_text(run)

            open_brace_count  += text.count('{{')
            close_brace_count += text.count('}}')

            text_to_replace += text

            # Once we have matching nodes, set text_to_replace as value to the
            # last run and remove previous runs
            if not open_brace_count == close_brace_count:
                runs_to_merge.append(run)
            elif runs_to_merge:
                set_run_text(run, text_to_replace)
                for r in runs_to_merge:
                    paragraph.removeChild(r)
                runs_to_merge = []
                text_to_replace = ''
        return _merge(run.nextSibling, text_to_replace, open_brace_count, close_brace_count, runs_to_merge)

    _merge(runs[0])


