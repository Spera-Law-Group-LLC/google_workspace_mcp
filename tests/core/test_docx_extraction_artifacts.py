# -*- coding: utf-8 -*-
"""RED tests for librechat-mcp#8 DOCX extraction artifacts.

All DOCX-like inputs are synthetic ZIP packages built in memory. No client or
production Office document is used or committed.
"""
import io
import zipfile
from dataclasses import dataclass

import pytest

from core import utils

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
X_NS = "urn:synthetic-custom-namespace"


def _docx(document_xml: str) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml.encode("utf-8"))
    return out.getvalue()


def _body(*paragraphs: str) -> str:
    return (
        f'<w:document xmlns:w="{W_NS}" xmlns:a="{A_NS}" xmlns:x="{X_NS}">'
        "<w:body>"
        + "".join(paragraphs)
        + "</w:body></w:document>"
    )


def _p(*inner: str) -> str:
    return "<w:p>" + "".join(inner) + "</w:p>"


def _r_text(text: str) -> str:
    return f"<w:r><w:t>{text}</w:t></w:r>"


def _r_preserve(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def _content(result: str) -> str:
    marker = "--- CONTENT ---"
    return result.split(marker, 1)[1].strip() if marker in result else result.strip()


def test_docx_run_boundaries_do_not_insert_spaces_or_strip_punctuation():
    data = _docx(_body(
        _p(_r_text("I"), _r_text("f")),
        _p(_r_text("th"), _r_text("ese")),
        _p(_r_text("rules"), _r_text(".")),
        _p(_r_text("90/25"), _r_text(".")),
        _p(_r_text("("), _r_text("Business"), _r_text(")")),
        _p(_r_text("ordinary"), _r_preserve(" "), _r_text("spacing")),
    ))

    result = utils.extract_office_xml_text(data, DOCX_MIME)

    assert _content(result) == "If\nthese\nrules.\n90/25.\n(Business)\nordinary spacing"


def test_docx_tabs_breaks_and_hyperlink_text_are_reconstructed_in_order():
    data = _docx(_body(
        _p(_r_text("Before"), "<w:r><w:tab/></w:r>", _r_text("After")),
        _p(_r_text("Line"), "<w:r><w:br/></w:r>", _r_text("Break")),
        _p('<w:hyperlink w:anchor="Synthetic"><w:r><w:t>Linked</w:t></w:r></w:hyperlink>', _r_text(" text")),
    ))

    result = utils.extract_office_xml_text(data, DOCX_MIME)

    assert _content(result) == "Before\tAfter\nLine\nBreak\nLinked text"


def test_docx_metadata_warns_that_extraction_is_untrusted_and_not_rendered_proof():
    data = _docx(_body(_p(_r_text("Visible content"))))

    result = utils.extract_office_xml_text(data, DOCX_MIME)

    assert result.startswith("--- EXTRACTION METADATA ---")
    assert "representation: docx_xml_extracted_text" in result
    assert "content_source: untrusted_external_document" in result
    assert "injection_risk: high" in result
    assert "proofreading_fidelity: unverified" in result
    assert "tracked_changes_view: proposed_final" in result
    assert "--- CONTENT ---" in result


def test_docx_word_extractor_uses_word_namespace_only():
    data = _docx(_body(
        _p(
            _r_text("Word text"),
            "<a:t>Drawing text must not appear</a:t>",
            "<x:t>Custom XML text must not appear</x:t>",
        )
    ))

    result = utils.extract_office_xml_text(data, DOCX_MIME)

    assert _content(result) == "Word text"
    assert "Drawing text" not in result
    assert "Custom XML" not in result


def test_docx_tracked_changes_policy_is_proposed_final_and_labeled():
    data = _docx(_body(_p(
        _r_text("Current "),
        '<w:ins w:id="1" w:author="Synthetic"><w:r><w:t>Inserted</w:t></w:r></w:ins>',
        '<w:del w:id="2" w:author="Synthetic"><w:r><w:delText>Deleted</w:delText></w:r></w:del>',
    )))

    result = utils.extract_office_xml_text(data, DOCX_MIME)

    assert "tracked_changes_view: proposed_final" in result
    assert _content(result) == "Current Inserted"
    assert "Deleted" not in result


@dataclass
class _FakeInfo:
    file_size: int
    compress_size: int


class _FakeZipFile:
    def __init__(self, _stream):
        self.read_called = False

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def namelist(self):
        return ["word/document.xml"]

    def getinfo(self, member):
        assert member == "word/document.xml"
        return _FakeInfo(file_size=51 * 1024 * 1024, compress_size=1024)

    def read(self, member):
        self.read_called = True
        raise AssertionError("oversized Office XML member was read before safety check")


def test_docx_zip_member_size_is_checked_before_read(monkeypatch):
    monkeypatch.setattr(utils.zipfile, "ZipFile", _FakeZipFile)

    result = utils.extract_office_xml_text(b"synthetic zip bytes", DOCX_MIME)

    assert result is None or "skipped" in result.lower() or "too large" in result.lower()
