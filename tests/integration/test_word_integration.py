from pathlib import Path

import pytest


pytestmark = pytest.mark.integration


@pytest.fixture(scope="session")
def word_client(real_msword_cli):
    client = real_msword_cli.WordClient(visible=False)
    try:
        yield client
    finally:
        for document in client.documents:
            try:
                document.close(force=True)
            except real_msword_cli.WordAPIError:
                pass
        client.quit()


@pytest.fixture
def temp_workspace(tmp_path):
    workspace = tmp_path / 'word-live'
    workspace.mkdir()
    return workspace


def test_word_client_can_create_save_and_reopen_document(word_client, temp_workspace):
    doc_path = temp_workspace / 'roundtrip.docx'

    doc = word_client.new(visible=False)
    doc.save(path=str(doc_path))
    doc.close(force=True)

    reopened = word_client.open(str(doc_path), visible=False)
    assert reopened.name.lower() == doc_path.name.lower()
    reopened.close(force=True)

    assert doc_path.exists()


def test_word_client_can_export_pdf(word_client, temp_workspace):
    doc_path = temp_workspace / 'export-source.docx'
    pdf_path = temp_workspace / 'exported.pdf'

    doc = word_client.new(visible=False)
    doc.save(path=str(doc_path))
    exported = doc.export_fixed_format(str(pdf_path))
    doc.close(force=True)

    assert Path(exported) == pdf_path.resolve()
    assert pdf_path.exists()


def test_word_client_can_compare_two_documents(word_client, temp_workspace):
    original_path = temp_workspace / 'original.docx'
    revised_path = temp_workspace / 'revised.docx'

    original = word_client.new(visible=False)
    original.save(path=str(original_path))
    original.close(force=True)

    revised = word_client.new(visible=False)
    revised.save(path=str(revised_path))
    revised.close(force=True)

    diff = word_client.compare(str(original_path), str(revised_path))
    assert diff.name
    diff.close(force=True)

