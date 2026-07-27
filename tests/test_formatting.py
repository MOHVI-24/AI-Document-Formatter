from backend.formatting import build_document, export_document, extract_upload


def test_offline_document_detects_meeting_notes_and_exports_pdf():
    source = "Agenda\n\nAttendees: Sam, Nia\n\nAction items\n- Send the draft by Friday"
    document = build_document(source, "auto", use_ai=False)
    content, content_type = export_document(document, "pdf", "#6d5dfc")
    assert document.kind == "meeting-notes"
    assert content.startswith(b"%PDF")
    assert content_type == "application/pdf"


def test_rejects_unsupported_upload():
    try:
        extract_upload("unsafe.exe", b"ignored")
    except ValueError as error:
        assert "Supported upload types" in str(error)
    else:
        raise AssertionError("unsupported upload should be rejected")
