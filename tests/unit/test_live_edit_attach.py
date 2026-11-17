"""CI-safe unit tests for the live-edit attach path — zero COM dependencies.

Covers the ROT multi-instance fix (docs/pitfalls.md #17): when a stale
zero-document engine instance is registered in the Running Object Table,
``Dispatch("Word.Application")`` returns the STALE one, and the user's
visible document becomes unreachable. The attach helper must prefer an
instance that actually holds documents.
"""

from docuhand.engine import live_edit


class _FakeApp:
    """Minimal stand-in for a Word/WPS Application object."""

    def __init__(self, doc_count: int, name: str = "Microsoft Word"):
        self._docs = _FakeDocs(doc_count)
        self.Name = name

    @property
    def Documents(self):
        return self._docs


class _FakeDocs:
    def __init__(self, count: int):
        self.Count = count


class _FakeMoniker:
    pass


class TestAppFromMoniker:
    def test_returns_none_when_binding_raises(self, monkeypatch):
        class _Boom:
            def GetObject(self, moniker):
                raise RuntimeError("no binding")

        monkeypatch.setattr(live_edit, "pythoncom", _Boom(), raising=False)
        assert live_edit._app_from_moniker(_Boom(), _FakeMoniker()) is None

    def test_returns_none_when_object_has_no_documents_collection(self, monkeypatch):
        class _NoDocs:
            pass

        class _Unk:
            def QueryInterface(self, iid):
                return _NoDocs()

        class _Rot:
            def GetObject(self, moniker):
                return _Unk()

        import win32com.client

        monkeypatch.setattr(
            win32com.client, "Dispatch", lambda disp: _NoDocs(), raising=False
        )
        assert live_edit._app_from_moniker(_Rot(), _FakeMoniker()) is None

    def test_returns_app_when_documents_collection_exists(self, monkeypatch):
        app = _FakeApp(2)

        class _Unk:
            def QueryInterface(self, iid):
                return object()

        class _Rot:
            def GetObject(self, moniker):
                return _Unk()

        import win32com.client

        monkeypatch.setattr(
            win32com.client, "Dispatch", lambda disp: app, raising=False
        )
        assert live_edit._app_from_moniker(_Rot(), _FakeMoniker()) is app


class TestAttachRunningApp:
    def test_prefers_instance_that_has_documents(self, monkeypatch):
        stale = _FakeApp(0)
        live = _FakeApp(3)
        monkeypatch.setattr(live_edit, "_rot_word_apps", lambda: [stale, live])
        assert live_edit._attach_running_app() is live

    def test_falls_back_to_first_candidate_when_none_has_documents(self, monkeypatch):
        stale = _FakeApp(0)
        monkeypatch.setattr(live_edit, "_rot_word_apps", lambda: [stale])

        def _boom(*a, **k):
            raise RuntimeError("no ProgID fallback")

        import win32com.client

        monkeypatch.setattr(win32com.client, "Dispatch", _boom, raising=False)
        # The ProgID fallback raises DocuhandError when nothing is running;
        # but with a candidate present we never reach it.
        assert live_edit._attach_running_app() is stale

    def test_structured_error_when_nothing_is_running(self, monkeypatch):
        monkeypatch.setattr(live_edit, "_rot_word_apps", lambda: [])

        def _boom(*a, **k):
            raise RuntimeError("no instance")

        import win32com.client

        monkeypatch.setattr(win32com.client, "Dispatch", _boom, raising=False)
        with pytest.raises(live_edit.DocuhandError) as exc:
            live_edit._attach_running_app()
        assert "No running Word/WPS instance" in str(exc.value)


import pytest  # noqa: E402  (placed after the helpers it decorates)
