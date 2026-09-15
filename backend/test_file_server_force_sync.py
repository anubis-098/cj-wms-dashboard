import asyncio
import importlib.util
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, main
from unittest.mock import AsyncMock, MagicMock, patch

import main as backend


class ForcedSyncTests(TestCase):
    def test_mirror_force_reloads_even_when_size_and_timestamp_match(self):
        spec = importlib.util.spec_from_file_location(
            "smd_mirror", Path(__file__).resolve().parents[1] / "scripts/smd-mirror-service.py"
        )
        mirror = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mirror)
        with TemporaryDirectory() as folder:
            source, destination = Path(folder) / "source", Path(folder) / "mirror"
            source.mkdir()
            workbook = source / "latest.xlsx"
            workbook.write_bytes(b"old")
            mirror.mirror_latest(source, destination)
            original = workbook.stat()
            workbook.write_bytes(b"new")
            os.utime(workbook, ns=(original.st_atime_ns, original.st_mtime_ns))
            self.assertFalse(mirror.mirror_latest(source, destination)["changed"])
            self.assertEqual((destination / workbook.name).read_bytes(), b"old")
            self.assertTrue(mirror.mirror_latest(source, destination, force=True)["changed"])
            self.assertEqual((destination / workbook.name).read_bytes(), b"new")

    def test_manual_endpoint_forces_sync(self):
        with patch.object(backend, "run_file_server_sync", new_callable=AsyncMock) as sync:
            sync.return_value = {"changed": True}
            asyncio.run(backend.sync_file_server_now())
            sync.assert_awaited_once_with(force=True)

    def test_force_reimports_existing_upload_and_invalidates_cache(self):
        with TemporaryDirectory() as folder:
            source = Path(folder) / "source.xlsx"
            source.write_bytes(b"workbook")
            stat = source.stat()
            state = {"source_path": str(source), "source_mtime_ns": stat.st_mtime_ns,
                     "source_size": stat.st_size, "upload_id": "existing-id"}
            upload = MagicMock(stored_filename="existing-id_source.xlsx")
            session = MagicMock()
            session.__enter__.return_value.get.return_value = upload
            with patch.multiple(backend, FILE_SERVER_SYNC_ENABLED=True, UPLOAD_ROOT=Path(folder),
                                file_server_sync_status={}, cached_dashboard={}), \
                 patch.object(backend, "find_latest_file_server_excel", return_value=source), \
                 patch.object(backend, "read_setting", return_value=state), \
                 patch.object(backend, "SessionLocal", return_value=session), \
                 patch.object(backend, "select_default_workbook_sheet", return_value="Sheet1"), \
                 patch.object(backend, "update_workspace_widgets_sheet", return_value=1), \
                 patch.object(backend, "process_excel_file", return_value={}) as parse, \
                 patch.object(backend, "write_setting") as save, \
                 patch.object(backend, "invalidate_excel_cache") as invalidate, \
                 patch.object(backend, "record_audit_event"):
                self.assertFalse(backend.sync_latest_file_server_excel()["changed"])
                parse.assert_not_called()
                result = backend.sync_latest_file_server_excel(force=True)
                self.assertTrue(result["changed"])
                self.assertEqual(result["upload_id"], "existing-id")
                parse.assert_called_once()
                invalidate.assert_called_once_with("existing-id")
                save.assert_called_once()
                self.assertEqual(upload.file_data, b"workbook")
                self.assertIsNotNone(backend.file_server_sync_status["last_synced_at"])

    def test_force_reaches_mirror_and_broadcasts_refresh(self):
        with patch.object(backend, "FILE_SERVER_MIRROR_URL", "http://mirror/sync"), \
             patch.object(backend, "urlopen") as request, \
             patch.object(backend, "sync_latest_file_server_excel", return_value={
                 "changed": True, "upload_id": "existing-id", "updated_widgets": 1
             }) as sync, patch.object(backend, "broadcast_workspace_event") as broadcast:
            request.return_value.__enter__.return_value.status = 200
            asyncio.run(backend.run_file_server_sync(force=True))
            self.assertEqual(request.call_args.args[0].get_header("X-sync-force"), "true")
            sync.assert_called_once_with(force=True)
            broadcast.assert_any_call("excel-upload-replaced", "existing-id")


if __name__ == "__main__":
    main()
