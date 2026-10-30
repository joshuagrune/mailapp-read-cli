"""Regression checks for unavailable automation and incomplete Sent metadata."""
import io
import subprocess
import unittest
from argparse import Namespace
from contextlib import redirect_stderr, redirect_stdout
from unittest.mock import patch

import apple_mail as cli


class AutomationRegressionTests(unittest.TestCase):
    def message(self):
        return cli.MessageRow(42, "2026-09-30", True, "Invoice", "Sent", "Business", attachments=[])

    def test_sandbox_syntax_error_is_reported_as_unavailable_bridge(self):
        failure = subprocess.CompletedProcess([], 1, "", "com.apple.hiservices-xpcservice Connection Invalid\nsyntax error (-2741)")
        with patch.object(cli.subprocess, "run", return_value=failure):
            with self.assertRaisesRegex(cli.MailAutomationUnavailable, "sandbox"):
                cli.run_osascript("read-only")

    def test_empty_index_is_not_presented_as_verified_absence(self):
        output = io.StringIO()
        with patch.object(cli, "fetch_message_meta", return_value=self.message()), redirect_stdout(output):
            cli.cmd_attachments(Namespace(id=42, live=False, json=False))
        self.assertIn("does not prove absence", output.getvalue())

    def test_live_attachment_read_can_recover_missing_index_metadata(self):
        output = io.StringIO()
        with patch.object(cli, "fetch_message_meta", return_value=self.message()), patch.object(cli, "run_osascript", return_value="MIME-Version: 1.0\nContent-Type: multipart/mixed; boundary=parts\n\n--parts\nContent-Type: application/pdf\nContent-Disposition: attachment; filename=invoice.pdf\n\ncontent\n--parts\nContent-Type: application/pdf\nContent-Disposition: attachment; filename=receipt.pdf\n\ncontent\n--parts--\n"), redirect_stdout(output):
            cli.cmd_attachments(Namespace(id=42, live=True, json=True))
        import json
        result = json.loads(output.getvalue())
        self.assertEqual(result["attachments"], ["invoice.pdf", "receipt.pdf"])
        self.assertEqual(result["source"], "mailapp_mime")

    def test_unavailable_sender_bridge_warns_once_without_per_message_retries(self):
        errors = io.StringIO()
        with patch.object(cli, "run_osascript", side_effect=cli.MailAutomationUnavailable("sandbox")), patch.object(cli, "fetch_sender") as fallback, redirect_stderr(errors):
            cli.enrich_senders([self.message(), self.message()])
        fallback.assert_not_called()
        self.assertEqual(errors.getvalue().count("Warning:"), 1)


if __name__ == "__main__":
    unittest.main()
