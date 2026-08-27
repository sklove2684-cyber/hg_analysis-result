import unittest

from honyu_app.services.database_service import DatabaseService


class DatabaseServiceContractTests(unittest.TestCase):
    def test_database_service_exposes_required_operations(self) -> None:
        required = {
            "check_connection",
            "check_duplicate",
            "save_analysis_batch",
            "replace_analysis_batch",
            "search_batches",
            "get_batch_detail",
            "add_peak_correction",
            "list_peak_corrections",
            "save_export_job",
        }
        self.assertLessEqual(required, set(DatabaseService.__dict__))
