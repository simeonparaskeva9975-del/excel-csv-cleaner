"""Tests for shared Stage 1 data models."""

import unittest

from cleaner_app.models import DataIssue, DatasetProfile


class ModelTests(unittest.TestCase):
    def test_dataset_profile_rejects_negative_dimensions(self) -> None:
        with self.assertRaises(ValueError):
            DatasetProfile(source_name="bad.csv", row_count=-1, column_count=3)

    def test_data_issue_requires_positive_count(self) -> None:
        with self.assertRaises(ValueError):
            DataIssue(issue_type="missing", message="Invalid test", count=0)


if __name__ == "__main__":
    unittest.main()
