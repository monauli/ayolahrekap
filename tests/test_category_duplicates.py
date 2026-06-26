from __future__ import annotations

import unittest

from app import duplicate_category_files_to_remove, find_duplicate_categories


def result(filename: str, category: str, first_date: str, last_date: str | None = None):
    return {
        "filename": filename,
        "_source_path": f"C:/data/{filename}",
        "categories": {category: 1},
        "first_date": first_date,
        "last_date": last_date or first_date,
    }


class CategoryDuplicateTests(unittest.TestCase):
    def test_same_category_and_month_is_duplicate(self) -> None:
        files = [
            result("topi-1.xlsx", "TOPI", "2026-05-01", "2026-05-31"),
            result("topi-2.xlsx", "Topi", "2026-05-01", "2026-05-31"),
        ]
        self.assertEqual(
            {"TOPI": ["topi-1.xlsx", "topi-2.xlsx"]},
            find_duplicate_categories(files),
        )

    def test_same_category_in_different_month_is_not_duplicate(self) -> None:
        files = [
            result("pickleball-mei.xlsx", "LAPANGAN PICKLEBALL", "2026-05-01", "2026-05-31"),
            result("pickleball-juni.xlsx", "LAPANGAN PICKLEBALL", "2026-06-01", "2026-06-30"),
        ]
        self.assertEqual({}, find_duplicate_categories(files))
        self.assertEqual([], duplicate_category_files_to_remove(files))

    def test_bulk_removal_keeps_first_file_for_each_duplicate(self) -> None:
        files = [
            result("bola-1.xlsx", "BOLA PADEL", "2026-05-01", "2026-05-31"),
            result("bola-2.xlsx", "BOLA PADEL", "2026-05-01", "2026-05-31"),
            result("bola-3.xlsx", "BOLA PADEL", "2026-05-01", "2026-05-31"),
        ]
        self.assertEqual(
            [
                ("C:/data/bola-2.xlsx", "bola-2.xlsx"),
                ("C:/data/bola-3.xlsx", "bola-3.xlsx"),
            ],
            duplicate_category_files_to_remove(files),
        )

    def test_bulk_removal_prioritizes_valid_raw_file(self) -> None:
        generated_summary = {
            "filename": "summary.xlsx",
            "_source_path": "C:/data/summary.xlsx",
            "categories": {"TOPI": 1, "GRIP": 1},
            "source_groups": {"TOPI": 1, "GRIP": 1},
            "first_date": "2026-05-01",
            "last_date": "2026-05-31",
        }
        raw_topi = result("topi.xlsx", "TOPI", "2026-05-01", "2026-05-31")
        raw_topi["source_groups"] = {"TOPI": 10}
        self.assertEqual(
            [("C:/data/summary.xlsx", "summary.xlsx")],
            duplicate_category_files_to_remove([generated_summary, raw_topi]),
        )


if __name__ == "__main__":
    unittest.main()
