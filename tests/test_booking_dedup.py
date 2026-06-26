from __future__ import annotations

import json
import sqlite3
import tempfile
import unittest
from pathlib import Path

import app


def booking_record(booking_id: str, revenue: int) -> dict[str, object]:
    return {
        "No": 36,
        "Booking ID": booking_id,
        "Venue": "BC Padel Club",
        "Court": "Pickleball Court No 1",
        "Customer Name": "sena",
        "Date of Booking": "2026-06-08",
        "Booking Period Start Time": "19:00:00",
        "Booking Period End Time": "20:00:00",
        "Payment Method": "MANUAL - QRIS",
        "Revenue Venue": revenue,
        "Total Booking Amount": revenue,
    }


class BookingDuplicateBookingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = app.DB_PATH
        app.DB_PATH = Path(self.temp_dir.name) / "rekap.db"

    def tearDown(self) -> None:
        app.DB_PATH = self.original_db_path
        self.temp_dir.cleanup()

    def test_duplicate_booking_id_rows_are_preserved(self) -> None:
        booking_id = "MN/2428/260607/0001862"
        first = booking_record(booking_id, 200_000)
        first["_source_row"] = 12
        second = booking_record(booking_id, 100_000)
        second["_source_row"] = 13
        app.save_records_to_db([first, second], "input.xlsx")

        records = app.load_records_from_db(None, None)
        self.assertEqual(2, len(records))
        self.assertEqual(300_000, sum(app.number_value(record["Revenue Venue"]) for record in records))
        self.assertEqual(1, len(app.find_duplicate_booking_ids(records)))

    def test_reupload_same_date_replaces_previous_snapshot(self) -> None:
        booking_id = "MN/2428/260607/0001862"
        app.save_records_to_db([booking_record(booking_id, 200_000)], "lama.xlsx")
        app.save_records_to_db([booking_record(booking_id, 100_000)], "koreksi.xlsx")

        records = app.load_records_from_db(None, None)
        self.assertEqual(1, len(records))
        self.assertEqual(100_000, records[0]["Revenue Venue"])

    def test_amount_is_part_of_row_identity(self) -> None:
        old = booking_record("MN/2428/260607/0001862", 200_000)
        corrected = booking_record("mn/2428/260607/0001862", 100_000)
        self.assertNotEqual(app.record_row_key(old), app.record_row_key(corrected))

    def test_existing_database_migration_preserves_duplicate_rows(self) -> None:
        booking_id = "MN/2428/260607/0001862"
        legacy_schema = app.BOOKINGS_SCHEMA.replace(
            "booking_id TEXT NOT NULL COLLATE NOCASE", "booking_id TEXT NOT NULL"
        )
        conn = sqlite3.connect(app.DB_PATH)
        try:
            with conn:
                conn.execute(legacy_schema)
                for row_key, revenue, source, uploaded_at in (
                    ("legacy-old", 200_000, "lama.xlsx", "2026-06-12T15:21:20"),
                    ("legacy-new", 100_000, "koreksi.xlsx", "2026-06-19T10:37:53"),
                ):
                    record = booking_record(booking_id, revenue)
                    conn.execute(
                        """
                        INSERT INTO bookings (
                            row_key, booking_id, booking_date, booking_year, booking_month,
                            booking_day, court, start_time, payment_method, channel, revenue,
                            fields_json, source_filename, uploaded_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            row_key,
                            booking_id,
                            "2026-06-08",
                            2026,
                            6,
                            8,
                            "Pickleball Court No 1",
                            "19:00:00",
                            "MANUAL - QRIS",
                            "Walk In",
                            revenue,
                            json.dumps(record),
                            source,
                            uploaded_at,
                        ),
                    )
        finally:
            conn.close()

        app.init_db()
        records = app.load_records_from_db(None, None)
        self.assertEqual(2, len(records))
        self.assertEqual(300_000, sum(app.number_value(record["Revenue Venue"]) for record in records))

    def test_existing_unique_booking_id_database_is_migrated(self) -> None:
        booking_id = "MN/2428/260607/0001862"
        legacy_schema = app.BOOKINGS_SCHEMA.replace(
            "booking_id TEXT NOT NULL COLLATE NOCASE", "booking_id TEXT NOT NULL COLLATE NOCASE UNIQUE"
        )
        conn = sqlite3.connect(app.DB_PATH)
        try:
            with conn:
                conn.execute(legacy_schema)
                record = booking_record(booking_id, 200_000)
                conn.execute(
                    """
                    INSERT INTO bookings (
                        row_key, booking_id, booking_date, booking_year, booking_month,
                        booking_day, court, start_time, payment_method, channel, revenue,
                        fields_json, source_filename, uploaded_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        "legacy-only",
                        booking_id,
                        "2026-06-08",
                        2026,
                        6,
                        8,
                        "Pickleball Court No 1",
                        "19:00:00",
                        "MANUAL - QRIS",
                        "Walk In",
                        200_000,
                        json.dumps(record),
                        "lama.xlsx",
                        "2026-06-12T15:21:20",
                    ),
                )
        finally:
            conn.close()

        app.init_db()
        first = booking_record(booking_id, 200_000)
        first["_source_row"] = 12
        second = booking_record(booking_id, 100_000)
        second["_source_row"] = 13
        app.save_records_to_db([first, second], "baru.xlsx")

        records = app.load_records_from_db(None, None)
        self.assertEqual(2, len(records))
        self.assertEqual(300_000, sum(app.number_value(record["Revenue Venue"]) for record in records))


if __name__ == "__main__":
    unittest.main()
