import io
import os
import tempfile
import unittest
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import app as target


class FakeGigaChat:
    def reply(self, messages):
        return "Тестовый ответ"


class FakeCalendar:
    def __init__(self, busy=None):
        self.busy = busy or []
        self.saved = []

    def search(self, start, end, event=True, expand=True):
        return [True] if any(a < end and b > start for a, b in self.busy) else []

    def save_event(self, event):
        self.saved.append(event)


class TestBot(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        target.DB = target.Path(self.tmp.name)
        target.gigachat = FakeGigaChat()
        target.app.config["TESTING"] = True
        target.app.secret_key = "test"
        self.client = target.app.test_client()
        self.calendar = FakeCalendar()
        self.original_calendar = target.yandex_calendar
        target.yandex_calendar = lambda: self.calendar

    def tearDown(self):
        target.yandex_calendar = self.original_calendar
        self.tmp.close()
        try:
            os.unlink(self.tmp.name)
        except OSError:
            pass

    def test_parse_relative_and_named_dates(self):
        now = datetime(2026, 9, 17, 12, 0, tzinfo=ZoneInfo("Europe/Moscow"))
        self.assertEqual(target.parse_requested_slot("Завтра в 20:00", now).isoformat(), "2026-09-18T20:00:00+03:00")
        self.assertEqual(target.parse_requested_slot("19 сентября в 11:30", now).isoformat(), "2026-09-19T11:30:00+03:00")
        self.assertEqual(target.parse_requested_slot("В пятницу в 18:00", now).isoformat(), "2026-09-18T18:00:00+03:00")

    def test_chat_checks_slot_and_books_from_contacts(self):
        with self.client.session_transaction() as session:
            session["expert_slug"] = "marketer"
            session["consultation_offered"] = True
        future = datetime.now(ZoneInfo("Europe/Moscow")) + timedelta(days=60)
        request_text = future.strftime("%d.%m.%Y в 20:00")
        first = self.client.post("/api/chat", json={"message": request_text})
        self.assertEqual(first.status_code, 200)
        self.assertIn("свободно", first.json["answer"])
        second = self.client.post("/api/chat", json={"message": "Анна Иванова, +7 999 123-45-67, anna@example.com"})
        self.assertEqual(second.status_code, 200)
        self.assertIn("Запись подтверждена", second.json["answer"])
        self.assertEqual(len(self.calendar.saved), 1)
        self.assertIn("Анна Иванова", self.calendar.saved[0])

    def test_busy_slot_offers_real_alternatives(self):
        tz = ZoneInfo("Europe/Moscow")
        busy_start = (datetime.now(tz) + timedelta(days=60)).replace(hour=20, minute=0, second=0, microsecond=0)
        self.calendar.busy = [(busy_start, busy_start + timedelta(hours=1))]
        with self.client.session_transaction() as session:
            session["expert_slug"] = "marketer"
        result = self.client.post("/api/chat", json={"message": busy_start.strftime("%d.%m.%Y в 20:00")})
        self.assertEqual(result.status_code, 200)
        self.assertIn("уже занято", result.json["answer"])
        self.assertIn((busy_start + timedelta(hours=1)).strftime("%d.%m в %H:%M"), result.json["answer"])
        chosen = self.client.post("/api/chat", json={"message": "Давайте в 21:00"})
        self.assertIn("свободно", chosen.json["answer"])
        self.assertIn("имя, телефон и email", chosen.json["answer"])

    def test_upload_and_chat(self):
        headers = {"X-Admin-Password": "admin123"}
        uploaded = self.client.post(
            "/api/admin/upload",
            headers=headers,
            data={"files": (io.BytesIO("Эксперт проводит консультации.".encode()), "base.txt")},
            content_type="multipart/form-data",
        )
        self.assertEqual(uploaded.status_code, 200)
        answer = self.client.post("/api/chat", json={"message": "Как проходит консультация?"})
        self.assertEqual(answer.json["answer"], "Тестовый ответ")


if __name__ == "__main__":
    unittest.main()
