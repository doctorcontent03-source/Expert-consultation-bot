import io
import os
import tempfile
import unittest
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import app as target


class FakeCalendar:
    def __init__(self, busy=None):
        self.busy = busy or []
        self.saved = []

    def search(self, start, end, event=True, expand=True):
        return [True] if any(a < end and b > start for a, b in self.busy) else []

    def save_event(self, event):
        self.saved.append(event)


class ScriptedGigaChat:
    def __init__(self, replies):
        self.replies = list(replies)
        self.prompts = []

    def reply(self, messages, model=None):
        self.prompts.append(messages[0]["content"])
        if not self.replies:
            raise RuntimeError("No scripted reply")
        value = self.replies.pop(0)
        if isinstance(value, Exception):
            raise value
        return value


def payload(reply, action="explore", intent="continue", evidence="", observations=None):
    observations = observations or {}
    fields = {}
    for key in ("contact", "need", "previous_experience", "desired_result"):
        fields[key] = observations.get(key, {"present": False, "evidence": ""})
    import json
    return json.dumps({
        "reply": reply,
        "action": action,
        "intent": intent,
        "intent_evidence": evidence,
        "observations": fields,
        "reply_assessment": {
            "based_on_client_meaning": True,
            "treats_message_as_feedback_to_expert": False,
            "asks_only_missing_information": True,
            "repeats_known_information": False,
            "performs_expert_work": False,
            "pressures_client": False,
            "offers_consultation": action == "offer_consultation",
            "uses_generic_self_promotion": False,
            "question_count": reply.count("?"),
        },
    }, ensure_ascii=False)


class TestBot(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        target.DB = target.Path(self.tmp.name)
        target.app.config["TESTING"] = True
        target.app.secret_key = "test"
        self.client = target.app.test_client()
        self.calendar = FakeCalendar()
        self.original_calendar = target.yandex_calendar
        self.original_gigachat = target.gigachat
        target.yandex_calendar = lambda: self.calendar
        con = target.db()
        con.execute(
            "insert into documents(id,name,text,created_at,expert_slug) values(?,?,?,?,?)",
            ("kb", "О Кирилле.txt", "Кирилл — психолог. Консультации проходят онлайн.", 1, "psychologist"),
        )
        con.commit()

    def tearDown(self):
        target.yandex_calendar = self.original_calendar
        target.gigachat = self.original_gigachat
        self.tmp.close()
        try:
            os.unlink(self.tmp.name)
        except OSError:
            pass

    def state(self, **changes):
        base = {
            "contact": False,
            "need": False,
            "previous_experience": False,
            "desired_result": False,
            "solution_explained": False,
            "interest_confirmed": False,
            "consultation_offered": False,
            "diagnostic_questions": 0,
            "asked_questions": [],
        }
        base.update(changes)
        return base

    # Universal sequence and grounded state
    def test_grounded_observations_only_accept_client_quote(self):
        state = self.state()
        updated = target.apply_grounded_observations(state, {
            "contact": {"present": True, "evidence": "мне 42"},
            "need": {"present": True, "evidence": "клиент потерял смысл"},
        }, "Мне 42, и я ни во что не верю")
        self.assertTrue(updated["contact"])
        self.assertFalse(updated["need"])

    def test_two_questions_required_before_consultation_invitation(self):
        state = self.state(contact=True, need=True, previous_experience=True, desired_result=True)
        self.assertEqual(target.expected_dialog_action(state, "continue", "", "Года два"), "explore")
        state["diagnostic_questions"] = 2
        self.assertEqual(target.expected_dialog_action(state, "continue", "", "Года два"), "offer_consultation")

    def test_missing_information_allows_only_three_questions(self):
        state = self.state(contact=True, need=True, diagnostic_questions=2)
        self.assertEqual(target.expected_dialog_action(state, "continue", "", "Не знаю"), "explore")
        state["diagnostic_questions"] = 3
        self.assertEqual(target.expected_dialog_action(state, "continue", "", "Не знаю"), "offer_consultation")

    def test_direct_question_is_answered_before_progression(self):
        state = self.state(contact=True, need=True, diagnostic_questions=1)
        text = "Сколько длится консультация?"
        self.assertEqual(target.expected_dialog_action(state, "question", text, text), "answer_information")

    def test_interest_cannot_trigger_offer_before_solution(self):
        text = "Да, мне интересно"
        state = self.state(contact=True, need=True, previous_experience=True, desired_result=True, diagnostic_questions=1)
        self.assertEqual(target.expected_dialog_action(state, "interest", text, text), "explore")

    def test_interest_after_solution_triggers_one_offer(self):
        text = "Да, мне интересно"
        state = self.state(solution_explained=True)
        self.assertEqual(target.expected_dialog_action(state, "interest", text, text), "offer_consultation")
        state["consultation_offered"] = True
        self.assertEqual(target.expected_dialog_action(state, "interest", text, text), "check_interest")

    def test_neutral_possible_is_not_interest_or_obstacle(self):
        state = self.state(solution_explained=True)
        self.assertEqual(target.expected_dialog_action(state, "continue", "", "Возможно"), "check_interest")

    def test_explicit_boundary_is_respected(self):
        text = "Я сейчас не хочу это обсуждать"
        self.assertEqual(target.expected_dialog_action(self.state(), "boundary", text, text), "respect_boundary")

    def test_semantic_correction_is_repaired_independent_of_wording(self):
        for text in ("Вы поняли меня не так", "Я говорила совсем о другом", "Это неверный вывод"):
            self.assertEqual(target.expected_dialog_action(self.state(), "correction", text, text), "repair_interpretation")

    def test_semantic_end_requires_grounded_evidence(self):
        self.assertEqual(target.expected_dialog_action(self.state(), "end", "другая цитата", "Да ерунда какая-то"), "explore")
        text = "На этом закончим"
        self.assertEqual(target.expected_dialog_action(self.state(), "end", text, text), "end_dialog")

    def test_semantic_rupture_repairs_contact_independent_of_wording(self):
        for text in ("Этот вопрос здесь неуместен", "Вы вообще меня слышите?", "Так беседовать невозможно"):
            self.assertEqual(target.expected_dialog_action(self.state(), "rupture", text, text), "repair_contact")

    def test_semantic_decline_is_respected_independent_of_wording(self):
        state = self.state(consultation_offered=True)
        for text in ("Нет", "Мне это не подходит", "Я не хочу записываться"):
            self.assertEqual(target.expected_dialog_action(state, "decline", text, text), "respect_decline")

    def test_decline_reply_cannot_repeat_or_pressure(self):
        data = target.parse_controller_payload(payload(
            "Хорошо, не буду настаивать.",
            action="respect_decline",
            intent="decline",
            evidence="Нет",
        ))
        issues = target.controller_reply_issues(data, "respect_decline", self.state(), [])
        self.assertEqual(issues, [])

    # First client turn is situation/question, never feedback to expert
    def test_first_turn_situation_cannot_be_treated_as_reaction(self):
        state = self.state()
        self.assertEqual(
            target.expected_dialog_action(state, "correction", "ерунда", "Да ерунда какая-то, ничего не хочу.", first_client_turn=True),
            "explore",
        )

    def test_first_turn_preserves_semantically_grounded_boundary_or_end(self):
        state = self.state()
        boundary = "Я не хочу это обсуждать"
        ending = "До свидания"
        self.assertEqual(target.expected_dialog_action(state, "boundary", boundary, boundary, first_client_turn=True), "respect_boundary")
        self.assertEqual(target.expected_dialog_action(state, "end", ending, ending, first_client_turn=True), "end_dialog")

    def test_first_turn_direct_question_still_gets_answer(self):
        state = self.state()
        text = "Сколько стоит консультация?"
        self.assertEqual(target.expected_dialog_action(state, "question", text, text, first_client_turn=True), "answer_information")

    def test_first_turn_validator_uses_semantic_assessment(self):
        bad = target.parse_controller_payload(payload("Ответ эксперта?"))
        bad["reply_assessment"]["treats_message_as_feedback_to_expert"] = True
        issues = target.controller_reply_issues(bad, "explore", self.state(), [], first_client_turn=True)
        self.assertIn("первая реплика ошибочно представлена как оценка слов эксперта", issues)

    def test_first_turn_prompt_marks_neutral_greeting_context(self):
        prompt = target.controller_prompt(self.state(), "База", [], "Да ерунда какая-то", first_client_turn=True)
        self.assertIn("ПЕРВАЯ РЕПЛИКА КЛИЕНТА", prompt)
        self.assertIn("содержательного высказывания", prompt)

    def test_sufficient_psychologist_information_does_not_require_extra_context(self):
        state = self.state(
            need=True,
            previous_experience=True,
            desired_result=True,
            diagnostic_questions=2,
        )
        self.assertEqual(target.expected_dialog_action(state, "continue", "", "Вернуть смысл жизни"), "offer_consultation")

    def test_generic_self_promotion_is_rejected_at_invitation_stage(self):
        data = target.parse_controller_payload(payload(
            "Моя помощь направлена на возвращение жизненной энергии.",
            action="offer_consultation",
        ))
        data["reply_assessment"]["offers_consultation"] = False
        data["reply_assessment"]["uses_generic_self_promotion"] = True
        issues = target.controller_reply_issues(data, "offer_consultation", self.state(), [])
        self.assertIn("консультация не предложена", issues)
        self.assertIn("вместо приглашения используется общая реклама помощи", issues)

    def test_repeated_assistant_reply_is_rejected_generically(self):
        previous = "Краткое отражение запроса и один вопрос?"
        data = target.parse_controller_payload(payload(previous))
        issues = target.controller_reply_issues(
            data,
            "explore",
            self.state(),
            [{"role": "assistant", "content": previous}],
        )
        self.assertIn("повторена предыдущая реплика", issues)

    # Reply validation
    def test_no_repeated_question(self):
        state = self.state(asked_questions=["Давно у вас такое состояние?"])
        data = target.parse_controller_payload(payload("Давно у вас такое состояние?"))
        issues = target.controller_reply_issues(data, "explore", state, [])
        self.assertIn("повторён уже заданный вопрос", issues)

    def test_no_early_consultation_offer(self):
        data = target.parse_controller_payload(payload("Давайте запишемся на консультацию?"))
        issues = target.controller_reply_issues(data, "explore", self.state(), [])
        self.assertIn("преждевременно предложена встреча", issues)

    def test_no_question_while_explaining_solution(self):
        data = target.parse_controller_payload(payload("На встрече смогу разобраться подробнее. Хотите?","explain_solution"))
        issues = target.controller_reply_issues(data, "explain_solution", self.state(), [])
        self.assertIn("задан вопрос на этапе без вопросов", issues)

    def test_answer_is_not_replaced_by_booking(self):
        data = target.parse_controller_payload(payload("Хотите записаться?","answer_information","question","Сколько стоит?"))
        issues = target.controller_reply_issues(data, "answer_information", self.state(), [])
        self.assertIn("ответ на вопрос заменён записью", issues)

    def test_expert_cannot_speak_in_third_person(self):
        data = target.parse_controller_payload(payload("Психолог поможет разобраться.","explain_solution"))
        issues = target.controller_reply_issues(data, "explain_solution", self.state(), [])
        self.assertIn("эксперт говорит о себе в третьем лице", issues)

    def test_team_voice_is_forbidden_by_system_rules(self):
        self.assertIn("не организация и не команда", target.SYSTEM_RULES)
        self.assertIn("будем рады", target.SYSTEM_RULES)

    def test_shown_fallback_question_always_advances_state(self):
        invalid = ScriptedGigaChat([
            "invalid json",
            "Сколько это длится? Как это влияет на вашу жизнь?",
        ])
        target.gigachat = invalid
        with target.app.test_request_context("/"):
            answer, action, state = target.generate_stateful_dialog_reply([], "Да ерунда какая-то, ничего не хочу.", "База")
        self.assertEqual(action, "explore")
        self.assertEqual(answer.count("?"), 1)
        self.assertEqual(state["diagnostic_questions"], 1)

    # Calendar and post-booking
    def test_parse_relative_and_named_dates(self):
        now = datetime(2026, 9, 17, 12, 0, tzinfo=ZoneInfo("Europe/Moscow"))
        self.assertEqual(target.parse_requested_slot("Завтра в 20:00", now).isoformat(), "2026-09-18T20:00:00+03:00")
        self.assertEqual(target.parse_requested_slot("В воскресенье в 11:00", now).isoformat(), "2026-09-20T11:00:00+03:00")

    def test_unrelated_date_does_not_start_booking(self):
        with target.app.test_request_context("/"):
            self.assertIsNone(target.chat_booking_answer("Завтра в 20:00 у меня начинается урок"))

    def test_direct_booking_request_starts_calendar_flow(self):
        with target.app.test_request_context("/"):
            answer = target.direct_booking_answer("Когда можно записаться на бесплатную консультацию?")
            self.assertIn("дату и время", answer)
            self.assertEqual(target.session["requested_booking_type"], "free")

    def test_free_slot_requests_contacts_and_rechecks_on_save(self):
        future = datetime.now(ZoneInfo("Europe/Moscow")) + timedelta(days=30)
        with target.app.test_request_context("/"):
            target.session["consultation_offered"] = True
            first = target.chat_booking_answer(future.strftime("%d.%m.%Y в %H:%M"))
            self.assertIn("свободно", first)
            self.assertIn("имя, телефон и email", first)
            second = target.chat_booking_answer("Анна, +7 999 123-45-67, anna@example.com")
            self.assertIn("Запись подтверждена", second)
            self.assertEqual(len(self.calendar.saved), 1)

    def test_busy_slot_offers_only_calendar_checked_alternatives(self):
        tz = ZoneInfo("Europe/Moscow")
        start = (datetime.now(tz) + timedelta(days=30)).replace(hour=11, minute=0, second=0, microsecond=0)
        self.calendar.busy = [(start, start + timedelta(hours=1))]
        with target.app.test_request_context("/"):
            target.session["consultation_offered"] = True
            answer = target.chat_booking_answer(start.strftime("%d.%m.%Y в %H:%M"))
            self.assertIn("уже занято", answer)
            self.assertTrue(target.session.get("offered_slots"))

    def test_booking_never_confirms_before_calendar_save(self):
        future = datetime.now(ZoneInfo("Europe/Moscow")) + timedelta(days=30)
        with target.app.test_request_context("/"):
            target.session["pending_booking"] = {"start": future.isoformat(), "type": "free"}
            self.calendar.busy = [(future, future + timedelta(hours=1))]
            answer = target.chat_booking_answer("Анна, +7 999 123-45-67, anna@example.com")
            self.assertNotIn("Запись подтверждена", answer)
            self.assertEqual(self.calendar.saved, [])

    def test_confirmed_booking_is_not_offered_again(self):
        with target.app.test_request_context("/"):
            target.session["last_booking"] = "Бесплатная консультация, 20.09.2026 в 11:00, 20 минут"
            self.assertIn("вижу вашу запись", target.direct_booking_answer("Я уже записалась"))
            final = target.completed_dialog_answer("Хорошо, до встречи")
            self.assertEqual(final, "До встречи! Хорошего дня.")
            self.assertTrue(target.session["dialog_closed"])

    # UI, persistence and admin
    def test_input_is_not_disabled_while_reply_pending(self):
        page = self.client.get("/").get_data(as_text=True)
        self.assertNotIn("input.disabled=true", page)
        self.assertIn("pending", page)

    def test_history_is_restored(self):
        page = self.client.get("/").get_data(as_text=True)
        self.assertIn("fetch('/api/history')", page)

    def test_reset_clears_dialog_and_controller_state(self):
        with self.client.session_transaction() as session:
            session["sid"] = "old"
            session["dialog_controller_state"] = {"need": True}
            session["consultation_offered"] = True
        con = target.db()
        con.execute("insert into messages values(?,?,?,?)", ("old", "user", "text", 1))
        con.commit()
        self.assertEqual(self.client.post("/api/reset").status_code, 200)
        with self.client.session_transaction() as session:
            self.assertNotIn("sid", session)
            self.assertNotIn("dialog_controller_state", session)

    def test_knowledge_base_survives_dialog_reset(self):
        self.client.post("/api/reset")
        con = target.db()
        self.assertEqual(con.execute("select count(*) from documents").fetchone()[0], 1)

    def test_admin_upload_and_list(self):
        headers = {"X-Admin-Password": "admin123"}
        response = self.client.post(
            "/api/admin/upload",
            headers=headers,
            data={"files": (io.BytesIO("Новый факт".encode()), "new.txt")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        listed = self.client.get("/api/admin", headers=headers)
        self.assertEqual(listed.status_code, 200)
        self.assertTrue(any(x["name"] == "new.txt" for x in listed.json["documents"]))


if __name__ == "__main__":
    unittest.main()
