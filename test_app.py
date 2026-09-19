import io
import os
import tempfile
import unittest
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import app as target


class FakeGigaChat:
    def reply(self, messages):
        prompt = messages[0]["content"]
        if "Извлеките из собственных слов клиента" in prompt:
            return '{"stages":{}}'
        if "Определите функцию реплики" in prompt:
            return '{"question_purpose":"answer_client_question","performs_expert_work":false,"asks_for_deliverable_details":false,"uses_unsupported_assumption":false,"repeats_answered_question":false,"claims_unverified_product":false,"makes_unverified_promise":false,"answers_client_question":true,"natural_and_clear":true}'
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

    def test_unrelated_date_does_not_start_booking(self):
        with target.app.test_request_context("/"):
            target.session["expert_slug"] = "marketer"
            self.assertIsNone(target.chat_booking_answer("Завтра в 20:00 у меня начинается урок"))

    def test_consultation_acceptance_moves_to_booking_once(self):
        with target.app.test_request_context("/"):
            target.session["expert_slug"] = "marketer"
            target.session["consultation_offered"] = True
            answer = target.accepted_consultation_answer("Хочу")
        self.assertIn("[[BOOK_FREE]]", answer)
        self.assertNotIn("[[BOOK_REGULAR]]", answer)

    def test_discovery_guard_blocks_early_offer(self):
        state = {key: False for key in target.EXPERT_PROFILES["marketer"]["discovery_stages"]}
        state.update(identity=True)
        answer = target.guard_discovery_answer(
            "Могу предложить бесплатную консультацию. Хотите попробовать?",
            state,
            "task",
            "На подготовку уходит много времени",
        )
        self.assertNotIn("консультац", answer.lower())

    def test_discovery_guard_removes_assumed_cooperation(self):
        state = {key: False for key in target.EXPERT_PROFILES["marketer"]["discovery_stages"]}
        state["identity"] = True
        answer = target.guard_discovery_answer(
            "Какой результат вы хотите получить, работая со мной?",
            state,
            "task",
            "Я репетитор английского языка",
        )
        self.assertNotIn("работая со мной", answer.lower())

    def test_discovery_focus_rephrases_after_confusion(self):
        profile = target.EXPERT_PROFILES["marketer"]
        class NaturalReply:
            def reply(self, messages):
                self.prompt = messages[0]["content"]
                return "Я спрашиваю именно о вашей повседневной работе. Что в ней сейчас больше всего мешает?"
        fake, old = NaturalReply(), target.gigachat
        target.gigachat = fake
        try:
            answer = target.generate_discovery_reply([], "В смысле?!", "", profile, "task")
        finally:
            target.gigachat = old
        self.assertIn("повседневной работе", answer)
        self.assertIn("Клиент выразил непонимание: да", fake.prompt)

    def test_discovery_focus_replaces_domain_consulting_question(self):
        profile = target.EXPERT_PROFILES["marketer"]
        class FocusedReply:
            def reply(self, messages):
                return "С курсом уже понятно, где теряется время. Как нейросети справлялись с этой задачей раньше?"
        old = target.gigachat
        target.gigachat = FocusedReply()
        try:
            answer = target.generate_discovery_reply([], "Хочу создать курс, но времени совсем нет", "", profile, "ai_experience")
        finally:
            target.gigachat = old
        self.assertIn("нейросети", answer)
        self.assertNotIn("темы", answer)
        self.assertEqual(answer.count("?"), 1)

    def test_discovery_guard_requires_solution_interest(self):
        state = {key: True for key in target.EXPERT_PROFILES["marketer"]["discovery_stages"]}
        state.update(solution_explained=True, solution_interest=False)
        answer = target.guard_discovery_answer(
            "Тогда предлагаю бесплатную консультацию. Хотите записаться?",
            state,
            None,
            "Понятно",
        )
        self.assertNotIn("запис", answer.lower())
        state["solution_interest"] = True
        allowed = target.guard_discovery_answer(
            "Тогда предлагаю бесплатную консультацию. Хотите записаться?",
            state,
            None,
            "Да, мне интересно",
        )
        self.assertIn("записаться", allowed)

    def test_discovery_json_parser(self):
        parsed = target.parse_json_object('```json\n{"identity":true,"goal":false}\n```')
        self.assertTrue(parsed["identity"])
        self.assertFalse(parsed["goal"])

    def test_state_machine_requires_grounded_stage_evidence(self):
        profile = target.EXPERT_PROFILES["marketer"]
        class Extractor:
            def __init__(self, response): self.response = response
            def reply(self, messages): return self.response
        old = target.gigachat
        try:
            with target.app.test_request_context("/"):
                target.gigachat = Extractor('{"stages":{"identity":{"complete":true,"evidence":"Я репетитор английского"},"task":{"complete":true,"evidence":"Я репетитор английского"},"ai_experience":{"complete":true,"evidence":"Я репетитор английского"}}}')
                state = target.assess_discovery([], "Я репетитор английского, работаю с детьми и взрослыми", profile)
        finally:
            target.gigachat = old
        self.assertTrue(state["identity"])
        self.assertFalse(state["task"])
        self.assertFalse(state["ai_experience"])
        self.assertFalse(state["solution_explained"])
        self.assertFalse(state["solution_interest"])

    def test_generation_failure_cannot_escape_state_machine(self):
        class BreakAfterAssessment:
            def reply(self, messages):
                if "Извлеките из собственных слов клиента" in messages[0]["content"]:
                    return '{"stages":{"identity":{"complete":true,"evidence":"Я репетитор английского"},"task":{"complete":false,"evidence":""},"ai_experience":{"complete":false,"evidence":""}}}'
                raise RuntimeError("generation unavailable")
        old = target.gigachat
        target.gigachat = BreakAfterAssessment()
        try:
            with self.client.session_transaction() as session:
                session["expert_slug"] = "marketer"
            response = self.client.post("/api/chat", json={"message":"Я репетитор английского, работаю с детьми и взрослыми"})
        finally:
            target.gigachat = old
        self.assertEqual(response.status_code, 200)
        self.assertNotIn("консультац", response.json["answer"].lower())
        self.assertEqual(response.json["answer"], target.safe_phase_reply("explore_task", target.EXPERT_PROFILES["marketer"]))

    def test_state_machine_accepts_need_only_from_client_words(self):
        profile = target.EXPERT_PROFILES["marketer"]
        class Extractor:
            def reply(self, messages):
                return '{"stages":{"identity":{"complete":true,"evidence":"Я репетитор"},"task":{"complete":true,"evidence":"Хотелось бы побыстрее готовиться к урокам"},"ai_experience":{"complete":false,"evidence":""}}}'
        old = target.gigachat
        target.gigachat = Extractor()
        try:
            with target.app.test_request_context("/"):
                history = [{"role":"user","content":"Я репетитор"},{"role":"assistant","content":"Что хотелось бы изменить?"}]
                state = target.assess_discovery(history, "Хотелось бы побыстрее готовиться к урокам", profile)
        finally:
            target.gigachat = old
        self.assertTrue(state["identity"])
        self.assertTrue(state["task"])
        self.assertFalse(state["ai_experience"])

    def test_interest_cannot_exist_before_solution_explanation(self):
        profile = target.EXPERT_PROFILES["marketer"]
        class Extractor:
            def reply(self, messages):
                return '{"stages":{"identity":{"complete":false,"evidence":""},"task":{"complete":false,"evidence":""},"ai_experience":{"complete":false,"evidence":""}}}'
        old = target.gigachat
        target.gigachat = Extractor()
        try:
            with target.app.test_request_context("/"):
                target.session["discovery_state"] = {"identity":True,"task":True,"ai_experience":True,"solution_explained":False,"solution_interest":False}
                state = target.assess_discovery([], "Да, интересно", profile)
                self.assertFalse(state["solution_interest"])
                target.session["discovery_state"]["solution_explained"] = True
                state = target.assess_discovery([], "Да, интересно", profile)
                self.assertTrue(state["solution_interest"])
        finally:
            target.gigachat = old

    def test_phase_controller_never_explains_solution_without_task(self):
        profile = target.EXPERT_PROFILES["marketer"]
        state = dict(identity=True, task=True, ai_experience=False, solution_explained=False, solution_interest=False)
        with target.app.test_request_context("/"):
            self.assertEqual(target.discovery_action(state, profile, "Нейросеть всё время теряет логику курса"), "explore_ai_experience")

    def test_phase_controller_repairs_misunderstood_question(self):
        profile = target.EXPERT_PROFILES["marketer"]
        state = dict(identity=True, task=False, ai_experience=False, solution_explained=False, solution_interest=False)
        with target.app.test_request_context("/"):
            self.assertEqual(target.discovery_action(state, profile, "Не поняла вопрос"), "repair_task")
            self.assertEqual(target.discovery_action(state, profile, "Вы издеваетесь? Какое отношение это имеет к вам?"), "repair_task")

    def test_phase_controller_explains_solution_when_discovery_complete(self):
        profile = target.EXPERT_PROFILES["marketer"]
        state = dict(identity=True, task=True, ai_experience=True, solution_explained=False, solution_interest=False)
        with target.app.test_request_context("/"):
            self.assertEqual(target.discovery_action(state, profile, "Не знаю, можно ли это исправить"), "explain_solution")

    def test_full_marketer_state_sequence(self):
        profile = target.EXPERT_PROFILES["marketer"]
        with target.app.test_request_context("/"):
            state = dict(identity=True, task=False, ai_experience=False, solution_explained=False, solution_interest=False)
            self.assertEqual(target.discovery_action(state, profile, "Я репетитор"), "explore_task")
            state["task"] = True
            self.assertEqual(target.discovery_action(state, profile, "Хочу быстрее готовиться"), "explore_ai_experience")
            state["ai_experience"] = True
            self.assertEqual(target.discovery_action(state, profile, "Пробовала, но переделываю"), "explain_solution")
            state["solution_explained"] = True
            self.assertEqual(target.discovery_action(state, profile, "Понятно"), "handle_solution_interest")
            state["solution_interest"] = True
            self.assertEqual(target.discovery_action(state, profile, "Да, интересно"), "offer_consultation")

    def test_psychologist_uses_same_state_machine(self):
        profile = target.EXPERT_PROFILES["psychologist"]
        with target.app.test_request_context("/"):
            state = dict(situation=True, duration_impact=False, solution_explained=False, solution_interest=False)
            self.assertEqual(target.discovery_action(state, profile, "Так уже давно"), "explore_duration_impact")
            state["duration_impact"] = True
            self.assertEqual(target.discovery_action(state, profile, "Это мешает жить"), "explain_solution")

    def test_phase_validation_detects_repeated_opening(self):
        history = [{"role": "assistant", "content": "Похоже, здесь теряется логика курса."}]
        issues = target.phase_reply_issues(
            "Похоже, здесь мало контекста. Расскажите подробнее, что именно не получилось?",
            "explain_solution",
            history,
        )
        self.assertIn("повтор того же начала реплики", issues)

    def test_semantic_review_blocks_expert_work_in_chat(self):
        issues = target.phase_review_issues({
            "question_purpose": "other",
            "performs_expert_work": True,
            "asks_for_deliverable_details": True,
            "uses_unsupported_assumption": False,
            "repeats_answered_question": False,
            "natural_and_clear": True,
        }, "explain_solution")
        self.assertIn("бот начинает выполнять работу живого эксперта", issues)
        self.assertIn("бот собирает данные для создания результата вместо продажи решения", issues)

    def test_semantic_review_blocks_invented_product_and_promises(self):
        issues = target.phase_review_issues({
            "question_purpose": "answer_client_question",
            "performs_expert_work": False,
            "asks_for_deliverable_details": False,
            "uses_unsupported_assumption": False,
            "repeats_answered_question": False,
            "claims_unverified_product": True,
            "makes_unverified_promise": True,
            "answers_client_question": False,
            "natural_and_clear": True,
        }, "answer_information")
        self.assertIn("бот выдаёт возможное направление решения за существующий продукт", issues)
        self.assertIn("бот обещает неподтверждённый результат", issues)
        self.assertIn("бот не ответил на прямой вопрос клиента", issues)

    def test_solution_question_is_recognized_as_informational(self):
        self.assertTrue(target.is_informational_question("Что за помощник?"))
        self.assertTrue(target.is_informational_question("А как он работает?"))

    def test_quality_filter_detects_result_promises(self):
        issues = target.quality_issues("Вы быстро получите качественные материалы и сэкономите время.")
        self.assertIn("неподтверждённое обещание результата", issues)

    def test_phase_validation_rejects_suggested_task_options(self):
        issues = target.phase_reply_issues(
            "Вам нужно автоматизировать подготовку уроков или улучшить контент для занятий?",
            "explore_task",
            [],
        )
        self.assertIn("догадка о задаче клиента вместо открытого вопроса", issues)
        self.assertIn("варианты ответа внутри вопроса", issues)

    def test_phase_validation_rejects_presumed_materials(self):
        issues = target.phase_reply_issues(
            "Какие конкретно учебные материалы вы хотели бы создавать?",
            "explore_task",
            [],
        )
        self.assertIn("догадка о задаче клиента вместо открытого вопроса", issues)

    def test_phase_validation_rejects_closed_problem_hypothesis(self):
        issues = target.phase_reply_issues(
            "Вам сложно подобрать подходящий способ работы с разными клиентами?",
            "explore_task",
            [],
        )
        self.assertIn("догадка о задаче клиента вместо открытого вопроса", issues)

    def test_safe_fallback_for_task_does_not_assume_ai_usage(self):
        answer = target.safe_phase_reply("explore_task")
        self.assertNotIn("нейросет", answer.lower())
        self.assertNotIn("вы уже", answer.lower())
        self.assertEqual(answer.count("?"), 1)

    def test_invalid_last_generation_cannot_reach_dialog(self):
        profile = target.EXPERT_PROFILES["marketer"]
        class AlwaysInvalid:
            def reply(self, messages):
                return "Здорово, что вы уже используете нейросети?"
        old = target.gigachat
        target.gigachat = AlwaysInvalid()
        try:
            answer = target.generate_phase_reply([], "Я репетитор", "", profile, "explore_task")
        finally:
            target.gigachat = old
        self.assertEqual(answer, target.safe_phase_reply("explore_task"))

    def test_semantic_review_allows_only_interest_question_after_solution(self):
        valid = target.phase_review_issues({
            "question_purpose": "check_solution_interest",
            "performs_expert_work": False,
            "asks_for_deliverable_details": False,
            "uses_unsupported_assumption": False,
            "repeats_answered_question": False,
            "natural_and_clear": True,
        }, "explain_solution")
        self.assertEqual(valid, [])
        wrong = target.phase_review_issues({
            "question_purpose": "discover_need",
            "performs_expert_work": False,
            "asks_for_deliverable_details": False,
            "uses_unsupported_assumption": False,
            "repeats_answered_question": False,
            "natural_and_clear": True,
        }, "explain_solution")
        self.assertTrue(any("другую функцию" in issue for issue in wrong))

    def test_shared_quality_filter_detects_cliches(self):
        issues = target.quality_issues("Понимаю вас. Это мощный инструмент. Что вы пробовали?")
        self.assertIn("шаблонная или канцелярская формулировка", issues)

    def test_shared_quality_filter_keeps_one_question(self):
        answer = target.keep_one_question("Как сейчас устроена работа? Используете ли вы нейросети? Расскажите подробнее.")
        self.assertEqual(answer.count("?"), 1)
        self.assertNotIn("Используете", answer)

    def test_shared_quality_filter_removes_unverified_followup(self):
        answer = target.remove_unverified_promises(
            "Встреча пройдёт через Телемост. Я свяжусь с вами накануне, чтобы прислать ссылку. До встречи!"
        )
        self.assertIn("Телемост", answer)
        self.assertNotIn("свяжусь", answer)

    def test_busy_slot_offers_real_alternatives(self):
        tz = ZoneInfo("Europe/Moscow")
        busy_start = (datetime.now(tz) + timedelta(days=60)).replace(hour=20, minute=0, second=0, microsecond=0)
        self.calendar.busy = [(busy_start, busy_start + timedelta(hours=1))]
        with self.client.session_transaction() as session:
            session["expert_slug"] = "marketer"
            session["consultation_offered"] = True
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
