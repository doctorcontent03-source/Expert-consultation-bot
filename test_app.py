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

    def test_home_script_keeps_message_separator_escaped(self):
        script = target.HOME_HTML.split("<script>", 1)[1].split("</script>", 1)[0]
        self.assertIn("pending.join('\\n\\n')", script)
        self.assertNotIn("pending.join('\n\n')", script)

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
            session["sid"] = "booking-dialog"
            session["consultation_offered"] = True
            session["consultation_offered_sid"] = "booking-dialog"
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
            target.session["sid"] = "booking-dialog"
            target.session["consultation_offered"] = True
            target.session["consultation_offered_sid"] = "booking-dialog"
            answer = target.accepted_consultation_answer("Хочу")
            self.assertIn("дату и время", answer)
            self.assertNotIn("[[BOOK_", answer)
            self.assertEqual(target.session.get("requested_booking_type"), "free")

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
        self.assertEqual(response.status_code, 502)
        self.assertIn("Не удалось сформировать корректный ответ", response.json["error"])

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
                target.session["sid"] = "same-dialog"
                target.session["discovery_sid"] = "same-dialog"
                target.session["discovery_state"] = {"identity":True,"task":True,"ai_experience":True,"solution_explained":False,"solution_interest":False}
                history = [{"role": "user", "content": "Предыдущая часть разговора"}]
                state = target.assess_discovery(history, "Да, интересно", profile)
                self.assertFalse(state["solution_interest"])
                target.session["discovery_state"]["solution_explained"] = True
                state = target.assess_discovery(history, "Да, интересно", profile)
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
        self.assertTrue(target.is_informational_question("Хочу. Вы имеете в виду какое решение?"))
        self.assertTrue(target.is_informational_question("Ассистента, приложение или чат-бота?"))
        self.assertTrue(target.is_informational_question("Я понимаю, это в ChatGPT?"))

    def test_question_does_not_cancel_explicit_solution_interest(self):
        self.assertTrue(target.explicit_solution_interest("Хочу. Вы имеете в виду какое решение?"))
        self.assertTrue(target.explicit_solution_interest("Ясно. Ну, может, это и помогло бы."))

    def test_generic_solution_does_not_close_explanation_stage(self):
        profile = target.EXPERT_PROFILES["marketer"]
        self.assertFalse(target.solution_was_explained(
            "Здесь может подойти ИИ-решение, настроенное под ваш рабочий процесс.",
            profile,
        ))
        self.assertTrue(target.solution_was_explained(
            "Возможное направление — ИИ-помощник для подготовки материалов.",
            profile,
        ))

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

    def test_task_discovery_requires_open_question(self):
        for answer in (
            "Вам приходится подбирать материалы для каждого ученика?",
            "Вы сталкиваетесь с проблемой подбора интересных тем?",
            "Выходит, вам важно найти актуальные задания?",
        ):
            issues = target.phase_reply_issues(answer, "explore_task", [])
            self.assertIn("закрытый или наводящий вопрос вместо открытого выяснения задачи", issues)

    def test_structured_discovery_recovers_after_closed_generated_questions(self):
        class ClosedThenStructured:
            def reply(self, messages):
                prompt = messages[0]["content"]
                if 'Верните только JSON: {"reaction":"","question":""}' in prompt:
                    return '{"reaction":"","question":"Что в вашей работе хотелось бы изменить в первую очередь?"}'
                if "Определите функцию реплики чат-бота" in prompt:
                    return '{"question_purpose":"discover_need","performs_expert_work":false,"asks_for_deliverable_details":false,"uses_unsupported_assumption":false,"repeats_answered_question":false,"claims_unverified_product":false,"makes_unverified_promise":false,"answers_client_question":true,"natural_and_clear":true}'
                return "Вам приходится подбирать материалы для каждого ученика?"
        old = target.gigachat
        target.gigachat = ClosedThenStructured()
        try:
            answer = target.generate_phase_reply(
                [],
                "Я репетитор английского языка и работаю с детьми и взрослыми.",
                "",
                target.EXPERT_PROFILES["marketer"],
                "explore_task",
            )
        finally:
            target.gigachat = old
        self.assertEqual(answer, "Что в вашей работе хотелось бы изменить в первую очередь?")

    def test_structured_discovery_rejects_repeated_invalid_wording(self):
        class AlwaysClosed:
            def reply(self, messages):
                return '{"reaction":"","question":"Вам трудно готовить материалы для разных учеников?"}'
        old = target.gigachat
        target.gigachat = AlwaysClosed()
        try:
            answer = target.generate_structured_discovery_reply(
                [],
                "Я репетитор английского языка.",
                target.EXPERT_PROFILES["marketer"],
                "explore_task",
            )
        finally:
            target.gigachat = old
        self.assertIsNone(answer)

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
        self.assertIsNone(answer)

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

    def test_style_issue_is_not_allowed_to_break_dialog(self):
        self.assertEqual(
            target.blocking_reply_issues(["шаблонная или канцелярская формулировка"]),
            [],
        )
        self.assertEqual(
            target.blocking_reply_issues(["не удалось проверить смысл реплики"]),
            [],
        )

    def test_state_machine_violation_still_blocks_reply(self):
        issues = target.blocking_reply_issues([
            "шаблонная или канцелярская формулировка",
            "догадка о задаче клиента вместо открытого вопроса",
        ])
        self.assertEqual(issues, ["догадка о задаче клиента вместо открытого вопроса"])

    def test_semantic_reviewer_cannot_veto_deterministically_valid_reply(self):
        class WrongReviewer:
            def reply(self, messages):
                prompt = messages[0]["content"]
                if "Определите функцию реплики чат-бота" in prompt:
                    return '{"question_purpose":"other","performs_expert_work":false,"asks_for_deliverable_details":false,"uses_unsupported_assumption":true,"repeats_answered_question":false,"claims_unverified_product":false,"makes_unverified_promise":false,"answers_client_question":true,"natural_and_clear":true}'
                return "Что в вашей работе сейчас отнимает больше всего времени?"
        old = target.gigachat
        target.gigachat = WrongReviewer()
        try:
            answer = target.generate_phase_reply(
                [],
                "Я репетитор английского языка.",
                "",
                target.EXPERT_PROFILES["marketer"],
                "explore_task",
            )
        finally:
            target.gigachat = old
        self.assertEqual(answer, "Что в вашей работе сейчас отнимает больше всего времени?")

    def test_unverified_product_claim_from_semantic_review_still_blocks(self):
        class ProductHallucinator:
            def reply(self, messages):
                prompt = messages[0]["content"]
                if "Определите функцию реплики чат-бота" in prompt:
                    return '{"question_purpose":"check_solution_interest","performs_expert_work":false,"asks_for_deliverable_details":false,"uses_unsupported_assumption":false,"repeats_answered_question":false,"claims_unverified_product":true,"makes_unverified_promise":false,"answers_client_question":true,"natural_and_clear":true}'
                return "У меня есть готовый ассистент, который анализирует популярные игры. Хотите посмотреть?"
        old = target.gigachat
        target.gigachat = ProductHallucinator()
        try:
            answer = target.generate_phase_reply(
                [],
                "Пыталась создать курс для подростков, но не получилось.",
                target.EXPERT_PROFILES["marketer"]["profile_context"],
                target.EXPERT_PROFILES["marketer"],
                "explain_solution",
            )
        finally:
            target.gigachat = old
        self.assertIsNone(answer)

    def test_expert_profile_uses_positive_offer_catalog_not_example_blacklist(self):
        context = target.EXPERT_PROFILES["marketer"]["profile_context"].lower()
        self.assertIn("готовый ассистент по созданию нестандартных курсов", context)
        self.assertIn("разработки персонального ассистента", context)
        self.assertNotIn("такие предложения делать нельзя", context)
        self.assertNotIn("нет сведений о", context)

    def test_confusion_is_a_separate_client_intent(self):
        class ConfusionClassifier:
            def reply(self, messages):
                return '{"intent":"confusion","subject":"solution","confidence":"high"}'
        old = target.gigachat
        target.gigachat = ConfusionClassifier()
        try:
            move = target.classify_client_move(
                [{"role": "assistant", "content": "Предыдущее объяснение решения."}],
                "Ничего не поняла.",
            )
        finally:
            target.gigachat = old
        self.assertEqual(move["intent"], "confusion")
        state = dict(identity=True, task=True, ai_experience=True, solution_explained=True, solution_interest=False)
        self.assertEqual(
            target.discovery_action(state, target.EXPERT_PROFILES["marketer"], "Ничего не поняла.", move),
            "answer_information",
        )

    def test_shared_quality_filter_rejects_team_voice_and_invented_specialization(self):
        issues = target.quality_issues(
            "Мы можем показать генератор, специально разработанный для преподавателей английского."
        )
        self.assertIn("эксперт говорит от имени команды, а не от первого лица", issues)
        self.assertIn("придумана неподтверждённая специализация продукта", issues)

    def test_bot_cannot_start_demo_inside_chat(self):
        issues = target.phase_reply_issues(
            "Я запущу помощника прямо здесь и сейчас и покажу пример задания.",
            "handle_solution_interest",
            [],
        )
        self.assertIn("бот пытается провести демонстрацию или работу эксперта внутри чата", issues)

    def test_sales_phase_cannot_request_client_deliverables(self):
        issues = target.phase_reply_issues(
            "Поделитесь типичными примерами заданий или тем, которые приходится готовить.",
            "offer_consultation",
            [],
        )
        self.assertIn("бот запрашивает материалы для выполнения работы эксперта", issues)
        self.assertIn("вместо предложения встречи бот выполняет другую задачу", issues)

    def test_refusal_stops_consultation_pressure(self):
        con = target.db()
        con.execute(
            "insert into messages(session_id,role,content,created_at) values(?,?,?,?)",
            ("refusal-dialog", "assistant", "Хотите записаться на бесплатную консультацию?", 1),
        )
        con.commit()
        with self.client.session_transaction() as session:
            session["expert_slug"] = "marketer"
            session["sid"] = "refusal-dialog"
            session["consultation_offered"] = True
            session["consultation_offered_sid"] = "refusal-dialog"
        first = self.client.post("/api/chat", json={"message": "Нет, спасибо."})
        self.assertEqual(first.json["answer"], "Хорошо, не буду настаивать.")
        second = self.client.post("/api/chat", json={"message": "Я уже сказала, нет."})
        self.assertNotIn("консультац", second.json["answer"].lower())
        self.assertNotIn("встреч", second.json["answer"].lower())

    def test_refusal_after_solution_interest_question_stops_sales_path(self):
        con = target.db()
        con.execute(
            "insert into messages(session_id,role,content,created_at) values(?,?,?,?)",
            ("solution-refusal", "assistant", "Хотите посмотреть, как такой помощник может работать в вашей ситуации?", 1),
        )
        con.commit()
        with self.client.session_transaction() as session:
            session["expert_slug"] = "marketer"
            session["sid"] = "solution-refusal"
        first = self.client.post("/api/chat", json={"message": "Нет, спасибо."})
        self.assertEqual(first.json["answer"], "Хорошо, не буду настаивать.")
        second = self.client.post("/api/chat", json={"message": "Я уже сказала, что нет."})
        self.assertEqual(second.json["answer"], "Хорошо, не буду возвращаться к этому предложению.")
        self.assertNotIn("решени", second.json["answer"].lower())

    def test_hesitation_after_offer_pauses_instead_of_repeating_offer(self):
        con = target.db()
        con.execute(
            "insert into messages(session_id,role,content,created_at) values(?,?,?,?)",
            ("hesitation-dialog", "assistant", "Могу показать это на бесплатной консультации. Хотите записаться?", 1),
        )
        con.commit()
        with self.client.session_transaction() as session:
            session["expert_slug"] = "marketer"
            session["sid"] = "hesitation-dialog"
            session["consultation_offered"] = True
            session["consultation_offered_sid"] = "hesitation-dialog"
        answer = self.client.post("/api/chat", json={"message": "Ну не знаю"}).json["answer"]
        self.assertIn("решать прямо сейчас не обязательно", answer)
        self.assertNotIn("Хотите записаться", answer)
        refusal = self.client.post("/api/chat", json={"message": "Теперь точно нет, спасибо."}).json["answer"]
        self.assertEqual(refusal, "Хорошо, не буду настаивать.")

    def test_not_interesting_inside_problem_description_is_not_refusal(self):
        history = [{
            "role": "assistant",
            "content": "А что в вашей работе сейчас хотелось бы упростить или изменить?",
        }]
        self.assertTrue(target.consultation_refusal(
            "Особенно сложно с подростками, которым вообще ничего не интересно."
        ))
        self.assertFalse(target.last_assistant_offered_consultation(history))

    def test_offer_flag_from_another_dialog_is_inactive(self):
        with target.app.test_request_context("/"):
            target.session["sid"] = "new-dialog"
            target.session["consultation_offered"] = True
            target.session["consultation_offered_sid"] = "old-dialog"
            self.assertFalse(target.consultation_offer_active())

    def test_semantic_client_move_controls_transition_without_phrase_matching(self):
        class SemanticClassifier:
            def reply(self, messages):
                return '{"intent":"interest","subject":"solution","confidence":"high"}'
        old = target.gigachat
        target.gigachat = SemanticClassifier()
        try:
            move = target.classify_client_move(
                [{"role": "assistant", "content": "Речь идёт о возможном ИИ-решении."}],
                "Возможно, в этом что-то есть.",
            )
        finally:
            target.gigachat = old
        self.assertEqual(move["intent"], "interest")
        state = dict(identity=True, task=True, ai_experience=True, solution_explained=True, solution_interest=True)
        self.assertEqual(
            target.discovery_action(state, target.EXPERT_PROFILES["marketer"], "Возможно, в этом что-то есть.", move),
            "offer_consultation",
        )

    def test_semantic_classifier_keeps_problem_description_out_of_refusal_state(self):
        class SemanticClassifier:
            def reply(self, messages):
                return '{"intent":"other","subject":"other","confidence":"high"}'
        old = target.gigachat
        target.gigachat = SemanticClassifier()
        try:
            move = target.classify_client_move(
                [{"role": "assistant", "content": "Что хотелось бы изменить в работе?"}],
                "Подросткам вообще ничего не интересно.",
            )
        finally:
            target.gigachat = old
        self.assertEqual(move["intent"], "other")

    def test_semantic_booking_request_skips_repeated_consultation_offer(self):
        class BookingClassifier:
            def reply(self, messages):
                prompt = messages[0]["content"]
                if "Определите функцию последней реплики клиента" in prompt:
                    return '{"intent":"booking_request","subject":"booking","confidence":"high"}'
                return "Могу показать это на бесплатной консультации. Хотите записаться?"
        con = target.db()
        con.execute(
            "insert into messages(session_id,role,content,created_at) values(?,?,?,?)",
            ("booking-intent", "assistant", "Могу показать это на бесплатной консультации. Хотите записаться?", 1),
        )
        con.commit()
        with self.client.session_transaction() as session:
            session["expert_slug"] = "marketer"
            session["sid"] = "booking-intent"
            session["consultation_offered"] = True
            session["consultation_offered_sid"] = "booking-intent"
        old = target.gigachat
        target.gigachat = BookingClassifier()
        try:
            response = self.client.post("/api/chat", json={"message": "Давайте. Когда?"})
        finally:
            target.gigachat = old
        self.assertEqual(response.json["answer"], "Назовите удобные дату и время — я сразу проверю их в календаре.")
        self.assertNotIn("Хотите записаться", response.json["answer"])

    def test_when_can_i_book_stays_in_chat_without_form(self):
        with target.app.test_request_context("/"):
            target.session["expert_slug"] = "marketer"
            answer = target.direct_booking_answer("И когда можно?")
            self.assertIn("Назовите удобные дату и время", answer)
            self.assertNotIn("[[BOOK_", answer)
            self.assertNotIn("форм", answer.lower())

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
            session["sid"] = "busy-dialog"
            session["consultation_offered"] = True
            session["consultation_offered_sid"] = "busy-dialog"
        result = self.client.post("/api/chat", json={"message": busy_start.strftime("%d.%m.%Y в 20:00")})
        self.assertEqual(result.status_code, 200)
        self.assertIn("уже занято", result.json["answer"])
        self.assertIn((busy_start + timedelta(hours=1)).strftime("%d.%m в %H:%M"), result.json["answer"])
        chosen = self.client.post("/api/chat", json={"message": "Давайте в 21:00"})
        self.assertIn("свободно", chosen.json["answer"])
        self.assertIn("имя, телефон и email", chosen.json["answer"])

    def test_home_restores_history_without_erasing_intro(self):
        page = self.client.get("/").get_data(as_text=True)
        self.assertIn("fetch('/api/history')", page)
        self.assertIn("(data.messages||[]).forEach", page)
        self.assertNotIn("chat.innerHTML=''", page)
        self.assertIn("[[BOOK_FREE]]", page)

    def test_reset_clears_dialog_state_and_messages(self):
        con = target.db()
        con.execute(
            "insert into messages(session_id,role,content,created_at) values(?,?,?,?)",
            ("old-dialog", "user", "Старое сообщение", 1),
        )
        con.commit()
        with self.client.session_transaction() as session:
            session["expert_slug"] = "marketer"
            session["sid"] = "old-dialog"
            session["discovery_state"] = {"identity": True, "task": True}
            session["consultation_offered"] = True
        response = self.client.post("/api/reset")
        self.assertEqual(response.status_code, 200)
        with self.client.session_transaction() as session:
            self.assertEqual(session.get("expert_slug"), "marketer")
            self.assertNotIn("sid", session)
            self.assertNotIn("discovery_state", session)
            self.assertNotIn("consultation_offered", session)
        self.assertEqual(
            con.execute("select count(*) from messages where session_id=?", ("old-dialog",)).fetchone()[0],
            0,
        )

    def test_stale_discovery_cookie_cannot_skip_new_dialog(self):
        class IdentityExtractor:
            def reply(self, messages):
                return '{"stages":{"identity":{"complete":true,"evidence":"Я репетитор по английскому, работаю с детьми от 10 лет и взрослыми"},"task":{"complete":false,"evidence":""},"ai_experience":{"complete":false,"evidence":""}}}'
        old = target.gigachat
        target.gigachat = IdentityExtractor()
        try:
            with target.app.test_request_context("/"):
                target.session["sid"] = "new-dialog"
                target.session["discovery_sid"] = "old-dialog"
                target.session["discovery_state"] = {
                    "identity": True,
                    "task": True,
                    "ai_experience": True,
                    "solution_explained": True,
                    "solution_interest": True,
                }
                state = target.assess_discovery(
                    [],
                    "Я репетитор по английскому, работаю с детьми от 10 лет и взрослыми.",
                    target.EXPERT_PROFILES["marketer"],
                )
                self.assertTrue(state["identity"])
                self.assertFalse(state["task"])
                self.assertFalse(state["ai_experience"])
                self.assertFalse(state["solution_explained"])
                self.assertFalse(state["solution_interest"])
                self.assertEqual(target.discovery_action(state, target.EXPERT_PROFILES["marketer"], ""), "explore_task")
        finally:
            target.gigachat = old

    def test_first_direct_identity_answer_closes_identity_stage_without_model_help(self):
        class EmptyExtractor:
            def reply(self, messages):
                return '{"stages":{"identity":{"complete":false,"evidence":""},"task":{"complete":false,"evidence":""},"ai_experience":{"complete":false,"evidence":""}}}'
        old = target.gigachat
        target.gigachat = EmptyExtractor()
        try:
            with target.app.test_request_context("/"):
                target.session["sid"] = "fresh-dialog"
                state = target.assess_discovery(
                    [],
                    "Я репетитор по английскому, работаю с детьми от 10 лет и взрослыми.",
                    target.EXPERT_PROFILES["marketer"],
                )
                self.assertTrue(state["identity"])
                self.assertFalse(state["task"])
                self.assertEqual(target.discovery_action(state, target.EXPERT_PROFILES["marketer"], ""), "explore_task")
        finally:
            target.gigachat = old

    def test_task_question_cannot_invent_problem_from_audience(self):
        issues = target.phase_reply_issues(
            "Расскажите подробнее, какие сложности возникают у вас при работе с разными возрастными группами?",
            "explore_task",
            [],
        )
        self.assertIn("проблема выведена из профессии или аудитории клиента", issues)

    def test_task_question_cannot_presuppose_a_difficulty(self):
        issues = target.phase_reply_issues(
            "Какую именно трудность вы испытываете при работе с учениками?",
            "explore_task",
            [],
        )
        self.assertIn("вопрос заранее приписывает клиенту проблему", issues)

    def test_discovery_rejects_internal_commentary_and_professional_smalltalk(self):
        issues = target.phase_reply_issues(
            "Похоже, клиент уверен в своей компетенции, однако стоит уточнить детали работы. Какой аспект вашей работы приносит наибольшее удовлетворение?",
            "explore_task",
            [],
        )
        self.assertIn("наружу выведено служебное рассуждение о клиенте", issues)
        self.assertIn("вопрос ушёл от рабочей задачи к общему разговору о профессии", issues)
        self.assertTrue(target.blocking_reply_issues(issues))

    def test_two_discovery_questions_are_blocking(self):
        issues = target.phase_reply_issues(
            "Какую трудность вы испытываете? Какие сложности возникают во время занятий?",
            "explore_task",
            [],
        )
        self.assertIn("больше одного вопроса", target.blocking_reply_issues(issues))

    def test_task_question_cannot_offer_answer_variants(self):
        issues = target.phase_reply_issues(
            "Что вам хотелось бы улучшить — сделать уроки интереснее, проще объяснять материал или найти новые методики?",
            "explore_task",
            [],
        )
        self.assertIn("варианты ответа внутри вопроса", issues)

    def test_reply_to_pending_task_advances_without_keyword_guessing(self):
        class EmptyExtractor:
            def reply(self, messages):
                return '{"stages":{}}'
        old = target.gigachat
        target.gigachat = EmptyExtractor()
        profile = target.EXPERT_PROFILES["marketer"]
        try:
            with target.app.test_request_context("/"):
                target.session["sid"] = "pending-task"
                target.session["discovery_sid"] = "pending-task"
                target.session["discovery_state"] = {"identity": True, "task": False, "ai_experience": False}
                target.session["pending_discovery_sid"] = "pending-task"
                target.session["pending_discovery_stage"] = "task"
                history = [
                    {"role": "user", "content": "Я репетитор английского."},
                    {"role": "assistant", "content": "Что в работе вы хотели бы изменить?"},
                ]
                text = "Быстрее готовиться к урокам и делать их актуальнее для подростков."
                state = target.assess_discovery(history, text, profile, {"intent": "other", "subject": "other"})
                self.assertTrue(state["task"])
                self.assertEqual(target.discovery_action(state, profile, text), "explore_ai_experience")
        finally:
            target.gigachat = old

    def test_refusal_to_answer_pauses_discovery(self):
        state = {"identity": True, "task": True, "ai_experience": False, "solution_explained": False, "solution_interest": False}
        action = target.discovery_action(
            state,
            target.EXPERT_PROFILES["marketer"],
            "Мне уже расхотелось отвечать на вопросы.",
            {"intent": "refusal", "subject": "other"},
        )
        self.assertEqual(action, "pause_discovery")

    def test_denial_of_suggested_problem_does_not_complete_need_stage(self):
        class MisleadingExtractor:
            def reply(self, messages):
                return '{"stages":{"identity":{"complete":true,"evidence":"Я репетитор по английскому"},"task":{"complete":true,"evidence":"сложности меня не пугают"},"ai_experience":{"complete":false,"evidence":""}}}'
        old = target.gigachat
        target.gigachat = MisleadingExtractor()
        profile = target.EXPERT_PROFILES["marketer"]
        try:
            with target.app.test_request_context("/"):
                target.session["sid"] = "denied-assumption"
                history = [
                    {"role": "user", "content": "Я репетитор по английскому, работаю с детьми и взрослыми."},
                    {"role": "assistant", "content": "Какие сложности возникают у вас во время занятий?"},
                ]
                state = target.assess_discovery(
                    history,
                    "Во время занятий — никаких. Я работаю 18 лет, сложности меня не пугают.",
                    profile,
                )
                self.assertTrue(state["identity"])
                self.assertFalse(state["task"])
                self.assertEqual(target.discovery_action(state, profile, ""), "explore_task")
        finally:
            target.gigachat = old

    def test_marketer_discovery_advances_from_direct_answers_without_model_labels(self):
        class EmptyExtractor:
            def reply(self, messages):
                return '{"stages":{"identity":{"complete":false,"evidence":""},"task":{"complete":false,"evidence":""},"ai_experience":{"complete":false,"evidence":""}}}'
        old = target.gigachat
        target.gigachat = EmptyExtractor()
        profile = target.EXPERT_PROFILES["marketer"]
        try:
            with target.app.test_request_context("/"):
                target.session["sid"] = "sequence-dialog"
                identity = "Я репетитор по английскому, работаю с детьми от 10 лет и взрослыми."
                state = target.assess_discovery([], identity, profile)
                self.assertEqual(target.discovery_action(state, profile, identity), "explore_task")

                history = [
                    {"role": "user", "content": identity},
                    {"role": "assistant", "content": "А что в вашей работе сейчас хотелось бы упростить или изменить?"},
                ]
                need = "Хотелось бы сократить время подготовки к урокам. Очень много времени занимает."
                state = target.assess_discovery(history, need, profile)
                self.assertTrue(state["task"])
                self.assertEqual(target.discovery_action(state, profile, need), "explore_ai_experience")

                history += [
                    {"role": "user", "content": need},
                    {"role": "assistant", "content": "Пробовали уже решать эту задачу с помощью нейросетей?"},
                ]
                attempt = "Пробовала, мне не понравилось. Всё равно многое приходится переделывать вручную."
                state = target.assess_discovery(history, attempt, profile)
                self.assertTrue(state["ai_experience"])
                self.assertEqual(target.discovery_action(state, profile, attempt), "explain_solution")
        finally:
            target.gigachat = old

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
