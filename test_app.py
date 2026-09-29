import io
import json
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
        self.models = []
        self.response_formats = []

    def reply(self, messages, model=None, response_format=None):
        self.prompts.append(messages[0]["content"])
        self.models.append(model)
        self.response_formats.append(response_format)
        if not self.replies:
            raise RuntimeError("No scripted reply")
        value = self.replies.pop(0)
        if isinstance(value, Exception):
            raise value
        return value


def payload(reply, action="explore", intent="continue", evidence="", observations=None, question_target=None, conversation_effect="continue"):
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
        "question_target": question_target or ("need" if action == "explore" else "none"),
        "question_scope": "personal_situation" if action == "explore" else "none",
        "conversation_effect": conversation_effect,
        "proposed_solution": {
            "type": "consultation" if action == "offer_consultation" else "none",
            "name": "консультация" if action == "offer_consultation" else "",
            "evidence": "",
        },
        "answer_grounding": {
            "source": "dialogue" if action == "answer_information" else "none",
            "evidence": evidence if action == "answer_information" else "",
        },
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
            "gender_matches_expert": True,
            "adds_or_repeats_consultation_offer": False,
            "leaks_internal_instructions": False,
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

    def test_complete_information_allows_consultation_after_one_question(self):
        state = self.state(contact=True, need=True, previous_experience=True, desired_result=True)
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

    def test_question_after_offer_is_answered_even_if_intent_was_misclassified(self):
        state = self.state(
            need=True,
            previous_experience=True,
            desired_result=True,
            consultation_offered=True,
        )
        text = "Как часто проходят встречи?"
        self.assertEqual(
            target.expected_dialog_action(state, "continue", "", text),
            "answer_information",
        )

    def test_consultation_is_not_offered_again_after_first_offer(self):
        state = self.state(
            need=True,
            previous_experience=True,
            desired_result=True,
            consultation_offered=True,
        )
        self.assertEqual(
            target.expected_dialog_action(state, "continue", "", "Понятно"),
            "check_interest",
        )

    def test_interest_does_not_block_offer_when_psychologist_request_is_complete(self):
        text = "Да, мне интересно"
        state = self.state(contact=True, need=True, previous_experience=True, desired_result=True, diagnostic_questions=1)
        self.assertEqual(target.expected_dialog_action(state, "interest", text, text), "offer_consultation")

    def test_interest_after_solution_triggers_one_offer(self):
        text = "Да, мне интересно"
        state = self.state(solution_explained=True)
        self.assertEqual(target.expected_dialog_action(state, "interest", text, text), "offer_consultation")
        state["consultation_offered"] = True
        self.assertEqual(target.expected_dialog_action(state, "interest", text, text), "check_interest")

    def test_semantic_booking_intent_starts_calendar_after_offer(self):
        state = self.state(consultation_offered=True)
        for text in ("Да, начинаем запись", "Хорошо, подберите время", "Я готова записаться"):
            self.assertEqual(
                target.expected_dialog_action(state, "booking", text, text),
                "start_booking",
            )

    def test_information_question_has_priority_over_booking_consent(self):
        text = (
            "Давайте. В зуме? Я предпочитаю Телемост. "
            "Сколько длится такая встреча?"
        )
        state = self.state(consultation_offered=True)
        self.assertEqual(
            target.expected_dialog_action(state, "booking_question", text, text),
            "answer_information",
        )

    def test_mixed_booking_reply_answers_questions_without_starting_calendar(self):
        text = (
            "Давайте. В зуме? Я предпочитаю Телемост. "
            "Сколько длится такая встреча?"
        )
        target.gigachat = ScriptedGigaChat([
            payload(
                "Да, можно встретиться через Телемост. "
                "Бесплатная консультация длится 15–20 минут. "
                "Подходящее время выберите сами в календаре.",
                action="start_booking",
                intent="booking_question",
                evidence=text,
            ),
        ])
        with target.app.test_request_context("/"):
            target.session["dialog_controller_state"] = self.state(
                consultation_offered=True,
            )
            answer, action, _ = target.generate_stateful_dialog_reply(
                [{"role": "assistant", "content": "Записаться можно?"}],
                text,
                "Встреча проходит онлайн и длится 15–20 минут.",
            )
        self.assertEqual(action, "answer_information")
        self.assertEqual(
            answer,
            "Да, можно встретиться через Телемост. "
            "Бесплатная консультация длится 15–20 минут.",
        )
        self.assertNotIn("календар", answer.lower())
        self.assertNotIn("дату и время", answer)

    def test_booking_intent_cannot_skip_consultation_offer(self):
        text = "Я готова записаться"
        self.assertNotEqual(
            target.expected_dialog_action(self.state(), "booking", text, text),
            "start_booking",
        )

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

    def test_concern_about_conditions_stays_in_information_cycle(self):
        text = "Надеюсь, это не растянется на годы."
        state = self.state(consultation_offered=True)
        self.assertEqual(
            target.expected_dialog_action(state, "concern", text, text),
            "answer_information",
        )

    def test_controller_prompt_distinguishes_concern_from_end(self):
        prompt = target.controller_prompt(
            self.state(consultation_offered=True),
            "База",
            [],
            "Надеюсь, это не растянется на годы.",
        )
        self.assertIn("продолжает информационный цикл", prompt)
        self.assertIn("end означает только явно выраженное намерение", prompt)

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

    def test_complete_information_does_not_force_second_diagnostic_question(self):
        state = self.state(
            need=True,
            previous_experience=True,
            desired_result=True,
            diagnostic_questions=1,
        )
        self.assertEqual(
            target.expected_dialog_action(state, "continue", "", "Хочу вернуть радость"),
            "offer_consultation",
        )

    def test_explore_question_cannot_target_information_already_received(self):
        state = self.state(
            need=True,
            previous_experience=True,
            desired_result=False,
            diagnostic_questions=2,
        )
        data = target.parse_controller_payload(payload(
            "Давно ли это продолжается?",
            question_target="previous_experience",
        ))
        issues = target.controller_reply_issues(data, "explore", state, [])
        self.assertIn("вопрос относится к уже полученной информации", issues)

    def test_explore_question_may_target_only_missing_information(self):
        state = self.state(
            need=True,
            previous_experience=True,
            desired_result=False,
            diagnostic_questions=2,
        )
        data = target.parse_controller_payload(payload(
            "Что вы хотели бы изменить?",
            question_target="desired_result",
        ))
        issues = target.controller_reply_issues(data, "explore", state, [])
        self.assertNotIn("вопрос относится к уже полученной информации", issues)

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

    def test_repeated_sentence_is_rejected_when_new_answer_is_appended(self):
        previous = "Предлагаю встретиться в Телемосте. Выберите удобное время."
        data = target.parse_controller_payload(payload(
            "Предлагаю встретиться в Телемосте. Бесплатная консультация длится 15-20 минут.",
            action="answer_information",
            intent="question",
            evidence="Сколько длится консультация?",
        ))
        issues = target.controller_reply_issues(
            data,
            "answer_information",
            self.state(consultation_offered=True),
            [{"role": "assistant", "content": previous}],
            client_text="Сколько длится консультация?",
        )
        self.assertIn("повторена предыдущая реплика", issues)

    def test_new_information_answer_without_old_platform_is_accepted(self):
        previous = "Предлагаю встретиться в Телемосте. Выберите удобное время."
        data = target.parse_controller_payload(payload(
            "Бесплатная консультация длится 15-20 минут.",
            action="answer_information",
            intent="question",
            evidence="Сколько длится консультация?",
        ))
        issues = target.controller_reply_issues(
            data,
            "answer_information",
            self.state(consultation_offered=True),
            [{"role": "assistant", "content": previous}],
            client_text="Сколько длится консультация?",
        )
        self.assertNotIn("повторена предыдущая реплика", issues)

    # Reply validation
    def test_no_repeated_question(self):
        state = self.state(asked_questions=["Давно у вас такое состояние?"])
        data = target.parse_controller_payload(payload("Давно у вас такое состояние?"))
        issues = target.controller_reply_issues(data, "explore", state, [])
        self.assertIn("повторён уже заданный вопрос", issues)

    def test_directive_and_question_are_normalized_to_one_request(self):
        reply = (
            "Расскажите подробнее, давно это ощущение появилось. "
            "Оно мешает повседневной жизни?"
        )
        self.assertEqual(target.semantic_information_request_count(reply), 2)
        normalized = target.normalize_reply_for_action(reply, "explore")
        self.assertEqual(
            normalized,
            "Расскажите подробнее, давно это ощущение появилось?",
        )
        data = target.parse_controller_payload(payload(normalized))
        issues = target.controller_reply_issues(
            data, "explore", self.state(), []
        )
        self.assertNotIn("задано больше одного смыслового вопроса", issues)

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

    def test_information_answer_cannot_close_active_dialogue(self):
        client_text = "Как часто проходят встречи?"
        data = target.parse_controller_payload(payload(
            "Обычно встречи проходят раз в неделю. До встречи! Хорошего дня.",
            action="answer_information",
            intent="question",
            evidence=client_text,
            conversation_effect="close",
        ))
        issues = target.controller_reply_issues(
            data,
            "answer_information",
            self.state(consultation_offered=True),
            [],
            client_text=client_text,
        )
        self.assertIn("реплика преждевременно завершает продолжающийся разговор", issues)

    def test_information_answer_keeps_dialogue_open_without_forced_follow_up(self):
        client_text = "Как часто проходят встречи?"
        data = target.parse_controller_payload(payload(
            "Обычно встречи проходят раз в неделю.",
            action="answer_information",
            intent="question",
            evidence=client_text,
        ))
        issues = target.controller_reply_issues(
            data,
            "answer_information",
            self.state(consultation_offered=True),
            [],
            client_text=client_text,
        )
        self.assertNotIn("реплика преждевременно завершает продолжающийся разговор", issues)

    def test_expert_cannot_speak_in_third_person(self):
        data = target.parse_controller_payload(payload("Психолог поможет разобраться.","explain_solution"))
        issues = target.controller_reply_issues(data, "explain_solution", self.state(), [])
        self.assertIn("эксперт говорит о себе в третьем лице", issues)

    def test_informal_address_is_rejected(self):
        data = target.parse_controller_payload(payload("Давай попробуем разобраться?"))
        issues = target.controller_reply_issues(data, "explore", self.state(), [])
        self.assertIn("нарушено обращение на вы", issues)

    def test_offer_does_not_repeat_clients_desired_result(self):
        client_text = "Вернуть веру в себя и радость жизни."
        data = target.parse_controller_payload(payload(
            "Моя помощь направлена на возвращение веры в себя и радости жизни. Предлагаю встретиться.",
            action="offer_consultation",
        ))
        issues = target.controller_reply_issues(
            data,
            "offer_consultation",
            self.state(),
            [],
            client_text=client_text,
        )
        self.assertIn("дословно пересказан ответ клиента", issues)

    def test_offer_may_connect_situation_with_expert_specialization(self):
        client_text = "Вернуть веру в себя и радость жизни."
        data = target.parse_controller_payload(payload(
            "Я работаю с такими ситуациями. Предлагаю обсудить вашу ситуацию на первой встрече.",
            action="offer_consultation",
        ))
        issues = target.controller_reply_issues(
            data,
            "offer_consultation",
            self.state(),
            [],
            client_text=client_text,
        )
        self.assertNotIn("дословно пересказан ответ клиента", issues)

    def test_team_voice_is_forbidden_by_system_rules(self):
        self.assertIn("не организация и не команда", target.SYSTEM_RULES)
        self.assertIn("будем рады", target.SYSTEM_RULES)

    def test_second_structured_attempt_advances_state(self):
        invalid = ScriptedGigaChat([
            "invalid json",
            payload("Давно у вас такое состояние?"),
        ])
        target.gigachat = invalid
        with target.app.test_request_context("/"):
            answer, action, state = target.generate_stateful_dialog_reply([], "Да ерунда какая-то, ничего не хочу.", "База")
        self.assertEqual(action, "explore")
        self.assertEqual(answer.count("?"), 1)
        self.assertEqual(state["diagnostic_questions"], 1)

    def test_structured_dialog_uses_lite_with_native_schema_first(self):
        scripted = ScriptedGigaChat([
            payload("Давно у вас такое состояние?"),
        ])
        target.gigachat = scripted
        with target.app.test_request_context("/"):
            target.generate_stateful_dialog_reply(
                [],
                "Да ерунда какая-то, ничего не хочу.",
                "База",
            )
        self.assertEqual(scripted.models, ["GigaChat"])
        self.assertEqual(scripted.response_formats, [target.CONTROLLER_RESPONSE_FORMAT])

    def test_dialog_timing_logs_do_not_change_generated_reply(self):
        target.gigachat = ScriptedGigaChat([
            payload("Давно у вас такое состояние?"),
        ])
        with target.app.test_request_context("/"), self.assertLogs(target.app.logger, level="INFO") as captured:
            answer, action, state = target.generate_stateful_dialog_reply(
                [],
                "Да ерунда какая-то, ничего не хочу.",
                "База",
            )
        self.assertEqual(answer, "Давно у вас такое состояние?")
        self.assertEqual(action, "explore")
        self.assertEqual(state["diagnostic_questions"], 1)
        joined = "\n".join(captured.output)
        self.assertIn("Dialog generation completed", joined)
        self.assertIn("attempts=1", joined)
        self.assertIn("prompt_chars=", joined)

    def test_rejected_dialog_log_contains_validation_reason(self):
        target.gigachat = ScriptedGigaChat([
            payload("Давайте запишемся на консультацию?"),
            payload("Давно у вас такое состояние?"),
        ])
        with target.app.test_request_context("/"), self.assertLogs(target.app.logger, level="WARNING") as captured:
            target.generate_stateful_dialog_reply(
                [],
                "Да ерунда какая-то, ничего не хочу.",
                "База",
            )
        joined = "\n".join(captured.output)
        self.assertIn("reason=validation", joined)
        self.assertIn("преждевременно предложена встреча", joined)

    def test_grounded_observations_survive_a_rejected_attempt(self):
        client_text = "Года два, но бывает то лучше, то хуже."
        scripted = ScriptedGigaChat([
            payload(
                "Давно у вас такое состояние?",
                observations={
                    "previous_experience": {"present": True, "evidence": "Года два"},
                },
                question_target="previous_experience",
            ),
            payload("Что хотелось бы изменить?", question_target="desired_result"),
        ])
        target.gigachat = scripted
        with target.app.test_request_context("/"):
            target.session["dialog_controller_state"] = self.state(
                need=True,
                diagnostic_questions=1,
                asked_questions=["Давно у вас такое состояние?"],
            )
            answer, action, state = target.generate_stateful_dialog_reply(
                [{"role": "assistant", "content": "Давно у вас такое состояние?"}],
                client_text,
                "База",
            )
        self.assertEqual(answer, "Что хотелось бы изменить?")
        self.assertEqual(action, "explore")
        self.assertTrue(state["previous_experience"])
        self.assertIn('"previous_experience": true', scripted.prompts[1])

    def test_empty_assistant_history_ignores_stale_session_state(self):
        scripted = ScriptedGigaChat([
            payload("Давно у вас такое состояние?"),
        ])
        target.gigachat = scripted
        with target.app.test_request_context("/"):
            target.session["dialog_controller_state"] = self.state(
                need=True,
                previous_experience=True,
                desired_result=True,
                diagnostic_questions=3,
                asked_questions=["Давно у вас такое состояние?"],
                consultation_offered=True,
            )
            answer, action, state = target.generate_stateful_dialog_reply(
                [],
                "Да ерунда какая-то, ничего не хочу.",
                "База",
            )
        self.assertEqual(answer, "Давно у вас такое состояние?")
        self.assertEqual(action, "explore")
        self.assertEqual(state["diagnostic_questions"], 1)
        self.assertFalse(state["consultation_offered"])

    def test_retry_prompt_keeps_controller_action_and_missing_state(self):
        client_text = "Никаких мыслей о будущем. Нет надежды."
        scripted = ScriptedGigaChat([
            payload("Предлагаю встретиться на консультации.", action="offer_consultation"),
            payload("Что хотелось бы изменить?", question_target="desired_result"),
        ])
        target.gigachat = scripted
        with target.app.test_request_context("/"):
            target.session["dialog_controller_state"] = self.state(
                need=True,
                previous_experience=True,
                diagnostic_questions=2,
            )
            answer, action, _ = target.generate_stateful_dialog_reply(
                [{"role": "assistant", "content": "Какие мысли возникают о будущем?"}],
                client_text,
                "База",
            )
        self.assertEqual(answer, "Что хотелось бы изменить?")
        self.assertEqual(action, "explore")
        self.assertIn("обязательное следующее действие: explore", scripted.prompts[1])
        self.assertIn("Недостающие элементы состояния: desired_result", scripted.prompts[1])

    def test_two_question_marks_are_normalized_in_main_pipeline(self):
        target.gigachat = ScriptedGigaChat([
            "invalid json",
            payload("Давно это продолжается? Как состояние влияет на вашу жизнь?"),
        ])
        with target.app.test_request_context("/"):
            answer, action, state = target.generate_stateful_dialog_reply(
                [],
                "Да ерунда какая-то, пустота, ничего не хочу.",
                "Кирилл — психолог.",
            )
        self.assertEqual(action, "explore")
        self.assertEqual(answer.count("?"), 1)
        self.assertEqual(state["diagnostic_questions"], 1)

    def test_information_answer_keeps_facts_and_removes_follow_up_question(self):
        target.gigachat = ScriptedGigaChat([
            payload(
                "Первая консультация проходит онлайн и длится 15–20 минут. "
                "Хотите записаться?",
                action="answer_information",
                intent="question",
                evidence="как проходит первая консультация и сколько она длится",
            ),
        ])
        with target.app.test_request_context("/"):
            answer, action, _ = target.generate_stateful_dialog_reply(
                [],
                "А как проходит первая консультация и сколько она длится?",
                "Первая консультация проходит онлайн и длится 15–20 минут.",
            )
        self.assertEqual(action, "answer_information")
        self.assertEqual(
            answer,
            "Первая консультация проходит онлайн и длится 15–20 минут.",
        )
        self.assertNotIn("?", answer)

    def test_information_reply_containing_only_follow_up_question_is_retried(self):
        target.gigachat = ScriptedGigaChat([
            payload(
                "Хотите записаться?",
                action="answer_information",
                intent="question",
                evidence="сколько она длится",
            ),
            payload(
                "Первая консультация длится 15–20 минут.",
                action="answer_information",
                intent="question",
                evidence="сколько она длится",
            ),
        ])
        with target.app.test_request_context("/"):
            answer, action, _ = target.generate_stateful_dialog_reply(
                [],
                "А сколько она длится?",
                "Первая консультация длится 15–20 минут.",
            )
        self.assertEqual(action, "answer_information")
        self.assertEqual(answer, "Первая консультация длится 15–20 минут.")

    def test_malformed_json_is_repaired_without_leaking_service_data(self):
        malformed = payload("Давно у вас такое состояние?").replace('":', '"\\:')
        target.gigachat = ScriptedGigaChat([malformed])
        with target.app.test_request_context("/"):
            answer, action, _ = target.generate_stateful_dialog_reply(
                [], "Да ерунда какая-то, ничего не хочу.", "База"
            )
        self.assertEqual(answer, "Давно у вас такое состояние?")
        self.assertEqual(action, "explore")
        self.assertFalse(answer.startswith("{"))

    def test_pr8_payload_without_reply_assessment_remains_valid(self):
        old_payload = __import__("json").loads(payload("Давно у вас такое состояние?"))
        old_payload.pop("reply_assessment")
        parsed = target.parse_controller_payload(
            __import__("json").dumps(old_payload, ensure_ascii=False)
        )
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed["reply_assessment"], {})

    def test_repeated_consultation_offer_is_retried_in_same_pipeline(self):
        repeated = payload("Хотите записаться?", action="answer_information", intent="question", evidence="Сколько длится?")
        repeated_data = __import__("json").loads(repeated)
        answer_payload = payload("Первая встреча длится 20 минут.", action="answer_information", intent="question", evidence="Сколько длится?")
        target.gigachat = ScriptedGigaChat([
            __import__("json").dumps(repeated_data, ensure_ascii=False),
            answer_payload,
        ])
        with target.app.test_request_context("/"):
            target.session["dialog_controller_state"] = self.state(consultation_offered=True)
            answer, action, _ = target.generate_stateful_dialog_reply(
                [], "Сколько длится?", "Кирилл — психолог. Первая встреча — 20 минут."
            )
        self.assertEqual(answer, "Первая встреча длится 20 минут.")
        self.assertEqual(action, "answer_information")

    def test_booking_intent_transfers_dialogue_to_calendar_flow(self):
        text = "Ну да, начинаем. Подберём время?"
        target.gigachat = ScriptedGigaChat([
            payload(
                "Сейчас пришлю ссылку на календарь.",
                action="start_booking",
                intent="booking",
                evidence=text,
            ),
        ])
        with self.client.session_transaction() as session:
            session["sid"] = "booking-dialog"
            session["dialog_controller_state"] = self.state(consultation_offered=True)
            session["consultation_offered"] = True
        con = target.db()
        con.execute(
            "insert into messages values(?,?,?,?)",
            ("booking-dialog", "assistant", "Предлагаю первую консультацию.", 1),
        )
        con.commit()
        response = self.client.post("/api/chat", json={"message": text})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.get_json()["answer"],
            "Назовите удобные дату и время — я проверю их в календаре.",
        )
        with self.client.session_transaction() as session:
            self.assertEqual(session["requested_booking_type"], "free")

    def test_dynamic_fallback_answers_when_all_controller_payloads_are_malformed(self):
        target.gigachat = ScriptedGigaChat([
            "not json",
            "still not json",
            "invalid",
            "invalid again",
            "Давно это состояние влияет на вашу повседневную жизнь?",
        ])
        with target.app.test_request_context("/"):
            answer, action, _ = target.generate_stateful_dialog_reply(
                [],
                "Да ерунда какая-то, пустота, ничего не хочу.",
                "Кирилл — психолог.",
            )
        self.assertEqual(action, "explore")
        self.assertEqual(
            answer,
            "Давно это состояние влияет на вашу повседневную жизнь?",
        )
        self.assertIsNone(target.gigachat.response_formats[-1])

    def test_dynamic_fallback_keeps_controller_action_after_rejected_replies(self):
        text = "Вернуть веру в себя и радость жизни."
        observations = {
            "desired_result": {"present": True, "evidence": text},
        }
        rejected = payload(
            " ".join(["лишнее"] * 60),
            action="offer_consultation",
            observations=observations,
        )
        target.gigachat = ScriptedGigaChat([
            rejected,
            rejected,
            rejected,
            rejected,
            "Я как раз работаю с такими состояниями. Предлагаю сначала встретиться на короткой бесплатной консультации и понять, подходим ли мы друг другу.",
        ])
        with target.app.test_request_context("/"):
            target.session["dialog_controller_state"] = self.state(
                contact=True,
                need=True,
                previous_experience=True,
                diagnostic_questions=2,
            )
            answer, action, _ = target.generate_stateful_dialog_reply(
                [
                    {"role": "assistant", "content": "Что бы вы хотели изменить?"},
                ],
                text,
                "Кирилл — психолог. Первая консультация бесплатная.",
            )
        self.assertEqual(action, "offer_consultation")
        self.assertIn("Предлагаю", answer)
        self.assertIsNone(target.gigachat.response_formats[-1])

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

    def test_booking_condition_question_does_not_start_calendar_flow(self):
        with target.app.test_request_context("/"):
            target.session["consultation_offered"] = True
            answer = target.direct_booking_answer(
                "То есть можно сначала записаться на бесплатную консультацию, а потом решить, хочу ли я продолжать?"
            )
            self.assertIsNone(answer)
            self.assertNotIn("awaiting_booking_type", target.session)

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

    def test_post_booking_question_is_not_mistaken_for_farewell(self):
        with target.app.test_request_context("/"):
            target.session["last_booking"] = "Бесплатная консультация, 28.09.2026 в 14:00, 20 минут"
            answer = target.completed_dialog_answer("Хорошо, спасибо. А сколько стоят сеансы?")
            self.assertIsNone(answer)
            self.assertFalse(target.session.get("dialog_closed"))

    # UI, persistence and admin
    def test_input_is_not_disabled_while_reply_pending(self):
        page = self.client.get("/").get_data(as_text=True)
        self.assertNotIn("input.disabled=true", page)
        self.assertIn("pending", page)

    def test_history_is_restored(self):
        page = self.client.get("/").get_data(as_text=True)
        self.assertIn("fetch('/api/history')", page)

    def test_post_booking_question_uses_simple_text_reply_without_controller_schema(self):
        target.gigachat = ScriptedGigaChat([
            "Один регулярный сеанс стоит 3500 рублей.",
        ])
        with self.client.session_transaction() as session:
            session["sid"] = "booked-dialog"
            session["last_booking"] = "Бесплатная консультация, 28.09.2026 в 15:00, 20 минут"
            session["consultation_offered"] = True
        con = target.db()
        con.execute(
            "insert into messages values(?,?,?,?)",
            ("booked-dialog", "assistant", "Запись подтверждена.", 1),
        )
        con.commit()
        response = self.client.post(
            "/api/chat",
            json={"message": "Хорошо, спасибо. А сколько стоят сеансы?"},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["answer"], "Один регулярный сеанс стоит 3500 рублей.")
        self.assertEqual(target.gigachat.response_formats, [None])
        self.assertIn("Запись клиента уже подтверждена", target.gigachat.prompts[0])

    def test_post_booking_reply_does_not_address_client_by_name(self):
        target.gigachat = ScriptedGigaChat([
            "Пожалуйста, Евгений. До свидания!",
            "Пожалуйста. До свидания!",
        ])
        history = [
            {
                "role": "user",
                "content": "Евгения, +79111234567, testclient14@yandex.ru",
            },
            {"role": "assistant", "content": "Запись подтверждена."},
        ]
        with target.app.test_request_context("/"):
            answer = target.generate_post_booking_reply(
                history,
                "Ясно, спасибо. До встречи!",
                [{"name": "FAQ.txt", "text": "Подготовка не требуется."}],
            )
        self.assertEqual(answer, "Пожалуйста. До свидания!")
        self.assertEqual(len(target.gigachat.prompts), 2)
        self.assertIn("Не обращайтесь к клиенту по имени", target.gigachat.prompts[0])

    def test_post_booking_retrieval_uses_full_knowledge_base_for_unresolved_question(self):
        target.gigachat = ScriptedGigaChat([
            "Один регулярный сеанс стоит 3500 рублей.",
        ])
        documents = [{
            "name": "FAQ.txt",
            "text": (
                "Сколько стоят регулярные сеансы? Один сеанс стоит 3500 рублей.\n\n"
                + "\n\n".join("Другой раздел базы знаний без сведений о цене." for _ in range(250))
            ),
        }]
        history = [
            {"role": "user", "content": "Сколько стоят регулярные сеансы?"},
            {"role": "assistant", "content": "Стоимость можно уточнить позднее."},
        ]
        with target.app.test_request_context("/"):
            answer = target.generate_post_booking_reply(
                history,
                "Я у вас и спрашиваю. Сколько стоят сеансы?",
                documents,
            )
        self.assertEqual(answer, "Один регулярный сеанс стоит 3500 рублей.")
        self.assertIn("Один сеанс стоит 3500 рублей", target.gigachat.prompts[0])

    def test_empty_server_history_clears_stale_closed_session(self):
        with self.client.session_transaction() as session:
            session["sid"] = "missing-history"
            session["dialog_closed"] = True
            session["last_booking"] = "Старая запись"
        response = self.client.get("/api/history")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {"messages": [], "closed": False})
        with self.client.session_transaction() as session:
            self.assertNotIn("dialog_closed", session)
            self.assertNotIn("last_booking", session)

    def test_first_message_is_not_blocked_by_stale_closed_session(self):
        target.gigachat = ScriptedGigaChat([
            payload("Давно у вас такое состояние?"),
        ])
        with self.client.session_transaction() as session:
            session["sid"] = "missing-history"
            session["dialog_closed"] = True
            session["last_booking"] = "Старая запись"
        response = self.client.post(
            "/api/chat",
            json={"message": "Да ерунда какая-то, пустота, ничего не хочу."},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["answer"], "Давно у вас такое состояние?")
        self.assertFalse(response.get_json()["closed"])

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

    def test_named_platform_cannot_be_replaced_by_similar_service(self):
        client_text = "А в Телемосте можно?"
        response = target.parse_controller_payload(payload(
            "Да, можно в Телеграме.",
            action="answer_information",
            intent="question",
            evidence=client_text,
        ))
        issues = target.controller_reply_issues(
            response,
            "answer_information",
            target.controller_state({}),
            [],
            client_text=client_text,
        )
        self.assertIn("подменено название варианта из вопроса клиента", issues)

    def test_named_platform_is_accepted_when_preserved(self):
        client_text = "А в Телемосте можно?"
        response = target.parse_controller_payload(payload(
            "Да, через Телемост можно провести встречу.",
            action="answer_information",
            intent="question",
            evidence=client_text,
        ))
        issues = target.controller_reply_issues(
            response,
            "answer_information",
            target.controller_state({}),
            [],
            client_text=client_text,
        )
        self.assertNotIn("подменено название варианта из вопроса клиента", issues)

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

    # Marketer profile remains separate from the psychologist profile.
    def test_marketer_has_separate_page_and_greeting(self):
        response = self.client.get("/marketer")
        self.assertEqual(response.status_code, 200)
        self.assertIn("чем вы занимаетесь и с кем работаете", response.get_data(as_text=True))
        self.assertIn("/api/marketer/chat", response.get_data(as_text=True))

    def test_marketer_knowledge_base_is_isolated(self):
        headers = {"X-Admin-Password": "admin123"}
        response = self.client.post(
            "/api/marketer/admin/upload",
            headers=headers,
            data={"files": (io.BytesIO("Екатерина — маркетолог".encode()), "marketing.txt")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        marketer_docs = self.client.get("/api/marketer/admin", headers=headers).json["documents"]
        psychologist_docs = self.client.get("/api/admin", headers=headers).json["documents"]
        self.assertTrue(any(x["name"] == "marketing.txt" for x in marketer_docs))
        self.assertFalse(any(x["name"] == "marketing.txt" for x in psychologist_docs))

    def test_marketer_explains_solution_before_consultation(self):
        state = self.state(need=True, previous_experience=True, desired_result=True)
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            self.assertEqual(
                target.expected_dialog_action(state, "continue", "", "Хочу наладить контент"),
                "explain_solution",
            )
            state["solution_explained"] = True
            text = "Да, это мне подходит"
            self.assertEqual(
                target.expected_dialog_action(state, "interest", text, text),
                "offer_consultation",
            )

    def test_marketer_prompt_keeps_ai_expert_role(self):
        state = self.state(need=True)
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            prompt = target.controller_prompt(
                state,
                "Екатерина создаёт персональных AI-ассистентов и готовые решения для экспертов.",
                [
                    {"role": "user", "content": "Я репетитор английского."},
                    {"role": "assistant", "content": "Что в работе хотелось бы изменить?"},
                ],
                "Долго готовлю материалы к урокам.",
            )
        self.assertIn("специалиста по нейросетям и ИИ-решениям", prompt)
        self.assertIn("Профессия клиента описывает только контекст", prompt)
        self.assertIn("не становится профессией эксперта", prompt)
        self.assertIn("Не запрашивайте частный пример", prompt)
        self.assertNotIn("Для маркетолога сначала объяснить", prompt)

    def test_marketer_stops_diagnosis_after_three_questions_and_explains_solution(self):
        state = self.state(need=True, diagnostic_questions=3)
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            self.assertEqual(
                target.expected_dialog_action(state, "continue", "", "Хочу быстрее готовиться"),
                "explain_solution",
            )

    def test_substantive_answer_completes_pending_dialog_target(self):
        state = self.state(need=True)
        state["pending_question_target"] = "previous_experience"
        updated = target.apply_pending_answer(
            state,
            "Пытаюсь создавать упражнения в нейросетях, но результат приходится переделывать.",
            "continue",
        )
        self.assertTrue(updated["previous_experience"])
        self.assertEqual(updated["pending_question_target"], "none")

    def test_marketer_offer_cannot_be_another_diagnostic_question(self):
        state = self.state(need=True, previous_experience=True, desired_result=True)
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            issues = target.controller_reply_issues(
                {
                    "reply": "Насколько глубоко вы изучали этот подход?",
                    "action": "offer_consultation",
                    "question_target": "none",
                    "conversation_effect": "continue",
                },
                "offer_consultation",
                state,
                [],
                client_text="Хочу освободить время",
            )
        self.assertIn("вместо предложения консультации продолжена диагностика", issues)

    def test_marketer_solution_must_be_kb_grounded_ai_offer(self):
        state = self.state(need=True, previous_experience=True, desired_result=True)
        context = "Готовый ассистент-методист помогает создавать нестандартные курсы. Возможна разработка персонального AI-ассистента под задачу клиента."
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            valid = {
                "reply": "Для вашей задачи подойдёт готовый ассистент-методист для нестандартных курсов.",
                "action": "explain_solution",
                "question_target": "none",
                "conversation_effect": "continue",
                "proposed_solution": {
                    "type": "ready_product",
                    "name": "Готовый ассистент-методист",
                    "evidence": "Готовый ассистент-методист помогает создавать нестандартные курсы",
                },
            }
            self.assertEqual(
                target.controller_reply_issues(valid, "explain_solution", state, [], context=context),
                [],
            )
            invalid = dict(valid)
            invalid["reply"] = "Я специализируюсь на разработке учебных курсов."
            invalid["proposed_solution"] = {
                "type": "none",
                "name": "",
                "evidence": "",
            }
            issues = target.controller_reply_issues(
                invalid, "explain_solution", state, [], context=context
            )
        self.assertIn("не выбран продукт или услуга специалиста по нейросетям", issues)
        self.assertIn("предлагаемое решение не подтверждено базой знаний", issues)

    def test_marketer_solution_cannot_lead_with_clients_course_as_the_offer(self):
        state = self.state(need=True, previous_experience=True, desired_result=True)
        context = "Методисты по английскому языку — готовые ассистенты для разработки нестандартных учебных программ."
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            issues = target.controller_reply_issues(
                {
                    "reply": "Вам подойдёт разработка индивидуального курса с использованием нейросетей. Методисты по английскому языку помогут организовать эту работу.",
                    "action": "explain_solution",
                    "question_target": "none",
                    "conversation_effect": "continue",
                    "proposed_solution": {
                        "type": "ready_product",
                        "name": "Методисты по английскому языку",
                        "evidence": "Методисты по английскому языку — готовые ассистенты для разработки нестандартных учебных программ",
                    },
                },
                "explain_solution",
                state,
                [],
                context=context,
            )
        self.assertIn(
            "объяснение начинается с результата клиента, а не с предлагаемого ИИ-решения",
            issues,
        )

    def test_marketer_rupture_after_solution_reopens_solution_selection(self):
        state = self.state(need=True, previous_experience=True, desired_result=True)
        state.update({
            "solution_explained": True,
            "solution_type": "ready_product",
            "solution_name": "неверно понятый продукт",
        })
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            repaired = target.advance_dialog_state(
                state,
                "repair_contact",
                "Да, предыдущее объяснение получилось нелогичным.",
            )
        self.assertFalse(repaired["solution_explained"])
        self.assertEqual(repaired["solution_type"], "none")
        self.assertEqual(repaired["solution_name"], "")

    def test_grounded_solution_fallback_uses_verified_evidence_instead_of_failing(self):
        raw = json.dumps({
            "reply": "Курс можно разработать с помощью нейросетей.",
            "type": "ready_product",
            "name": "Методист по английскому языку",
            "evidence": "Методисты по английскому языку — готовые ассистенты для разработки нестандартных учебных программ.",
        }, ensure_ascii=False)
        target.gigachat = ScriptedGigaChat([raw, raw])
        context = "Методисты по английскому языку — готовые ассистенты для разработки нестандартных учебных программ."
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            reply, solution = target.grounded_solution_fallback(
                [],
                "Мне нужен курс, но нет времени его разрабатывать.",
                context,
                ("GigaChat", "GigaChat-2-Max"),
            )
        self.assertEqual(reply, context)
        self.assertEqual(solution["type"], "ready_product")
        self.assertIn("Методист", solution["name"])

    def test_marketer_rejects_questions_about_clients_professional_domain(self):
        state = self.state(contact=True)
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            issues = target.controller_reply_issues(
                {
                    "reply": "Какие задачи по обучению сейчас перед вами стоят?",
                    "action": "explore",
                    "question_target": "need",
                    "question_scope": "client_domain",
                    "conversation_effect": "continue",
                },
                "explore",
                state,
                [],
                client_text="Я репетитор английского.",
            )
        self.assertIn(
            "диагностический вопрос вышел за пределы рабочего процесса клиента",
            issues,
        )

    def test_marketer_accepts_work_process_question(self):
        state = self.state(contact=True)
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            issues = target.controller_reply_issues(
                {
                    "reply": "Какую часть вашей работы хотелось бы упростить или ускорить?",
                    "action": "explore",
                    "question_target": "need",
                    "question_scope": "work_process",
                    "conversation_effect": "continue",
                },
                "explore",
                state,
                [],
                client_text="Я репетитор английского.",
            )
        self.assertEqual(issues, [])

    def test_marketer_prompt_treats_confusion_about_question_as_rupture(self):
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            prompt = target.controller_prompt(
                self.state(contact=True),
                "Екатерина — специалист по нейросетям.",
                [{"role": "assistant", "content": "Какие задачи по обучению перед вами стоят?"}],
                "Не поняла вопрос. Вы же не методист.",
            )
        self.assertIn("это обратная связь о ходе беседы", prompt)
        self.assertIn("используйте rupture и repair_contact".lower(), prompt.lower())

    def test_marketer_product_answer_requires_knowledge_base_evidence(self):
        context = "Ассистенты живут в ChatGPT. Пользоваться ассистентами можно в любой версии."
        state = self.state(need=True, previous_experience=True, desired_result=True)
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            valid = {
                "reply": "Ассистенты живут в ChatGPT. Пользоваться ассистентами можно в любой версии.",
                "action": "answer_information",
                "question_target": "none",
                "question_scope": "none",
                "conversation_effect": "continue",
                "answer_grounding": {
                    "source": "knowledge_base",
                    "evidence": "Ассистенты живут в ChatGPT. Пользоваться ассистентами можно в любой версии.",
                },
            }
            self.assertEqual(
                target.controller_reply_issues(
                    valid,
                    "answer_information",
                    state,
                    [],
                    client_text="В какой нейросети он работает?",
                    context=context,
                ),
                [],
            )
            invented = dict(valid)
            invented["reply"] = "Это самостоятельный продукт, не связанный с ChatGPT."
            invented["answer_grounding"] = {
                "source": "knowledge_base",
                "evidence": "самостоятельный продукт, не связанный с ChatGPT",
            }
            issues = target.controller_reply_issues(
                invented,
                "answer_information",
                state,
                [],
                client_text="Он создан в ChatGPT?",
                context=context,
            )
        self.assertIn("информационный ответ не подтверждён базой знаний", issues)

    def test_marketer_rejects_claim_not_entailed_by_correct_evidence(self):
        context = "Ассистенты живут в ChatGPT. Пользоваться ассистентами можно в любой версии."
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            issues = target.controller_reply_issues(
                {
                    "reply": "Ассистенты используют платформу, аналогичную ChatGPT. Достаточно бесплатной версии ChatGPT.",
                    "action": "answer_information",
                    "question_target": "none",
                    "question_scope": "none",
                    "conversation_effect": "continue",
                    "answer_grounding": {
                        "source": "knowledge_base",
                        "evidence": context,
                    },
                },
                "answer_information",
                self.state(),
                [],
                client_text="Они работают в ChatGPT?",
                context=context,
            )
        self.assertIn(
            "информационный ответ добавляет сведения, которых нет в подтверждающей цитате",
            issues,
        )

    def test_marketer_information_fallback_rejects_invented_platform(self):
        target.gigachat = ScriptedGigaChat([
            '{"reply":"Это отдельный продукт вне ChatGPT.","source":"knowledge_base","evidence":"отдельный продукт вне ChatGPT"}',
            '{"reply":"Ассистенты работают в ChatGPT; достаточно бесплатной версии.","source":"knowledge_base","evidence":"Ассистенты живут в ChatGPT"}',
            '{"answers":true}',
        ])
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            answer = target.grounded_information_fallback(
                [],
                "Он работает в ChatGPT?",
                "Ассистенты живут в ChatGPT. Пользоваться ассистентами можно в любой версии.",
                ("GigaChat", "GigaChat-2-Max"),
            )
        self.assertIn("ChatGPT", answer)
        self.assertNotIn("отдельный продукт", answer)

    def test_evidence_must_answer_the_clients_exact_question(self):
        target.gigachat = ScriptedGigaChat([
            '{"answers":false}',
            '{"answers":true}',
        ])
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            adult_answered = target.evidence_answers_question(
                "Ассистент разрабатывает программы для взрослых?",
                "Методисты по английскому языку — готовые ассистенты для детских курсов и подготовки к экзаменам.",
                "GigaChat-2-Max",
            )
            platform_answered = target.evidence_answers_question(
                "Ассистенты работают в ChatGPT?",
                "Ассистенты живут в ChatGPT.",
                "GigaChat-2-Max",
            )
        self.assertFalse(adult_answered)
        self.assertTrue(platform_answered)

    def test_marketer_has_one_booking_type_and_reserves_sixty_minutes(self):
        with target.app.test_request_context("/marketer"):
            target.g.expert_slug = "marketer"
            target.sset("consultation_offered", True)
            self.assertEqual(
                target.direct_booking_answer("Давайте."),
                "Назовите удобные дату и время — я проверю их в календаре.",
            )
            self.assertEqual(target.booking_config("free")[1], 60)
            self.assertEqual(target.booking_config("regular")[1], 60)
        page = self.client.get("/marketer/booking")
        body = page.get_data(as_text=True)
        self.assertIn("обычно занимает 30–40 минут", body)
        self.assertIn("календаре резервируется 60 минут", body)

    def test_marketer_and_psychologist_session_states_are_separate(self):
        with target.app.test_request_context("/"):
            target.g.expert_slug = "psychologist"
            target.sset("dialog_closed", True)
            target.g.expert_slug = "marketer"
            self.assertIsNone(target.sget("dialog_closed"))


if __name__ == "__main__":
    unittest.main()
