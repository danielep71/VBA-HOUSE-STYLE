"""Tests for skills/vba-house-style/scripts/check_vba_style.py (stdlib only).

Run from the repository root:
    python -m unittest discover -s tests
"""
from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills" / "vba-house-style"
spec = importlib.util.spec_from_file_location(
    "check_vba_style", SKILL / "scripts" / "check_vba_style.py")
cvs = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cvs)

BANNER = "'" + "-" * 78
FIELD_TEXT = {f: "None" for f in cvs.FIELDS}
FIELD_TEXT["UPDATED"] = "2026-01-15"


def header(fields: list[str], title: str | None = None) -> list[str]:
    """A header block with the given labels, each followed by its text."""
    out = ["'", BANNER]
    if title:
        out += [title, BANNER]
    for n, field in enumerate(fields):
        if n:
            out.append("'")
        out += [f"' {field}", f"'   {FIELD_TEXT.get(field, 'Text')}"]
    return out + [BANNER]


def module(name: str, body: list[str], component: list[str] | None = None) -> str:
    comp = component if component is not None else header(cvs.FIELDS, f"' MODULE: {name}")
    lines = [f'Attribute VB_Name = "{name}"', "", "Option Explicit", ""] + comp + [""] + body
    return "\r\n".join(lines) + "\r\n"


DECLARE = ("DECLARE", [
    '    Const PROC_NAME             As String = "M_Test.M_Test_Run"', "",
    "    Dim Count                   As Long         'Counter"])
INIT = ("INITIALIZE", ["    'Enable controlled error handling",
                       "        On Error GoTo ErrorHandler"])
TASK = ("DO WORK", ["    'Count once", "        Debug.Print 1"])
EXIT = ("EXIT PROCEDURE", ["    'Exit before the error handler", "        Exit Sub"])
HANDLER = ("ERROR HANDLER", ["ErrorHandler:", "    'Raise a descriptive error",
                             "        Err.Raise Err.Number, PROC_NAME, Err.Description"])


def procedure(signature: str, fields: list[str], code_lines: int = 1) -> list[str]:
    title = "'" + " " * 35 + "DO WORK"
    code = ["        Debug.Print 1"] * code_lines
    return [signature, ""] + header(fields, title) + [""] + code + ["", "End Sub"]


class CheckerTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def run_check(self, text: str, name: str = "M_Test", suffix: str = ".bas",
                  raw: bytes | None = None) -> list[str]:
        path = self.dir / f"{name}{suffix}"
        path.write_bytes(raw if raw is not None else text.encode("latin-1"))
        return [f"{level} {msg}" for level, _, msg in cvs.check(path, 120)]

    def assertFinding(self, findings: list[str], fragment: str) -> None:
        self.assertTrue(any(fragment in f for f in findings),
                        f"{fragment!r} not in {findings}")

    def assertNoFinding(self, findings: list[str], fragment: str) -> None:
        self.assertFalse(any(fragment in f for f in findings),
                         f"{fragment!r} unexpectedly in {findings}")

    # The worked example is the reference: it must stay clean.
    def test_example_is_clean(self) -> None:
        path = SKILL / "references" / "M_Example.bas"
        self.assertEqual(cvs.check(path, 120), [])

    def test_clean_module(self) -> None:
        body = procedure("Public Sub M_Test_Run()", cvs.FIELDS)
        self.assertEqual(self.run_check(module("M_Test", body)), [])

    # Mechanical errors.
    def test_lf_line_endings(self) -> None:
        text = module("M_Test", []).replace("\r\n", "\n")
        self.assertFinding(self.run_check(text), "ERROR line endings are not all CRLF")

    def test_missing_final_crlf(self) -> None:
        text = module("M_Test", ["    'Last line"])[:-2]
        self.assertFinding(self.run_check(text), "ERROR file does not end with CRLF")

    def test_vb_name_mismatch(self) -> None:
        self.assertFinding(self.run_check(module("M_Other", [])), "ERROR VB_Name")

    def test_missing_option_explicit(self) -> None:
        text = module("M_Test", []).replace("Option Explicit", "")
        self.assertFinding(self.run_check(text), "ERROR missing Option Explicit")

    def test_tab_and_trailing_whitespace(self) -> None:
        findings = self.run_check(module("M_Test", ["\tDim A As Long  "]))
        self.assertFinding(findings, "ERROR tab character")
        self.assertFinding(findings, "ERROR trailing whitespace")

    def test_short_banner(self) -> None:
        findings = self.run_check(module("M_Test", ["'" + "-" * 70]))
        self.assertFinding(findings, "ERROR banner is 71 chars, expected 79")

    # Headers.
    def test_missing_component_header(self) -> None:
        findings = self.run_check(module("M_Test", [], component=["' M_Test"]))
        self.assertFinding(findings, "no ' MODULE: / ' CLASS: component header")

    def test_component_header_missing_field(self) -> None:
        comp = header([f for f in cvs.FIELDS if f != "NOTES"], "' MODULE: M_Test")
        findings = self.run_check(module("M_Test", [], component=comp))
        self.assertFinding(findings, "component header missing NOTES")

    def test_public_needs_all_nine(self) -> None:
        body = procedure("Public Sub M_Test_Run()", ["PURPOSE", "UPDATED"])
        self.assertFinding(self.run_check(module("M_Test", body)), "M_Test_Run: header missing")

    def test_small_private_may_be_compact(self) -> None:
        body = procedure("Private Sub M_Test_Run()", ["PURPOSE", "UPDATED"], code_lines=15)
        self.assertEqual(self.run_check(module("M_Test", body)), [])

    def test_compact_still_needs_purpose_and_updated(self) -> None:
        body = procedure("Private Sub M_Test_Run()", ["PURPOSE"])
        self.assertFinding(self.run_check(module("M_Test", body)), "header missing UPDATED")

    def test_large_private_needs_all_nine(self) -> None:
        body = procedure("Private Sub M_Test_Run()", ["PURPOSE", "UPDATED"], code_lines=16)
        self.assertFinding(self.run_check(module("M_Test", body)), "header missing WHY THIS EXISTS")

    def test_fields_out_of_order(self) -> None:
        fields = list(cvs.FIELDS)
        fields[0], fields[1] = fields[1], fields[0]
        body = procedure("Public Sub M_Test_Run()", fields)
        self.assertFinding(self.run_check(module("M_Test", body)), "header fields out of order")

    def test_extra_section_before_updated(self) -> None:
        fields = cvs.FIELDS[:-1] + ["PUBLIC SURFACE", "UPDATED"]
        comp = header(fields, "' MODULE: M_Test")
        self.assertEqual(self.run_check(module("M_Test", [], component=comp)), [])

    def test_extra_section_too_early(self) -> None:
        fields = ["PUBLIC SURFACE"] + cvs.FIELDS
        comp = header(fields, "' MODULE: M_Test")
        findings = self.run_check(module("M_Test", [], component=comp))
        self.assertFinding(findings, "extra section PUBLIC SURFACE must follow NOTES")

    def test_blank_lines_between_procedures(self) -> None:
        first = procedure("Public Sub M_Test_A()", cvs.FIELDS)
        second = procedure("Public Sub M_Test_B()", cvs.FIELDS)
        findings = self.run_check(module("M_Test", first + ["", ""] + second))
        self.assertFinding(findings, "M_Test_B: 2 blank lines after previous End, expected 1")

    # Line length.
    def test_long_code_line_warns(self) -> None:
        line = "        Total = " + " + ".join(["ValueNumber"] * 12)
        self.assertFinding(self.run_check(module("M_Test", [line])), "WARN line is")

    def test_long_string_literal_is_exempt(self) -> None:
        line = '        Message = "' + "x" * 150 + '"'
        self.assertNoFinding(self.run_check(module("M_Test", [line])), "line is")

    def test_long_declaration_comment_is_exempt(self) -> None:
        line = "    Private mIsBusy" + " " * 20 + "As Boolean" + " " * 20 + "'" + "y" * 100
        self.assertNoFinding(self.run_check(module("M_Test", [line])), "line is")

    # Statement-level conventions.
    def statements(self, *code: str, signature: str = "Public Sub M_Test_A()") -> list[str]:
        proc = procedure(signature, cvs.FIELDS)
        body = ["        " + c for c in code]
        return self.run_check(module("M_Test", proc[:-2] + body + proc[-2:]))

    def test_next_without_counter_warns(self) -> None:
        self.assertFinding(self.statements("Next"), "Next without the counter name")

    def test_next_with_counter_and_resume_next_are_silent(self) -> None:
        findings = self.statements("Next Index", "On Error Resume Next", "Resume Next")
        self.assertNoFinding(findings, "Next without")

    def test_colon_separator_warns(self) -> None:
        self.assertFinding(self.statements("A = 1: B = 2"), "statement separator")

    def test_colon_in_label_string_comment_and_named_argument_is_silent(self) -> None:
        findings = self.statements(
            'Message = "a: b"', "Call Work(Name:=1)", "X = 1 'note: text", "Done:")
        self.assertNoFinding(findings, "statement separator")

    def test_instr_without_compare_warns(self) -> None:
        self.assertFinding(self.statements("N = VBA.InStr(Text, Part)"),
                           "InStr without an explicit compare mode")
        self.assertFinding(self.statements("N = InStr(1, Text, Part)"),
                           "InStr without an explicit compare mode")

    def test_strcomp_without_compare_warns(self) -> None:
        self.assertFinding(self.statements("N = VBA.StrComp(A, B)"),
                           "StrComp without an explicit compare mode")

    def test_instr_and_strcomp_with_compare_are_silent(self) -> None:
        findings = self.statements(
            "N = VBA.InStr(Text, Part, vbTextCompare)",
            "N = VBA.InStr(1, Text, Part, vbBinaryCompare)",
            "N = VBA.StrComp(A, B, Mode)", "N = VBA.InStr(1, Text, Part, CompareMode)",
            "N = Left$(Text, VBA.InStr(1, Text, Part, vbBinaryCompare))",
            "N = Item.InStr(Text)")
        self.assertNoFinding(findings, "explicit compare mode")

    def test_instr_with_variable_compare_is_silent(self) -> None:
        findings = self.statements("N = VBA.InStr(Text, Part, Mode)")
        self.assertNoFinding(findings, "explicit compare mode")

    def test_compare_in_input_names_does_not_count(self) -> None:
        findings = self.statements("N = VBA.StrComp(TextToCompare, Expected)",
                                   "N = VBA.InStr(1, CompareText, Needle)")
        self.assertEqual(sum("explicit compare mode" in f for f in findings), 2, findings)

    def test_named_compare_argument_is_silent(self) -> None:
        findings = self.statements("N = VBA.StrComp(A, B, Compare:=vbTextCompare)",
                                   "N = VBA.InStr(Start:=1, String1:=A, String2:=B, "
                                   "Compare:=vbBinaryCompare)")
        self.assertNoFinding(findings, "explicit compare mode")

    def test_date_literal_is_not_a_separator(self) -> None:
        findings = self.statements("Deadline = #11/21/2026 12:30:00#",
                                   "Deadline = #12:30:00 PM#", "Deadline = #1/2/2026#")
        self.assertNoFinding(findings, "statement separator")

    def test_file_number_markers_are_not_a_date_literal(self) -> None:
        findings = self.statements("Print #1, Value: Print #2, OtherValue")
        self.assertFinding(findings, "statement separator")
        findings = self.statements("Print #1, VBA.StrComp(A, B): Print #2, C")
        self.assertFinding(findings, "StrComp without an explicit compare mode")

    def test_month_name_date_literal_is_not_a_separator(self) -> None:
        findings = self.statements("Deadline = #January 5, 2026 9:30 AM#",
                                   "Deadline = #2026-01-05 09:30:00#")
        self.assertNoFinding(findings, "statement separator")

    def test_attribute_records_are_not_statements(self) -> None:
        text = ('Attribute VB_Name = "M_Test"\r\nAttribute VB_Description = ""\r\n'
                'Attribute VB_ProcData.VB_Invoke_Func = " \\n14"\r\nOption Explicit\r\n')
        findings = self.run_check(text)
        self.assertNoFinding(findings, "vbNullString")
        self.assertNoFinding(findings, "statement separator")

    def test_rem_comment_is_not_code(self) -> None:
        findings = self.statements("Rem Example: VBA.StrComp(A, B)", "Rem Next",
                                   "Remainder = 1")
        self.assertNoFinding(findings, "statement separator")
        self.assertNoFinding(findings, "explicit compare mode")
        self.assertNoFinding(findings, "Next without")

    def test_statement_before_rem_is_still_checked(self) -> None:
        findings = self.statements("A = 1: Rem note")
        self.assertFinding(findings, "statement separator")

    def test_empty_string_comparison_warns(self) -> None:
        self.assertFinding(self.statements('If Name = "" Then'), "vbNullString")
        self.assertFinding(self.statements('If Name <> "" Then'), "vbNullString")

    def test_empty_string_in_const_and_literal_text_is_silent(self) -> None:
        findings = self.statements('Const EMPTY_TEXT As String = ""', 'Message = "a = """""',
                                   "If Name = vbNullString Then")
        self.assertNoFinding(findings, "vbNullString")

    def test_parameter_without_byval_or_byref_warns(self) -> None:
        findings = self.statements(signature="Public Sub M_Test_A(Name As String)")
        self.assertFinding(findings, "parameter Name has no ByVal or ByRef")

    def test_parameter_without_as_warns(self) -> None:
        findings = self.statements(signature="Public Sub M_Test_A(ByVal Name)")
        self.assertFinding(findings, "parameter Name has no As type")

    def test_function_without_as_warns(self) -> None:
        findings = self.run_check(module("M_Test", procedure(
            "Public Function M_Test_F(ByVal Name As String)", cvs.FIELDS)))
        self.assertFinding(findings, "no As return type")

    def test_complete_signatures_are_silent(self) -> None:
        findings = self.statements(signature=(
            "Public Sub M_Test_A(ByVal Name As String, ByRef Items() As Long, _\r\n"
            "    Optional ByVal Mode As Long = 0, ParamArray Rest() As Variant)"))
        self.assertNoFinding(findings, "parameter")
        self.assertNoFinding(findings, "return type")

    def test_empty_parameter_list_is_silent(self) -> None:
        self.assertNoFinding(self.statements(), "parameter")

    def test_statement_findings_do_not_change_errors(self) -> None:
        findings = self.statements("Next", "A = 1: B = 2")
        self.assertFalse(any(f.startswith("ERROR") for f in findings), findings)

    # Body sections.
    def sections(self, *parts: tuple[str, list[str]],
                 signature: str = "Public Sub M_Test_Run()",
                 title: str = "'" + " " * 35 + "DO WORK", blank: bool = True) -> list[str]:
        proc = [signature, ""] + header(cvs.FIELDS, title) + ([""] if blank else [])
        for name, code in parts:
            proc += [BANNER, f"' {name}", BANNER] + code + [""]
        end = "End Function" if " Function " in signature else "End Sub"
        return self.run_check(module("M_Test", proc + [end]))

    def test_full_body_layout_is_silent(self) -> None:
        self.assertEqual(self.sections(DECLARE, INIT, TASK, EXIT, HANDLER), [])

    def test_declare_without_error_handler_is_silent(self) -> None:
        self.assertEqual(self.sections(DECLARE, TASK), [])

    def test_task_sections_only_are_silent(self) -> None:
        self.assertEqual(self.sections(TASK, ("RETURN VALUE", ["        Debug.Print 2"])), [])

    def test_one_line_getter_is_exempt(self) -> None:
        findings = self.run_check(module("M_Test", procedure(
            "Public Function M_Test_Get() As Long", cvs.FIELDS)[:-1] + ["End Function"]))
        self.assertEqual(findings, [])

    def test_compact_private_is_exempt(self) -> None:
        proc = procedure("Private Sub M_Test_Run()", ["PURPOSE", "UPDATED"])
        code = ["    Dim Count As Long", "    On Error GoTo Fail", "        Count = 1",
                "        Exit Sub", "Fail:", "        Count = 0"]
        self.assertEqual(self.run_check(module("M_Test", proc[:-3] + code + proc[-2:])), [])

    def test_left_aligned_procedure_title_is_not_a_section(self) -> None:
        findings = self.sections(DECLARE, TASK, title="' DO WORK")
        self.assertNoFinding(findings, "out of order")
        self.assertNoFinding(findings, "DECLARE does not")

    def test_section_directly_after_the_header_is_found(self) -> None:
        findings = self.sections(("DO WORK", ["    Dim Count As Long", "        Count = 1"]),
                                 blank=False)
        self.assertFinding(findings, "no DECLARE section")
        findings = self.sections(DECLARE, TASK, blank=False)
        self.assertEqual(findings, [])

    def test_local_variables_without_declare_warn(self) -> None:
        findings = self.sections(("DO WORK", ["    Dim Count As Long", "        Count = 1"]))
        self.assertFinding(findings, "M_Test_Run: no DECLARE section")

    def test_error_handler_without_declare_warns(self) -> None:
        findings = self.sections(INIT, TASK, EXIT, HANDLER)
        self.assertFinding(findings, "no DECLARE section for local variables or error handler")

    def test_declare_must_start_with_proc_name(self) -> None:
        declare = ("DECLARE", DECLARE[1][2:] + ["", DECLARE[1][0]])
        findings = self.sections(declare, TASK)
        self.assertFinding(findings, "DECLARE does not start with Const PROC_NAME As String")

    def test_proc_name_must_match_component_and_procedure(self) -> None:
        declare = ("DECLARE", [DECLARE[1][0].replace("M_Test_Run", "M_Test_Old")])
        findings = self.sections(declare, TASK)
        self.assertFinding(findings, 'PROC_NAME is "M_Test.M_Test_Old", expected "M_Test.M_Test_Run"')

    def test_proc_name_on_a_continuation_line_is_read(self) -> None:
        split = ["    Const PROC_NAME As String = _", '        "M_Test.M_Test_Old"']
        findings = self.sections(("DECLARE", split), TASK)
        self.assertFinding(findings, 'PROC_NAME is "M_Test.M_Test_Old"')
        split[1] = '        "M_Test.M_Test_Run"'
        self.assertEqual(self.sections(("DECLARE", split), TASK), [])

    def test_on_error_outside_initialize_warns(self) -> None:
        task = ("DO WORK", INIT[1] + TASK[1])
        findings = self.sections(DECLARE, task, EXIT, HANDLER)
        self.assertFinding(findings, "On Error GoTo ErrorHandler is not in INITIALIZE")

    def test_later_on_error_goto_outside_initialize_is_silent(self) -> None:
        task = ("DO WORK", ["        On Error Resume Next", "        Count = 1",
                            "        On Error GoTo ErrorHandler"])
        self.assertEqual(self.sections(DECLARE, INIT, task, EXIT, HANDLER), [])

    def test_indented_error_label_warns(self) -> None:
        handler = ("ERROR HANDLER", ["    ErrorHandler:"] + HANDLER[1][1:])
        findings = self.sections(DECLARE, INIT, TASK, EXIT, handler)
        self.assertFinding(findings, "label ErrorHandler: is not at column 0")

    def test_section_out_of_order_warns(self) -> None:
        findings = self.sections(DECLARE, TASK, INIT, EXIT, HANDLER)
        self.assertFinding(findings, "section INITIALIZE is out of order (after DO WORK)")

    def test_error_handler_not_last_warns(self) -> None:
        findings = self.sections(DECLARE, INIT, TASK, HANDLER, EXIT)
        self.assertFinding(findings, "section EXIT PROCEDURE is out of order (after ERROR HANDLER)")

    def test_missing_clean_exit_warns(self) -> None:
        findings = self.sections(DECLARE, INIT, TASK, HANDLER)
        self.assertFinding(findings, "no Exit Sub in EXIT PROCEDURE before ErrorHandler:")

    def test_clean_exit_names_the_procedure_kind(self) -> None:
        declare = ("DECLARE", [DECLARE[1][0].replace("M_Test_Run", "M_Test_Get")])
        findings = self.sections(declare, INIT, TASK, HANDLER,
                                 signature="Public Function M_Test_Get() As Long")
        self.assertFinding(findings, "no Exit Function in EXIT PROCEDURE")

    def test_exit_outside_exit_procedure_warns(self) -> None:
        task = ("DO WORK", TASK[1] + ["        Exit Sub"])
        findings = self.sections(DECLARE, INIT, task, HANDLER)
        self.assertFinding(findings, "no Exit Sub in EXIT PROCEDURE")

    def test_body_findings_do_not_change_errors(self) -> None:
        findings = self.sections(TASK, INIT, HANDLER)
        self.assertTrue(findings)
        self.assertFalse(any(f.startswith("ERROR") for f in findings), findings)

    # VBA. qualification.
    def test_each_unqualified_library_call_warns(self) -> None:
        for name in cvs.VBA_QUALIFIED:
            with self.subTest(name=name):
                findings = self.statements(f"Result = {name}(Text)")
                self.assertFinding(findings, f"use VBA.{name} instead of {name}")

    def test_each_name_is_silent_qualified_as_member_in_text_or_declared(self) -> None:
        for name in cvs.VBA_QUALIFIED:
            base = name.rstrip("$")
            with self.subTest(name=name):
                findings = self.statements(
                    f"Result = VBA.{name}(Text)", f"Result = Item.{name}(Text)",
                    f"Result = Item!{name}", f"Result = .{name}(Text)",
                    f'Message = "{name}(Text)"', f"'{name}(Text)",
                    f"Result = 1 ' {name}(Text)", f"Run {base}:=1")
                self.assertNoFinding(findings, "use VBA.")
                findings = self.statements(f"Dim {base} As Long", f"Result = {name}(1)")
                self.assertNoFinding(findings, "use VBA.")

    def test_names_declared_by_the_file_are_silent(self) -> None:
        proc = procedure("Private Function Trim(ByVal Text As String) As String",
                         ["PURPOSE", "UPDATED"])
        decl = ["Private Declare PtrSafe Function CStr Lib \"x\" () As Long",
                "Public Event LCase(ByVal Len As Long)",
                "Private Enum E_Kind", "    CLng = 1", "End Enum",
                "Private Type T_Item", "    CDbl As Double", "End Type", ""]
        body = ["        Result = Trim$(Text) & Len(Text) & CStr() & CLng & CDbl"]
        findings = self.run_check(module("M_Test", decl + proc[:-2] + body + proc[-2:]))
        self.assertNoFinding(findings, "use VBA.")
        findings = self.statements("Result = Len(Text)",
                                   signature="Public Sub M_Test_A(ByVal Len As Long)")
        self.assertNoFinding(findings, "use VBA.")

    def test_name_case_and_spelling_are_reported(self) -> None:
        findings = self.statements("Result = ucase$(Text)")
        self.assertFinding(findings, "use VBA.UCase$ instead of ucase$")

    def test_mid_statement_and_bare_left_right_are_silent(self) -> None:
        findings = self.statements('Mid$(Text, 1, 1) = "x"', 'If Ok Then Mid$(Text, 1) = "y"',
                                   "Left = 10", "Result = Right(Text, 2)")
        self.assertNoFinding(findings, "use VBA.")
        findings = self.statements("Result = Mid$(Text, 1, 1)")
        self.assertFinding(findings, "use VBA.Mid$ instead of Mid$")

    def test_open_len_clause_and_len_of_a_record_are_silent(self) -> None:
        findings = self.statements("Dim Record As T_Record", "Dim Records(1 To 2) As T_Record",
                                   "Open Path For Random As #1 Len = Len(Record)",
                                   "Size = LenB(Records(1))")
        self.assertNoFinding(findings, "use VBA.")
        findings = self.statements("Result = Len(Text)",
                                   signature="Public Sub M_Test_A(ByRef Text As String)")
        self.assertFinding(findings, "use VBA.Len instead of Len")

    def test_several_names_on_one_line_each_warn(self) -> None:
        findings = self.statements("Result = UCase$(Trim$(Text)) & Len(Text)")
        for name in ("UCase$", "Trim$", "Len"):
            self.assertFinding(findings, f"use VBA.{name} instead of {name}")

    def test_qualification_findings_do_not_change_errors(self) -> None:
        findings = self.statements("Result = Len(Text)")
        self.assertFalse(any(f.startswith("ERROR") for f in findings), findings)

    # Tokenizer.
    def kinds(self, line: str) -> list[tuple[str, str]]:
        return cvs.tokenize(line)

    def test_tokens_join_back_to_the_line(self) -> None:
        for line in ['A = "x: y" \'note', "Print #1, V: Print #2, W", "X = #1/2/2026# ' d",
                     'Msg = "say ""hi"""', "A = 1: Rem note", "Remainder = 1", ""]:
            self.assertEqual("".join(t for _, t in self.kinds(line)), line)

    def test_token_kinds(self) -> None:
        self.assertEqual(self.kinds('X = "a" \'c'), [
            (cvs.CODE, "X = "), (cvs.STRING, '"a"'), (cvs.CODE, " "), (cvs.COMMENT, "'c")])
        self.assertEqual(self.kinds("D = #1/2/2026#"), [
            (cvs.CODE, "D = "), (cvs.DATE, "#1/2/2026#")])

    def test_apostrophe_inside_a_string_is_not_a_comment(self) -> None:
        self.assertEqual(self.kinds('M = "it\'s"'), [
            (cvs.CODE, "M = "), (cvs.STRING, '"it\'s"')])
        self.assertNoFinding(self.statements('Message = "it\'s: fine"'), "statement separator")

    def test_doubled_quotes_stay_in_one_string(self) -> None:
        self.assertEqual(self.kinds('M = "say ""hi"": ok"')[1], (cvs.STRING, '"say ""hi"": ok"'))

    def test_rem_after_colon_is_a_comment(self) -> None:
        self.assertEqual(self.kinds("A = 1: Rem note: x"), [
            (cvs.CODE, "A = 1: "), (cvs.COMMENT, "Rem note: x")])
        self.assertEqual(self.kinds("Remainder = 1"), [(cvs.CODE, "Remainder = 1")])
        self.assertEqual(self.kinds("Total = Remainder"), [(cvs.CODE, "Total = Remainder")])

    def test_file_numbers_in_open_input_close_are_code(self) -> None:
        line = 'Open Path For Input As #1: Input #1, A: Close #1'
        self.assertEqual(self.kinds(line), [(cvs.CODE, line)])
        self.assertFinding(self.statements(line), "statement separator")

    def test_comment_text_counts_in_full_for_line_length(self) -> None:
        line = "    X = 1 '" + 'say "' + "y" * 120 + '"'
        self.assertFinding(self.run_check(module("M_Test", [line])), "WARN line is")

    def test_rem_trailing_a_declaration_is_exempt_from_line_length(self) -> None:
        line = "    Private mIsBusy" + " " * 20 + "As Boolean" + " " * 5 + ": Rem " + "y" * 100
        self.assertNoFinding(self.run_check(module("M_Test", [line])), "line is")

    # --fix.
    def test_fix_changes_whitespace_only(self) -> None:
        raw = b'Attribute VB_Name = "M_Test"\nOption Explicit  \n\tDim A As Long\t\n'
        fixed = cvs.fix(raw)
        self.assertEqual(fixed, b'Attribute VB_Name = "M_Test"\r\nOption Explicit\r\n'
                                b'    Dim A As Long\r\n')

    def test_fix_keeps_frm_designer_block(self) -> None:
        designer = b"VERSION 5.00\r\nBegin {C62A69F0} UF_Test \r\n\tCaption = 1   \r\nEnd\r\n"
        code = b'Attribute VB_Name = "UF_Test"\r\nOption Explicit   \r\n'
        fixed = cvs.fix(designer + code)
        self.assertTrue(fixed.startswith(designer))
        self.assertTrue(fixed.endswith(b"Option Explicit\r\n"))


if __name__ == "__main__":
    unittest.main()
