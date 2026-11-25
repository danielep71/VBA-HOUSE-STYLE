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
FIELD_TEXT["RETURNS"] = "Nothing"
MODULE_TEXT = {"INPUTS": "None at module level"}


def header(fields: list[str], title: str | None = None,
           text: dict[str, str] | None = None) -> list[str]:
    """A header block with the given labels, each followed by its text."""
    values = {**FIELD_TEXT, **(text or {})}
    out = ["'", BANNER]
    if title:
        out += [title, BANNER]
    for n, field in enumerate(fields):
        if n:
            out.append("'")
        out += [f"' {field}", f"'   {values.get(field, 'Text')}"]
    return out + [BANNER]


def module(name: str, body: list[str], component: list[str] | None = None) -> str:
    comp = component if component is not None else header(
        cvs.FIELDS, f"' MODULE: {name}", MODULE_TEXT)
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
        comp = header(fields, "' MODULE: M_Test", MODULE_TEXT)
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

    # The VBA. qualifier is optional: neither form is reported (#20).
    def test_qualified_and_plain_library_calls_are_both_silent(self) -> None:
        for name in ("Len", "LenB", "Trim$", "UCase$", "LCase", "Left$", "Right$", "Mid$",
                     "InStr", "CStr", "CLng", "CDate"):
            with self.subTest(name=name):
                findings = self.statements(f"Result = {name}(Text)",
                                           f"Result = VBA.{name}(Text)")
                self.assertNoFinding(findings, "VBA.")
                self.assertNoFinding(findings, "instead of")

    def test_library_calls_do_not_change_errors(self) -> None:
        findings = self.statements("Result = Len(Text)", "Result = VBA.Len(Text)")
        self.assertFalse(any(f.startswith("ERROR") for f in findings), findings)

    # Select Case and local names.
    def test_select_case_without_case_else_warns(self) -> None:
        findings = self.statements("Select Case Count", "    Case 1", "        Debug.Print 1",
                                   "End Select")
        self.assertFinding(findings, "WARN Select Case without Case Else")

    def test_select_case_with_case_else_is_silent(self) -> None:
        findings = self.statements("Select Case Count", "    Case 1: Debug.Print 1",
                                   "    Case Else", "        Debug.Print 2", "End Select")
        self.assertNoFinding(findings, "Case Else")

    def test_nested_select_case_is_matched_to_its_own_end(self) -> None:
        inner_missing = ["Select Case Count", "    Case 1",
                         "        Select Case Total", "            Case 2",
                         "        End Select", "    Case Else", "End Select"]
        findings = self.statements(*inner_missing)
        self.assertEqual(sum("Case Else" in f for f in findings), 1, findings)
        outer_missing = ["Select Case Count", "    Case 1",
                         "        Select Case Total", "            Case Else",
                         "        End Select", "End Select"]
        findings = self.statements(*outer_missing)
        self.assertEqual(sum("Case Else" in f for f in findings), 1, findings)
        both = ["Select Case Count", "    Case 1", "        Select Case Total",
                "            Case Else", "        End Select", "    Case Else", "End Select"]
        self.assertNoFinding(self.statements(*both), "Case Else")

    def test_hungarian_local_warns(self) -> None:
        for name in ("strName", "lngCount", "intIndex", "objShape", "blnDone",
                     "dblRate", "varValue", "rngData", "wsData", "wbBook"):
            with self.subTest(name=name):
                findings = self.statements(f"Dim {name} As Variant")
                self.assertFinding(findings, f"WARN local variable {name} has a Hungarian prefix")

    def test_local_that_is_not_pascal_case_warns(self) -> None:
        findings = self.statements("Dim Total As Long, shapeIndex As Long",
                                   "Static Shape_Count As Long")
        self.assertFinding(findings, "local variable shapeIndex is not PascalCase")
        self.assertFinding(findings, "local variable Shape_Count is not PascalCase")
        self.assertNoFinding(findings, "local variable Total")

    def test_words_constants_module_members_and_parameters_are_silent(self) -> None:
        findings = self.statements("Dim Strength As Long", "Dim Intent As String",
                                   "Dim Variance As Double, Items() As String",
                                   "Const MAX_ITEMS As Long = 10",
                                   signature="Private Sub mLabel_Click(ByVal strButton As Long)")
        self.assertNoFinding(findings, "local variable")
        decl = ["Private mName As String", "Public gCount As Long",
                "Private Const MAX_COUNT As Long = 5", "Dim mTotal As Long", ""]
        proc = procedure("Public Sub M_Test_A()", cvs.FIELDS)
        self.assertNoFinding(self.run_check(module("M_Test", decl + proc)), "local variable")

    def test_select_and_naming_findings_do_not_change_errors(self) -> None:
        findings = self.statements("Dim strName As String", "Select Case strName", "End Select")
        self.assertTrue(findings)
        self.assertFalse(any(f.startswith("ERROR") for f in findings), findings)

    # Visibility, constant names and declaration alignment.
    def declarations(self, *decl: str) -> list[str]:
        """Check a module whose declarations section holds the given lines."""
        proc = procedure("Public Sub M_Test_A()", cvs.FIELDS)
        return self.run_check(module("M_Test", list(decl) + [""] + proc))

    def test_signature_without_visibility_warns(self) -> None:
        for signature, name in (("Sub M_Test_A()", "M_Test_A"),
                                ("Function M_Test_A() As Long", "M_Test_A"),
                                ("Property Get Value() As Long", "Value"),
                                ("Static Sub M_Test_A()", "M_Test_A")):
            with self.subTest(signature=signature):
                findings = self.statements(signature=signature)
                self.assertFinding(findings, f"WARN {name}: no Public, Private or Friend")

    def test_signature_with_visibility_is_silent(self) -> None:
        for signature in ("Public Sub M_Test_A()", "Private Sub Class_Initialize()",
                          "Friend Property Get Value() As Long",
                          "Public Static Function M_Test_A() As Long",
                          "private sub M_Test_A()"):
            with self.subTest(signature=signature):
                self.assertNoFinding(self.statements(signature=signature), "no Public")

    def test_visibility_follows_a_continued_signature(self) -> None:
        proc = procedure("Sub M_Test_A(ByVal A As Long, _\r\n    ByVal B As Long)",
                         cvs.FIELDS)
        findings = self.run_check(module("M_Test", proc))
        self.assertFinding(findings, "M_Test_A: no Public, Private or Friend")

    def test_constant_not_upper_snake_case_warns(self) -> None:
        findings = self.declarations(
            "    Private Const MaxItems  As Long = 5",
            "    Public Const maxCount   As Long = 5",
            "    Const Win_Style         As Long = 1",
            "    Global Const A_B, Mixed_Case = 3")
        for name in ("MaxItems", "maxCount", "Win_Style", "Mixed_Case"):
            self.assertFinding(findings, f"WARN constant {name} is not UPPER_SNAKE_CASE")
        self.assertNoFinding(findings, "constant A_B")

    def test_upper_snake_case_and_local_constants_are_silent(self) -> None:
        findings = self.declarations(
            "    Private Const EX_ERR_NO_OWNER As Long = 1001",
            "    Public Const GWL_STYLE        As Long = -16",
            "    Private Const LIMIT2          As Long = 2",
            "    Private Const SMALL&          = 1",
            "    #Const DEBUGGING = 1")
        self.assertNoFinding(findings, "UPPER_SNAKE_CASE")
        findings = self.statements("Const LocalLimit As Long = 1")
        self.assertNoFinding(findings, "UPPER_SNAKE_CASE")
        self.assertEqual(self.sections(DECLARE, TASK), [])

    def test_misaligned_as_warns_once_per_group(self) -> None:
        findings = self.declarations(
            "    Private mFirst      As String",
            "    Private mSecond  As String",
            "    Private mThird         As String")
        self.assertFinding(findings, "WARN As is not aligned with its group (column 22, expected 25)")
        self.assertEqual(sum("As is not aligned" in f for f in findings), 1, findings)

    def test_misaligned_trailing_comment_warns(self) -> None:
        findings = self.declarations(
            "    Private mFirst      As String       'First",
            "    Private mSecond     As String     'Second")
        self.assertFinding(findings, "WARN trailing comment is not aligned with its group "
                                     "(column 39, expected 41)")
        self.assertNoFinding(findings, "As is not aligned")

    def test_aligned_groups_are_silent_across_blank_lines(self) -> None:
        findings = self.declarations(
            "    Private Const EX_A     As Long = 1     'First",
            "    Private Const EX_BB    As String = \"x\" 'Second",
            "",
            "    'Error numbers",
            "    Private Const EX_ERR_X As Long = 1001  'Other column, other group",
            "    Private Const EX_ERR_Y As Long = 1002  'Other column, other group")
        self.assertNoFinding(findings, "not aligned")

    def test_one_line_group_and_lines_without_a_comment_are_silent(self) -> None:
        findings = self.declarations(
            "    Private mOnly       As String       'Alone",
            "",
            "    Private mFirst      As String",
            "    Private mSecond     As String       'Has one")
        self.assertNoFinding(findings, "not aligned")

    def test_groups_end_at_comments_statements_and_continuations(self) -> None:
        findings = self.declarations(
            "    Private mFirst      As String",
            "    'A comment ends the group",
            "    Private mSecond  As String",
            "    Private Declare Function Beep Lib \"kernel32\" () As Long",
            "    Private mThird  As String",
            "    Private Const EX_TEXT As String = _",
            "        \"long\"",
            "    Private mFourth     As String")
        self.assertNoFinding(findings, "not aligned")

    def test_enum_type_and_conditional_members_are_exempt(self) -> None:
        findings = self.declarations(
            "    Public Type EX_Rect",
            "        Left As Long",
            "        TopEdge          As Long",
            "    End Type",
            "    Public Enum EX_Mode",
            "        EX_ModeA = 0    'A",
            "        EX_ModeB = 1  'B",
            "    End Enum",
            "    #If VBA7 Then",
            "    Private Declare PtrSafe Function Beep Lib \"kernel32\" ( _",
            "        ByVal Freq As Long) As Long",
            "    #End If")
        self.assertNoFinding(findings, "not aligned")

    def test_declare_section_alignment_is_checked(self) -> None:
        bad = ("DECLARE", [
            '    Const PROC_NAME             As String = "M_Test.M_Test_Run"', "",
            "    Dim Count                   As Long         'Counter",
            "    Dim Total              As Long    'Sum"])
        findings = self.sections(bad, TASK)
        self.assertFinding(findings, "M_Test_Run: As is not aligned with its group")
        self.assertFinding(findings, "M_Test_Run: trailing comment is not aligned")
        self.assertEqual(self.sections(DECLARE, TASK), [])

    def test_alignment_outside_declare_sections_is_not_checked(self) -> None:
        task = ("DO WORK", ["    Dim Count As Long", "    Dim Total          As Long"])
        findings = self.sections(DECLARE, task)
        self.assertNoFinding(findings, "not aligned")

    def test_declaration_findings_do_not_change_errors(self) -> None:
        findings = self.declarations("    Private Const MaxItems As Long = 5",
                                     "    Private mA    As Long", "    Private mB  As Long")
        self.assertTrue(findings)
        self.assertFalse(any(f.startswith("ERROR") for f in findings), findings)

    # Banner labels and procedure titles.
    def body_with(self, *extra: str, title: str | None = None) -> list[str]:
        """Check a procedure whose body holds the given column-0 lines."""
        proc = procedure("Public Sub M_Test_A()", cvs.FIELDS)
        if title is not None:
            proc = [proc[0], ""] + header(cvs.FIELDS, title) + proc[-4:]
        return self.run_check(module("M_Test", proc[:-2] + list(extra) + proc[-2:]))

    def test_title_not_uppercase_warns(self) -> None:
        findings = self.body_with(title="'" + " " * 35 + "Do Work")
        self.assertFinding(findings, 'WARN M_Test_A: procedure title "Do Work" is not UPPERCASE')

    def test_title_not_centered_warns(self) -> None:
        for lead in (0, 1, 34, 36):
            with self.subTest(lead=lead):
                findings = self.body_with(title="'" + " " * lead + "DO WORK")
                self.assertFinding(findings, f'procedure title "DO WORK" is not centered '
                                             f"({lead} leading spaces, expected 35)")

    def test_centered_titles_of_odd_and_even_length_are_silent(self) -> None:
        for title in ("DO WORK", "DO WORKS", "A", "X" * 78):
            with self.subTest(title=title):
                lead = (78 - len(title)) // 2
                findings = self.body_with(title="'" + " " * lead + title)
                self.assertNoFinding(findings, "procedure title")

    def test_header_without_a_title_banner_is_silent(self) -> None:
        findings = self.run_check(module("M_Test", [
            "Public Sub M_Test_A()", ""] + header(cvs.FIELDS) + ["", "End Sub"]))
        self.assertNoFinding(findings, "procedure title")

    def test_section_label_not_uppercase_or_not_left_aligned_warns(self) -> None:
        findings = self.body_with(BANNER, "' Do work", BANNER, "    'Count once",
                                  "        Debug.Print 1",
                                  BANNER, "'DECLARE", BANNER,
                                  BANNER, "'   EXIT PROCEDURE", BANNER)
        self.assertFinding(findings, 'WARN M_Test_A: section label "Do work" is not UPPERCASE')
        self.assertFinding(findings, 'section label "DECLARE" is not left-aligned')
        self.assertFinding(findings, 'section label "EXIT PROCEDURE" is not left-aligned')

    def test_section_labels_are_checked_when_updated_is_missing(self) -> None:
        proc = procedure("Public Sub M_Test_A()", cvs.FIELDS[:-1])
        findings = self.run_check(module(
            "M_Test", proc[:-2] + [BANNER, "' Do work", BANNER] + proc[-2:]))
        self.assertFinding(findings, "header missing UPDATED")
        self.assertFinding(findings, 'WARN M_Test_A: section label "Do work" is not UPPERCASE')

    def test_header_fields_are_not_read_as_labels_when_updated_is_missing(self) -> None:
        proc = procedure("Public Sub M_Test_A()", cvs.FIELDS[:-1])
        findings = self.run_check(module("M_Test", proc))
        self.assertNoFinding(findings, "section label")

    def test_section_label_in_the_module_top_is_checked(self) -> None:
        findings = self.declarations(BANNER, "' Private constants", BANNER,
                                     "    Private Const EX_A As Long = 1")
        self.assertFinding(findings, 'WARN section label "Private constants" is not UPPERCASE')

    def test_standard_and_task_labels_are_silent(self) -> None:
        sections = [(name, ["    'Do it", "        Debug.Print 1"]) for name in (
            "DECLARE", "INITIALIZE", "VALIDATE INPUTS", "STEP 2 / 3", "LOAD SETTINGS & STATE",
            "EXIT PROCEDURE", "ERROR HANDLER")]
        findings = self.sections(*sections)
        self.assertNoFinding(findings, "section label")

    def test_broken_banner_blocks_warn(self) -> None:
        findings = self.body_with(BANNER, "' DO WORK", "    'Count once",
                                  "        Debug.Print 1", "' EXIT PROCEDURE", BANNER)
        self.assertFinding(findings, 'section label "DO WORK" has no closing banner line')
        self.assertFinding(findings, 'section label "EXIT PROCEDURE" has no opening banner line')

    def test_group_banner_component_header_and_comments_are_silent(self) -> None:
        group = ["'", BANNER, "'", "'" + " " * 35 + "SETTINGS", "'", BANNER, "'"]
        decl = ["    Private mName As String", ""] + group
        proc = procedure("Public Sub M_Test_A()", cvs.FIELDS)
        text = module("M_Test", decl + proc[:-2] + ["    'A note", "' column-0 note"] + proc[-2:])
        self.assertNoFinding(self.run_check(text), "section label")
        eq = BANNER.replace("-", "=")
        comp = header(cvs.FIELDS, "' MODULE: M_Test")
        comp = [eq if ln == BANNER else ln for ln in comp]
        self.assertNoFinding(self.run_check(module("M_Test", [], component=comp)),
                             "section label")

    def test_banner_findings_do_not_change_errors(self) -> None:
        findings = self.body_with(BANNER, "' Do work", BANNER,
                                  title="' Do Work")
        self.assertTrue(findings)
        self.assertFalse(any(f.startswith("ERROR") for f in findings), findings)

    # Statement indentation (#25).
    NESTED = ["    'Walk the items",
              "        For Index = 1 To 3",
              "            If Index = 1 Then",
              "                Debug.Print 1",
              "            ElseIf Index = 2 Then",
              "                Debug.Print 2",
              "            Else",
              "                Debug.Print 3",
              "            End If",
              "            Select Case Index",
              "                Case 1",
              "                    Debug.Print 1",
              "                Case Else",
              "                    Debug.Print 2",
              "            End Select",
              "            If Index = 3 Then Exit For",
              "            With Application",
              "                .StatusBar = False",
              "            End With",
              "            Do While Index < 0",
              "                Index = Index + 1",
              "            Loop",
              "        Next Index"]

    def test_nested_blocks_at_their_levels_are_silent(self) -> None:
        self.assertEqual(self.sections(DECLARE, ("DO WORK", self.NESTED)), [])

    def test_statement_out_of_indentation_warns(self) -> None:
        for at, line in ((0, "      Debug.Print 0"), (3, "            Debug.Print 1"),
                         (11, "                Debug.Print 1")):
            with self.subTest(line=line):
                code = self.NESTED[:at + 1] + [line] + self.NESTED[at + 1:]
                findings = self.sections(DECLARE, ("DO WORK", code))
                got = len(line) - len(line.lstrip())
                self.assertEqual(sum("indented" in f for f in findings), 1, findings)
                self.assertFinding(findings, f"WARN M_Test_Run: statement indented {got} "
                                             "spaces, expected")

    def test_closing_and_middle_lines_follow_their_opening(self) -> None:
        for old, new in (("            Else", "                Else"),
                         ("            End If", "                End If"),
                         ("                Case Else", "            Case Else"),
                         ("            End Select", "                End Select"),
                         ("        Next Index", "            Next Index")):
            with self.subTest(line=new):
                code = [new if ln == old else ln for ln in self.NESTED]
                findings = self.sections(DECLARE, ("DO WORK", code))
                self.assertFinding(findings, "statement indented")

    def test_one_finding_per_section_and_kind(self) -> None:
        code = ["    'Count", "    Debug.Print 1", "    Debug.Print 2",
                "  'Off", "    Debug.Print 3"]
        findings = self.sections(DECLARE, ("DO WORK", code), ("RETURN VALUE", code))
        self.assertEqual(sum("statement indented 4 spaces, expected 8; 2 more in DO WORK" in f
                             for f in findings), 1, findings)
        self.assertEqual(sum("intent comment indented 2 spaces, expected 4" in f
                             for f in findings), 2, findings)
        self.assertFinding(findings, "2 more in RETURN VALUE")

    def test_intent_comment_out_of_indentation_warns(self) -> None:
        for lead in (0, 2, 8):
            with self.subTest(lead=lead):
                code = [" " * lead + "'Count once", "        Debug.Print 1"]
                findings = self.sections(("DO WORK", code))
                self.assertFinding(findings, f"intent comment indented {lead} spaces, expected 4")

    def test_comments_inside_blocks_and_commented_out_code_are_silent(self) -> None:
        code = ["    'Loop", "        For Index = 1 To 3",
                "                'A note at any column inside a block",
                "            Debug.Print Index", "        Next Index",
                "'        Debug.Print Index", "    'A multi-line intent comment",
                "    'continues at 4 spaces"]
        self.assertEqual(self.sections(DECLARE, ("DO WORK", code)), [])

    def test_one_line_if_continuations_labels_and_directives_are_silent(self) -> None:
        code = ["    'Stop early", "        If Index = 0 Then Exit Sub",
                "    'Report the totals",
                "        Debug.Print Index, _",
                "            Index + 1",
                "    'Branch on a long condition",
                "        If Index > 1 And _",
                "                Index < 9 Then",
                "            Debug.Print Index",
                "        End If",
                "#If VBA7 Then",
                "        Debug.Print 1",
                "#End If"]
        self.assertEqual(self.sections(DECLARE, ("DO WORK", code)), [])
        self.assertEqual(self.sections(DECLARE, INIT, TASK, EXIT, HANDLER), [])

    def test_declare_and_lines_before_the_first_section_are_not_checked(self) -> None:
        declare = ("DECLARE", [DECLARE[1][0], "", "  Dim Count As Long"])
        self.assertNoFinding(self.sections(declare, TASK), "indented")
        findings = self.statements("Debug.Print 1", "  Debug.Print 2")
        self.assertNoFinding(findings, "indented")

    def test_unclassifiable_or_unbalanced_procedures_are_silent(self) -> None:
        for code in (["      A = 1: B = 2"],
                     ["10      Debug.Print 1"],
                     ["        If Index = 1 Then", "      Debug.Print 1"],
                     ["        Next Index", "      Debug.Print 1"],
                     ["        Case 1", "      Debug.Print 1"]):
            with self.subTest(code=code):
                findings = self.sections(DECLARE, ("DO WORK", code))
                self.assertNoFinding(findings, "indented")

    def test_blocks_chosen_by_conditional_compilation_are_silent(self) -> None:
        code = ["    'Pick the branch", "#If VBA7 Then", "        If Ready Then", "#Else",
                "        If Done Then", "#End If", "            Debug.Print 1",
                "#If VBA7 Then", "        End If", "#Else", "        End If", "#End If",
                "    'Off", "      Debug.Print 2"]
        self.assertNoFinding(self.sections(DECLARE, ("DO WORK", code)), "indented")
        plain = ["    'Pick the branch", "#If VBA7 Then", "        Debug.Print 1", "#Else",
                 "      Debug.Print 2", "#End If"]
        self.assertFinding(self.sections(DECLARE, ("DO WORK", plain)),
                           "statement indented 6 spaces, expected 8")

    def test_next_closing_two_loops(self) -> None:
        code = ["    'Walk the grid", "        For Row = 1 To 2", "            For Col = 1 To 2",
                "                Debug.Print Row, Col", "        Next Col, Row",
                "    'Done", "        Debug.Print 0"]
        self.assertEqual(self.sections(DECLARE, ("DO WORK", code)), [])

    def test_indentation_findings_do_not_change_errors(self) -> None:
        findings = self.sections(DECLARE, ("DO WORK", ["  'Off", "    Debug.Print 1"]))
        self.assertTrue(findings)
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

    # Header field values (#29).
    def header_check(self, signature: str = "Public Sub M_Test_A()",
                     text: dict[str, str] | None = None,
                     fields: list[str] | None = None, end: str = "End Sub",
                     component: list[str] | None = None, name: str = "M_Test",
                     suffix: str = ".bas") -> list[str]:
        title = "'" + " " * 35 + "DO WORK"
        proc = [signature, ""] + header(fields or cvs.FIELDS, title, text) + [
            "", "        Debug.Print 1", "", end]
        return self.run_check(module(name, proc, component), name, suffix)

    PARAMS_SIGNATURE = ("Public Sub M_Test_A( _\r\n"
                        "    ByVal Name As String, _\r\n"
                        "    Optional ByVal Mode As Long = 0, _\r\n"
                        "    ParamArray Rest() As Variant)")
    PARAMS_INPUTS = ("Name\r\n'     Person to greet\r\n'\r\n"
                     "'   Mode\r\n'     Optional style\r\n'       - 0 plain\r\n'\r\n"
                     "'   Rest\r\n'     Extra values")

    def test_header_values_that_follow_the_skill_are_silent(self) -> None:
        findings = self.header_check(self.PARAMS_SIGNATURE, {"INPUTS": self.PARAMS_INPUTS})
        self.assertEqual(findings, [])

    def test_updated_that_is_not_a_date_warns(self) -> None:
        for value in ("2026-02-30", "15/01/2026", "2026-1-5", "Today"):
            with self.subTest(value=value):
                findings = self.header_check(text={"UPDATED": value})
                self.assertFinding(findings, f"UPDATED {value} is not a YYYY-MM-DD date")

    def test_real_dates_are_silent(self) -> None:
        for value in ("2024-02-29", "2026-12-31"):
            with self.subTest(value=value):
                self.assertNoFinding(self.header_check(text={"UPDATED": value}), "UPDATED")

    def test_component_updated_is_checked(self) -> None:
        comp = header(cvs.FIELDS, "' MODULE: M_Test", {**MODULE_TEXT, "UPDATED": "soon"})
        findings = self.run_check(module("M_Test", [], component=comp))
        self.assertFinding(findings, "component UPDATED soon is not a YYYY-MM-DD date")

    def test_sub_that_does_not_return_nothing_warns(self) -> None:
        findings = self.header_check(text={"RETURNS": "The count"})
        self.assertFinding(findings, "RETURNS of a Sub is not Nothing")

    def test_returns_of_functions_and_nothing_with_text_are_silent(self) -> None:
        findings = self.header_check("Public Function M_Test_A() As Long",
                                     {"RETURNS": "The count"}, end="End Function")
        self.assertNoFinding(findings, "RETURNS")
        findings = self.header_check(text={"RETURNS": "Nothing; the result goes to the sheet"})
        self.assertNoFinding(findings, "RETURNS")

    def test_inputs_missing_a_parameter_warns(self) -> None:
        findings = self.header_check(self.PARAMS_SIGNATURE, {"INPUTS": "Name\r\n'\r\n'   Rest"})
        self.assertFinding(findings, "M_Test_A: INPUTS does not list Mode")

    def test_inputs_listing_a_name_outside_the_signature_warns(self) -> None:
        text = {"INPUTS": self.PARAMS_INPUTS + "\r\n'\r\n'   Count"}
        findings = self.header_check(self.PARAMS_SIGNATURE, text)
        self.assertFinding(findings, "M_Test_A: INPUTS lists Count, not in the signature")
        findings = self.header_check(text={"INPUTS": "Count"})
        self.assertFinding(findings, "INPUTS lists Count, not in the signature")

    def test_inputs_none_with_parameters_warns(self) -> None:
        findings = self.header_check(self.PARAMS_SIGNATURE)
        self.assertFinding(findings, "INPUTS is None but the signature declares Name, Mode, Rest")

    def test_inputs_that_cannot_be_read_and_case_differences_are_silent(self) -> None:
        signature = "Public Sub M_Test_A(ByVal Name As String)"
        for inputs in ("Name - the person to greet", "name", "None of the callers pass it"):
            with self.subTest(inputs=inputs):
                findings = self.header_check(signature, {"INPUTS": inputs})
                if inputs.startswith("None"):
                    self.assertFinding(findings, "INPUTS is None")
                else:
                    self.assertNoFinding(findings, "INPUTS")

    def test_property_let_lists_its_value_parameter(self) -> None:
        findings = self.header_check("Public Property Let Count(ByVal Value As Long)",
                                     {"INPUTS": "Value"}, end="End Property")
        self.assertNoFinding(findings, "INPUTS")
        self.assertNoFinding(findings, "RETURNS")

    def test_compact_header_checks_only_the_fields_present(self) -> None:
        signature = "Private Sub M_Test_A(ByVal Name As String)"
        findings = self.header_check(signature, fields=["PURPOSE", "UPDATED"])
        self.assertNoFinding(findings, "INPUTS")
        findings = self.header_check(signature, {"RETURNS": "The name"},
                                     fields=["PURPOSE", "RETURNS", "UPDATED"])
        self.assertFinding(findings, "RETURNS of a Sub is not Nothing")

    def test_component_inputs_wording_warns(self) -> None:
        comp = header(cvs.FIELDS, "' MODULE: M_Test")
        findings = self.run_check(module("M_Test", [], component=comp))
        self.assertFinding(findings, "component INPUTS is not 'None at module level'")
        comp = header(cvs.FIELDS, "' CLASS: C_Test", MODULE_TEXT)
        findings = self.run_check(module("C_Test", [], component=comp), "C_Test", ".cls")
        self.assertFinding(findings, "component INPUTS is not 'None at class level'")

    def test_component_inputs_at_module_and_class_level_are_silent(self) -> None:
        self.assertNoFinding(self.run_check(module("M_Test", [])), "INPUTS")
        comp = header(cvs.FIELDS, "' CLASS: C_Test", {"INPUTS": "None at class level."})
        findings = self.run_check(module("C_Test", [], component=comp), "C_Test", ".cls")
        self.assertNoFinding(findings, "INPUTS")

    def test_field_text_not_at_three_spaces_warns(self) -> None:
        for text in ("'  Two spaces", "'    Four spaces", "'Flush"):
            with self.subTest(text=text):
                comp = header(cvs.FIELDS, "' MODULE: M_Test", MODULE_TEXT)
                comp[comp.index("' PURPOSE") + 1] = text
                findings = self.run_check(module("M_Test", [], component=comp))
                self.assertFinding(findings, "component PURPOSE text not at 3 spaces")

    def test_deeper_continuation_lines_are_silent(self) -> None:
        text = {"BEHAVIOR": "Manages:\r\n'     - a prefix\r\n'       detail"}
        self.assertNoFinding(self.header_check(text=text), "text not at 3 spaces")

    def test_header_value_findings_do_not_change_errors(self) -> None:
        text = {"UPDATED": "soon", "RETURNS": "A value", "INPUTS": "Count"}
        findings = self.header_check(text=text)
        self.assertTrue(findings)
        self.assertFalse(any(f.startswith("ERROR") for f in findings), findings)

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
