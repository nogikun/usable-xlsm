Attribute VB_Name = "UsableXlsmAssert"
Option Explicit

' Assertion helpers for the usable-xlsm test harness.
'
' A failed assertion raises ASSERT_ERROR rather than returning a flag, so the
' rest of the test body is skipped once an expectation is already broken. The
' generated runner catches it with On Error Resume Next, which is also what
' keeps a failing test from raising a modal dialog and hanging Excel.
'
' The harness injects this module automatically. Copy it into your project only
' if you want the VBE to resolve these names while you are writing tests.

Public Const ASSERT_ERROR As Long = vbObjectError + 9001

Public Sub Fail(Optional ByVal message As String = "explicit failure")
    Err.Raise ASSERT_ERROR, "Fail", message
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String = "")
    If condition Then Exit Sub
    Err.Raise ASSERT_ERROR, "AssertTrue", Describe(message, "expected True but got False")
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String = "")
    If Not condition Then Exit Sub
    Err.Raise ASSERT_ERROR, "AssertFalse", Describe(message, "expected False but got True")
End Sub

Public Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, Optional ByVal message As String = "")
    If ValuesEqual(actual, expected) Then Exit Sub
    Err.Raise ASSERT_ERROR, "AssertEqual", _
        Describe(message, "expected " & Display(expected) & " but got " & Display(actual))
End Sub

Public Sub AssertNotEqual(ByVal actual As Variant, ByVal unexpected As Variant, Optional ByVal message As String = "")
    If Not ValuesEqual(actual, unexpected) Then Exit Sub
    Err.Raise ASSERT_ERROR, "AssertNotEqual", _
        Describe(message, "expected a value other than " & Display(unexpected))
End Sub

' Floating point arithmetic rarely lands on an exact value, so comparing
' calculated Doubles with AssertEqual produces failures that are not real.
'
' Note the argument order: message comes third, matching every other assertion
' here. Putting tolerance there instead meant that the natural positional call
' AssertAlmostEqual a, b, "note" pushed a String into a Double and failed with
' a type mismatch that said nothing about the real mistake.
Public Sub AssertAlmostEqual(ByVal actual As Double, ByVal expected As Double, _
                             Optional ByVal message As String = "", _
                             Optional ByVal tolerance As Double = 0.0000001)
    If Abs(actual - expected) <= tolerance Then Exit Sub
    Err.Raise ASSERT_ERROR, "AssertAlmostEqual", _
        Describe(message, "expected " & CStr(expected) & " +/- " & CStr(tolerance) & _
                 " but got " & CStr(actual))
End Sub

Public Sub AssertIsNothing(ByVal value As Variant, Optional ByVal message As String = "")
    If IsObject(value) Then
        If value Is Nothing Then Exit Sub
    End If
    Err.Raise ASSERT_ERROR, "AssertIsNothing", Describe(message, "expected Nothing")
End Sub

Public Sub AssertNotNothing(ByVal value As Variant, Optional ByVal message As String = "")
    If IsObject(value) Then
        If Not value Is Nothing Then Exit Sub
    End If
    Err.Raise ASSERT_ERROR, "AssertNotNothing", Describe(message, "expected an object")
End Sub

Private Function Describe(ByVal message As String, ByVal fallback As String) As String
    If Len(message) = 0 Then
        Describe = fallback
    Else
        Describe = message & " (" & fallback & ")"
    End If
End Function

Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsObject(a) Or IsObject(b) Then
        If Not (IsObject(a) And IsObject(b)) Then Exit Function
        ValuesEqual = (a Is b)
    ElseIf IsNull(a) Or IsNull(b) Then
        ValuesEqual = (IsNull(a) And IsNull(b))
    ElseIf IsEmpty(a) Or IsEmpty(b) Then
        ValuesEqual = (IsEmpty(a) And IsEmpty(b))
    ElseIf IsNumeric(a) And IsNumeric(b) Then
        ValuesEqual = (CDbl(a) = CDbl(b))
    Else
        ValuesEqual = (CStr(a) = CStr(b))
    End If
End Function

Private Function Display(ByVal value As Variant) As String
    If IsObject(value) Then
        If value Is Nothing Then
            Display = "Nothing"
        Else
            Display = "<" & TypeName(value) & ">"
        End If
    ElseIf IsNull(value) Then
        Display = "Null"
    ElseIf IsEmpty(value) Then
        Display = "Empty"
    ElseIf VarType(value) = vbString Then
        Display = """" & CStr(value) & """"
    Else
        Display = CStr(value)
    End If
End Function
