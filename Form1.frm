VERSION 5.00
Begin VB.Form FormUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Utama"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Static Sub StrSort(words() As String, _
   Ascending As Boolean, AllLowerCase As Boolean)


Dim I As Integer
Dim J As Integer
Dim NumInArray, LowerBound As Integer
NumInArray = UBound(words)
LowerBound = LBound(words)
For I = LowerBound To NumInArray
    J = 0
    For J = LowerBound To NumInArray
        If AllLowerCase = True Then
            If Ascending = True Then
                If StrComp(LCase(words(I)), _
                     LCase(words(J))) = -1 Then
                    Call Swap(words(I), words(J))
                End If
            Else
                If StrComp(LCase(words(I)), _
                       LCase(words(J))) = 1 Then
                    Call Swap(words(I), words(J))
                End If
            End If
        Else
            If Ascending = True Then
                If StrComp(words(I), words(J)) = -1 Then
                    Call Swap(words(I), words(J))
                End If
            Else
                If StrComp(words(I), _
                    words(J)) = 1 Then
                    Call Swap(words(I), words(J))
                End If
            End If
        End If
    Next J
Next I
End Sub

Public Static Sub NumSort(nums() As Variant, Ascending As Boolean)

'Pass in numeric array you want to sort by reference and
'read it back.  The array should be declared as an array
'of variants

'Set Ascending to True to sort ascending,
'false to sort descending

    Dim I As Integer
    Dim J As Integer
    Dim NumInArray, LowerBound As Integer
    NumInArray = UBound(nums)
    LowerBound = LBound(nums)
    For I = LowerBound To NumInArray
        J = 0
        For J = LowerBound To NumInArray
            If Ascending = True Then
                If nums(I) < nums(J) Then
                    NumSwap nums(I), nums(J)
                End If
            Else
                If nums(I) > nums(J) Then
                    NumSwap nums(I), nums(J)
                End If
            End If
        Next J
    Next I
End Sub

Private Sub NumSwap(var1 As Variant, var2 As Variant)
    Dim x As Variant
    x = var1
    var1 = var2
    var2 = x
End Sub

Private Sub Swap(var1 As String, var2 As String)
    Dim x As String
    x = var1
    var1 = var2
    var2 = x
End Sub

Private Sub Form_Load()
Dim iArray(9) As Variant
Dim sArray(9) As String
Dim iCtr As Integer


For iCtr = 0 To 9
  iArray(iCtr) = RandomNumber(10000, 0)
  sArray(iCtr) = Chr(RandomNumber(90, 65))
Next

StrSort sArray, False, False
NumSort iArray, False

For iCtr = 0 To 9
  Debug.Print iArray(iCtr)
  Debug.Print sArray(iCtr)
Next
End Sub
