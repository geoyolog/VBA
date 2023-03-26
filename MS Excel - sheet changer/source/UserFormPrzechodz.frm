VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPrzechodz 
   Caption         =   "Przejdü do arkusza"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   OleObjectBlob   =   "UserFormPrzechodz.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPrzechodz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub przejscie()
    Worksheets(ListBox1.Value).Activate
    If (CB_A1.Value) Then
        Range("A1").Activate
    End If
End Sub

Private Sub BUT_info_Click()
MsgBox ("Author: PK" & vbCrLf & "Special thanks to MR"  & vbCrLf & "ver. 1.0")
End Sub

Private Sub BUT_odswiez_Click()
    Call UserForm_Initialize
End Sub

Private Sub BUT_przejdz_Click()
    Call przejscie
End Sub
Private Sub BUT_zamknij_Click()
    Unload Me
End Sub

Private Sub CB_auto_przechodz_Change()

    If (CB_auto_przechodz.Value) Then
        BUT_przejdz.Locked = True
    Else
        BUT_przejdz.Locked = False
    End If
    
End Sub
Private Sub ListBox1_Click()
    
    If (CB_auto_przechodz.Value) Then
        Call przejscie
    End If
        
End Sub

Private Sub UserForm_Initialize()
Dim I_ile As Long

ListBox1.Clear

If (Application.Workbooks.Count) Then
    ListBox1.Visible = True
    BUT_przejdz.Visible = True
    
    For I_ile = 1 To ActiveWorkbook.Sheets.Count
        ListBox1.AddItem ActiveWorkbook.Sheets(I_ile).Name
    Next I_ile

    ListBox1.Value = ActiveSheet.Name
Else
    MsgBox ("Nie znaleziono otwartego skoroszytu")
    ListBox1.Visible = False
    BUT_przejdz.Visible = False
    
End If





End Sub


