VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf��������� 
   Caption         =   "��������� �������"
   ClientHeight    =   3860
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7490
   OleObjectBlob   =   "uf���������.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf���������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ������() As String
  letter = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
  Number = "00112233445566778899"
  symbol = "-_-_-_"
  
  If cb�����.Value Then
    material = material & Number
    lb������� = lb������ & Number
  ElseIf cb�����.Value Then
    material = material & letter
    lb������� = lb������ & letter
  ElseIf cb�����.Value Then
    material = material & symbol
    lb������� = lb������ & symbol
  ElseIf cb�����.Value And cb�����.Value Then
    material = material & Number & letter
    lb������� = lb������ & Number & letter
  ElseIf cb�����.Value And cb�����.Value Then
    material = material & Number & symbol
    lb������� = lb������ & Number & symbol
  ElseIf cb�����.Value And cb�����.Value Then
    material = material & symbol & letter
    lb������� = lb������ & symbol & letter
  ElseIf cb�����.Value And cb�����.Value And cb�����.Value Then
    material = material & symbol & letter & Number
    lb������� = lb������ & symbol & letter & Number
  Else: MsgBox "�������� ��������!", vbExclamation, "������"
  End If
  
  If tb�����.Text = "" Then
    MsgBox "�������� ����� ������!", vbExclamation, "������"
  Else:
    dlmat = Len(material)
    Randomize
    For i = 1 To CInt(tb�����.Text)
        ������ = ������ & Mid(material, 1 + Int(dlmat * Rnd), 1)
    Next i
  End If
End Function
Private Sub CommandButton1_Click()
If cb�����.Value = False Then
  Dim MyData As New DataObject
  MyData.SetText tb����������.Text
  MyData.PutInClipboard
End If
End Sub

Private Sub sb�������_Change()
 tb�����.Text = sb�������.Value
End Sub

Private Sub UserForm_Click()
  tb���������� = ������
End Sub
Private Sub UserForm_Activate()
  MsgBox "����� ���������� � ���������! ��� ��������� ������ ������� � ����� ������ �����.", vbInformation, "�����������"
End Sub
