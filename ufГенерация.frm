VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufГенерация 
   Caption         =   "Генератор паролей"
   ClientHeight    =   3860
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7490
   OleObjectBlob   =   "ufГенерация.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufГенерация"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function Пароль() As String
  letter = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
  Number = "00112233445566778899"
  symbol = "-_-_-_"
  
  If cbЦифры.Value Then
    material = material & Number
    lbВыбСимв = lbВыбСим & Number
  ElseIf cbБуквы.Value Then
    material = material & letter
    lbВыбСимв = lbВыбСим & letter
  ElseIf cbЗнаки.Value Then
    material = material & symbol
    lbВыбСимв = lbВыбСим & symbol
  ElseIf cbЦифры.Value And cbБуквы.Value Then
    material = material & Number & letter
    lbВыбСимв = lbВыбСим & Number & letter
  ElseIf cbЦифры.Value And cbЗнаки.Value Then
    material = material & Number & symbol
    lbВыбСимв = lbВыбСим & Number & symbol
  ElseIf cbЗнаки.Value And cbБуквы.Value Then
    material = material & symbol & letter
    lbВыбСимв = lbВыбСим & symbol & letter
  ElseIf cbЗнаки.Value And cbБуквы.Value And cbЦифры.Value Then
    material = material & symbol & letter & Number
    lbВыбСимв = lbВыбСим & symbol & letter & Number
  Else: MsgBox "Выберите материал!", vbExclamation, "Ошибка"
  End If
  
  If tbДлина.Text = "" Then
    MsgBox "Выберите длину пароля!", vbExclamation, "Ошибка"
  Else:
    dlmat = Len(material)
    Randomize
    For i = 1 To CInt(tbДлина.Text)
        Пароль = Пароль & Mid(material, 1 + Int(dlmat * Rnd), 1)
    Next i
  End If
End Function
Private Sub CommandButton1_Click()
If cbЗапом.Value = False Then
  Dim MyData As New DataObject
  MyData.SetText tbИтогПароль.Text
  MyData.PutInClipboard
End If
End Sub

Private Sub sbСчетчик_Change()
 tbДлина.Text = sbСчетчик.Value
End Sub

Private Sub UserForm_Click()
  tbИтогПароль = Пароль
End Sub
Private Sub UserForm_Activate()
  MsgBox "Добро пожаловать в программу! Для генерации пароля нажмите в любое пустое место.", vbInformation, "Приветствие"
End Sub
