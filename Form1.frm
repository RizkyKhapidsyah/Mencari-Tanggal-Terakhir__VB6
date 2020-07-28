VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencari Tanggal Terakhir"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Tanggal As Date
Dim JumlahTanggal As Byte
On Error Resume Next
   Tanggal = InputBox$("Masukkan Sebuah Tanggal", _
                     "Cari Jumlah Tanggal")
   JumlahTanggal = CekTanggal(Tanggal)
   MsgBox JumlahTanggal
End Sub

Function CekTanggal(strTanggal As Date) As Integer
Dim strTgl As String, intKabisat As Integer
Dim dd As Integer, mm As Integer, yyyy As Integer
 On Error GoTo Pesan
 strTgl = Format(strTanggal, "dd/mm/yyyy")
 'Konversikan ke string
 dd = Int(Left(strTgl, 2))    'Ambil 2 angka pertama
                              'untuk tanggal
 mm = Int(Mid(strTgl, 4, 2))  'Ambil 2 angka di tengah
                              'untuk bulan
 yyyy = Int(Right(strTgl, 4)) 'Ambil 4 angka terakhir
                              'untuk tahun
 intKabisat = yyyy Mod 4      'Set variabel kabisat
 'Lakukan pemeriksaan untuk memperoleh jumlah tanggal
    If ((dd >= 1) And (dd <= 31)) And ((mm = 1) _
      Or (mm = 3) Or (mm = 5) Or (mm = 7) Or (mm = 8) _
      Or (mm = 10) Or (mm = 12)) Then
      CekTanggal = 31
    ElseIf ((dd >= 1) And (dd <= 30)) And ((mm = 4) _
      Or (mm = 6) Or (mm = 9) Or (mm = 11)) Then
      CekTanggal = 30
    ElseIf ((dd >= 1) And (dd <= 28)) And (mm = 2) _
      And (intKabisat <> 0) Then
      CekTanggal = 28
    ElseIf (dd = 29) And (mm = 2) And (intKabisat = 0) Then
      CekTanggal = 29
    Else
      CekTanggal = 29
   End If
   Exit Function
Pesan:
   MsgBox "Tanggal atau formatnya salah!", _
          vbCritical, "Error Tanggal"
End Function


