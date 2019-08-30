VERSION 5.00
Begin VB.Form frmAES 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Criptografia AES"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescripFinal 
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox txtTexto 
      Height          =   375
      Left            =   1920
      MaxLength       =   16
      TabIndex        =   13
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdDescriptografar 
      Caption         =   "Descriptografar"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtDescriptografia 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
      Width           =   4695
   End
   Begin VB.CommandButton cmdCriptografar 
      Caption         =   "Criptografar"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtCriptografia 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox txtKey 
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   1920
      MaxLength       =   32
      TabIndex        =   6
      Text            =   "000102030405060708090a0b0c0d0e0f"
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Text            =   "00112233445566778899aabbccddeeff"
      Top             =   600
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSubByte 
      Caption         =   "subByte"
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Texto Descriptografado:"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Texto p/ criptografar:"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Texto Descriptografado (Hexa):"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Texto criptografado:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Key:"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Texto p/ criptografar (Hexa):"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmAES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const U32_CVT = 4294967296#
Private mSBox       As Variant
Private mSBoxInv    As Variant
Private mState      As Variant
Private mNr         As Long
Private mNb         As Long
Private mNk         As Long
Private mWord       As Variant
Private mstrKey     As String
Private mRCon       As Variant
Private mWordInv    As Variant
Private m_lOnBits(30)   As Long
Private m_l2Power(30)   As Long
Private m_bytOnBits(7)  As Byte
Private m_byt2Power(7)  As Byte

Private m_InCo(3) As Byte

Private m_fbsub(255)    As Byte
Private m_rbsub(255)    As Byte
Private m_ptab(255)     As Byte
Private m_ltab(255)     As Byte
Private m_ftable(255)   As Long
Private m_rtable(255)   As Long
Private m_rco(29)       As Long

Private m_Nk        As Long
Private m_Nb        As Long
Private m_Nr        As Long
Private m_fi(23)    As Byte
Private m_ri(23)    As Byte
Private m_fkey(119) As Long
Private m_rkey(119) As Long


Private Sub Command1_Click()
Dim oAes        As New clsAES
Dim k           As String
Dim p           As String
Dim b           As Byte
Dim vKey(31)    As Byte
Dim vPlain(31)  As Byte
Dim n           As Long
Dim i           As Long
Dim r           As Variant

k = "000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f"
p = "112233445566778899aabbccddeeff0000000000000000000000000000000000"
n = 0
For i = 1 To 32 Step 2
    If Mid(k, i, 2) = "00" Then
        b = 0
    Else
        b = HexToDec(Mid(k, i, 2))
    End If
    vKey(n) = b
    n = n + 1
    
Next

n = 0
For i = 1 To 32 Step 2
    If Mid(p, i, 2) = "00" Then
        b = 0
    Else
        b = HexToDec(Mid(p, i, 2))
    End If
    vPlain(n) = b
    n = n + 1
    
Next

'Call oAes.gkey(4, 8, v)
'Call oAes.gentables
r = oAes.DecryptData(vPlain, vKey)


End Sub

Private Sub cmdCriptografar_Click()
Dim lngRound    As Long
Dim i           As Long
Dim c           As String

'testes
'text: 3243f6a8885a308d313198a2e0370734
'key.: 2b7e151628aed2a6abf7158809cf4f3c
'crip: 3925841d02dc09fbdc118597196a0b32

'text: 00112233445566778899aabbccddeeff
'key.: 000102030405060708090a0b0c0d0e0f
'crip: 69c4e0d86a7b0430d8cdb78070b4c55a
    


Call IniciarSBox
Call IniciarMState
Call KeyExpansion

Call AddRound(0)

For lngRound = 1 To mNr - 1
    Call SubBytes
    Call ShiftRows
    Call MixColumns
    Call AddRound(lngRound)
Next

Call SubBytes
Call ShiftRows
Call AddRound(lngRound)

c = ""
For i = 0 To 15
   c = c & LCase(mState(i))
Next
txtCriptografia.Text = c

txtDescriptografia.Text = ""
txtDescripFinal.Text = ""

End Sub

Private Sub cmdDescriptografar_Click()
Dim lngRound    As Long
Dim i           As Long
Dim c           As String

Call InvIniciarSBox
Call IniciarSBox
Call InvIniciarMState 'ok
Call KeyExpansion
Call InvKeyExpansion

Call AddRound(mNr)

For lngRound = mNr - 1 To 1 Step -1
    Call InvShiftRows 'ok
    Call InvSubBytes 'ok - mesmo
    Call AddRound(lngRound)
    Call InvMixColumns
Next

Call InvShiftRows
Call InvSubBytes
Call AddRound(lngRound)

c = ""
For i = 0 To 15
   c = c & LCase(mState(i))
Next
txtDescriptografia.Text = c

txtDescripFinal.Text = ConverterTextoHexToChar(c)

End Sub

Private Sub cmdSubByte_Click()
Caption = SubByteByte(Text1.Text)
End Sub

Private Sub SubBytes()
'entradas: array com bytes em hexa do texto a ser cifrado
'saidas  : o mesmo array convertido com subbytes
Dim i As Long

For i = 0 To 15
    mState(i) = SubByteByte(CStr(mState(i)))
Next

End Sub
Private Sub InvSubBytes()
'entradas: array com bytes em hexa do texto a ser cifrado
'saidas  : o mesmo array convertido com subbytes
Dim i As Long

For i = 0 To 15
    mState(i) = InvSubByteByte(CStr(mState(i)))
Next

End Sub

Private Function HexToDec(strNumero As Variant) As String
    '*****************************************************************
    ' Propósito : Converter o Nº. de HexaDecimal para a base Decimal
    ' Entrada   : O Nº. a ser convertido
    ' Retornou  : O Nº. convertido
    ' Autor     : Douglas Ferreira
    ' Criação   : 03-08-2002
    '*****************************************************************
    Dim lngExpoente As Long
    Dim k           As Long
    Dim dblResult   As Double
    Dim lngCadaNum  As Long
    
    strNumero = Trim(strNumero)
    lngExpoente = Len(strNumero) - 1
    
    For k = 1 To Len(strNumero)
        lngCadaNum = letraHexaToNum(Mid(strNumero, k, 1))
        dblResult = dblResult + CDbl(lngCadaNum * 16 ^ lngExpoente)
        lngExpoente = lngExpoente - 1
    Next
    
    HexToDec = Trim(Str(dblResult))

End Function

Private Function letraHexaToNum(varLetra As Variant) As Long
    Select Case UCase(varLetra)
        Case "A"
            letraHexaToNum = 10
        Case "B"
            letraHexaToNum = 11
        Case "C"
            letraHexaToNum = 12
        Case "D"
            letraHexaToNum = 13
        Case "E"
            letraHexaToNum = 14
        Case "F"
            letraHexaToNum = 15
        Case Else
            letraHexaToNum = Val(varLetra)
    End Select
End Function

Private Sub IniciarSBox()
    ReDim mSBox(15, 15)
    mSBox(0, 0) = "63"
    mSBox(0, 1) = "7c"
    mSBox(0, 2) = "77"
    mSBox(0, 3) = "7b"
    mSBox(0, 4) = "f2"
    mSBox(0, 5) = "6b"
    mSBox(0, 6) = "6f"
    mSBox(0, 7) = "c5"
    mSBox(0, 8) = "30"
    mSBox(0, 9) = "01"
    mSBox(0, 10) = "67"
    mSBox(0, 11) = "2b"
    mSBox(0, 12) = "fe"
    mSBox(0, 13) = "d7"
    mSBox(0, 14) = "ab"
    mSBox(0, 15) = "76"
    
    mSBox(1, 0) = "ca"
    mSBox(1, 1) = "82"
    mSBox(1, 2) = "c9"
    mSBox(1, 3) = "7d"
    mSBox(1, 4) = "fa"
    mSBox(1, 5) = "59"
    mSBox(1, 6) = "47"
    mSBox(1, 7) = "f0"
    mSBox(1, 8) = "ad"
    mSBox(1, 9) = "d4"
    mSBox(1, 10) = "a2"
    mSBox(1, 11) = "af"
    mSBox(1, 12) = "9c"
    mSBox(1, 13) = "a4"
    mSBox(1, 14) = "72"
    mSBox(1, 15) = "c0"
    
    mSBox(2, 0) = "b7"
    mSBox(2, 1) = "fd"
    mSBox(2, 2) = "93"
    mSBox(2, 3) = "26"
    mSBox(2, 4) = "36"
    mSBox(2, 5) = "3f"
    mSBox(2, 6) = "f7"
    mSBox(2, 7) = "cc"
    mSBox(2, 8) = "34"
    mSBox(2, 9) = "a5"
    mSBox(2, 10) = "e5"
    mSBox(2, 11) = "f1"
    mSBox(2, 12) = "71"
    mSBox(2, 13) = "d8"
    mSBox(2, 14) = "31"
    mSBox(2, 15) = "15"

    mSBox(3, 0) = "04"
    mSBox(3, 1) = "c7"
    mSBox(3, 2) = "23"
    mSBox(3, 3) = "c3"
    mSBox(3, 4) = "18"
    mSBox(3, 5) = "96"
    mSBox(3, 6) = "05"
    mSBox(3, 7) = "9a"
    mSBox(3, 8) = "07"
    mSBox(3, 9) = "12"
    mSBox(3, 10) = "80"
    mSBox(3, 11) = "e2"
    mSBox(3, 12) = "eb"
    mSBox(3, 13) = "27"
    mSBox(3, 14) = "b2"
    mSBox(3, 15) = "75"
    
    mSBox(4, 0) = "09"
    mSBox(4, 1) = "83"
    mSBox(4, 2) = "2c"
    mSBox(4, 3) = "1a"
    mSBox(4, 4) = "1b"
    mSBox(4, 5) = "6e"
    mSBox(4, 6) = "5a"
    mSBox(4, 7) = "a0"
    mSBox(4, 8) = "52"
    mSBox(4, 9) = "3b"
    mSBox(4, 10) = "d6"
    mSBox(4, 11) = "b3"
    mSBox(4, 12) = "29"
    mSBox(4, 13) = "e3"
    mSBox(4, 14) = "2f"
    mSBox(4, 15) = "84"

    mSBox(5, 0) = "53"
    mSBox(5, 1) = "d1"
    mSBox(5, 2) = "00"
    mSBox(5, 3) = "ed"
    mSBox(5, 4) = "20"
    mSBox(5, 5) = "fc"
    mSBox(5, 6) = "b1"
    mSBox(5, 7) = "5b"
    mSBox(5, 8) = "6a"
    mSBox(5, 9) = "cb"
    mSBox(5, 10) = "be"
    mSBox(5, 11) = "39"
    mSBox(5, 12) = "4a"
    mSBox(5, 13) = "4c"
    mSBox(5, 14) = "58"
    mSBox(5, 15) = "cf"
   
    mSBox(6, 0) = "d0"
    mSBox(6, 1) = "ef"
    mSBox(6, 2) = "aa"
    mSBox(6, 3) = "fb"
    mSBox(6, 4) = "43"
    mSBox(6, 5) = "4d"
    mSBox(6, 6) = "33"
    mSBox(6, 7) = "85"
    mSBox(6, 8) = "45"
    mSBox(6, 9) = "f9"
    mSBox(6, 10) = "02"
    mSBox(6, 11) = "7f"
    mSBox(6, 12) = "50"
    mSBox(6, 13) = "3c"
    mSBox(6, 14) = "9f"
    mSBox(6, 15) = "a8"
   
    mSBox(7, 0) = "51"
    mSBox(7, 1) = "a3"
    mSBox(7, 2) = "40"
    mSBox(7, 3) = "8f"
    mSBox(7, 4) = "92"
    mSBox(7, 5) = "9d"
    mSBox(7, 6) = "38"
    mSBox(7, 7) = "f5"
    mSBox(7, 8) = "bc"
    mSBox(7, 9) = "b6"
    mSBox(7, 10) = "da"
    mSBox(7, 11) = "21"
    mSBox(7, 12) = "10"
    mSBox(7, 13) = "ff"
    mSBox(7, 14) = "f3"
    mSBox(7, 15) = "d2"
    
    mSBox(8, 0) = "cd"
    mSBox(8, 1) = "0c"
    mSBox(8, 2) = "13"
    mSBox(8, 3) = "ec"
    mSBox(8, 4) = "5f"
    mSBox(8, 5) = "97"
    mSBox(8, 6) = "44"
    mSBox(8, 7) = "17"
    mSBox(8, 8) = "c4"
    mSBox(8, 9) = "a7"
    mSBox(8, 10) = "7e"
    mSBox(8, 11) = "3d"
    mSBox(8, 12) = "64"
    mSBox(8, 13) = "5d"
    mSBox(8, 14) = "19"
    mSBox(8, 15) = "73"

    mSBox(9, 0) = "60"
    mSBox(9, 1) = "81"
    mSBox(9, 2) = "4f"
    mSBox(9, 3) = "dc"
    mSBox(9, 4) = "22"
    mSBox(9, 5) = "2a"
    mSBox(9, 6) = "90"
    mSBox(9, 7) = "88"
    mSBox(9, 8) = "46"
    mSBox(9, 9) = "ee"
    mSBox(9, 10) = "b8"
    mSBox(9, 11) = "14"
    mSBox(9, 12) = "de"
    mSBox(9, 13) = "5e"
    mSBox(9, 14) = "0b"
    mSBox(9, 15) = "db"
    
    mSBox(10, 0) = "e0"
    mSBox(10, 1) = "32"
    mSBox(10, 2) = "3a"
    mSBox(10, 3) = "0a"
    mSBox(10, 4) = "49"
    mSBox(10, 5) = "06"
    mSBox(10, 6) = "24"
    mSBox(10, 7) = "5c"
    mSBox(10, 8) = "c2"
    mSBox(10, 9) = "d3"
    mSBox(10, 10) = "ac"
    mSBox(10, 11) = "62"
    mSBox(10, 12) = "91"
    mSBox(10, 13) = "95"
    mSBox(10, 14) = "e4"
    mSBox(10, 15) = "79"

    mSBox(11, 0) = "e7"
    mSBox(11, 1) = "c8"
    mSBox(11, 2) = "37"
    mSBox(11, 3) = "6d"
    mSBox(11, 4) = "8d"
    mSBox(11, 5) = "d5"
    mSBox(11, 6) = "4e"
    mSBox(11, 7) = "a9"
    mSBox(11, 8) = "6c"
    mSBox(11, 9) = "56"
    mSBox(11, 10) = "f4"
    mSBox(11, 11) = "ea"
    mSBox(11, 12) = "65"
    mSBox(11, 13) = "7a"
    mSBox(11, 14) = "ae"
    mSBox(11, 15) = "08"

    mSBox(12, 0) = "ba"
    mSBox(12, 1) = "78"
    mSBox(12, 2) = "25"
    mSBox(12, 3) = "2e"
    mSBox(12, 4) = "1c"
    mSBox(12, 5) = "a6"
    mSBox(12, 6) = "b4"
    mSBox(12, 7) = "c6"
    mSBox(12, 8) = "e8"
    mSBox(12, 9) = "dd"
    mSBox(12, 10) = "74"
    mSBox(12, 11) = "1f"
    mSBox(12, 12) = "4b"
    mSBox(12, 13) = "bd"
    mSBox(12, 14) = "8b"
    mSBox(12, 15) = "8a"

    mSBox(13, 0) = "70"
    mSBox(13, 1) = "3e"
    mSBox(13, 2) = "b5"
    mSBox(13, 3) = "66"
    mSBox(13, 4) = "48"
    mSBox(13, 5) = "03"
    mSBox(13, 6) = "f6"
    mSBox(13, 7) = "0e"
    mSBox(13, 8) = "61"
    mSBox(13, 9) = "35"
    mSBox(13, 10) = "57"
    mSBox(13, 11) = "b9"
    mSBox(13, 12) = "86"
    mSBox(13, 13) = "c1"
    mSBox(13, 14) = "1d"
    mSBox(13, 15) = "9e"

    mSBox(14, 0) = "e1"
    mSBox(14, 1) = "f8"
    mSBox(14, 2) = "98"
    mSBox(14, 3) = "11"
    mSBox(14, 4) = "69"
    mSBox(14, 5) = "d9"
    mSBox(14, 6) = "8e"
    mSBox(14, 7) = "94"
    mSBox(14, 8) = "9b"
    mSBox(14, 9) = "1e"
    mSBox(14, 10) = "87"
    mSBox(14, 11) = "e9"
    mSBox(14, 12) = "ce"
    mSBox(14, 13) = "55"
    mSBox(14, 14) = "28"
    mSBox(14, 15) = "df"

    mSBox(15, 0) = "8c"
    mSBox(15, 1) = "a1"
    mSBox(15, 2) = "89"
    mSBox(15, 3) = "0d"
    mSBox(15, 4) = "bf"
    mSBox(15, 5) = "e6"
    mSBox(15, 6) = "42"
    mSBox(15, 7) = "68"
    mSBox(15, 8) = "41"
    mSBox(15, 9) = "99"
    mSBox(15, 10) = "2d"
    mSBox(15, 11) = "0f"
    mSBox(15, 12) = "b0"
    mSBox(15, 13) = "54"
    mSBox(15, 14) = "bb"
    mSBox(15, 15) = "16"


End Sub

Private Sub InvIniciarSBox()
    ReDim mSBoxInv(15, 15)
    mSBoxInv(0, 0) = "52"
    mSBoxInv(0, 1) = "09"
    mSBoxInv(0, 2) = "6a"
    mSBoxInv(0, 3) = "d5"
    mSBoxInv(0, 4) = "30"
    mSBoxInv(0, 5) = "36"
    mSBoxInv(0, 6) = "a5"
    mSBoxInv(0, 7) = "38"
    mSBoxInv(0, 8) = "bf"
    mSBoxInv(0, 9) = "40"
    mSBoxInv(0, 10) = "a3"
    mSBoxInv(0, 11) = "9e"
    mSBoxInv(0, 12) = "81"
    mSBoxInv(0, 13) = "f3"
    mSBoxInv(0, 14) = "d7"
    mSBoxInv(0, 15) = "fb"
    
    mSBoxInv(1, 0) = "7c"
    mSBoxInv(1, 1) = "e3"
    mSBoxInv(1, 2) = "39"
    mSBoxInv(1, 3) = "82"
    mSBoxInv(1, 4) = "9b"
    mSBoxInv(1, 5) = "2f"
    mSBoxInv(1, 6) = "ff"
    mSBoxInv(1, 7) = "87"
    mSBoxInv(1, 8) = "34"
    mSBoxInv(1, 9) = "8e"
    mSBoxInv(1, 10) = "43"
    mSBoxInv(1, 11) = "44"
    mSBoxInv(1, 12) = "c4"
    mSBoxInv(1, 13) = "de"
    mSBoxInv(1, 14) = "e9"
    mSBoxInv(1, 15) = "cb"
    
    mSBoxInv(2, 0) = "54"
    mSBoxInv(2, 1) = "7b"
    mSBoxInv(2, 2) = "94"
    mSBoxInv(2, 3) = "32"
    mSBoxInv(2, 4) = "a6"
    mSBoxInv(2, 5) = "c2"
    mSBoxInv(2, 6) = "23"
    mSBoxInv(2, 7) = "3d"
    mSBoxInv(2, 8) = "ee"
    mSBoxInv(2, 9) = "4c"
    mSBoxInv(2, 10) = "95"
    mSBoxInv(2, 11) = "0b"
    mSBoxInv(2, 12) = "42"
    mSBoxInv(2, 13) = "fa"
    mSBoxInv(2, 14) = "c3"
    mSBoxInv(2, 15) = "4e"
    
    mSBoxInv(3, 0) = "08"
    mSBoxInv(3, 1) = "2e"
    mSBoxInv(3, 2) = "a1"
    mSBoxInv(3, 3) = "66"
    mSBoxInv(3, 4) = "28"
    mSBoxInv(3, 5) = "d9"
    mSBoxInv(3, 6) = "24"
    mSBoxInv(3, 7) = "b2"
    mSBoxInv(3, 8) = "76"
    mSBoxInv(3, 9) = "5b"
    mSBoxInv(3, 10) = "a2"
    mSBoxInv(3, 11) = "49"
    mSBoxInv(3, 12) = "6d"
    mSBoxInv(3, 13) = "8b"
    mSBoxInv(3, 14) = "d1"
    mSBoxInv(3, 15) = "25"

    mSBoxInv(4, 0) = "72"
    mSBoxInv(4, 1) = "f8"
    mSBoxInv(4, 2) = "f6"
    mSBoxInv(4, 3) = "64"
    mSBoxInv(4, 4) = "86"
    mSBoxInv(4, 5) = "68"
    mSBoxInv(4, 6) = "98"
    mSBoxInv(4, 7) = "16"
    mSBoxInv(4, 8) = "d4"
    mSBoxInv(4, 9) = "a4"
    mSBoxInv(4, 10) = "5c"
    mSBoxInv(4, 11) = "cc"
    mSBoxInv(4, 12) = "5d"
    mSBoxInv(4, 13) = "65"
    mSBoxInv(4, 14) = "b6"
    mSBoxInv(4, 15) = "92"
    
    mSBoxInv(5, 0) = "6c"
    mSBoxInv(5, 1) = "70"
    mSBoxInv(5, 2) = "48"
    mSBoxInv(5, 3) = "50"
    mSBoxInv(5, 4) = "fd"
    mSBoxInv(5, 5) = "ed"
    mSBoxInv(5, 6) = "b9"
    mSBoxInv(5, 7) = "da"
    mSBoxInv(5, 8) = "5e"
    mSBoxInv(5, 9) = "15"
    mSBoxInv(5, 10) = "46"
    mSBoxInv(5, 11) = "57"
    mSBoxInv(5, 12) = "a7"
    mSBoxInv(5, 13) = "8d"
    mSBoxInv(5, 14) = "9d"
    mSBoxInv(5, 15) = "84"
    
    mSBoxInv(6, 0) = "90"
    mSBoxInv(6, 1) = "d8"
    mSBoxInv(6, 2) = "ab"
    mSBoxInv(6, 3) = "00"
    mSBoxInv(6, 4) = "8c"
    mSBoxInv(6, 5) = "bc"
    mSBoxInv(6, 6) = "d3"
    mSBoxInv(6, 7) = "0a"
    mSBoxInv(6, 8) = "f7"
    mSBoxInv(6, 9) = "e4"
    mSBoxInv(6, 10) = "58"
    mSBoxInv(6, 11) = "05"
    mSBoxInv(6, 12) = "b8"
    mSBoxInv(6, 13) = "b3"
    mSBoxInv(6, 14) = "45"
    mSBoxInv(6, 15) = "06"

    mSBoxInv(7, 0) = "d0"
    mSBoxInv(7, 1) = "2c"
    mSBoxInv(7, 2) = "1e"
    mSBoxInv(7, 3) = "8f"
    mSBoxInv(7, 4) = "ca"
    mSBoxInv(7, 5) = "3f"
    mSBoxInv(7, 6) = "0f"
    mSBoxInv(7, 7) = "02"
    mSBoxInv(7, 8) = "c1"
    mSBoxInv(7, 9) = "af"
    mSBoxInv(7, 10) = "bd"
    mSBoxInv(7, 11) = "03"
    mSBoxInv(7, 12) = "01"
    mSBoxInv(7, 13) = "13"
    mSBoxInv(7, 14) = "8a"
    mSBoxInv(7, 15) = "6b"
    
    mSBoxInv(8, 0) = "3a"
    mSBoxInv(8, 1) = "91"
    mSBoxInv(8, 2) = "11"
    mSBoxInv(8, 3) = "41"
    mSBoxInv(8, 4) = "4f"
    mSBoxInv(8, 5) = "67"
    mSBoxInv(8, 6) = "dc"
    mSBoxInv(8, 7) = "ea"
    mSBoxInv(8, 8) = "97"
    mSBoxInv(8, 9) = "f2"
    mSBoxInv(8, 10) = "cf"
    mSBoxInv(8, 11) = "ce"
    mSBoxInv(8, 12) = "f0"
    mSBoxInv(8, 13) = "b4"
    mSBoxInv(8, 14) = "e6"
    mSBoxInv(8, 15) = "73"

    mSBoxInv(9, 0) = "96"
    mSBoxInv(9, 1) = "ac"
    mSBoxInv(9, 2) = "74"
    mSBoxInv(9, 3) = "22"
    mSBoxInv(9, 4) = "e7"
    mSBoxInv(9, 5) = "ad"
    mSBoxInv(9, 6) = "35"
    mSBoxInv(9, 7) = "85"
    mSBoxInv(9, 8) = "e2"
    mSBoxInv(9, 9) = "f9"
    mSBoxInv(9, 10) = "37"
    mSBoxInv(9, 11) = "e8"
    mSBoxInv(9, 12) = "1c"
    mSBoxInv(9, 13) = "75"
    mSBoxInv(9, 14) = "df"
    mSBoxInv(9, 15) = "6e"

    mSBoxInv(10, 0) = "47"
    mSBoxInv(10, 1) = "f1"
    mSBoxInv(10, 2) = "1a"
    mSBoxInv(10, 3) = "71"
    mSBoxInv(10, 4) = "1d"
    mSBoxInv(10, 5) = "29"
    mSBoxInv(10, 6) = "c5"
    mSBoxInv(10, 7) = "89"
    mSBoxInv(10, 8) = "6f"
    mSBoxInv(10, 9) = "b7"
    mSBoxInv(10, 10) = "62"
    mSBoxInv(10, 11) = "0e"
    mSBoxInv(10, 12) = "aa"
    mSBoxInv(10, 13) = "18"
    mSBoxInv(10, 14) = "be"
    mSBoxInv(10, 15) = "1b"
    
    mSBoxInv(11, 0) = "fc"
    mSBoxInv(11, 1) = "56"
    mSBoxInv(11, 2) = "3e"
    mSBoxInv(11, 3) = "4b"
    mSBoxInv(11, 4) = "c6"
    mSBoxInv(11, 5) = "d2"
    mSBoxInv(11, 6) = "79"
    mSBoxInv(11, 7) = "20"
    mSBoxInv(11, 8) = "9a"
    mSBoxInv(11, 9) = "db"
    mSBoxInv(11, 10) = "c0"
    mSBoxInv(11, 11) = "fe"
    mSBoxInv(11, 12) = "78"
    mSBoxInv(11, 13) = "cd"
    mSBoxInv(11, 14) = "5a"
    mSBoxInv(11, 15) = "f4"
    
    mSBoxInv(12, 0) = "1f"
    mSBoxInv(12, 1) = "dd"
    mSBoxInv(12, 2) = "a8"
    mSBoxInv(12, 3) = "33"
    mSBoxInv(12, 4) = "88"
    mSBoxInv(12, 5) = "07"
    mSBoxInv(12, 6) = "c7"
    mSBoxInv(12, 7) = "31"
    mSBoxInv(12, 8) = "b1"
    mSBoxInv(12, 9) = "12"
    mSBoxInv(12, 10) = "10"
    mSBoxInv(12, 11) = "59"
    mSBoxInv(12, 12) = "27"
    mSBoxInv(12, 13) = "80"
    mSBoxInv(12, 14) = "ec"
    mSBoxInv(12, 15) = "5f"

    mSBoxInv(13, 0) = "60"
    mSBoxInv(13, 1) = "51"
    mSBoxInv(13, 2) = "7f"
    mSBoxInv(13, 3) = "a9"
    mSBoxInv(13, 4) = "19"
    mSBoxInv(13, 5) = "b5"
    mSBoxInv(13, 6) = "4a"
    mSBoxInv(13, 7) = "0d"
    mSBoxInv(13, 8) = "2d"
    mSBoxInv(13, 9) = "e5"
    mSBoxInv(13, 10) = "7a"
    mSBoxInv(13, 11) = "9f"
    mSBoxInv(13, 12) = "93"
    mSBoxInv(13, 13) = "c9"
    mSBoxInv(13, 14) = "9c"
    mSBoxInv(13, 15) = "ef"
    
    mSBoxInv(14, 0) = "a0"
    mSBoxInv(14, 1) = "e0"
    mSBoxInv(14, 2) = "3b"
    mSBoxInv(14, 3) = "4d"
    mSBoxInv(14, 4) = "ae"
    mSBoxInv(14, 5) = "2a"
    mSBoxInv(14, 6) = "f5"
    mSBoxInv(14, 7) = "b0"
    mSBoxInv(14, 8) = "c8"
    mSBoxInv(14, 9) = "eb"
    mSBoxInv(14, 10) = "bb"
    mSBoxInv(14, 11) = "3c"
    mSBoxInv(14, 12) = "83"
    mSBoxInv(14, 13) = "53"
    mSBoxInv(14, 14) = "99"
    mSBoxInv(14, 15) = "61"
    
    mSBoxInv(15, 0) = "17"
    mSBoxInv(15, 1) = "2b"
    mSBoxInv(15, 2) = "04"
    mSBoxInv(15, 3) = "7e"
    mSBoxInv(15, 4) = "ba"
    mSBoxInv(15, 5) = "77"
    mSBoxInv(15, 6) = "d6"
    mSBoxInv(15, 7) = "26"
    mSBoxInv(15, 8) = "e1"
    mSBoxInv(15, 9) = "69"
    mSBoxInv(15, 10) = "14"
    mSBoxInv(15, 11) = "63"
    mSBoxInv(15, 12) = "55"
    mSBoxInv(15, 13) = "21"
    mSBoxInv(15, 14) = "0c"
    mSBoxInv(15, 15) = "7d"






        
End Sub


Function SubByteByte(strByte As String) As String
'**************************************************
'entradas: byte em hexa
'saidas  : byte convertido com subbytes
'pre req.: iniciarmsbox
'criacao : 25-05-16
'autor   : Douglas
'**************************************************
Dim s As String
Dim s1 As String
Dim s2 As String

s = Right("00" & strByte, 2)

s1 = UCase(Left(s, 1))
If s1 = "A" Then s1 = "10"
If s1 = "B" Then s1 = "11"
If s1 = "C" Then s1 = "12"
If s1 = "D" Then s1 = "13"
If s1 = "E" Then s1 = "14"
If s1 = "F" Then s1 = "15"

s2 = UCase(Right(s, 1))
If s2 = "A" Then s2 = "10"
If s2 = "B" Then s2 = "11"
If s2 = "C" Then s2 = "12"
If s2 = "D" Then s2 = "13"
If s2 = "E" Then s2 = "14"
If s2 = "F" Then s2 = "15"
SubByteByte = CStr(mSBox(Val(s1), Val(s2)))

End Function

Function InvSubByteByte(strByte As String) As String
'**************************************************
'entradas: byte em hexa
'saidas  : byte convertido com subbytes
'pre req.: iniciarmsbox
'criacao : 25-05-16
'autor   : Douglas
'**************************************************
Dim s As String
Dim s1 As String
Dim s2 As String

s = Right("00" & strByte, 2)

s1 = UCase(Left(s, 1))
If s1 = "A" Then s1 = "10"
If s1 = "B" Then s1 = "11"
If s1 = "C" Then s1 = "12"
If s1 = "D" Then s1 = "13"
If s1 = "E" Then s1 = "14"
If s1 = "F" Then s1 = "15"

s2 = UCase(Right(s, 1))
If s2 = "A" Then s2 = "10"
If s2 = "B" Then s2 = "11"
If s2 = "C" Then s2 = "12"
If s2 = "D" Then s2 = "13"
If s2 = "E" Then s2 = "14"
If s2 = "F" Then s2 = "15"
InvSubByteByte = CStr(mSBoxInv(Val(s1), Val(s2)))

End Function

Private Sub IniciarMState()
Dim i As Long
Dim j As Long
Dim s As String
ReDim mState(15)
s = Right(String(32, "0") & txtInput.Text, 32)

For i = 2 To 32 Step 2
    mState((i / 2) - 1) = Mid(s, i - 1, 2)
Next

End Sub

Private Sub InvIniciarMState()
Dim i As Long
Dim j As Long
Dim s As String
ReDim mState(15)
s = Right(String(32, "0") & txtCriptografia.Text, 32)

For i = 2 To 32 Step 2
    mState((i / 2) - 1) = Mid(s, i - 1, 2)
Next

End Sub



'Private Sub IniciarMKey()
'Dim i As Long
'Dim j As Long
'Dim s As String
'ReDim mKey(15)
's = Left(txtKey.Text & String(32, "0"), 32)
'
'For i = 2 To 32 Step 2
'    mKey((i / 2) - 1) = Mid(s, i - 1, 2)
'Next
'
'End Sub

Private Sub AddRound(lngRound As Long)
Dim i As Long
Dim j As Long
'GoTo a:
'For i = 0 To 15
'    If lngRound = 0 Then
'        Debug.Print (i)
'        mState(i) = Right("00" & Conversion.Hex(HexToDec(CStr(mState(i))) Xor HexToDec(CStr(mKey(i)))), 2)
'    Else
'        For j = 1 To 8 Step 2
'            mState(i) = Right("00" & Conversion.Hex(HexToDec(CStr(mState(i))) Xor HexToDec(Mid(mWord(lngRound), j, 2))), 2)
'        Next
'    End If
'Next
'Exit Sub

'a:
i = 0
For r = lngRound * 4 To lngRound * 4 + 3
     For j = 1 To 8 Step 2
        mState(i) = Right("00" & Conversion.Hex(HexToDec(CStr(mState(i))) Xor HexToDec(Mid(mWord(r), j, 2))), 2)
        i = i + 1
     Next
Next


End Sub
Private Sub KeyExpansion()
Dim strTemp As String
Dim i       As Long
Dim j       As Long
Dim strB    As String

ReDim mWord((mNb * (mNr + 1)) - 1)

mstrKey = Right(String(32, "0") & txtKey.Text, 32)

mWord(0) = Mid(mstrKey, 1, 8)
mWord(1) = Mid(mstrKey, 9, 8)
mWord(2) = Mid(mstrKey, 17, 8)
mWord(3) = Mid(mstrKey, 25, 8)

ReDim mRCon(mNr + 1)
mRCon(1) = "01000000"
mRCon(2) = "02000000"
mRCon(3) = "04000000"
mRCon(4) = "08000000"
mRCon(5) = "10000000"
mRCon(6) = "20000000"
mRCon(7) = "40000000"
mRCon(8) = "80000000"
mRCon(9) = "1b000000"
mRCon(10) = "36000000"


For i = mNk To (mNb * (mNr + 1)) - 1
    strTemp = mWord(i - 1)
    If i Mod mNk = 0 Then
        strTemp = UnsignedXor(HexToDec(subWord(rotWord(strTemp))), HexToDec(mRCon(i / mNk)))
        strTemp = DecToHex(CDbl(strTemp))
        
    ElseIf (mNk > 6 And i Mod mNk = 4) Then
        strTemp = subWord(strTemp)
    End If
    mWord(i) = DecToHex(UnsignedXor(HexToDec(CStr(mWord(i - mNk))), HexToDec(strTemp)))
    mWord(i) = Right(String(8, "0") & mWord(i), 8)
Next
End Sub

Private Sub InvKeyExpansion()
Dim strTemp     As String
Dim i           As Long
Dim j           As Long
Dim k           As Long
Dim s           As String
Dim strB        As String



'ReDim mWord((mNb * (mNr + 1)) - 1)
ReDim mWordInv((mNb * (mNr + 1)) - 1)

'mstrKey = Right(String(32, "0") & txtKey.Text, 32)
'
'mWordInv(0) = Mid(mstrKey, 1, 8)
'mWordInv(1) = Mid(mstrKey, 9, 8)
'mWordInv(2) = Mid(mstrKey, 17, 8)
'mWordInv(3) = Mid(mstrKey, 25, 8)
'
'ReDim mRCon(mNr + 1)
'mRCon(1) = "01000000"
'mRCon(2) = "02000000"
'mRCon(3) = "04000000"
'mRCon(4) = "08000000"
'mRCon(5) = "10000000"
'mRCon(6) = "20000000"
'mRCon(7) = "40000000"
'mRCon(8) = "80000000"
'mRCon(9) = "1b000000"
'mRCon(10) = "36000000"
'
'k = (mNb * (mNr + 1)) - 1
'For i = mNk To (mNb * (mNr + 1)) - 1
'    strTemp = mWordInv(i - 1)
'    If i Mod mNk = 0 Then
'        strTemp = UnsignedXor(HexToDec(subWord(rotWord(strTemp))), HexToDec(mRCon(i / mNk)))
'        strTemp = DecToHex(CDbl(strTemp))
'
'    ElseIf (mNk > 6 And i Mod mNk = 4) Then
'        strTemp = subWord(strTemp)
'    End If
'    mWordInv(i) = DecToHex(UnsignedXor(HexToDec(CStr(mWordInv(i - mNk))), HexToDec(strTemp)))
'    mWordInv(i) = Right(String(8, "0") & mWordInv(i), 8)
'Next

'Inverter
k = UBound(mWord) - 3
For i = 0 To UBound(mWord) - 3
    mWordInv(i) = mWord(k)
    k = k - 1
Next

End Sub


Private Sub Form_Activate()
Call Class_Initialize
Call gentables
End Sub

Private Sub Form_Load()
'->128 bits
'Nro de rounds
mNr = 10 '-> Para 128 bits

'Nro de palavras de 32 bits
mNk = 4

'Nro de colunas do bloco
mNb = 4

'-> 192 bits
'mNR = 12 '-> Para 192 bits
'mNk = 6
'mNb = 4

'-> 256 bits
'mNR = 14 '-> Para 256 bits
'mNk = 8
'mNb = 4

m_byt2Power(0) = 1          ' 00000001
m_byt2Power(1) = 2          ' 00000010
m_byt2Power(2) = 4          ' 00000100
m_byt2Power(3) = 8          ' 00001000
m_byt2Power(4) = 16         ' 00010000
m_byt2Power(5) = 32         ' 00100000
m_byt2Power(6) = 64         ' 01000000
m_byt2Power(7) = 128        ' 10000000

m_bytOnBits(0) = 1          ' 00000001
m_bytOnBits(1) = 3          ' 00000011
m_bytOnBits(2) = 7          ' 00000111
m_bytOnBits(3) = 15         ' 00001111
m_bytOnBits(4) = 31         ' 00011111
m_bytOnBits(5) = 63         ' 00111111
m_bytOnBits(6) = 127        ' 01111111
m_bytOnBits(7) = 255        ' 11111111
End Sub

Private Function rotWord(strW As String) As String
rotWord = Mid(strW, 3, 6) & Left(strW, 2)
End Function

Private Function subWord(strW As String)
Dim i As Long
Dim s As String
s = ""
For i = 1 To 8 Step 2
    s = s & SubByteByte(Mid(strW, i, 2))
Next
subWord = s
End Function

Public Function UnsignedXor(uint0 As Double, uint1 As Double) As Double
UnsignedXor = LongToUDbl(UDblToLong(uint0) Xor UDblToLong(uint1))
End Function

Public Function UDblToLong(Udbl As Double) As Long
'Returns vb Long value compatible with 'C' 32-bit unsigned
'DWORD, ULONG types for vb Double in range 0 to 4,294,967,295.
Select Case Udbl
Case 0# To 2147483647#
UDblToLong = CLng(Udbl) 'Ok - return input

Case 2147483648# To 4294967295# 'convert to Long in range
'-2,147,483,648 to -1
UDblToLong = CLng(Udbl - U32_CVT)

Case Else 'won't work - force error if input <0 or >4294967295
Err.Raise 6, , "UDblToLong overflow" & vbCrLf & "Value is out of range for DWORD, ULONG type"
End Select
End Function

Public Function LongToUDbl(Lnum As Long) As Double
'Returns positive vb Double value for 'C' DWORD, ULONG
'unsigned 32-bit types stored in vb Long.
If Lnum < 0& Then 'convert to positive Double in range
'2,147,483,648 to 4,294,967,295.
LongToUDbl = U32_CVT + CDbl(Lnum)
Else
LongToUDbl = CDbl(Lnum) 'Ok - return input
End If
End Function

Private Function DecToHex(dblNumero As Double) As String
    '*****************************************************************
    ' Propósito : Converter o Nº. de Decimal para a base HexaDecimal
    ' Entrada   : O Nº. a ser convertido
    ' Retornou  : O Nº. convertido
    ' Autor     : Douglas Ferreira
    ' Criação   : 03-08-2002
    '*****************************************************************
    
    Dim dblResult        As Double
    Dim arrRestos        As Variant
    Dim arrInt           As Variant
    Dim k                As Long
    Dim i                As Long
    Dim strHexa          As String
    Dim dblRInt          As Double
    ReDim arrRestos(0 To 200)
    
    On Error GoTo Erro
    
    
    If dblNumero < 10 Then
        DecToHex = Trim(Str(dblNumero))
        GoSub fim
    ElseIf dblNumero < 16 Then
        DecToHex = letraHexa(dblNumero)
        GoSub fim
    End If
    
    dblResult = dblNumero
    Do While dblResult >= 16
        
        arrInt = Split(Format(dblResult, "0.00"), ",")
        dblRInt = arrInt(0)
        DBLrdivint = Split(Format(dblRInt / 16, "0.00"), ",")(0)
        arrRestos(k) = dblRInt - (16 * DBLrdivint)
        '   arrRestos(k) = dblResult Mod 16
        dblResult = Int(dblResult / 16)
        
        If arrRestos(k) > 9 Then
            arrRestos(k) = letraHexa(CDbl(arrRestos(k)))
        End If
        
        k = k + 1
    Loop
    
    arrRestos(k) = dblResult
    
    If arrRestos(k) > 9 Then
        arrRestos(k) = letraHexa(CDbl(arrRestos(k)))
    End If
    
    For i = k To 0 Step -1
       strHexa = strHexa & arrRestos(i)
    Next
    DecToHex = Trim(strHexa)
    
fim:
    Exit Function
    
Erro:
    MsgBox "Ocorreu o seguinte erro ao converter de Decimal para HexaDecimal: " & Err.Number & " - " & Err.Description
    Resume fim
    Resume
End Function

Private Function letraHexa(dblNumero As Double) As String
    Select Case dblNumero
        Case 10
            letraHexa = "A"
        Case 11
            letraHexa = "B"
        Case 12
            letraHexa = "C"
        Case 13
            letraHexa = "D"
        Case 14
            letraHexa = "E"
        Case 15
            letraHexa = "F"
    End Select
End Function

Private Sub ShiftRows()
    Dim a As String

    a = mState(1)
    mState(1) = mState(5)
    mState(5) = mState(9)
    mState(9) = mState(13)
    mState(13) = a
    
    a = mState(2)
    mState(2) = mState(10)
    mState(10) = a
    
    a = mState(3)
    mState(3) = mState(15)
    mState(15) = mState(11)
    mState(11) = mState(7)
    mState(7) = a
    
    a = mState(6)
    mState(6) = mState(14)
    mState(14) = a

End Sub

Private Sub InvShiftRows()
    Dim a As String

    a = mState(1)
    mState(1) = mState(13)
    mState(13) = mState(9)
    mState(9) = mState(5)
    mState(5) = a
    
    a = mState(2)
    mState(2) = mState(10)
    mState(10) = a
    
    a = mState(3)
    mState(3) = mState(7)
    mState(7) = mState(11)
    mState(11) = mState(15)
    mState(15) = a
    
    a = mState(6)
    mState(6) = mState(14)
    mState(14) = a

End Sub

Private Function DecToBin(dblNumero As Double) As String
    '*****************************************************************
    ' Propósito : Converter o Nº. de Decimal para a base Binária
    ' Entrada   : O Nº. a ser convertido
    ' Retornou  : O Nº. convertido
    ' Autor     : Douglas Ferreira
    ' Criação   : 03-08-2002
    '*****************************************************************
    
    Dim dblResult        As Double
    Dim arrRestos        As Variant
    Dim k                As Long
    Dim i                As Long
    Dim strBin           As String
    ReDim arrRestos(0 To 200)
    
    On Error GoTo Erro
    
    dblResult = dblNumero
    Do While dblResult >= 1
        
        arrRestos(k) = dblResult Mod 2
        dblResult = Int(dblResult / 2)
        
        k = k + 1
    Loop
    
    If dblResult > 0 Then arrRestos(k) = dblResult
    
    For i = k To 0 Step -1
        strBin = strBin & arrRestos(i)
    Next
    DecToBin = Trim(Str(strBin))
    
fim:
    Exit Function
    
Erro:
    MsgBox "Ocorreu o seguinte erro ao converter de decimal para Binário: " & Err.Number & " - " & Err.Description
    Resume fim
    Resume
End Function

Private Function BinToDec(strNumero As String) As String
    '*****************************************************************
    ' Propósito : Converter o Nº. de Binário para a base Decimal
    ' Entrada   : O Nº. a ser convertido
    ' Retornou  : O Nº. convertido
    ' Autor     : Douglas Ferreira
    ' Criação   : 03-08-2002
    '*****************************************************************
    Dim lngExpoente As Long
    Dim k           As Long
    Dim dblResult   As Double
    Dim lngCadaNum  As Long
    
    strNumero = Trim(strNumero)
    lngExpoente = Len(strNumero) - 1
    
    For k = 1 To Len(strNumero)
        lngCadaNum = letraHexaToNum(Mid(strNumero, k, 1))
        dblResult = dblResult + CDbl(lngCadaNum * 2 ^ lngExpoente)
        lngExpoente = lngExpoente - 1
    Next
    
    BinToDec = Trim(Str(dblResult))

End Function


Private Sub MixColumns()
Dim i   As Long
Dim a   As Double
Dim b   As Double
Dim c   As Double
Dim d   As Double
Dim s   As Double
Dim f   As Double
Dim l   As Long
Dim aa  As String
Dim bb  As String
Dim cc  As String
Dim dd  As String

f = HexToDec("1B")

'coluna 1
aa = CStr(mState(0))
bb = CStr(mState(1))
cc = CStr(mState(2))
dd = CStr(mState(3))

a = Misturar(f, aa, 2)
b = Misturar(f, bb, 3)
c = Misturar(f, cc, 1)
d = Misturar(f, dd, 1)


s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(0) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 1)
b = Misturar(f, bb, 2)
c = Misturar(f, cc, 3)
d = Misturar(f, dd, 1)


s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(1) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 1)
b = Misturar(f, bb, 1)
c = Misturar(f, cc, 2)
d = Misturar(f, dd, 3)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(2) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 3)
b = Misturar(f, bb, 1)
c = Misturar(f, cc, 1)
d = Misturar(f, dd, 2)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(3) = Right("0" & DecToHex(s), 2)

'coluna 2
aa = CStr(mState(4))
bb = CStr(mState(5))
cc = CStr(mState(6))
dd = CStr(mState(7))

a = Misturar(f, aa, 2)
b = Misturar(f, bb, 3)
c = Misturar(f, cc, 1)
d = Misturar(f, dd, 1)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(4) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 1)
b = Misturar(f, bb, 2)
c = Misturar(f, cc, 3)
d = Misturar(f, dd, 1)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(5) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 1)
b = Misturar(f, bb, 1)
c = Misturar(f, cc, 2)
d = Misturar(f, dd, 3)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(6) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 3)
b = Misturar(f, bb, 1)
c = Misturar(f, cc, 1)
d = Misturar(f, dd, 2)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(7) = Right("0" & DecToHex(s), 2)

'coluna 3
aa = CStr(mState(8))
bb = CStr(mState(9))
cc = CStr(mState(10))
dd = CStr(mState(11))

a = Misturar(f, aa, 2)
b = Misturar(f, bb, 3)
c = Misturar(f, cc, 1)
d = Misturar(f, dd, 1)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(8) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 1)
b = Misturar(f, bb, 2)
c = Misturar(f, cc, 3)
d = Misturar(f, dd, 1)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(9) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 1)
b = Misturar(f, bb, 1)
c = Misturar(f, cc, 2)
d = Misturar(f, dd, 3)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(10) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 3)
b = Misturar(f, bb, 1)
c = Misturar(f, cc, 1)
d = Misturar(f, dd, 2)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(11) = Right("0" & DecToHex(s), 2)

'coluna 4
aa = CStr(mState(12))
bb = CStr(mState(13))
cc = CStr(mState(14))
dd = CStr(mState(15))

a = Misturar(f, aa, 2)
b = Misturar(f, bb, 3)
c = Misturar(f, cc, 1)
d = Misturar(f, dd, 1)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(12) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 1)
b = Misturar(f, bb, 2)
c = Misturar(f, cc, 3)
d = Misturar(f, dd, 1)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(13) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 1)
b = Misturar(f, bb, 1)
c = Misturar(f, cc, 2)
d = Misturar(f, dd, 3)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(14) = Right("0" & DecToHex(s), 2)

a = Misturar(f, aa, 3)
b = Misturar(f, bb, 1)
c = Misturar(f, cc, 1)
d = Misturar(f, dd, 2)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(15) = Right("0" & DecToHex(s), 2)

End Sub


Private Sub InvMixColumns()
Dim i   As Long
Dim a   As Double
Dim b   As Double
Dim c   As Double
Dim d   As Double
Dim s   As Double
Dim f   As Double
Dim l   As Long
Dim aa  As String
Dim bb  As String
Dim cc  As String
Dim dd  As String

f = HexToDec("1B")

'coluna 1
aa = CStr(mState(0))
bb = CStr(mState(1))
cc = CStr(mState(2))
dd = CStr(mState(3))

a = bmul(HexToDec(aa), (14))
b = bmul(HexToDec(bb), (11))
c = bmul(HexToDec(cc), (13))
d = bmul(HexToDec(dd), (9))

'a = Misturar(f, aa, 14)
'b = Misturar(f, bb, 11)
'c = Misturar(f, cc, 13)
'd = Misturar(f, dd, 9)

s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(0) = Right("0" & DecToHex(s), 2)
'ini test
'MsgBox xtime(255)
'MsgBox product(57, 13)
'MsgBox bmul(87, 19)
'fin

a = bmul(HexToDec(aa), (9))
b = bmul(HexToDec(bb), (14))
c = bmul(HexToDec(cc), (11))
d = bmul(HexToDec(dd), (13))

'a = Misturar(f, aa, 9)
'b = Misturar(f, bb, 14)
'c = Misturar(f, cc, 11)
'd = Misturar(f, dd, 13)

s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(1) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (13))
b = bmul(HexToDec(bb), (9))
c = bmul(HexToDec(cc), (14))
d = bmul(HexToDec(dd), (11))

'a = Misturar(f, aa, 13)
'b = Misturar(f, bb, 9)
'c = Misturar(f, cc, 14)
'd = Misturar(f, dd, 11)

s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(2) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (11))
b = bmul(HexToDec(bb), (13))
c = bmul(HexToDec(cc), (9))
d = bmul(HexToDec(dd), (14))

'a = Misturar(f, aa, 11)
'b = Misturar(f, bb, 13)
'c = Misturar(f, cc, 9)
'd = Misturar(f, dd, 14)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(3) = Right("0" & DecToHex(s), 2)

'coluna 2
aa = CStr(mState(4))
bb = CStr(mState(5))
cc = CStr(mState(6))
dd = CStr(mState(7))

a = bmul(HexToDec(aa), (14))
b = bmul(HexToDec(bb), (11))
c = bmul(HexToDec(cc), (13))
d = bmul(HexToDec(dd), (9))

'a = Misturar(f, aa, 14)
'b = Misturar(f, bb, 11)
'c = Misturar(f, cc, 13)
'd = Misturar(f, dd, 9)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(4) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (9))
b = bmul(HexToDec(bb), (14))
c = bmul(HexToDec(cc), (11))
d = bmul(HexToDec(dd), (13))

'a = Misturar(f, aa, 9)
'b = Misturar(f, bb, 14)
'c = Misturar(f, cc, 11)
'd = Misturar(f, dd, 13)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(5) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (13))
b = bmul(HexToDec(bb), (9))
c = bmul(HexToDec(cc), (14))
d = bmul(HexToDec(dd), (11))

'a = Misturar(f, aa, 13)
'b = Misturar(f, bb, 9)
'c = Misturar(f, cc, 14)
'd = Misturar(f, dd, 11)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(6) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (11))
b = bmul(HexToDec(bb), (13))
c = bmul(HexToDec(cc), (9))
d = bmul(HexToDec(dd), (14))

'a = Misturar(f, aa, 11)
'b = Misturar(f, bb, 13)
'c = Misturar(f, cc, 9)
'd = Misturar(f, dd, 14)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(7) = Right("0" & DecToHex(s), 2)

'coluna 3
aa = CStr(mState(8))
bb = CStr(mState(9))
cc = CStr(mState(10))
dd = CStr(mState(11))

a = bmul(HexToDec(aa), (14))
b = bmul(HexToDec(bb), (11))
c = bmul(HexToDec(cc), (13))
d = bmul(HexToDec(dd), (9))

'a = Misturar(f, aa, 14)
'b = Misturar(f, bb, 11)
'c = Misturar(f, cc, 13)
'd = Misturar(f, dd, 9)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(8) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (9))
b = bmul(HexToDec(bb), (14))
c = bmul(HexToDec(cc), (11))
d = bmul(HexToDec(dd), (13))

'a = Misturar(f, aa, 9)
'b = Misturar(f, bb, 14)
'c = Misturar(f, cc, 11)
'd = Misturar(f, dd, 13)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(9) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (13))
b = bmul(HexToDec(bb), (9))
c = bmul(HexToDec(cc), (14))
d = bmul(HexToDec(dd), (11))

'a = Misturar(f, aa, 13)
'b = Misturar(f, bb, 9)
'c = Misturar(f, cc, 14)
'd = Misturar(f, dd, 11)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(10) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (11))
b = bmul(HexToDec(bb), (13))
c = bmul(HexToDec(cc), (9))
d = bmul(HexToDec(dd), (14))

'a = Misturar(f, aa, 11)
'b = Misturar(f, bb, 13)
'c = Misturar(f, cc, 9)
'd = Misturar(f, dd, 14)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(11) = Right("0" & DecToHex(s), 2)

'coluna 4
aa = CStr(mState(12))
bb = CStr(mState(13))
cc = CStr(mState(14))
dd = CStr(mState(15))

a = bmul(HexToDec(aa), (14))
b = bmul(HexToDec(bb), (11))
c = bmul(HexToDec(cc), (13))
d = bmul(HexToDec(dd), (9))

'a = Misturar(f, aa, 14)
'b = Misturar(f, bb, 11)
'c = Misturar(f, cc, 13)
'd = Misturar(f, dd, 9)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(12) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (9))
b = bmul(HexToDec(bb), (14))
c = bmul(HexToDec(cc), (11))
d = bmul(HexToDec(dd), (13))

'a = Misturar(f, aa, 9)
'b = Misturar(f, bb, 14)
'c = Misturar(f, cc, 11)
'd = Misturar(f, dd, 13)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(13) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (13))
b = bmul(HexToDec(bb), (9))
c = bmul(HexToDec(cc), (14))
d = bmul(HexToDec(dd), (11))

'a = Misturar(f, aa, 13)
'b = Misturar(f, bb, 9)
'c = Misturar(f, cc, 14)
'd = Misturar(f, dd, 11)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(14) = Right("0" & DecToHex(s), 2)

a = bmul(HexToDec(aa), (11))
b = bmul(HexToDec(bb), (13))
c = bmul(HexToDec(cc), (9))
d = bmul(HexToDec(dd), (14))

'a = Misturar(f, aa, 11)
'b = Misturar(f, bb, 13)
'c = Misturar(f, cc, 9)
'd = Misturar(f, dd, 14)
s = UnsignedXor(a, b)
s = UnsignedXor(s, c)
s = UnsignedXor(s, d)
mState(15) = Right("0" & DecToHex(s), 2)

End Sub

Private Function Misturar(f As Double, strNum As String, lngFator As Long) As Double
Dim a As Double
Dim s As String
Dim b As Boolean

If strNum = "00" Then
    Misturar = 0
    Exit Function
End If

b = False
a = HexToDec(strNum)
a = DecToBin(a)
s = Right(String(8, "0") & a, 8)
If Left(s, 1) = "1" Then
    b = True
End If
a = BinToDec(CStr(a))

If lngFator = 1 Then
    a = HexToDec(strNum)
ElseIf lngFator = 2 Then
    a = HexToDec(strNum) * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
ElseIf lngFator = 3 Then
    a = HexToDec(strNum) * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    a = UnsignedXor(a, HexToDec(strNum))
    If b Then a = UnsignedXor(a, f)

ElseIf lngFator = 9 Then
    a = HexToDec(strNum) * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    a = UnsignedXor(HexToDec(strNum), a)
    
ElseIf lngFator = 11 Then
    a = HexToDec(strNum) * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    a = UnsignedXor(HexToDec(strNum), a)

    
ElseIf lngFator = 13 Then
    a = HexToDec(strNum) * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    
    
    a = UnsignedXor(HexToDec(strNum), a)
    
       
ElseIf lngFator = 14 Then
    a = HexToDec(strNum) * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
    '--
    a = a * 2
    a = DecToBin(a)
    a = Right(a, 8)
    a = BinToDec(CStr(a))
    If b Then
        a = UnsignedXor(a, f)
    End If
    
    a = DecToBin(a)
    s = Right(String(8, "0") & a, 8)
    If Left(s, 1) = "1" Then
        b = True
    Else
        b = False
    End If
    a = BinToDec(CStr(a))
    '--
    
   ' a = UnsignedXor(HexToDec(strNum), a)
    
End If

Misturar = a

End Function

Private Function xtime(ByVal a As Byte) As Byte
    Dim b As Byte
    
    If (a And &H80) Then
        b = &H1B
    Else
        b = 0
    End If
    
    a = LShiftByte(a, 1)
    a = a Xor b
    
    xtime = a
End Function

Private Function LShiftByte(ByVal bytValue As Byte, _
                            ByVal bytShiftBits As Byte) As Byte
    If bytShiftBits = 0 Then
        LShiftByte = bytValue
        Exit Function
    ElseIf bytShiftBits = 7 Then
        If bytValue And 1 Then
            LShiftByte = &H80
        Else
            LShiftByte = 0
        End If
        Exit Function
    ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
        Err.Raise 6
    End If
    
    LShiftByte = ((bytValue And m_bytOnBits(7 - bytShiftBits)) * _
        m_byt2Power(bytShiftBits))
End Function

Private Function bmul(ByVal x As Byte, _
                      y As Byte) As Byte
    If x <> 0 And y <> 0 Then
        bmul = m_ptab((CLng(m_ltab(x)) + CLng(m_ltab(y))) Mod 255)
    Else
        bmul = 0
    End If
End Function

Public Sub gentables()
    Dim i       As Long
    Dim y       As Byte
    Dim b(3)    As Byte
    Dim ib      As Byte
    
    m_ltab(0) = 0
    m_ptab(0) = 1
    m_ltab(1) = 0
    m_ptab(1) = 3
    m_ltab(3) = 1
    
    For i = 2 To 255
        m_ptab(i) = m_ptab(i - 1) Xor xtime(m_ptab(i - 1))
        m_ltab(m_ptab(i)) = i
    Next
    
    m_fbsub(0) = &H63
    m_rbsub(&H63) = 0
    
    For i = 1 To 255
        ib = i
        y = ByteSub(ib)
        m_fbsub(i) = y
        m_rbsub(y) = i
    Next
    
        y = 1
    For i = 0 To 29
        m_rco(i) = y
        y = xtime(y)
    Next
    
    For i = 0 To 255
        y = m_fbsub(i)
        b(3) = y Xor xtime(y)
        b(2) = y
        b(1) = y
        b(0) = xtime(y)
        m_ftable(i) = Pack(b)
        
        y = m_rbsub(i)
        b(3) = bmul(m_InCo(0), y)
        b(2) = bmul(m_InCo(1), y)
        b(1) = bmul(m_InCo(2), y)
        b(0) = bmul(m_InCo(3), y)
        m_rtable(i) = Pack(b)
    Next
End Sub

Private Function ByteSub(ByVal x As Byte) As Byte
    Dim y As Byte
    
    y = m_ptab(255 - m_ltab(x))
    x = y
    x = RotateLeftByte(x, 1)
    y = y Xor x
    x = RotateLeftByte(x, 1)
    y = y Xor x
    x = RotateLeftByte(x, 1)
    y = y Xor x
    x = RotateLeftByte(x, 1)
    y = y Xor x
    y = y Xor &H63
    
    ByteSub = y
End Function
Private Function Pack(b() As Byte) As Long
    Dim lCount As Long
    Dim lTemp  As Long
    
    For lCount = 0 To 3
        lTemp = b(lCount)
        Pack = Pack Or LShift(lTemp, (lCount * 8))
    Next
End Function
Private Function RotateLeft(ByVal lValue As Long, _
                            ByVal iShiftBits As Integer) As Long
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

''*******************************************************************************
'' RotateLeftByte (FUNCTION)
'*******************************************************************************
Private Function RotateLeftByte(ByVal bytValue As Byte, _
                                ByVal bytShiftBits As Byte) As Byte
    RotateLeftByte = LShiftByte(bytValue, bytShiftBits) Or _
        RShiftByte(bytValue, (8 - bytShiftBits))
End Function

Private Function RShiftByte(ByVal bytValue As Byte, _
                            ByVal bytShiftBits As Byte) As Byte
    If bytShiftBits = 0 Then
        RShiftByte = bytValue
        Exit Function
    ElseIf bytShiftBits = 7 Then
        If bytValue And &H80 Then
            RShiftByte = 1
        Else
            RShiftByte = 0
        End If
        Exit Function
    ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
        Err.Raise 6
    End If
    
    RShiftByte = bytValue \ m_byt2Power(bytShiftBits)
End Function


Private Function LShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * _
            m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * _
            m_l2Power(iShiftBits))
    End If
End Function

Private Function product(ByVal x As Long, _
                         ByVal y As Long) As Long
    Dim xb(3) As Byte
    Dim yb(3) As Byte
    
    Unpack x, xb
    Unpack y, yb
    product = bmul(xb(0), yb(0)) Xor bmul(xb(1), yb(1)) Xor bmul(xb(2), yb(2)) _
        Xor bmul(xb(3), yb(3))
End Function
Private Sub Unpack(ByVal a As Long, _
                   b() As Byte)
    b(0) = a And m_lOnBits(7)
    b(1) = RShift(a, 8) And m_lOnBits(7)
    b(2) = RShift(a, 16) And m_lOnBits(7)
    b(3) = RShift(a, 24) And m_lOnBits(7)
End Sub

Private Function RShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    
    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Sub Class_Initialize()
    m_InCo(0) = &HB
    m_InCo(1) = &HD
    m_InCo(2) = &H9
    m_InCo(3) = &HE
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - BITS SET FROM RIGHT
    m_bytOnBits(0) = 1          ' 00000001
    m_bytOnBits(1) = 3          ' 00000011
    m_bytOnBits(2) = 7          ' 00000111
    m_bytOnBits(3) = 15         ' 00001111
    m_bytOnBits(4) = 31         ' 00011111
    m_bytOnBits(5) = 63         ' 00111111
    m_bytOnBits(6) = 127        ' 01111111
    m_bytOnBits(7) = 255        ' 11111111
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - POWERS OF 2
    m_byt2Power(0) = 1          ' 00000001
    m_byt2Power(1) = 2          ' 00000010
    m_byt2Power(2) = 4          ' 00000100
    m_byt2Power(3) = 8          ' 00001000
    m_byt2Power(4) = 16         ' 00010000
    m_byt2Power(5) = 32         ' 00100000
    m_byt2Power(6) = 64         ' 01000000
    m_byt2Power(7) = 128        ' 10000000
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - BITS SET FROM RIGHT
    m_lOnBits(0) = 1            ' 00000000000000000000000000000001
    m_lOnBits(1) = 3            ' 00000000000000000000000000000011
    m_lOnBits(2) = 7            ' 00000000000000000000000000000111
    m_lOnBits(3) = 15           ' 00000000000000000000000000001111
    m_lOnBits(4) = 31           ' 00000000000000000000000000011111
    m_lOnBits(5) = 63           ' 00000000000000000000000000111111
    m_lOnBits(6) = 127          ' 00000000000000000000000001111111
    m_lOnBits(7) = 255          ' 00000000000000000000000011111111
    m_lOnBits(8) = 511          ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - POWERS OF 2
    m_l2Power(0) = 1            ' 00000000000000000000000000000001
    m_l2Power(1) = 2            ' 00000000000000000000000000000010
    m_l2Power(2) = 4            ' 00000000000000000000000000000100
    m_l2Power(3) = 8            ' 00000000000000000000000000001000
    m_l2Power(4) = 16           ' 00000000000000000000000000010000
    m_l2Power(5) = 32           ' 00000000000000000000000000100000
    m_l2Power(6) = 64           ' 00000000000000000000000001000000
    m_l2Power(7) = 128          ' 00000000000000000000000010000000
    m_l2Power(8) = 256          ' 00000000000000000000000100000000
    m_l2Power(9) = 512          ' 00000000000000000000001000000000
    m_l2Power(10) = 1024        ' 00000000000000000000010000000000
    m_l2Power(11) = 2048        ' 00000000000000000000100000000000
    m_l2Power(12) = 4096        ' 00000000000000000001000000000000
    m_l2Power(13) = 8192        ' 00000000000000000010000000000000
    m_l2Power(14) = 16384       ' 00000000000000000100000000000000
    m_l2Power(15) = 32768       ' 00000000000000001000000000000000
    m_l2Power(16) = 65536       ' 00000000000000010000000000000000
    m_l2Power(17) = 131072      ' 00000000000000100000000000000000
    m_l2Power(18) = 262144      ' 00000000000001000000000000000000
    m_l2Power(19) = 524288      ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
End Sub
Private Function InvMixCol(ByVal x As Long) As Long
    Dim y       As Long
    Dim m       As Long
    Dim b(3)    As Byte
    
    m = Pack(m_InCo)
    b(3) = product(m, x)
    m = RotateLeft(m, 24)
    b(2) = product(m, x)
    m = RotateLeft(m, 24)
    b(1) = product(m, x)
    m = RotateLeft(m, 24)
    b(0) = product(m, x)
    y = Pack(b)
    
    InvMixCol = y
End Function




Private Sub txtTexto_Change()

txtInput.Text = ConverterTextoCharToHex(txtTexto.Text)
txtDescriptografia.Text = ""
txtDescripFinal.Text = ""
txtCriptografia.Text = ""
End Sub

Private Function ConverterTextoHexToChar(h As String) As String
Dim i       As Long
Dim a       As Double
Dim s       As String

n = 0
s = ""
For i = 1 To Len(h) Step 2
    If Mid(h, i, 2) <> "00" Then
        a = HexToDec(Mid(h, i, 2))
        s = s & Chr(a)
    End If
Next
ConverterTextoHexToChar = s
End Function
Private Function ConverterTextoCharToHex(c As String) As String
Dim i       As Long
Dim a       As Double
Dim h       As String
Dim s       As String

n = 0
s = ""
For i = 1 To Len(c)
   a = Asc(Mid(c, i, 1))
   h = DecToHex(a)
   s = s & Right("00" & h, 2)
Next

s = Right(String(32, "0") & s, 32)

ConverterTextoCharToHex = s

End Function
