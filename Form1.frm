VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   4185
   ClientTop       =   2925
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6195
   Begin VB.CommandButton Command3 
      Caption         =   "Check"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   735
      Left            =   2880
      TabIndex        =   4
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Text            =   "160700004"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Text            =   "SAL"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Text            =   "2294"
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SALSN As ADODB.Connection
Dim Pos As ADODB.Connection

Dim ERRMSGID$, BUFSQLID$

Const ServerName = "KDBA"
Const CurPath = "C:\SEDILIB"

Private Sub Command1_Click()
    SUBENDSR
End Sub

Private Sub Command2_Click()
    Dim FileName$, fh%
    Dim rd As ADODB.Recordset
    Dim rdS As ADODB.Recordset
    Dim COMPANY, S5_TYPE, GlbSAO03
    Const PrintDetails = False
    
    COMPANY = Trim$(Text1)
    S5_TYPE = Trim$(Text2)
    GlbSAO03 = Trim$(Text3)
    
    BUFSQLID = "Select * From EISAO Where SAO01='" & COMPANY & "' "
    BUFSQLID = BUFSQLID & "And SAO02='" & S5_TYPE & "' "
    BUFSQLID = BUFSQLID & "And SAO03=" & GlbSAO03
    ERRMSGID = ADO_OpenRecSet(SALSN, rd, "", BUFSQLID)
    
    FileName = CurPath & "\InvPrt\INPUT\ElecInvoice.txt"
    
    fh = FreeFile
    Open FileName For Append As #fh
    If (Mid(rd("SAO06"), 5, 2) Mod 2) = 1 Then
       Print #fh, "1 " & Mid$(rd("SAO06"), 2, 3) & "年" & Mid$(rd("SAO06"), 5, 2) & "-" & Format$(Mid(rd("SAO06"), 5, 2) + 1, "00")
    Else
       Print #fh, "1 " & Mid$(rd("SAO06"), 2, 3) & "年" & Format$(Mid(rd("SAO06"), 5, 2) - 1, "00") & "-" & Mid$(rd("SAO06"), 5, 2)
    End If

    Print #fh, "2 " & Mid$(rd("SAO13"), 1, 2) & "-" & Mid$(rd("SAO13"), 3)
    Print #fh, "3 " & Mid(rd("SAO06"), 2, 3) + 1911 & "-" & Format$(Mid(rd("SAO06"), 5, 2), "00") & "-" & Format$(Mid(rd("SAO06"), 7, 2), "00")
    Print #fh, "4 " & Mid$(rd("SAO83"), 1, 2) & ":" & Mid$(rd("SAO83"), 3, 2) & ":" & Mid$(rd("SAO83"), 5, 2)
    Print #fh, "5 "                                  '發票格式
    Print #fh, "6 " & rd("SAO49")                    '隨機碼
    Print #fh, "7 " & rd("SAO35")
    Print #fh, "8 " & "70384140"                     '賣方統編
    Print #fh, "9 " & rd("SAO16")
    
    If (Mid(rd("SAO06"), 5, 2) Mod 2) = 1 Then
       Print #fh, "10 " & Mid$(rd("SAO06"), 2, 3) & Format$(Mid(rd("SAO06"), 5, 2) + 1, "00") & rd("SAO13") & rd("SAO49")
    Else
       Print #fh, "10 " & Mid$(rd("SAO06"), 2, 3) & Mid$(rd("SAO06"), 5, 2) & rd("SAO13") & rd("SAO49")
    End If
    
    Print #fh, "11 " & rd("SAO13") & Mid$(rd("SAO06"), 2) & rd("SAO49") & String$(8 - Len(rd("SAO31")), "0") & rd("SAO31") & _
                       String$(8 - Len(rd("SAO35")), "0") & rd("SAO35") & "70384140" & IIf(Trim$(rd("SAO16")) = "", String$(8, "0"), Trim$(rd("SAO16"))) & _
                       "012345678AES234567891234" & ":" & String$(10, "*") & ":1:1:1:" & Trim$(rd("SAO69")) & ":1"
    Print #fh, "12 **" & ":" & rd("SAO35")
    Print #fh, "13 店" & Trim$(rd("SAO01")) & "-序" & rd("SAO03")
    Print #fh, "14 0"
    Print #fh, "15 " & rd("SAO03")
    
    Print #fh, String$(28, "-")
    
    If Trim(rd("SAO16")) <> "" Then Print #fh, Space(8) & "銷貨明細表"
    
    Print #fh, Mid$(Trim(rd("SAO69")), 1, 12)
    Print #fh, Space(14) & "X1" & Space(12 - 3 - Len(rd("SAO35"))) & "$" & rd("SAO35") & "TX"
    Print #fh,
    Print #fh, "銷售額" & Space(28 - 7 - Len(rd("SAO31"))) & "$" & rd("SAO31")
    Print #fh, "稅額" & Space(28 - 5 - Len(rd("SAO33"))) & "$" & rd("SAO33")
    Print #fh, "總計金額" & Space(28 - 9 - Len(rd("SAO35"))) & "$" & rd("SAO35")
    
    If PrintDetails Then
       BUFSQLID = "Select SAS11,SAS12,SAS17 From EISAS Where SAS01='" & COMPANY & "' "
       BUFSQLID = BUFSQLID & "And SAS02='" & S5_TYPE & "' "
       BUFSQLID = BUFSQLID & "And SAS03=" & GlbSAO03 & " Order By SAS04"
       ERRMSGID = ADO_OpenRecSet(SALSN, rdS, "", BUFSQLID)

       Print #fh,

       If Trim(rd("SAO16")) = "" Then Print #fh, Space(9) & "銷貨明細"
       Do While Not rdS.EOF
           Print #fh, Mid$(Trim(rdS("SAS12")), 1, 12)
           Print #fh, Space(6) & Trim$(rdS("SAS11")) & Space(28 - 6 - Len(Trim$(rdS("SAS11"))) - 1 - Len(rdS("SAS17"))) & "X" & rdS("SAS17")

           rdS.MoveNext
       Loop
       rdS.Close
    End If
    
    Print #fh,
    Print #fh, "交易時間:" & Mid(rd("SAO82"), 2, 3) + 1911 & "-" & Format$(Mid(rd("SAO82"), 5, 2), "00") & "-" & Format$(Mid(rd("SAO82"), 7, 2), "00") & _
                           " " & Mid$(rd("SAO83"), 1, 2) & ":" & Mid$(rd("SAO83"), 3, 2) & ":" & Mid$(rd("SAO83"), 5, 2)
    Print #fh, "店" & Trim$(rd("SAO01")) & "-" & rd("SAO13")
    Print #fh, "交易序號:" & rd("SAO03")
    
    Close #fh
    
    rd.Close
End Sub

Private Sub Command3_Click()
    CHKInvoPrint
End Sub

Private Sub Form_Load()
    Dim objPrc As Object
    Dim strSQL As String
    
    strSQL = "Select * from Win32_Process Where Name = 'ElecInvoPrint.exe'"
    
    If GetObject("winmgmts:").ExecQuery(strSQL).Count > 0 Then
       MsgBox "已經有執行！"
    Else
       MsgBox "程式尚未啟動！"
    End If
    
    
    'Shell CurPath & "\InvPrt\ElecInvoPrint.exe"
    
    SUBINZSR

    ERRMSGID = ADO_OpenDB(ServerName, "POS", Pos)
    ERRMSGID = ADO_OpenDB(ServerName, "SALSN", SALSN)

End Sub

Private Sub SUBINZSR()
    Plib.ServerLogin = "sa"
    Plib.ServerPaswd = ""
End Sub

Sub SUBENDSR()
    Dim WshShell
    
    Screen.MousePointer = 11
    
    'Set WshShell = CreateObject("WScript.Shell")
    'WshShell.SendKeys "^+%{F8}"

    Pos.Close
    SALSN.Close

    End

    Screen.MousePointer = 0
End Sub

Sub CHKInvoPrint()
    Dim WshShell
    Dim sFile$, txtLine$
    
    sFile = CurPath & "\InvPrt\EVENTLOG\Log\ExecuteResult.log"
    If Dir$(sFile) <> "" Then Kill sFile
       
    
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "^+%{F10}"
    
    Sleep 2000
    
    If Dir$(sFile) = "" Then
       MsgBox "檢查電子發票列印常駐程式是否有執行！"
    Else
       File_UTF8_2_ANSI sFile, sFile
    
       Open sFile For Input As #1          '開啟文字檔
       Line Input #1, txtLine              '逐行讀取
       Line Input #1, txtLine
       Line Input #1, txtLine
       
       If Right$(Trim(txtLine), 1) = "0" Then
          Line Input #1, txtLine
          Line Input #1, txtLine
          MsgBox txtLine
       End If
       Close #1
    End If

End Sub

Private Sub File_UTF8_2_ANSI(strSrc As String, strDst As String)
    Dim strData As String
    With CreateObject("ADODB.Stream")
        .Open
        .Charset = "UTF-8"
        .LoadFromFile strSrc
        strData = .ReadText
        .Position = 0
        .Charset = "BIG5"
        .WriteText strData
        .SaveToFile strDst, 2
        .Close
    End With
End Sub
