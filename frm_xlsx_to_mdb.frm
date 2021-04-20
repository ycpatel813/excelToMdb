VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_xlsx_to_mdb 
   Caption         =   "Excel To MDB"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_convert 
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   8655
   End
   Begin VB.TextBox txt_filename 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   10455
   End
   Begin VB.CommandButton cmd_select_excel 
      Caption         =   "Select Excel File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   10455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl_msg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   10335
   End
   Begin VB.Label lbl_sheetname 
      Caption         =   "No File Selected..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   10335
   End
   Begin VB.Label lbl_file 
      Caption         =   "No File Selected..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   10335
   End
End
Attribute VB_Name = "frm_xlsx_to_mdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_convert_Click()
If lbl_file = "No File Selected..." Then
    MsgBox "Please select file which needs to convert."
    GoTo Finish
End If
   

' objects you need:
Dim srcConn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim dstConn As New ADODB.Connection
Dim rsDst As New ADODB.Recordset
Dim sheetName As String

'Example connection with Excel - HDR is discussed below
srcConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=" + lbl_file + ";" & _
        "Extended Properties=""Excel 12.0;"";"

'Enter Sheet Name here
sheetName = "Sheet1"

rs.Open "SELECT * FROM [" & sheetName & "$]", _
    srcConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    

Dim one As Long
Dim two As String
Dim file As String

file = App.Path + "\created.mdb"

FileCopy App.Path + "\sample.mdb", file

' Example connection with your destination database
dstConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & file & ";" & _
        "Persist Security Info=False"



rsDst.Open "sheet1", _
    dstConn, adOpenDynamic, adLockOptimistic, adCmdTable
    
' Import
Do Until rs.EOF
    If (Not IsNull(rs.Fields.Item(0)) And Not IsNull(rs.Fields.Item(1))) Then
            
            one = rs.Fields.Item(0)
            two = rs.Fields.Item(1)
            
            rsDst.AddNew
                rsDst.Fields("one") = one
                rsDst.Fields("two") = two
            rsDst.Update
    End If
    Index = Index + 1
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
srcConn.Close
Set srcConn = Nothing
dstConn.Close
Set dstConn = Nothing
MsgBox "File is created successfully."
Finish:
End Sub

Private Sub cmd_select_excel_Click()
CommonDialog1.Filter = "Apps (*.xlsx)|*.xlsx|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
lbl_file = CommonDialog1.FileName
End Sub

