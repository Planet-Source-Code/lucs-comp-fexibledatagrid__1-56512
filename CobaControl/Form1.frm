VERSION 5.00
Object = "*\ACtlProj.vbp"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7440
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   4800
      Width           =   1575
   End
   Begin CtlProj.LucsDatagrid LucsDatagrid1 
      Height          =   3855
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6800
      ConnectionString=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\CobaControl\DBCtl.mdb;Persist Security Info=False"
      RecordSource    =   "Absensi"
      RecordSourceCmb1=   "karyawan"
      ListField1      =   "nama"
      BoundColumn1    =   "employeeid"
      Datafield1      =   "employeeid"
      colCmb1         =   1
      FieldDT1        =   "Tanggal"
      ColumnDT1       =   0
      ColumnChk1      =   4
      ColumnChk2      =   5
      ColumnChk3      =   6
      FieldChk1       =   "UM"
      FieldChk2       =   "UMlbr"
      FieldChk3       =   "Cuti"
      DataFieldTxt1   =   "Lembur"
      ValueCmb1       =   "BARI0100104"
      ValueDT1        =   38255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort Descending"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sort Ascending"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form1.frx":0000
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "NAMA"
      BoundColumn     =   "EmployeeID"
      Text            =   ""
      Object.DataMember      =   "Karyawan"
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Release Filter"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "employeeid"
      DataSource      =   "LucsDatagrid1"
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Value Datacombo1 :"
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
If Check1.Value = 1 Then
    Me.DataCombo1.Enabled = False
    
    LucsDatagrid1.RebindDG
Else
    Me.DataCombo1.Enabled = True
    Me.DataCombo1.Text = ""
End If

End Sub

Private Sub Command1_Click()
LucsDatagrid1.SortAsc
End Sub

Private Sub Command2_Click()
LucsDatagrid1.SortDesc
End Sub

Private Sub Command3_Click()
Me.LucsDatagrid1.UpdateBatch
End Sub

Private Sub DataCombo1_Click(Area As Integer)
LucsDatagrid1.rs.Filter = "[employeeid] = '" & Me.DataCombo1.BoundText & "'"

End Sub

Private Sub Form_Load()
Me.LucsDatagrid1.Columns(0).Width = 2000
Me.LucsDatagrid1.Columns(1).Width = 1500
Me.LucsDatagrid1.Columns(2).Width = 1000
End Sub


Private Sub LucsDatagrid1_Chk1Click()
If Me.LucsDatagrid1.ValueChk1 = 1 Then
    Me.LucsDatagrid1.Text = 9
Else
    Me.LucsDatagrid1.Text = 0
End If
End Sub

Private Sub LucsDatagrid1_Cmb1Change()
Text2.Text = Me.LucsDatagrid1.ValueCmb1
End Sub
