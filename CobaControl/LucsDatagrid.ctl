VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl LucsDatagrid 
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   DataSourceBehavior=   1  'vbDataSource
   ScaleHeight     =   4050
   ScaleWidth      =   7095
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24838144
      CurrentDate     =   38260
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   315
      Left            =   4080
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
End
Attribute VB_Name = "LucsDatagrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_DataFieldTxt1 = ""
Const m_def_ColumnTxt1 = -1
Const m_def_colCmb3 = -1
Const m_def_ListField3 = ""
Const m_def_BoundColumn3 = ""
Const m_def_RecordSourceCmb3 = ""
Const m_def_DataField3 = ""
Const m_def_FieldChk1 = ""
Const m_def_FieldChk2 = ""
Const m_def_FieldChk3 = ""
Const m_def_ColumnChk1 = -1
Const m_def_ColumnChk2 = -1
Const m_def_ColumnChk3 = -1
Const m_def_FieldDT1 = ""
Const m_def_ColumnDT1 = -1
Const m_def_colCmb1 = -1
Const m_def_colCmb2 = -1
Const m_def_ConnectionString = ""
Const m_def_RecordSource = ""
Const m_def_RecordSourceCmb1 = ""
Const m_def_ListField1 = ""
Const m_def_BoundColumn1 = ""
Const m_def_Datafield1 = ""
Const m_def_RecordSourceCmb2 = ""
Const m_def_Listfield2 = ""
Const m_def_BoundColumn2 = ""
Const m_def_Datafield2 = ""
'Property Variables:
Dim m_DataFieldTxt1 As String
Dim m_ColumnTxt1 As Long
Dim m_colCmb3 As Long
Dim m_ListField3 As String
Dim m_BoundColumn3 As String
Dim m_RecordSourceCmb3 As String
Dim m_DataField3 As String
Dim m_FieldChk1 As String
Dim m_FieldChk2 As String
Dim m_FieldChk3 As String
Dim m_ColumnChk1 As Long
Dim m_ColumnChk2 As Long
Dim m_ColumnChk3 As Long
Dim m_FieldDT1 As String
Dim m_ColumnDT1 As Long
Dim m_colCmb1 As Long
Dim m_colCmb2 As Long
Dim m_ConnectionString As String
Dim m_RecordSource As String
Dim m_RecordSourceCmb1 As String
Dim m_ListField1 As String
Dim m_BoundColumn1 As String
Dim m_Datafield1 As String
Dim m_RecordSourceCmb2 As String
Dim m_Listfield2 As String
Dim m_BoundColumn2 As String
Dim m_Datafield2 As String

Private cn As ADODB.Connection
Public rs As ADODB.Recordset
Private rs1 As New ADODB.Recordset
Private rs2 As New ADODB.Recordset
Private rs3 As New ADODB.Recordset
'Event Declarations:
Event Chk2Click() 'MappingInfo=Check2,Check2,-1,Click
Attribute Chk2Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Chk3Click() 'MappingInfo=Check3,Check3,-1,Click
Attribute Chk3Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Cmb1Change() 'MappingInfo=DataCombo1,DataCombo1,-1,Change
Attribute Cmb1Change.VB_Description = "Indicates that the contents of a control have changed."
Event Cmb2Change() 'MappingInfo=DataCombo2,DataCombo2,-1,Change
Attribute Cmb2Change.VB_Description = "Indicates that the contents of a control have changed."
Event Cmb3Change() 'MappingInfo=DataCombo3,DataCombo3,-1,Change
Attribute Cmb3Change.VB_Description = "Indicates that the contents of a control have changed."
Event DT1AfterUpdate() 'MappingInfo=DTPicker1,DTPicker1,-1,Change
Attribute DT1AfterUpdate.VB_Description = "Occurs when the user selects a new date or changes a date in the edit portion of the control."
Event Chk1Click() 'MappingInfo=Check1,Check1,-1,Click
Attribute Chk1Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
'Event AfterUpdate() 'MappingInfo=DataGrid1,DataGrid1,-1,AfterUpdate



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ConnectionString() As String
    ConnectionString = m_ConnectionString
End Property

Public Property Let ConnectionString(ByVal New_ConnectionString As String)
    m_ConnectionString = New_ConnectionString
    PropertyChanged "ConnectionString"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get RecordSource() As String
    RecordSource = m_RecordSource
End Property

Public Property Let RecordSource(ByVal New_RecordSource As String)
    m_RecordSource = New_RecordSource
    PropertyChanged "RecordSource"
End Property
'


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get RecordSourceCmb1() As String
    RecordSourceCmb1 = m_RecordSourceCmb1
End Property

Public Property Let RecordSourceCmb1(ByVal New_RecordSourceCmb1 As String)
    m_RecordSourceCmb1 = New_RecordSourceCmb1
    PropertyChanged "RecordSourceCmb1"
End Property
'

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ListField1() As String
    ListField1 = m_ListField1
End Property

Public Property Let ListField1(ByVal New_ListField1 As String)
    m_ListField1 = New_ListField1
    PropertyChanged "ListField1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get BoundColumn1() As String
    BoundColumn1 = m_BoundColumn1
End Property

Public Property Let BoundColumn1(ByVal New_BoundColumn1 As String)
    m_BoundColumn1 = New_BoundColumn1
    PropertyChanged "BoundColumn1"
End Property
'

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Datafield1() As String
    Datafield1 = m_Datafield1
End Property

Public Property Let Datafield1(ByVal New_Datafield1 As String)
    m_Datafield1 = New_Datafield1
    PropertyChanged "Datafield1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get RecordSourceCmb2() As String
    RecordSourceCmb2 = m_RecordSourceCmb2
End Property

Public Property Let RecordSourceCmb2(ByVal New_RecordSourceCmb2 As String)
    m_RecordSourceCmb2 = New_RecordSourceCmb2
    PropertyChanged "RecordSourceCmb2"
End Property
'

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Listfield2() As String
    Listfield2 = m_Listfield2
End Property

Public Property Let Listfield2(ByVal New_Listfield2 As String)
    m_Listfield2 = New_Listfield2
    PropertyChanged "Listfield2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get BoundColumn2() As String
    BoundColumn2 = m_BoundColumn2
End Property

Public Property Let BoundColumn2(ByVal New_BoundColumn2 As String)
    m_BoundColumn2 = New_BoundColumn2
    PropertyChanged "BoundColumn2"
End Property
'


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Datafield2() As String
    Datafield2 = m_Datafield2
End Property

Public Property Let Datafield2(ByVal New_Datafield2 As String)
    m_Datafield2 = New_Datafield2
    PropertyChanged "Datafield2"
End Property




Private Sub DataGrid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
DataCombo1.Visible = False
DataCombo2.Visible = False
DataCombo3.Visible = False
DTPicker1.Visible = False
Check1.Visible = False
Check2.Visible = False
Check3.Visible = False
Text1.Visible = False
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hihi
If m_colCmb1 >= 0 Then
    If DataGrid1.Col = m_colCmb1 Then
        With DataCombo1
            DataCombo1.Visible = True
            .Left = DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left
            .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row)
            .Width = DataGrid1.Columns(m_colCmb1).Width + Screen.TwipsPerPixelX
            .Height = DataGrid1.RowHeight + Screen.TwipsPerPixelY
        End With
    Else
        DataCombo1.Visible = False
    End If
End If

If m_colCmb2 >= 0 Then
    If DataGrid1.Col = m_colCmb2 Then
        With DataCombo2
            DataCombo2.Visible = True
            .Left = DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left
            .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row)
            .Width = DataGrid1.Columns(m_colCmb2).Width + Screen.TwipsPerPixelX
            .Height = DataGrid1.RowHeight + Screen.TwipsPerPixelY
        End With
            
    Else
        DataCombo2.Visible = False
    End If
End If

If m_colCmb3 >= 0 Then
    If DataGrid1.Col = m_colCmb2 Then
        With DataCombo3
            DataCombo3.Visible = True
            .Left = DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left
            .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row)
            .Width = DataGrid1.Columns(m_colCmb3).Width + Screen.TwipsPerPixelX
            .Height = DataGrid1.RowHeight + Screen.TwipsPerPixelY
        End With
            
    Else
        DataCombo3.Visible = False
    End If
End If

If m_ColumnDT1 >= 0 Then
    If DataGrid1.Col = m_ColumnDT1 Then
        With DTPicker1
            .Visible = True
            .Left = DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left
            .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row)
            .Width = DataGrid1.Columns(DataGrid1.Col).Width + Screen.TwipsPerPixelX
            .Height = DataGrid1.RowHeight + Screen.TwipsPerPixelY
        End With
    Else
        DTPicker1.Visible = False
    End If
End If

If m_ColumnChk1 >= 0 Then
    If DataGrid1.Col = m_ColumnChk1 Then
        With Check1
            .Visible = True
            .Left = DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left
            .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row)
            .Width = DataGrid1.Columns(DataGrid1.Col).Width + Screen.TwipsPerPixelX
            .Height = DataGrid1.RowHeight + Screen.TwipsPerPixelY
        End With
    Else
        Check1.Visible = False
    End If
End If

If m_ColumnChk2 >= 0 Then
    If DataGrid1.Col = m_ColumnChk2 Then
        With Check2
            .Visible = True
            .Left = DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left
            .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row)
            .Width = DataGrid1.Columns(DataGrid1.Col).Width + Screen.TwipsPerPixelX
            .Height = DataGrid1.RowHeight + Screen.TwipsPerPixelY
        End With
    
    Else
        Check2.Visible = False
    End If
End If

If m_ColumnChk3 >= 0 Then
    If DataGrid1.Col = m_ColumnChk3 Then
        With Check3
            .Visible = True
            .Left = DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left
            .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row)
            .Width = DataGrid1.Columns(DataGrid1.Col).Width + Screen.TwipsPerPixelX
            .Height = DataGrid1.RowHeight + Screen.TwipsPerPixelY
        End With
    Else
        Check3.Visible = False
    End If
End If

If m_ColumnTxt1 >= 0 Then
    If DataGrid1.Col = m_ColumnTxt1 Then
        With Text1
            .Visible = True
            .Left = DataGrid1.Left + DataGrid1.Columns(DataGrid1.Col).Left
            .Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row)
            .Width = DataGrid1.Columns(DataGrid1.Col).Width + Screen.TwipsPerPixelX
            .Height = DataGrid1.RowHeight + Screen.TwipsPerPixelY
        End With
    Else
        Text1.Visible = False
    End If
End If
Exit Sub
hihi:
DataCombo1.Visible = False
DataCombo2.Visible = False
DataCombo3.Visible = False
DTPicker1.Visible = False
Check1.Visible = False
Check2.Visible = False
Check3.Visible = False
Text1.Visible = False
End Sub



Private Sub DataGrid1_Scroll(Cancel As Integer)
DataCombo1.Visible = False
DataCombo2.Visible = False
DataCombo3.Visible = False
DTPicker1.Visible = False
Check1.Visible = False
Check2.Visible = False
Check3.Visible = False
Text1.Visible = False
End Sub

Private Sub Text1_Change()
On Error Resume Next
rs.Fields(m_DataFieldTxt1) = Text1
End Sub

Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
Set cn = New ADODB.Connection
cn.Open m_ConnectionString

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.LockType = adLockBatchOptimistic
rs.ActiveConnection = cn
rs.Open m_RecordSource

If m_RecordSourceCmb1 <> "" Then
    
    With rs1
        'Set rs1 = New ADODB.Recordset
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .ActiveConnection = cn
        .Open m_RecordSourceCmb1
    End With
    With DataCombo1
        Set .RowSource = rs1
        .ListField = m_ListField1
        Set .DataSource = rs
        .BoundColumn = m_BoundColumn1
        .DataField = m_Datafield1
    End With
End If

If m_RecordSourceCmb3 <> "" Then
    
    With rs3
        'Set rs1 = New ADODB.Recordset
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .ActiveConnection = cn
        .Open m_RecordSourceCmb3
    End With
    With DataCombo3
        Set .RowSource = rs3
        .ListField = m_ListField3
        Set .DataSource = rs
        .BoundColumn = m_BoundColumn3
        .DataField = m_DataField3
    End With
End If

If m_RecordSourceCmb2 <> "" Then
    With rs2
        'Set rs2 = New ADODB.Recordset
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .ActiveConnection = cn
        .Open m_RecordSourceCmb2
    End With
    With DataCombo2
        Set .RowSource = rs2
        Set .DataSource = rs
        .ListField = m_Listfield2
        .BoundColumn = m_BoundColumn2
        .DataField = m_Datafield2
    End With
End If

If m_ColumnDT1 >= 0 Then
    If m_FieldDT1 <> "" Then
        Set DTPicker1.DataSource = rs
        DTPicker1.DataField = m_FieldDT1
    Else
        DTPicker1.Visible = False
        
    End If
End If

If m_ColumnChk1 >= 0 Then
    If m_FieldChk1 <> "" Then
        Set Check1.DataSource = rs
        Check1.DataField = m_FieldChk1
    Else
        Check1.Visible = False
        
    End If
End If

If m_ColumnChk2 >= 0 Then
    If m_FieldChk2 <> "" Then
        Set Check2.DataSource = rs
        Check2.DataField = m_FieldChk2
    Else
        Check2.Visible = False
        
    End If
End If

If m_ColumnChk3 >= 0 Then
    If m_FieldChk3 <> "" Then
        Set Check3.DataSource = rs
        Check3.DataField = m_FieldChk3
    Else
        Check3.Visible = False
        
    End If
End If

If m_ColumnTxt1 >= 0 Then
    If m_DataFieldTxt1 <> "" Then
        Set Text1.DataSource = rs
        Text1.DataField = m_DataFieldTxt1
    Else
        Text1.Visible = False
    End If
End If

Set Data = rs
Set DataGrid1.DataSource = rs
DataGrid1.ReBind
End Sub

Sub RebindDG()
rs.Filter = ""
Set DataGrid1.DataSource = Nothing
Set DataGrid1.DataSource = rs

DataCombo1.Visible = False
DataCombo2.Visible = False
DataCombo3.Visible = False
DTPicker1.Visible = False
Check1.Visible = False
Check2.Visible = False
Check3.Visible = False

End Sub


Sub UpdateBatch()

rs.ActiveConnection = cn
rs.UpdateBatch
End Sub

Sub SortAsc()
On Error GoTo hihi
Dim strSort As String
strSort = DataGrid1.Columns(DataGrid1.Col).DataField
strSort = strSort & " Asc"
rs.Sort = strSort '
Exit Sub
hihi:
MsgBox "Please Click on row"
End Sub

Sub SortDesc()
On Error GoTo hihi
Dim strSort As String
strSort = DataGrid1.Columns(DataGrid1.Col).DataField
strSort = strSort & " Desc"
rs.Sort = strSort
Exit Sub
hihi:
MsgBox "Please Click on row"

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ConnectionString = m_def_ConnectionString
    m_RecordSource = m_def_RecordSource
    m_RecordSourceCmb1 = m_def_RecordSourceCmb1
    m_ListField1 = m_def_ListField1
    m_BoundColumn1 = m_def_BoundColumn1
    m_Datafield1 = m_def_Datafield1
    m_RecordSourceCmb2 = m_def_RecordSourceCmb2

    m_Listfield2 = m_def_Listfield2
    m_BoundColumn2 = m_def_BoundColumn2

    m_Datafield2 = m_def_Datafield2
    m_colCmb1 = m_def_colCmb1
    m_colCmb2 = m_def_colCmb2
    m_FieldDT1 = m_def_FieldDT1
    m_ColumnDT1 = m_def_ColumnDT1
    m_ColumnChk1 = m_def_ColumnChk1
    m_ColumnChk2 = m_def_ColumnChk2
    m_ColumnChk3 = m_def_ColumnChk3
    m_FieldChk1 = m_def_FieldChk1
    m_FieldChk2 = m_def_FieldChk2
    m_FieldChk3 = m_def_FieldChk3
    m_colCmb3 = m_def_colCmb3
    m_ListField3 = m_def_ListField3
    m_BoundColumn3 = m_def_BoundColumn3
    m_RecordSourceCmb3 = m_def_RecordSourceCmb3
    m_DataField3 = m_def_DataField3
    m_ColumnTxt1 = m_def_ColumnTxt1
    m_DataFieldTxt1 = m_def_DataFieldTxt1
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_ConnectionString = PropBag.ReadProperty("ConnectionString", m_def_ConnectionString)
    m_RecordSource = PropBag.ReadProperty("RecordSource", m_def_RecordSource)
    m_RecordSourceCmb1 = PropBag.ReadProperty("RecordSourceCmb1", m_def_RecordSourceCmb1)
    m_ListField1 = PropBag.ReadProperty("ListField1", m_def_ListField1)
    m_BoundColumn1 = PropBag.ReadProperty("BoundColumn1", m_def_BoundColumn1)
    m_Datafield1 = PropBag.ReadProperty("Datafield1", m_def_Datafield1)
    m_RecordSourceCmb2 = PropBag.ReadProperty("RecordSourceCmb2", m_def_RecordSourceCmb2)
    m_Listfield2 = PropBag.ReadProperty("Listfield2", m_def_Listfield2)
    m_BoundColumn2 = PropBag.ReadProperty("BoundColumn2", m_def_BoundColumn2)
    m_Datafield2 = PropBag.ReadProperty("Datafield2", m_def_Datafield2)
    m_colCmb1 = PropBag.ReadProperty("colCmb1", m_def_colCmb1)
    m_colCmb2 = PropBag.ReadProperty("colCmb2", m_def_colCmb2)
    m_FieldDT1 = PropBag.ReadProperty("FieldDT1", m_def_FieldDT1)
    m_ColumnDT1 = PropBag.ReadProperty("ColumnDT1", m_def_ColumnDT1)
    m_ColumnChk1 = PropBag.ReadProperty("ColumnChk1", m_def_ColumnChk1)
    m_ColumnChk2 = PropBag.ReadProperty("ColumnChk2", m_def_ColumnChk2)
    m_ColumnChk3 = PropBag.ReadProperty("ColumnChk3", m_def_ColumnChk3)
    m_FieldChk1 = PropBag.ReadProperty("FieldChk1", m_def_FieldChk1)
    m_FieldChk2 = PropBag.ReadProperty("FieldChk2", m_def_FieldChk2)
    m_FieldChk3 = PropBag.ReadProperty("FieldChk3", m_def_FieldChk3)
    m_colCmb3 = PropBag.ReadProperty("colCmb3", m_def_colCmb3)
    m_ListField3 = PropBag.ReadProperty("ListField3", m_def_ListField3)
    m_BoundColumn3 = PropBag.ReadProperty("BoundColumn3", m_def_BoundColumn3)
    m_RecordSourceCmb3 = PropBag.ReadProperty("RecordSourceCmb3", m_def_RecordSourceCmb3)
    m_DataField3 = PropBag.ReadProperty("DataField3", m_def_DataField3)
    Text1.Text = PropBag.ReadProperty("Text", "")
    m_ColumnTxt1 = PropBag.ReadProperty("ColumnTxt1", m_def_ColumnTxt1)
    m_DataFieldTxt1 = PropBag.ReadProperty("DataFieldTxt1", m_def_DataFieldTxt1)
    Check1.Value = PropBag.ReadProperty("ValueChk1", 0)
    Check2.Value = PropBag.ReadProperty("ValueChk2", 0)
    Check3.Value = PropBag.ReadProperty("ValueChk3", 0)
    DataCombo1.BoundText = PropBag.ReadProperty("ValueCmb1", "&h80000005&")
    DataCombo2.BoundText = PropBag.ReadProperty("ValueCmb2", "&h80000005&")
    DataCombo3.BoundText = PropBag.ReadProperty("ValueCmb3", "&h80000005&")
    DTPicker1.Value = PropBag.ReadProperty("ValueDT1", 30 / 9 / 2004)
End Sub

Private Sub UserControl_Resize()
DataGrid1.Top = 0
DataGrid1.Left = 0
DataGrid1.Width = UserControl.Width
DataGrid1.Height = UserControl.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ConnectionString", m_ConnectionString, m_def_ConnectionString)
    Call PropBag.WriteProperty("RecordSource", m_RecordSource, m_def_RecordSource)
    Call PropBag.WriteProperty("RecordSourceCmb1", m_RecordSourceCmb1, m_def_RecordSourceCmb1)
    Call PropBag.WriteProperty("ListField1", m_ListField1, m_def_ListField1)
    Call PropBag.WriteProperty("BoundColumn1", m_BoundColumn1, m_def_BoundColumn1)
    Call PropBag.WriteProperty("Datafield1", m_Datafield1, m_def_Datafield1)
    Call PropBag.WriteProperty("RecordSourceCmb2", m_RecordSourceCmb2, m_def_RecordSourceCmb2)
    Call PropBag.WriteProperty("Listfield2", m_Listfield2, m_def_Listfield2)
    Call PropBag.WriteProperty("BoundColumn2", m_BoundColumn2, m_def_BoundColumn2)
    Call PropBag.WriteProperty("Datafield2", m_Datafield2, m_def_Datafield2)
    Call PropBag.WriteProperty("colCmb1", m_colCmb1, m_def_colCmb1)
    Call PropBag.WriteProperty("colCmb2", m_colCmb2, m_def_colCmb2)
    Call PropBag.WriteProperty("FieldDT1", m_FieldDT1, m_def_FieldDT1)
    Call PropBag.WriteProperty("ColumnDT1", m_ColumnDT1, m_def_ColumnDT1)
    Call PropBag.WriteProperty("ColumnChk1", m_ColumnChk1, m_def_ColumnChk1)
    Call PropBag.WriteProperty("ColumnChk2", m_ColumnChk2, m_def_ColumnChk2)
    Call PropBag.WriteProperty("ColumnChk3", m_ColumnChk3, m_def_ColumnChk3)
    Call PropBag.WriteProperty("FieldChk1", m_FieldChk1, m_def_FieldChk1)
    Call PropBag.WriteProperty("FieldChk2", m_FieldChk2, m_def_FieldChk2)
    Call PropBag.WriteProperty("FieldChk3", m_FieldChk3, m_def_FieldChk3)
    Call PropBag.WriteProperty("colCmb3", m_colCmb3, m_def_colCmb3)
    Call PropBag.WriteProperty("ListField3", m_ListField3, m_def_ListField3)
    Call PropBag.WriteProperty("BoundColumn3", m_BoundColumn3, m_def_BoundColumn3)
    Call PropBag.WriteProperty("RecordSourceCmb3", m_RecordSourceCmb3, m_def_RecordSourceCmb3)
    Call PropBag.WriteProperty("DataField3", m_DataField3, m_def_DataField3)
    Call PropBag.WriteProperty("Text", Text1.Text, "")
    Call PropBag.WriteProperty("ColumnTxt1", m_ColumnTxt1, m_def_ColumnTxt1)
    Call PropBag.WriteProperty("DataFieldTxt1", m_DataFieldTxt1, m_def_DataFieldTxt1)
    Call PropBag.WriteProperty("ValueChk1", Check1.Value, 0)
    Call PropBag.WriteProperty("ValueChk2", Check2.Value, 0)
    Call PropBag.WriteProperty("ValueChk3", Check3.Value, 0)
    Call PropBag.WriteProperty("ValueCmb1", DataCombo1.BoundText, "&h80000005&")
    Call PropBag.WriteProperty("ValueCmb2", DataCombo2.BoundText, "&h80000005&")
    Call PropBag.WriteProperty("ValueCmb3", DataCombo3.BoundText, "&h80000005&")
    Call PropBag.WriteProperty("ValueDT1", DTPicker1.Value, 30 / 9 / 2004)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get colCmb1() As Long
    colCmb1 = m_colCmb1
End Property

Public Property Let colCmb1(ByVal New_colCmb1 As Long)
    m_colCmb1 = New_colCmb1
    PropertyChanged "colCmb1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get colCmb2() As Long
    colCmb2 = m_colCmb2
End Property

Public Property Let colCmb2(ByVal New_colCmb2 As Long)
    m_colCmb2 = New_colCmb2
    PropertyChanged "colCmb2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FieldDT1() As String
    FieldDT1 = m_FieldDT1
End Property

Public Property Let FieldDT1(ByVal New_FieldDT1 As String)
    m_FieldDT1 = New_FieldDT1
    PropertyChanged "FieldDT1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,-1
Public Property Get ColumnDT1() As Long
    ColumnDT1 = m_ColumnDT1
End Property

Public Property Let ColumnDT1(ByVal New_ColumnDT1 As Long)
    m_ColumnDT1 = New_ColumnDT1
    PropertyChanged "ColumnDT1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,-1
Public Property Get ColumnChk1() As Long
    ColumnChk1 = m_ColumnChk1
End Property

Public Property Let ColumnChk1(ByVal New_ColumnChk1 As Long)
    m_ColumnChk1 = New_ColumnChk1
    PropertyChanged "ColumnChk1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,-1
Public Property Get ColumnChk2() As Long
    ColumnChk2 = m_ColumnChk2
End Property

Public Property Let ColumnChk2(ByVal New_ColumnChk2 As Long)
    m_ColumnChk2 = New_ColumnChk2
    PropertyChanged "ColumnChk2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,-1
Public Property Get ColumnChk3() As Long
    ColumnChk3 = m_ColumnChk3
End Property

Public Property Let ColumnChk3(ByVal New_ColumnChk3 As Long)
    m_ColumnChk3 = New_ColumnChk3
    PropertyChanged "ColumnChk3"
End Property

Public Property Get FieldChk1() As String
    FieldChk1 = m_FieldChk1
End Property

Public Property Let FieldChk1(ByVal New_FieldChk1 As String)
    m_FieldChk1 = New_FieldChk1
    PropertyChanged "FieldChk1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,-1
Public Property Get FieldChk2() As String
    FieldChk2 = m_FieldChk2
End Property

Public Property Let FieldChk2(ByVal New_FieldChk2 As String)
    m_FieldChk2 = New_FieldChk2
    PropertyChanged "FieldChk2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,-1
Public Property Get FieldChk3() As String
    FieldChk3 = m_FieldChk3
End Property

Public Property Let FieldChk3(ByVal New_FieldChk3 As String)
    m_FieldChk3 = New_FieldChk3
    PropertyChanged "FieldChk3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,-1
Public Property Get colCmb3() As Long
    colCmb3 = m_colCmb3
End Property

Public Property Let colCmb3(ByVal New_colCmb3 As Long)
    m_colCmb3 = New_colCmb3
    PropertyChanged "colCmb3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ListField3() As String
    ListField3 = m_ListField3
End Property

Public Property Let ListField3(ByVal New_ListField3 As String)
    m_ListField3 = New_ListField3
    PropertyChanged "ListField3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get BoundColumn3() As String
    BoundColumn3 = m_BoundColumn3
End Property

Public Property Let BoundColumn3(ByVal New_BoundColumn3 As String)
    m_BoundColumn3 = New_BoundColumn3
    PropertyChanged "BoundColumn3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get RecordSourceCmb3() As String
    RecordSourceCmb3 = m_RecordSourceCmb3
End Property

Public Property Let RecordSourceCmb3(ByVal New_RecordSourceCmb3 As String)
    m_RecordSourceCmb3 = New_RecordSourceCmb3
    PropertyChanged "RecordSourceCmb3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DataField3() As String
    DataField3 = m_DataField3
End Property

Public Property Let DataField3(ByVal New_DataField3 As String)
    m_DataField3 = New_DataField3
    PropertyChanged "DataField3"
End Property
'
'Private Sub DataGrid1_AfterUpdate()
'    RaiseEvent AfterUpdate
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get ColumnTxt1() As Long
    ColumnTxt1 = m_ColumnTxt1
End Property

Public Property Let ColumnTxt1(ByVal New_ColumnTxt1 As Long)
    m_ColumnTxt1 = New_ColumnTxt1
    PropertyChanged "ColumnTxt1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DataFieldTxt1() As String
    DataFieldTxt1 = m_DataFieldTxt1
End Property

Public Property Let DataFieldTxt1(ByVal New_DataFieldTxt1 As String)
    m_DataFieldTxt1 = New_DataFieldTxt1
    PropertyChanged "DataFieldTxt1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Check1,Check1,-1,Value
Public Property Get ValueChk1() As Integer
Attribute ValueChk1.VB_Description = "Returns/sets the value of an object."
    ValueChk1 = Check1.Value
End Property

Public Property Let ValueChk1(ByVal New_ValueChk1 As Integer)
    Check1.Value() = New_ValueChk1
    PropertyChanged "ValueChk1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Check2,Check2,-1,Value
Public Property Get ValueChk2() As Integer
Attribute ValueChk2.VB_Description = "Returns/sets the value of an object."
    ValueChk2 = Check2.Value
End Property

Public Property Let ValueChk2(ByVal New_ValueChk2 As Integer)
    Check2.Value() = New_ValueChk2
    PropertyChanged "ValueChk2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Check3,Check3,-1,Value
Public Property Get ValueChk3() As Integer
Attribute ValueChk3.VB_Description = "Returns/sets the value of an object."
    ValueChk3 = Check3.Value
End Property

Public Property Let ValueChk3(ByVal New_ValueChk3 As Integer)
    Check3.Value() = New_ValueChk3
    PropertyChanged "ValueChk3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo1,DataCombo1,-1,BoundText
Public Property Get ValueCmb1() As String
Attribute ValueCmb1.VB_Description = "Returns/sets the value of the data field named in the BoundColumn property."
    ValueCmb1 = DataCombo1.BoundText
End Property

Public Property Let ValueCmb1(ByVal New_ValueCmb1 As String)
    DataCombo1.BoundText() = New_ValueCmb1
    PropertyChanged "ValueCmb1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo2,DataCombo2,-1,BoundText
Public Property Get ValueCmb2() As String
Attribute ValueCmb2.VB_Description = "Returns/sets the value of the data field named in the BoundColumn property."
    ValueCmb2 = DataCombo2.BoundText
End Property

Public Property Let ValueCmb2(ByVal New_ValueCmb2 As String)
    DataCombo2.BoundText() = New_ValueCmb2
    PropertyChanged "ValueCmb2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo3,DataCombo3,-1,BoundText
Public Property Get ValueCmb3() As String
Attribute ValueCmb3.VB_Description = "Returns/sets the value of the data field named in the BoundColumn property."
    ValueCmb3 = DataCombo3.BoundText
End Property

Public Property Let ValueCmb3(ByVal New_ValueCmb3 As String)
    DataCombo3.BoundText() = New_ValueCmb3
    PropertyChanged "ValueCmb3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DTPicker1,DTPicker1,-1,Value
Public Property Get ValueDT1() As Variant
Attribute ValueDT1.VB_Description = "Returns/sets the current date."
    ValueDT1 = DTPicker1.Value
End Property

Public Property Let ValueDT1(ByVal New_ValueDT1 As Variant)
    DTPicker1.Value() = New_ValueDT1
    PropertyChanged "ValueDT1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataGrid1,DataGrid1,-1,Columns
Public Property Get Columns() As Columns
Attribute Columns.VB_Description = "Contains a collection of grid columns"
    Set Columns = DataGrid1.Columns
End Property

Private Sub Check1_Click()
    RaiseEvent Chk1Click
End Sub

Private Sub Check2_Click()
    RaiseEvent Chk2Click
End Sub

Private Sub Check3_Click()
    RaiseEvent Chk3Click
End Sub

Private Sub DataCombo1_Change()
    RaiseEvent Cmb1Change
End Sub

Private Sub DataCombo2_Change()
    RaiseEvent Cmb2Change
End Sub

Private Sub DataCombo3_Change()
    RaiseEvent Cmb3Change
End Sub

Private Sub DTPicker1_Change()
    RaiseEvent DT1AfterUpdate
End Sub

