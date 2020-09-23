VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E06827&
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E06827&
      Caption         =   "Open record from MSHFGrid when row is clicked."
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2940
      TabIndex        =   8
      Top             =   5430
      Width           =   3765
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E06827&
      Caption         =   "Open record from DataGrid when row is clicked."
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2655
      TabIndex        =   7
      Top             =   1395
      Width           =   3585
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H009AC9A4&
      Caption         =   "Export the MSFlexGrid to Excel"
      Height          =   285
      Left            =   3330
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3210
      Width           =   3000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1185
      Left            =   150
      TabIndex        =   2
      Top             =   1890
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   2090
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   0
      ForeColor       =   65280
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ID1"
         Caption         =   "ID1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "First Name"
         Caption         =   "First Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Last Name"
         Caption         =   "Last Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "City"
         Caption         =   "City"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "State"
         Caption         =   "State"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Zip"
         Caption         =   "Zip"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   2670
      Top             =   225
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1980
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   3493
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   4259584
      BackColorFixed  =   0
      ForeColorFixed  =   1017855
      BackColorBkg    =   0
      BackColorUnpopulated=   0
      GridColor       =   3513087
      GridColorUnpopulated=   33023
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form1.frx":0015
      Height          =   1725
      Left            =   90
      TabIndex        =   0
      Top             =   3555
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   3043
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   65280
      BackColorFixed  =   8421504
      ForeColorFixed  =   65535
      BackColorSel    =   16711680
      GridColor       =   4235263
      AllowUserResizing=   1
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   630
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Purchases"
      Top             =   255
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0029
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   210
      TabIndex        =   9
      Top             =   60
      Width           =   14595
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MSHFlexGrid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   135
      TabIndex        =   5
      Top             =   5535
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "MSFlexGrid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   3195
      Width           =   2310
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DataGrid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   135
      TabIndex        =   3
      Top             =   1545
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1

Dim SortCol As Variant
Dim SortDate As Variant

Dim PstrColNum As Integer
Dim bSortAsc As Boolean
Dim MSHFColHeadClick As Boolean
Dim MSFColHeadClick As Boolean

' 10/12/06 Updated
' Changes / Additions
' MSFlexgrid and MSHFlexgrid only sort when column is clicked (Should have been there in the first place)
' Database paths in controls were changed to work wherever you extract the files to. (App.Path) Should have been there in the first place.
' Export to excel now has autofit columns
' Alternating row colors in Excel if desired
' With alternating rows colors font color will change based on index color used for alternate rows
' May need to be adjusted for you. Black and white are the only two colors that I am using.



Private Sub Command3_Click()
    ' Look at code in module has more notes
    ' These colors are based off of my computer there may be some difference on yours.
    '1 = black
    '2 = white
    '3 = Red
    '4 = lime green
    '5 = blue
    '6 = yellow
    '7 = purple
    '8 = light bright blue
    '9 = maroon
    '10 = green
    '11 = dark blue
    '12 = browngreen
    '13 = Dark purple
    '14 = Dark green
    '15 = grey
    '16 = Dark grey
    '17 = dark violet
    '18 = maroon light
    '19 = tan
    '20 = powder blue
    '21 = Dark pruple different than 13
    '22 = salmon
    '23 = blue
    '24= grey
    '25 = dark blue
    '26 = deep pink
    ' 35 = light xp office blue
    ' 37 = light blue green
    
   ' FlexGrd_SaveToExcel MSFlexGrid1, "The Header", "The Footer", 1, 16, App.Path & "\ms_masthead_10x7a_ltr.bmp", , , 28, 4
   ' FlexGrd_SaveToExcel MSFlexGrid1, "The Header", "The Footer", 1, 16, App.Path & "\ms_masthead_10x7a_ltr.bmp", , , 1, 4
   FlexGrd_SaveToExcel MSFlexGrid1, "The Header", "The Footer", 1, 16, App.Path & "\ms_masthead_10x7a_ltr.bmp", , , 37, 35, True
End Sub

Private Sub Command4_Click()
    Load adoClsPurchases
    adoClsPurchases.Show
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
 
    Dim TheColName As String
    Static bSortAsc As Boolean
    Static PrevColName As String
   
   ' The "[" and "]" are crucial for sorting with field names that have spaces
    TheColName = "[" & DataGrid1.Columns(ColIndex).DataField & "]"
    
       ''''''''''''''''''''''''''''''
   'Sort Ascending the first time column is click.
   ' If the same column is clicked then sort descending

    If TheColName = PrevColName Then
        If bSortAsc Then
            Adodc1.Recordset.Sort = TheColName & " DESC"
            bSortAsc = False
        Else
            Adodc1.Recordset.Sort = TheColName
            bSortAsc = True
        End If
    Else
        Adodc1.Recordset.Sort = TheColName
        bSortAsc = True
    End If
    
    PrevColName = TheColName
    
End Sub

Private Sub DataGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

'using column 0 which is not visible as the column Unique id
igRow = DataGrid1.RowContaining(Y) 'Set the Cell Row number
If igRow > -1 And Check1.Value = 1 Then
vargBookmark = DataGrid1.RowBookmark(igRow) 'Set Bookmark Value


idnum = ""
idnum = DataGrid1.Columns(0).CellValue(vargBookmark)

'Creates new adoCustomer form
Dim DocumentCount As Long

    DocumentCount = DocumentCount + 1
    Set adoClsCustomer = New adoClsCustomer
 
    adoClsCustomer.Visible = True
''''''''''''''''''''''''''''''''''''
Load adoClsCustomer
adoClsCustomer.Show


End If

End Sub

Private Sub Form_Load()
Dim db As Connection

  Set db = New Connection

  db.CursorLocation = adUseClient
  db.Provider = "Microsoft.Jet.OLEDB.4.0"
  db.Open (App.Path & "\AdataBasetest.mdb")
 'Connecting ado control to the database and the Customer table
  Connect_AdoControl Adodc1, App.Path & "\AdataBasetest.mdb", "SELECT * FROM Customer", False, ""
 'Set database for Data control
  Data1.DatabaseName = (App.Path & "\AdataBasetest.mdb")
  
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "Select Purchases.* From Purchases ORDER BY PO_Number DESC", db, adOpenStatic, adLockOptimistic
    ' Set Data Source for MSHFlexGrid
  Set MSHFlexGrid1.DataSource = adoPrimaryRS

' Hooking the form for mouse wheel scroll
 Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Call WheelUnHook(Me.hWnd)
End Sub

Private Sub MSFlexGrid1_DblClick()

If MSFColHeadClick = False Then Exit Sub
' Module function
MSGridSort MSFlexGrid1

End Sub





' Here you can add scrolling support to controls that don't normally respond
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
    If TypeOf ctl Is MSHFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then HorFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
     If TypeOf ctl Is DataGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then DataGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub


Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
         ' Checks to see if the Header column was clicked
    If (MSFlexGrid1.RowHeight(0) < Y) Then
        MSFColHeadClick = False
        Else
        MSFColHeadClick = True
        End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
    ' We do not want to sort if we are trying to open a record
    ' Most users have a tendency to double click
     ' If Check2.Value = 1 Or MSHFColHeadClick = False Then Exit Sub
    If MSHFColHeadClick = False Then Exit Sub
            
    Dim TheColName As String
    Static bSortAsc As Boolean
    Static PrevColName As String
   
   ' The "[" and "]" are crucial for sorting with field names that have spaces
   
    TheColName = "[" & MSHFlexGrid1.DataField(0, MSHFlexGrid1.Col) & "]"
    
   ''''''''''''''''''''''''''''''
   'Sort Ascending the first time column is click.
   ' If the same column is clicked then sort descending
    If TheColName = PrevColName Then


        If bSortAsc Then
            adoPrimaryRS.Sort = TheColName & " DESC"
            bSortAsc = False
        Else
            adoPrimaryRS.Sort = TheColName
            bSortAsc = True
        End If
    Else
        adoPrimaryRS.Sort = TheColName
        bSortAsc = True
    End If
    ''''''''''''''''''''''''''''''''
    PrevColName = TheColName
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Checks to see if the Header column was clicked
    If (MSHFlexGrid1.RowHeight(0) < Y) Then
        MSHFColHeadClick = False
        Else
        MSHFColHeadClick = True
        End If
        
End Sub

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' The check box value must be checked or the data grid will act as normal so you can sort or do what ever.
    ' All of the columns are locked so that when you click on the data grid the adoClsPurchases form is on top. Otherwise it is behind and you have to find it.
    If Check2.Value = 1 And MSHFColHeadClick = False Then  ' Opens the record in adoClsPurchases form
        'Sets Purchase order number to bring up. Using filter property of clsPurchases which I added
         ponum = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
        ' Each time the datagrid is click when check box 2 is checked it will open a new form not change existing
         Dim DocumentCount1 As Long
         
             DocumentCount1 = DocumentCount1 + 1
             Set adoClsPurchases = New adoClsPurchases
        
             adoClsPurchases.Visible = True
         
         Load adoClsPurchases
         adoClsPurchases.Show
    End If

  
End Sub
