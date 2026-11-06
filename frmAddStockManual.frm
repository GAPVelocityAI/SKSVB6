VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddStockManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Update"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6360
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   5640
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   20
      Top             =   6240
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11165
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUnit 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtProductName 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtQuantityPerUnit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtValues 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtValues 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtValues 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search product "
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   6135
      Begin VB.CommandButton cmdProducts 
         Caption         =   "..."
         Height          =   315
         Left            =   5400
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Name "
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Code"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lvProducts 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Existence"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Ordered"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Qty per Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblNewQuantity 
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   5445
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Stock quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5445
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Unit"
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Product"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   1100
   End
   Begin VB.Label Label7 
      Caption         =   "Units"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Unit Price"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   4965
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Price"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4965
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4485
      Width           =   1100
   End
   Begin VB.Label Label3 
      Caption         =   "Select a product first"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddStockManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private editingData As Boolean
Private currentIdProduct As String
Private currentQuantityPerUnit As String
Private currentUnit As String
Private currentProductName As String
Private currentPriceReference As Double
Private codeGeneratedChange As Boolean
Private quantity As Double
Private stockPrice As Double, unitPrice As Double

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdProducts_Click()
frmProducts.Show vbModal
txtCode = frmProducts.CurrentProductID
txtName = ""
DoSearchProduct
End Sub

Private Sub cmdSave_Click()
Dim newStockId As Integer
Dim newManualLogId As Integer
Dim newStockLogId As Integer
editingData = False
On Error GoTo HandleError
ExecuteSql "Select * from Stocks"
rs.AddNew
rs("ProductID") = currentIdProduct
rs("Stock") = txtValues(0)
rs("InitialStock") = txtValues(0)
rs("DateStarted") = CStr(Date)
rs("DateModified") = CStr(Date)
rs("User") = UserId
rs("UnitPrice") = txtValues(2)
rs("StockPrice") = txtValues(1)
rs.Update
ExecuteSql "SELECT last_insert_rowid()"
newStockId = rs(0)

ExecuteSql "Select * from ManualStocks"
rs.AddNew
rs("StockID") = newStockId
rs("Quantity") = txtValues(0)
rs("Price") = txtValues(1)
rs("User") = UserId
rs("Date") = CStr(Date)
rs("Action") = "ADD"
rs.Update
ExecuteSql "SELECT last_insert_rowid()"
newManualLogId = rs(0)

ExecuteSql "Select * from StockLog"
rs.AddNew
rs("User") = UserId
rs("Date") = CStr(Date)
rs("Quantity") = txtValues(0)
rs("StockPrice") = txtValues(1)
rs("ProductID") = currentIdProduct
rs("StockID") = newStockId
rs("DocType") = "MANUAL"
rs("DocID") = newManualLogId
rs.Update
ExecuteSql "SELECT last_insert_rowid()"
newStockLogId = rs(0)

ExecuteSql "Update Products Set UnitsInStock = UnitsInStock + " & txtValues(0) & _
" Where ProductId = '& currentIdProduct &'"

If MsgBox("Data added successfully" & vbCrLf & "Would you like to add a new stock manually?", vbYesNo + vbQuestion, "New data") = vbYes Then
    ClearFields
Else
    Unload Me
End If

Exit Sub
HandleError:
MsgBox "An error has occurred adding the data." & vbCrLf & "Error: (" & err.Number & ") " & err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Load()
editingData = False
codeGeneratedChange = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If editingData Then
    Dim res As VbMsgBoxResult
    res = MsgBox("Do you want to save the edited data?", vbYesNoCancel + vbQuestion, "Save data")
    If res = vbYes Then
        cmdSave_Click
    ElseIf res <> vbNo Then
        Cancel = True
    End If
End If
End Sub

Private Sub lvProducts_Click()
RetrieveDataProduct
End Sub

Private Sub lvProducts_ItemClick(ByVal Item As MSComctlLib.ListItem)
RetrieveDataProduct
End Sub

Private Sub txtCode_Change()
DoSearchProduct
End Sub

'Private Sub txtCode_KeyPress(KeyAscii As Integer)
'KeyAscii = UpCase(KeyAscii)
'End Sub

Private Sub txtCode_LostFocus()
If lvProducts.ListItems.Count = 1 Then
    RetrieveDataProduct
End If
End Sub

Private Sub txtName_Change()
DoSearchProduct
End Sub


Private Sub DoSearchProduct()
Dim filter As String
filter = ""
If txtCode <> Empty Then
    filter = "ProductId LIKE '%" & txtCode & "%'"
End If
If txtName <> Empty Then
    If filter <> Empty Then
        filter = filter & " AND "
    End If
    filter = filter & "ProductName LIKE '%" & txtName & "%'"
End If
If filter <> Empty Then
    filter = "Where " & filter
End If
ExecuteSql "Select ProductID, ProductName, UnitPrice, UnitsInStock, UnitsOnOrder, QuantityPerUnit, Unit from Products " & filter
lvProducts.ListItems.Clear
If rs.RecordCount = 0 Then
    LogStatus "There are no records with the selected criteria", Me
Else
    Dim x As ListItem
    While Not rs.EOF
        Set x = lvProducts.ListItems.Add(, , rs.Fields(0))
        For i = 1 To (rs.Fields.Count - 1)
            If Not IsNull(rs.Fields(i)) Then
                x.SubItems(i) = rs.Fields(i)
            End If
        Next i
        rs.MoveNext
    Wend
    If lvProducts.ListItems.Count = 1 Then
        lvProducts.SelectedItem = lvProducts.ListItems(1)
        'RetrieveDataProduct
    End If
End If
End Sub

Private Sub RetrieveDataProduct()
If editingData Then
    If MsgBox("Do you want to cancel previous edited data?", vbYesNo + vbQuestion, "Data edition") <> vbYes Then
        Exit Sub
    End If
End If

If lvProducts.SelectedItem <> Empty Then
    With lvProducts.SelectedItem
        currentIdProduct = lvProducts.SelectedItem
        If .SubItems(5) <> Empty Then currentQuantityPerUnit = .SubItems(5)
        If .SubItems(6) <> Empty Then currentUnit = .SubItems(6)
        currentProductName = .SubItems(1)
        currentPriceReference = .SubItems(2)
    End With
    txtProductName = currentProductName
    txtQuantityPerUnit = currentQuantityPerUnit
    txtUnit = currentUnit
    txtValues(0).Locked = False
    txtValues(1).Locked = False
    txtValues(2).Locked = False
    txtValues(0) = 1
    txtValues(1) = currentPriceReference
    txtValues(2) = currentPriceReference
    txtValues(0).SetFocus
    SelectAll txtValues(0)
    editingData = False
End If
End Sub


Private Sub txtName_LostFocus()
If lvProducts.ListItems.Count = 1 Then
    RetrieveDataProduct
End If
End Sub

Private Sub txtValues_Change(Index As Integer)
If Not codeGeneratedChange Then
    editingData = True
    codeGeneratedChange = True
    If txtValues(0) <> Empty Then quantity = CDbl(txtValues(0))
    If txtValues(1) <> Empty Then stockPrice = CDbl(txtValues(1))
    If txtValues(2) <> Empty Then unitPrice = CDbl(txtValues(2))
    Select Case Index
        Case 0:
            txtValues(1) = unitPrice * quantity
        Case 1:
            txtValues(2) = stockPrice / quantity
        Case 2:
            txtValues(1) = unitPrice * quantity
    End Select
    lblNewQuantity = Format(CDbl(quantity * currentQuantityPerUnit), "##,###.00") & currentUnit
    codeGeneratedChange = False
End If
End Sub

Private Sub txtValues_GotFocus(Index As Integer)
SelectAll txtValues(Index)
End Sub

Private Sub txtValues_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case vbKeyBack, vbKeyClear, vbKeyDelete
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub ClearFields()
codeGeneratedChange = True
txtValues(0).Locked = True
txtValues(1).Locked = True
txtValues(2).Locked = True
txtValues(0) = ""
txtValues(1) = ""
txtValues(2) = ""
txtCode = ""
txtName = ""
txtUnit = ""
txtProductName = ""
txtQuantityPerUnit = ""
lvProducts.ListItems.Clear
txtCode.SetFocus
editingData = False
codeGeneratedChange = False
lblNewQuantity = ""
ClearLogStatus Me
End Sub
