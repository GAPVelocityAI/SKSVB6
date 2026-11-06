VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAdjustStockManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Adjust"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   6390
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   8070
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11218
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtStockID 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5600
      Width           =   1215
   End
   Begin VB.TextBox txtOriginalPrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6040
      Width           =   1215
   End
   Begin VB.TextBox txtValues 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   0
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6040
      Width           =   1215
   End
   Begin VB.TextBox txtQuantityPerUnit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5600
      Width           =   1215
   End
   Begin VB.TextBox txtProductName 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox txtUnit 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtValues 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   1
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtOriginalQuantity 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   6135
      Begin MSComctlLib.ListView lvStocks 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2990
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Stock ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Current Stock"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Initial Stock"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Stock Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Created"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Modified"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "User"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search product "
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   6135
      Begin VB.TextBox txtCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdProducts 
         Caption         =   "..."
         Height          =   315
         Left            =   5400
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Code"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   7560
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvProducts 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2566
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
         Text            =   "Quantity per Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label14 
      Caption         =   "Adjusted quantity"
      Height          =   255
      Left            =   3240
      TabIndex        =   32
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblNewQuantity 
      Height          =   255
      Left            =   4680
      TabIndex        =   31
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Stock quantity"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblCurrentQuantity 
      Height          =   255
      Left            =   1680
      TabIndex        =   29
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Stock ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Product"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Unit"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "New Quantity"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   6525
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "New Price"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   6085
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   6525
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Old Price"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   6085
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Select a product first"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdjustStockManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private editingData As Boolean
Private currentIdProduct As String
Private currentIdStock As Integer
Private currentQuantityPerUnit As String
Private currentUnit As String
Private currentProductName As String
Private currentStockPrice As Double
Private currentStock As Double
Private changedStockPrice As Double
Private changedStock As Double
Private codeGeneratedChange As Boolean
Private quantity As Double
Private stockPrice As Double, unitPrice As Double

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim newStockId As Integer
Dim newManualLogId As Integer
Dim newStockLogId As Integer
editingData = False
On Error GoTo HandleError

Dim deltaStockPrice As Double
Dim deltaStock As Double
changedStockPrice = CDbl(txtValues(0))
changedStock = CDbl(txtValues(1))

deltaStockPrice = changedStockPrice - currentStockPrice
deltaStock = changedStock - currentStock

If deltaStockPrice = 0 And deltaStock = 0 Then
    LogStatus "There is no modification of the Stock, the data won't be saved", Me
    Exit Sub
End If
' UPDATE
ExecuteSql "Update Stocks Set StockPrice = " & changedStockPrice & _
", Stock = " & changedStock & " Where StockId = " & currentIdStock

' NEW
ExecuteSql "Select * from ManualStocks"
rs.AddNew
rs("StockID") = currentIdStock
rs("Quantity") = deltaStock
rs("Price") = deltaStockPrice
rs("User") = UserId
rs("Date") = CStr(Date)
rs("Action") = "MOD"
rs.Update
ExecuteSql "SELECT last_insert_rowid()"
newManualLogId = rs(0)

'NEW
ExecuteSql "Select * from StockLog"
rs.AddNew
rs("User") = UserId
rs("Date") = CStr(Date)
rs("Quantity") = deltaStock
rs("StockPrice") = deltaStockPrice
rs("ProductID") = currentIdProduct
rs("StockID") = currentIdStock
rs("DocType") = "MANUAL"
rs("DocID") = newManualLogId
rs.Update
ExecuteSql "SELECT last_insert_rowid()"
newStockLogId = rs(0)

ExecuteSql "Update Products Set UnitsInStock = UnitsInStock + " & deltaStock & _
" Where ProductId = '& currentIdProduct &'"

If MsgBox("Data modified successfully" & vbCrLf & "Would you like to modify another stock manually?", vbYesNo + vbQuestion, "Modify data") = vbYes Then
    ClearFields
Else
    Unload Me
End If

Exit Sub
HandleError:
MsgBox "An error has occurred adding the data." & vbCrLf & "Error: (" & err.Number & ") " & err.Description, vbCritical, "Error"
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

Private Sub cmdProducts_Click()
frmProducts.Show vbModal
txtCode = frmProducts.CurrentProductID
txtName = ""
DoSearchProduct
End Sub

Private Sub lvProducts_ItemClick(ByVal Item As MSComctlLib.ListItem)
DoSearchStocks
End Sub

Private Sub lvStocks_ItemClick(ByVal Item As MSComctlLib.ListItem)
RetrieveDataProduct
End Sub

Private Sub txtCode_Change()
DoSearchProduct
End Sub

'Private Sub txtCode_KeyPress(KeyAscii As Integer)
'KeyAscii = UpCase(KeyAscii)
'End Sub

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
    LogStatus "There are no records with the selected criteria"
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
    End If
End If
End Sub


Private Sub DoSearchStocks()
If lvProducts.SelectedItem Is Nothing Then
    Exit Sub
End If
If editingData Then
    If MsgBox("Do you want to cancel previous edited data?", vbYesNo + vbQuestion, "Data edition") <> vbYes Then
        Exit Sub
    End If
End If
Dim productId As String
Dim productName As String
productId = lvProducts.SelectedItem
productName = lvProducts.SelectedItem.SubItems(1)
currentIdProduct = lvProducts.SelectedItem
With lvProducts.SelectedItem
    If Not IsEmpty(.SubItems(5)) Then currentQuantityPerUnit = .SubItems(5)
    If Not IsEmpty(.SubItems(6)) Then currentUnit = .SubItems(6)
    currentProductName = .SubItems(1)
End With
txtProductName = productName
txtUnit = currentUnit
txtQuantityPerUnit = currentQuantityPerUnit

ExecuteSql "Select StockID, Stock, InitialStock, UnitPrice, " & _
"StockPrice, DateStarted, DateModified, User From Stocks " & _
 " Where ProductId = '" & productId & "'"
lvStocks.ListItems.Clear
If rs.RecordCount = 0 Then
    LogStatus "There are no stock records of the product " & productName
    RetrieveDataProduct
Else
    Dim x As ListItem
    While Not rs.EOF
        Set x = lvStocks.ListItems.Add(, , rs.Fields(0))
        For i = 1 To (rs.Fields.Count - 1)
            x.SubItems(i) = rs.Fields(i)
        Next i
        rs.MoveNext
    Wend
    If lvStocks.ListItems.Count = 1 Then
        lvStocks.SelectedItem = lvStocks.ListItems(1)
    End If
End If
End Sub

Private Sub RetrieveDataProduct()
If editingData Then
    If MsgBox("Do you want to cancel previous edited data?", vbYesNo + vbQuestion, "Data edition") <> vbYes Then
        Exit Sub
    End If
End If

Dim setEmpty As Boolean
setEmpty = True
If Not lvStocks.SelectedItem Is Nothing Then
    If lvStocks.SelectedItem <> Empty Then
        currentIdStock = lvStocks.SelectedItem
        With lvStocks.SelectedItem
            currentStock = .SubItems(1)
            currentStockPrice = .SubItems(4)
        End With
        codeGeneratedChange = True
        txtOriginalQuantity = currentStock
        txtOriginalPrice = currentStockPrice
        txtStockID = currentIdStock
        txtValues(0) = currentStockPrice
        txtValues(1) = currentStock
        lblNewQuantity = Format(CDbl(currentStock * currentQuantityPerUnit), "##,###.00") & currentUnit
        lblCurrentQuantity = Format(CDbl(currentStock * currentQuantityPerUnit), "##,###.00") & currentUnit
        codeGeneratedChange = False
        setEmpty = False
        txtValues(0).Locked = False
        txtValues(1).Locked = False
        txtValues(0).SetFocus
    End If
End If
If setEmpty Then
    codeGeneratedChange = True
    txtOriginalQuantity.text = ""
    txtOriginalPrice.text = ""
    txtStockID.text = ""
    txtValues(0).text = ""
    txtValues(1).text = ""
    lblNewQuantity.Caption = ""
    lblCurrentQuantity.Caption = ""
    codeGeneratedChange = False
End If
editingData = False

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

Private Sub txtValues_Change(Index As Integer)
If Not codeGeneratedChange Then
    editingData = True
    codeGeneratedChange = True
    If txtValues(0) <> Empty Then changedStockPrice = CDbl(txtValues(0))
    If txtValues(1) <> Empty Then changedStock = CDbl(txtValues(1))
    Select Case Index
        Case 1:
            If changedStock > currentStock Then
                changedStock = currentStock
                LogStatus "Cannot pass the original stock, to add more, add a new stock manually", Me
                txtValues(1) = changedStock
            End If
    End Select
    lblNewQuantity = Format(CDbl(changedStock * currentQuantityPerUnit), "##,###.00") & currentUnit
    lblCurrentQuantity = Format(CDbl(currentStock * currentQuantityPerUnit), "##,###.00") & currentUnit
    codeGeneratedChange = False
End If
End Sub

Private Sub ClearFields()
codeGeneratedChange = True
txtValues(0) = ""
txtValues(1) = ""
txtValues(0).Locked = True
txtValues(1).Locked = True
txtCode = ""
txtName = ""
txtUnit = ""
txtStockID = ""
txtOriginalPrice = ""
txtOriginalQuantity = ""
txtProductName = ""
txtQuantityPerUnit = ""
lvProducts.ListItems.Clear
lvStocks.ListItems.Clear
lblCurrentQuantity = ""
lblNewQuantity = ""
txtCode.SetFocus
editingData = False
codeGeneratedChange = False
ClearLogStatus Me
End Sub
