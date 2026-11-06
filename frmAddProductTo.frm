VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddProductTo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Product Item"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6360
   Begin VB.CheckBox chkAll 
      Caption         =   "Check All"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove "
      Height          =   375
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   7320
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search product "
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmdProducts 
         Caption         =   "..."
         Height          =   315
         Left            =   5400
         TabIndex        =   5
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
      Begin MSComctlLib.ListView lvProducts 
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4260
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
      Begin VB.Label Label4 
         Caption         =   "Name"
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
   Begin MSComctlLib.ListView lvProductsBy 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
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
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Quantity per Unit"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label lblProductsRelated 
      Caption         =   "Products"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   6135
   End
End
Attribute VB_Name = "frmAddProductTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Id As Integer
Public ObjectReferred As String
Public Table As String
Public ColumnName As String

Public SavedChanges As Boolean
Private productsStored As Collection
Private productsToDelete As Collection
Private productsToAdd As Collection
Private editingData As Boolean
Private currentIdProduct As String

Private codeGeneratedChange As Boolean

Private Sub chkAll_Click()
Dim check As Boolean
check = chkAll.value = 1
For i = 1 To lvProductsBy.ListItems.Count
    lvProductsBy.ListItems(i).Checked = check
Next i
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdProducts_Click()
frmProducts.Show vbModal
txtCode = frmProducts.CurrentProductID
txtName = ""
DoSearchProduct
End Sub

Private Sub cmdRemove_Click()
Dim productIdToDelete As String
For i = lvProductsBy.ListItems.Count To 1 Step -1
    If lvProductsBy.ListItems(i).Checked Then
        productIdToDelete = lvProductsBy.ListItems.Item(i)
        
        If Exists(productsStored, productIdToDelete) Then
            If Exists(productsToAdd, productIdToDelete) Then
                productsToDelete.Remove productIdToDelete
            Else
                AddToCollection productsToDelete, productIdToDelete
            End If
        Else
            If Exists(productsToAdd, currentIdProduct) Then
                productsToAdd.Remove currentIdProduct
            End If
        End If
        
        lvProductsBy.ListItems.Remove i
        editingData = True
    End If
Next i
End Sub

Private Sub cmdSave_Click()

If productsToAdd.Count = 0 And productsToDelete.Count = 0 Then
    editingData = True
    MsgBox "No data to be saved", vbOKOnly + vbInformation, "No data modified"
    Unload Me
    Exit Sub
End If
SavedChanges = True
Dim productCode
For Each productCode In productsToAdd
    ExecuteSql "Insert into " & Table & "(" & ColumnName & ", ProductID) Values (" & Id & ", '" & productCode & "')"
Next
For Each productCode In productsToDelete
    ExecuteSql "Delete from " & Table & " Where " & ColumnName & " = " & Id & " And ProductID = '" & productCode & "'"
Next

editingData = False
MsgBox "Data was succesfully saved", vbOKOnly + vbInformation, "New data"
Unload Me
Exit Sub
HandleError:
MsgBox "An error has occurred adding the data." & vbCrLf & "Error: (" & err.Number & ") " & err.Description, vbCritical, "Error"
End Sub

Public Sub LoadData()
editingData = False
editingData = False
codeGeneratedChange = False
Me.Caption = "Add product(s) to " & ObjectReferred
lblProductsRelated = "Products related to " & ObjectReferred
Set productsStored = New Collection
Set productsToDelete = New Collection
Set productsToAdd = New Collection
LoadProductsById
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

Private Sub lvProducts_ItemClick(ByVal Item As MSComctlLib.ListItem)
AddProductToSet
End Sub

Private Sub txtCode_Change()
DoSearchProduct
End Sub

'Private Sub txtCode_KeyPress(KeyAscii As Integer)
'KeyAscii = UpCase(KeyAscii)
'End Sub

Private Sub txtCode_LostFocus()
If lvProducts.ListItems.Count = 1 Then
    AddProductToSet
End If
End Sub

Private Sub txtName_Change()
DoSearchProduct
End Sub

Private Sub LoadProductsById()
Dim productCode As String
ExecuteSql "Select p.ProductID, p.ProductName, p.UnitPrice, p.QuantityPerUnit, p.Unit from Products as p, " & Table _
& " as pb Where pb." & ColumnName & " = " & Id & " And pb.ProductId = p.ProductId"

LogStatus "There are " & rs.RecordCount & " records with the selected criteria", Me
If rs.RecordCount > 0 Then
    Dim x As ListItem
    While Not rs.EOF
        'productCode = CStr(rs.Fields(0).value)
productCode = rs.Fields(0)
        AddToCollection productsStored, productCode
        Set x = lvProductsBy.ListItems.Add(, , productCode)
        For i = 1 To 2
            If Not IsNull(rs.Fields(i)) Then
                x.SubItems(i) = rs.Fields(i)
            End If
        Next i
        x.SubItems(3) = rs.Fields(3) & rs.Fields(4)
        rs.MoveNext
    Wend
End If
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
LogStatus "There are " & rs.RecordCount & " records with the selected criteria", Me
If rs.RecordCount > 0 Then
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

Private Sub AddProductToSet()

If lvProducts.SelectedItem <> Empty Then
    Dim y As ListItem
    Set y = lvProducts.SelectedItem
    currentIdProduct = lvProducts.SelectedItem
    Dim i As Integer
    Dim found As Boolean
    found = False
    For i = 1 To lvProductsBy.ListItems.Count
        If lvProductsBy.ListItems(i) = currentIdProduct Then
            lvProductsBy.SelectedItem = lvProductsBy.ListItems(i)
            found = True
            Exit For
        ElseIf lvProductsBy.ListItems(i) > currentIdProduct Then
            Exit For
        End If
    Next i
    If Not found Then
        editingData = True
        If Not Exists(productsStored, currentIdProduct) Then
            If Exists(productsToDelete, currentIdProduct) Then
                productsToDelete.Remove currentIdProduct
            Else
                AddToCollection productsToAdd, currentIdProduct
            End If
        Else
            If Exists(productsToDelete, currentIdProduct) Then
                productsToDelete.Remove currentIdProduct
            End If
        End If
        Dim x As ListItem
        Set x = lvProductsBy.ListItems.Add(i, , currentIdProduct)
        x.SubItems(1) = y.SubItems(1)
        x.SubItems(2) = y.SubItems(2)
        x.SubItems(3) = y.SubItems(5) & y.SubItems(6)
    End If
End If
End Sub

Private Sub ClearFields()
codeGeneratedChange = True
txtCode = ""
txtName = ""
lvProducts.ListItems.Clear
lvProductsBy.ListItems.Clear
txtCode.SetFocus
editingData = False
codeGeneratedChange = False
End Sub
