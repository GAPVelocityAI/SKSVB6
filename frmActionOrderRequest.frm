VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmActionOrderRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Invoice"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   7845
   Begin VB.TextBox txtPromisedBy 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtRequiredBy 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtReceivedBy 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "&Create"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtReceived 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   540
      Width           =   1575
   End
   Begin VB.TextBox txtChangedBy 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtChanged 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   540
      Width           =   1575
   End
   Begin VB.TextBox txtOrderID 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtNotes 
      BackColor       =   &H80000004&
      Height          =   660
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   6855
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox txtTotalTax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox txtFreightCharge 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox txtSalesTax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid fgDetails 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BorderStyle     =   0
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   7635
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13785
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Customer"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   7575
      Begin VB.TextBox txtCustomerContact 
         BackColor       =   &H80000004&
         Height          =   300
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtCustomerCompany 
         BackColor       =   &H80000004&
         Height          =   300
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Contact:"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Required"
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Promised"
      Height          =   255
      Left            =   3960
      TabIndex        =   35
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "By"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Status:"
      Height          =   255
      Left            =   4800
      TabIndex        =   32
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "Requested:"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblChangedBy 
      Caption         =   "By"
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Order Id:"
      Height          =   255
      Left            =   180
      TabIndex        =   27
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblChanged 
      Caption         =   "Changed:"
      Height          =   255
      Left            =   4800
      TabIndex        =   26
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Freight"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Total:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Total Tax:"
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Sub Total:"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Sales Tax:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   495
   End
End
Attribute VB_Name = "frmActionOrderRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private currentSubTotal As Double
Private currentTotal As Double
Private currentTax As Double
Private currentFreightCharge As Double
Private currentTotalTax As Double

Public Action As Integer

Public OrderId As Integer

Private Sub cmdApprove_Click()
On Error GoTo HandleError
If UCase(txtStatus) = "APPROVED" Then
    LogStatus "Order is already approved, not need to be approved again", Me
    Exit Sub
End If

If UCase(txtStatus) = "CANCELLED" Then
    LogStatus "Order was already cancelled by " & txtChangedBy & " on " & txtChanged & ", it cannot be approved", Me
    Exit Sub
End If

' UPDATE
ExecuteSql "Update OrderRequests Set Status = 'APPROVED', ChangedBy = '" & UserId & "', ChangedDate = '" & Date & "'" & _
" Where OrderId = " & OrderId

LoadData
MsgBox "The order was successfully approved"
Unload Me

Exit Sub
HandleError:
MsgBox "An error has occurred adding the data." & vbCrLf & "Error: (" & err.Number & ") " & err.Description, vbCritical, "Error"

End Sub

Private Sub cmdCancel_Click()
On Error GoTo HandleError
If UCase(txtStatus) = "CANCELLED" Then
    LogStatus "Order was already cancelled, not need to be cancelled again", Me
    Exit Sub
End If
If UCase(txtStatus) = "APPROVED" Then
    LogStatus "Order was already approved by " & txtChangedBy & " on " & txtChanged & ", it cannot be cancelled", Me
    Exit Sub
End If


If MsgBox("Do you want to cancel the order request?", vbYesNo + vbQuestion, "Confirm cancellation") <> vbYes Then
    Exit Sub
End If

' UPDATE
ExecuteSql "Update OrderRequests Set Status = 'CANCELLED', ChangedBy = '" & UserId & "', ChangedDate = '" & Date & "'" & _
" Where OrderId = " & OrderId

LoadData
MsgBox "The order was successfully cancelled"
Unload Me

Exit Sub
HandleError:
MsgBox "An error has occurred adding the data." & vbCrLf & "Error: (" & err.Number & ") " & err.Description, vbCritical, "Error"

End Sub

Private Sub Form_Load()
' LoadData
If Action <> 0 Then
   
    Select Case (Action)
        Case 1:
           ' cmdApprove_Click
        Case 2:
           ' cmdCancel_Click
    End Select
End If
End Sub

Public Sub LoadData()
currentSubTotal = 0
currentTotalTax = 0
ExecuteSql "Select o.OrderDate, u.Fullname, o.Status, c.CompanyName, c.ContactFirstName + ' ' + c.ContactLastName as Contact, o.ChangedDate, o.ChangedBy, o.FreightCharge, o.SalesTaxRate, o.RequiredByDate, o.PromisedByDate, o.Notes " & _
"From OrderRequests as o, Users as u, Customers as c " & _
"Where o.OrderID = " & OrderId & " And u.Username = o.EmployeeId And c.CustomerId = o.CustomerId"
If rs.EOF Then
    LogStatus "The order with the ID '" & OrderId & "' does not exist", Me
    Exit Sub
End If
txtOrderID = OrderId
txtReceived = rs("OrderDate")
txtReceivedBy = rs("Fullname")
If rs("Notes") <> Null Then txtNotes = rs("Notes")
txtFreightCharge = rs("FreightCharge")
currentFreightCharge = rs("FreightCharge")
txtSalesTax = rs("SalesTaxRate")
currentTax = rs("SalesTaxRate")
txtCustomerCompany = rs("CompanyName")
txtCustomerContact = rs("Contact")
txtStatus = rs("Status")
txtRequiredBy = rs("RequiredByDate")
txtPromisedBy = rs("PromisedByDate")
If rs("ChangedDate") <> Null Then txtChanged = rs("ChangedDate")
If rs("ChangedBy") <> Null Then txtChangedBy = rs("ChangedBy")

Dim isRequested As Boolean
isRequested = txtStatus = "REQUESTED"
lblChanged.Visible = Not isRequested
lblChangedBy.Visible = Not isRequested
txtChanged.Visible = Not isRequested
txtChangedBy.Visible = Not isRequested
cmdApprove.Enabled = True ' Requested
cmdCancel.Enabled = True ' Requested

If txtStatus = "APPROVED" Then
    lblChanged = "Approved Date:"
    lblChangedBy = "Approved By:"
Else
    lblChanged = "Cancelled Date:"
    lblChangedBy = "Cancelled By:"
End If
LoadDetails
DisplayTotals
End Sub

Private Sub DisplayTotals()
currentTotal = currentFreightCharge + currentSubTotal + currentTotalTax
txtSubTotal = Format(currentSubTotal, "#,##0.00")
txtTotalTax = Format(currentTotalTax, "#,##0.00")
txtTotal = Format(currentTotal, "#,##0.00")
End Sub


Private Sub AddToTotals(current As Double)
currentSubTotal = currentSubTotal + current
currentTotalTax = currentSubTotal * currentTax
currentTotal = currentFreightCharge + currentSubTotal + currentTotalTax
txtSubTotal = Format(currentSubTotal, "#,##0.00")
txtTotalTax = Format(currentTotalTax, "#,##0.00")
txtTotal = Format(currentTotal, "#,##0.00")
End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub LoadDetails()

ExecuteSql "Select d.Quantity, p.ProductID, p.ProductName, d.UnitPrice, d.SalePrice, p.UnitsInStock, p.UnitsOnOrder, cast(p.QuantityPerUnit as text) + p.Unit, d.LineTotal From Products as p, OrderRequestDetails as d " & "Where d.OrderID = " & OrderId & " And d.ProductId = p.ProductId"

Dim lng As Long
Dim intLoopCount As Integer
With fgDetails
    .Rows = 0
    .Cols = 9
    .FixedCols = 0
    .AddItem "Quantity" & vbTab & "Code" & vbTab & "Product" & vbTab & "UnitPrice" & vbTab & "Price" & vbTab & "Existence" & vbTab & "Ordered" & vbTab & "Quantity per unit" & vbTab & "Line Total"
    .Rows = rs.RecordCount + 1
    If .Rows = 1 Then .FixedRows = 0 Else .FixedRows = 1
    Dim i As Integer
    Dim j As Integer
    i = 1
    While Not rs.EOF
        For j = 1 To rs.Fields.Count
            If Not IsNull(rs.Fields(j - 1)) Then
                .TextMatrix(i, j - 1) = rs.Fields(j - 1)
            End If
        Next j
        AddToTotals rs("LineTotal")
        rs.MoveNext
        i = i + 1
    Wend
End With

End Sub

