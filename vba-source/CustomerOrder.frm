VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerOrder 
   Caption         =   "Customer Order"
   ClientHeight    =   4788
   ClientLeft      =   90
   ClientTop       =   408
   ClientWidth     =   4650
   OleObjectBlob   =   "CustomerOrder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustomerOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddOrder_Click()
'select product and quanitity
    If ProdList.Value = "" Or OrderQtxt.Value = 0 Then
        MsgBox "Select a product and a quantity", vbCritical, "ERROR"
        Exit Sub
    End If
    
    'lookup product number using name
    Dim Prow As Variant
    Dim PID As String
    Prow = Application.Match(ProdList.Value, Range("Product"), 0)
    If IsError(Prow) Then
        MsgBox "Product not found", vbCritical, "ERROR"
        Exit Sub
    End If
    
    PID = Range("ProdID").Cells(Prow, 1).Value
    
    'check if inventory is less than zero
    Dim RemInv As Integer
    RemInv = Application.WorksheetFunction.VLookup(PID, Range("ProdPrice"), 5, False)
    If IsError(RemInv) Then
        MsgBox "Product inventory not found", vbCritical, "ERROR"
        Exit Sub
    End If
    
    If RemInv - OrderQtxt < 0 Then
        MsgBox "Not enough inventory to add this order.", vbCritical, "ERROR"
        Exit Sub
    End If
    
    'place order if inventory is greater than zero
    Dim numO As Integer
    numO = Range("TempNumOrders").Value + 1
    Range("TempNumOrders") = numO
    NumOshow.Value = numO
    
    'add order to hidden sheet
    Range("TempOrderStart").Offset(numO, 0) = PID
    Range("TempOrderStart").Offset(numO, 1) = OrderQtxt.Value
    
End Sub
Private Sub CustCancel_Click()
    Unload Me
End Sub
Private Sub CustDone_Click()
    If CustLast.Value = "" Or CustPhone.Value = "" Or CustStore.Value = "" Then
        MsgBox "Please enter customer name, phone number, and store", vbCritical, "ERROR"
        Exit Sub
    End If
    
    Dim numCust As Integer
    numCust = Application.WorksheetFunction.CountA(Worksheets("Data").Range(Worksheets("Data").Range("CustStart").Offset(1, 0), Worksheets("Data").Range("CustStart").End(xlDown)))
    
    'update Data sheet
    'copy and paste the first table row as a new row
    Range(Worksheets("Data").Range("CustStart").Offset(1, 0), Worksheets("Data").Range("CustStart").Offset(1, 4)).Copy
    Worksheets("Data").Range("CustStart").Offset(numCust + 1, 0).PasteSpecial xlPasteAll
    
    'enter the customer info in the new row
    With Worksheets("Data").Range("CustStart")
        .Offset(numCust + 1, 0) = "ORD" & 1000 + numCust + 1
        .Offset(numCust + 1, 1) = CustLast.Value
        .Offset(numCust + 1, 2) = CustPhone.Value
        .Offset(numCust + 1, 3) = CustStore.Value
    End With
    
    'based on the number of orders placed, create new order rows
    Dim numO As Integer
    numO = Range("TempNumOrders").Value
    
    'copy and paste the first order row as new order rows
    Range(Worksheets("Data").Range("OStart").Offset(1, 0), Worksheets("Data").Range("OStart").Offset(1, 4)).Copy
    Range(Worksheets("Data").Range("OStart").End(xlDown).Offset(1, 0), Worksheets("Data").Range("OStart").End(xlDown).Offset(numO, 4)).PasteSpecial xlPasteAll
    
    'copy the order info to the new rows
    Range(Range("TempOrderStart").Offset(1, 0), Range("TempOrderStart").Offset(numO, 1)).Copy
    Worksheets("Data").Range("OStart").End(xlDown).Offset(-numO + 1, 2).PasteSpecial xlPasteValues
    
    'enter the order number and location for each new row
    Range(Worksheets("Data").Range("OStart").End(xlDown).Offset(-numO + 1, 0), Worksheets("Data").Range("OStart").End(xlDown)) = "ORD" & 1000 + numCust + 1
    Range(Worksheets("Data").Range("OStart").End(xlDown).Offset(-numO + 1, 1), Worksheets("Data").Range("OStart").End(xlDown).Offset(0, 1)) = CustStore.Value
    
    'copy the price formula to the new rows
    'Range("OStart").Offset(1, 4).Copy
    'Range(Range("OStart").Offset(0, 4).End(xlDown).Offset(1, 0), Range("OStart").End(xlDown).Offset(0, 4)).PasteSpecial xlPasteFormulas
    
    'update range names for the order table columns
    Worksheets("Data").Visible = True
    Worksheets("Data").Activate
    Application.DisplayAlerts = False
    Range(Range("OStart"), Range("OStart").End(xlDown).Offset(0, 4)).Select
    Selection.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False
    Worksheets("Data").Visible = False
    
    
    
    Unload Me
    
End Sub
Private Sub CustPhone_AfterUpdate()
    If IsNumeric(CustPhone.Value) = False Then
        MsgBox "Must enter a valid phone number"
        Exit Sub
    End If
    
    If Len(CustPhone.Value) <> 10 Then
        MsgBox "Must enter a valid phone number"
        Exit Sub
    End If
    
    CustPhone.Value = Format(CustPhone.Value, "(###) ###-####")
End Sub
Private Sub OrderQspin_Change()
    OrderQtxt.Value = OrderQspin.Value
End Sub
Private Sub CustomerOrder_Initialize()
    Range("TempNumOrders").ClearContents
    Range(Range("TempOrderStart").Offset(1, 0), Range("TempOrderStart").End(xlDown).Offset(0, 1)).ClearContents
    OrderQtxt.Value = 0
    NumOshow.Value = 0
    MsgBox "Add multiple product orders then enter customer info and complete order.", vbOKOnly, "Place Orders"
End Sub
