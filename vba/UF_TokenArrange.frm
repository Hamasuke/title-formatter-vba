VERSION 5.00
Begin VB.UserForm UF_TokenArrange
   Caption         =   "UF_TokenArrange"
   ClientHeight    =   3000
   ClientWidth     =   4800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_TokenArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private batchSheet As String
Private batchRows() As Long
Private batchIndex As Long

Private WithEvents lbPool As MSForms.ListBox
Private WithEvents lbOrder As MSForms.ListBox
Private WithEvents btnAdd As MSForms.CommandButton
Private WithEvents btnBack As MSForms.CommandButton
Private WithEvents btnOK As MSForms.CommandButton
Private lblPreviewSrc As MSForms.Label
Private lblPreview As MSForms.Label
Private lblInfo As MSForms.Label

Public Sub InitBatch(ByVal sheetName As String, ByRef rowsArr() As Long, ByVal startIndex As Long)
    batchSheet = sheetName
    batchRows = rowsArr
    batchIndex = startIndex
    EnsureUI
    LoadTarget
End Sub

Private Sub EnsureUI()
    If lbPool Is Nothing Then
        Me.Caption = "タイトル整形"
        Me.Width = 420
        Me.Height = 300

        Set lbPool = Me.Controls.Add("Forms.ListBox.1", "lbPool")
        lbPool.Left = 10: lbPool.Top = 70
        lbPool.Width = 150: lbPool.Height = 150
        lbPool.MultiSelect = fmMultiSelectExtended

        Set lbOrder = Me.Controls.Add("Forms.ListBox.1", "lbOrder")
        lbOrder.Left = 260: lbOrder.Top = 70
        lbOrder.Width = 150: lbOrder.Height = 150

        Set btnAdd = Me.Controls.Add("Forms.CommandButton.1", "btnAdd")
        btnAdd.Left = 170: btnAdd.Top = 100
        btnAdd.Caption = ">>追加"

        Set btnBack = Me.Controls.Add("Forms.CommandButton.1", "btnBack")
        btnBack.Left = 170: btnBack.Top = 140
        btnBack.Caption = "<<戻す"

        Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
        btnOK.Left = 170: btnOK.Top = 220
        btnOK.Caption = "確定"

        Set lblInfo = Me.Controls.Add("Forms.Label.1", "lblInfo")
        lblInfo.Left = 10: lblInfo.Top = 10
        lblInfo.Width = 380

        Set lblPreviewSrc = Me.Controls.Add("Forms.Label.1", "lblPreviewSrc")
        lblPreviewSrc.Left = 10: lblPreviewSrc.Top = 30
        lblPreviewSrc.Width = 380

        Set lblPreview = Me.Controls.Add("Forms.Label.1", "lblPreview")
        lblPreview.Left = 10: lblPreview.Top = 45
        lblPreview.Width = 380
    End If
End Sub

Private Sub LoadTarget()
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim srcText As String
    Dim tokens As Variant
    Dim i As Long

    rowNum = batchRows(batchIndex)
    Set ws = ThisWorkbook.Worksheets(batchSheet)

    srcText = CStr(ws.Cells(rowNum, 1).Value)
    tokens = Tokenize(srcText)

    lbPool.Clear
    lbOrder.Clear

    For i = LBound(tokens) To UBound(tokens)
        lbPool.AddItem tokens(i)
    Next i

    lblInfo.Caption = "対象: " & batchSheet & "!A" & rowNum & _
                      " (" & (batchIndex + 1) & " / " & (UBound(batchRows) - LBound(batchRows) + 1) & ")"
    UpdatePreview
End Sub

Private Function Tokenize(ByVal text As String) As Variant
    Dim t As String
    Dim arr As Variant

    t = Replace(text, ChrW(&H3000), " ")
    t = Trim(t)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    If t = "" Then
        Tokenize = Array()
    Else
        arr = Split(t, " ")
        Tokenize = arr
    End If
End Function

Private Sub btnAdd_Click()
    MoveSelected lbPool, lbOrder, True
    UpdatePreview
End Sub

Private Sub btnBack_Click()
    Dim i As Long
    Dim sel As Boolean
    For i = 0 To lbOrder.ListCount - 1
        If lbOrder.Selected(i) Then sel = True: Exit For
    Next i
    If Not sel Then
        If lbOrder.ListCount = 0 Then Exit Sub
        lbOrder.Selected(lbOrder.ListCount - 1) = True
    End If
    MoveSelected lbOrder, lbPool, True
    UpdatePreview
End Sub

Private Sub btnOK_Click()
    Dim ws As Worksheet
    Dim rowNum As Long

    rowNum = batchRows(batchIndex)
    Set ws = ThisWorkbook.Worksheets(batchSheet)
    ws.Cells(rowNum, 2).Value = JoinListBox(lbOrder)

    batchIndex = batchIndex + 1
    If batchIndex > UBound(batchRows) Then
        Unload Me
    Else
        LoadTarget
    End If
End Sub

Private Sub lbPool_Change()
    UpdatePreview
End Sub

Private Sub lbOrder_Change()
    UpdatePreview
End Sub

Private Sub MoveSelected(fromList As MSForms.ListBox, toList As MSForms.ListBox, Optional removeFromSource As Boolean = True)
    Dim i As Long
    For i = fromList.ListCount - 1 To 0 Step -1
        If fromList.Selected(i) Then
            toList.AddItem fromList.List(i)
            If removeFromSource Then fromList.RemoveItem i
        End If
    Next i
End Sub

Private Function JoinListBox(lb As MSForms.ListBox) As String
    Dim i As Long
    Dim arr() As String
    If lb.ListCount = 0 Then
        JoinListBox = ""
        Exit Function
    End If
    ReDim arr(0 To lb.ListCount - 1)
    For i = 0 To lb.ListCount - 1
        arr(i) = lb.List(i)
    Next i
    JoinListBox = Join(arr, " ")
End Function

Private Sub UpdatePreview()
    lblPreviewSrc.Caption = JoinListBox(lbPool)
    lblPreview.Caption = JoinListBox(lbOrder)
End Sub
