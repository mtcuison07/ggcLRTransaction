'Note:
'    Example and Usage of Overriding ProcessCmdKey
'       See the function [Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean]

Imports ggcAppDriver
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmLRMaster
    Private WithEvents p_oLoanMstr As LRMaster
    Private pnLoadx As Integer
    Private poControl As Control
    Private p_nButton As Integer
    Private p_bOnSeek As Boolean
    Private p_sReferNox As String
    Private p_sBnkActID As String
    Private p_sCheckNox As String
    Private p_dCheckDte As String

    'Property ShowMessage()
    Public WriteOnly Property LoanObject() As LRMaster
        Set(ByVal value As LRMaster)
            p_oLoanMstr = value
        End Set
    End Property

    Public ReadOnly Property Cancelled() As Boolean
        Get
            Return p_nButton <> 1
        End Get
    End Property

    Public ReadOnly Property ReferNo() As String
        Get
            If p_nButton = 1 Then
                Return p_sReferNox
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property BankID() As String
        Get
            If p_nButton = 1 Then
                Return p_sBnkActID
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property CheckNo() As String
        Get
            If p_nButton = 1 Then
                Return p_sCheckNox
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property CheckDate() As String
        Get
            If p_nButton = 1 Then
                Return p_dCheckDte
            Else
                Return ""
            End If
        End Get
    End Property

    Private Sub frmLRMaster_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Debug.Print("frmLRMaster_Activated")
        If pnLoadx = 1 Then

            Call loadMaster(Me)

            txtField07.Focus()
            pnLoadx = 2
            txtField08.ReadOnly = True
        End If
    End Sub

    Private Sub frmLRMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Debug.Print("frmLRMaster_Load")
        If pnLoadx = 0 Then

            'Set event Handler for txtField
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "GotFocus", AddressOf txtField_GotFocus)
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "LostFocus", AddressOf txtField_LostFocus)
            Call grpCancelHandler(Me, GetType(TextBox), "txtField", "Validating", AddressOf txtField_Validating)
            Call grpKeyHandler(Me, GetType(TextBox), "txtField", "KeyDown", AddressOf txtField_KeyDown)

            Call grpEventHandler(Me, GetType(Button), "cmdButtn", "Click", AddressOf cmdButton_Click)
            pnLoadx = 1
        End If
    End Sub

    'Handles GotFocus Events for txtField & txtItems
    Private Sub txtField_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loIndex As Integer
        loIndex = Val(Mid(sender.Name, 9))
        Console.WriteLine("»Got Focus: " & sender.Name)

        If Mid(sender.Name, 1, 8) = "txtField" Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)
                If Not loTxt.ReadOnly Then
                    Select Case loIndex
                        Case 7, 16, 18
                            loTxt.Text = Format(p_oLoanMstr.Master(loIndex), "yyyy/MM/dd")
                    End Select

                    loTxt.BackColor = Color.Azure
                    loTxt.SelectAll()
                End If

                poControl = loTxt
            End If
    End Sub

    'Handles LostFocus Events for txtField & txtItems
    Private Sub txtField_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)

        Console.WriteLine("Lost Focus: " & sender.Name)

        Dim loIndex As Integer
        loIndex = Val(Mid(sender.Name, 9))

        If Mid(sender.Name, 1, 8) = "txtField" Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)

            If Not loTxt.ReadOnly Then
                p_oLoanMstr.Master(loIndex) = loTxt.Text
                Select Case loIndex
                    Case 7, 16, 18
                        loTxt.Text = Format(p_oLoanMstr.Master(loIndex), "MMMM dd, yyyy")
                    Case 8, 9, 10, 11, 12, 13, 14 ', 19
                        If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                            loTxt.Text = ""
                        Else
                            loTxt.Text = Format(p_oLoanMstr.Master(loIndex), xsDECIMAL)
                        End If

                        If loIndex = 19 Then
                            If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                                txtOther01.Text = "0.00"
                                txtOther02.Text = "0.00"
                            Else
                                txtOther01.Text = Format(p_oLoanMstr.Master("nIntTotal") / p_oLoanMstr.Master("nAcctTerm"), xsDECIMAL)
                                txtOther02.Text = Format(CDec(txtField19.Text) + CDec(txtOther01.Text), xsDECIMAL)
                            End If
                        End If

                    Case 17, 20, 15 'Term, Penalty Rate, Interest Rate
                        If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                            loTxt.Text = ""
                        Else
                            loTxt.Text = Format(p_oLoanMstr.Master(loIndex), xsDECIMAL)
                        End If

                        'kalyptus - 2017.02.17 11:09am
                        'For some unknown reason(i can not trace the reason), 
                        'the txtField19 is assigned with the nIntTotal. To solve the problem
                        'I add this line to reload the value of txtField(19)...
                        txtField19.Text = Format(p_oLoanMstr.Master(19), xsDECIMAL)

                        If loIndex = 17 Or loIndex = 15 Then
                            If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                                txtOther01.Text = "0.00"
                                txtOther02.Text = "0.00"
                            Else
                                txtOther01.Text = Format(p_oLoanMstr.Master("nIntTotal") / p_oLoanMstr.Master("nAcctTerm"), xsDECIMAL)
                                txtOther02.Text = Format(CDec(txtField19.Text) + CDec(txtOther01.Text), xsDECIMAL)
                            End If
                        End If
                End Select

                loTxt.BackColor = SystemColors.Window
                'poControl = Nothing
            End If
        End If
    End Sub

    'Handles Validating Events for txtField & txtItems
    Private Sub txtField_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        'Dim loIndex As Integer
        'loIndex = Val(Mid(sender.Name, 9))
        'If Mid(sender.Name, 1, 8) = "txtField" Then
        'Dim loTxt As TextBox
        'loTxt = CType(sender, System.Windows.Forms.TextBox)
        'p_oClient.Master(loIndex) = loTxt.Text
        'ElseIf Mid(sender.Name, 1, 8) = "cmbField" Then
        'Dim loCmb As ComboBox
        'loCmb = DirectCast(sender, ComboBox)
        'p_oClient.Master(loIndex) = loCmb.SelectedIndex
        'End If
    End Sub

    Private Sub txtField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)
            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtField" Then
                Select Case loIndex
                    Case 85, 83
                        p_oLoanMstr.SearchMaster(loIndex, loTxt.Text)
                End Select
            End If

            If TypeOf poControl Is TextBox Then
                SelectNextControl(loTxt, True, True, True, True)
            End If
        End If
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 1 ' Ok
                If isEntryOk() Then
                    'Show the form that will display the releasing of 
                    'amount including the Voucher and Check Info.
                    Dim loFrm As frmReleaseInfo
                    loFrm = New frmReleaseInfo
                    loFrm.AppDriver = p_oLoanMstr.AppDriver
                    loFrm.txtField00.Text = p_oLoanMstr.Master("sAcctNmbr")
                    loFrm.txtField80.Text = p_oLoanMstr.Master("sClientNm")
                    loFrm.txtField81.Text = p_oLoanMstr.Master("sAddressX")
                    loFrm.txtField02.Text = p_oLoanMstr.Master("nTakeHome")
                    loFrm.ShowDialog()

                    If loFrm.Cancelled Then Exit Sub

                    p_sBnkActID = loFrm.BankID
                    p_sReferNox = loFrm.txtField01.Text
                    p_sCheckNox = loFrm.txtField03.Text
                    p_dCheckDte = loFrm.txtField04.Text

                    loFrm = Nothing

                    p_nButton = 1
                    Me.Hide()
                End If
            Case 2 ' Cancel Update
                p_nButton = 2
                Me.Hide()
        End Select
    End Sub

    Private Sub loadMaster(ByVal loControl As Control)
        Dim loTxt As Control

        For Each loTxt In loControl.Controls
            If loTxt.HasChildren Then
                Call loadMaster(loTxt)
            Else
                If (TypeOf loTxt Is TextBox) Then
                    Dim loIndex As Integer
                    loIndex = Val(Mid(loTxt.Name, 9))
                    Dim loBox As TextBox
                    loBox = CType(loTxt, TextBox)
                    If LCase(Mid(loBox.Name, 1, 8)) = "txtfield" Then
                        If p_oLoanMstr.EditMode <> xeEditMode.MODE_UNKNOWN Then
                            Select Case loIndex
                                Case 7, 16, 18
                                    If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                                        loTxt.Text = ""
                                    Else
                                        loTxt.Text = Format(p_oLoanMstr.Master(loIndex), "MMMM dd, yyyy")
                                    End If
                                Case 8, 9, 10, 11, 12, 13, 14, 19, 36

                                    If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                                        loTxt.Text = ""
                                    Else
                                        loTxt.Text = Format(p_oLoanMstr.Master(loIndex), xsDECIMAL)
                                    End If

                                    If loIndex = 19 Then
                                        Dim lsValue As String = p_oLoanMstr.Master(loIndex)
                                        If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                                            txtOther01.Text = "0.00"
                                            txtOther02.Text = "0.00"
                                        Else
                                            txtOther01.Text = Format(p_oLoanMstr.Master("nIntTotal") / p_oLoanMstr.Master("nAcctTerm"), xsDECIMAL)
                                            txtOther02.Text = Format(CDec(txtField19.Text) + CDec(txtOther01.Text), xsDECIMAL)
                                        End If
                                    End If

                                Case 17, 20, 15 'Term, Penalty Rate, Interest Rate
                                    If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                                        loTxt.Text = ""
                                    Else
                                        loTxt.Text = Format(p_oLoanMstr.Master(loIndex), xsDECIMAL)
                                    End If
                                Case Else
                                    If IsDBNull(p_oLoanMstr.Master(loIndex)) Then
                                        loBox.Text = ""
                                    Else
                                        loBox.Text = p_oLoanMstr.Master(loIndex)
                                    End If
                            End Select
                        Else
                            loBox.Text = ""
                        End If
                    End If 'LCase(Mid(loBox.Name, 1, 8)) = "txtfield"
                End If '(TypeOf loTxt Is TextBox)
            End If 'If loTxt.HasChildren
        Next 'loTxt In loControl.Controls
    End Sub

    Private Function isEntryOk() As Boolean
        If p_oLoanMstr.Master("nPrincipl") <= 1000 Then
            MsgBox("Invalid principal detected!", vbOKOnly, "LR Master")
            Return False
        ElseIf p_oLoanMstr.Master("nAcctTerm") < 1 Then
            MsgBox("Invalid Term detected...", vbOKOnly, "LR Master")
            Return False
        ElseIf p_oLoanMstr.Master("sCompnyID") = "" Then
            MsgBox("Invalid Company detected...", vbOKOnly, "LR Master")
            Return False
        End If

        Return True
    End Function

    Private Sub p_oLoanMstr_MasterRetrieved(Index As Integer, Value As Object) Handles p_oLoanMstr.MasterRetrieved
        Dim loTxt As TextBox
        'Find TextBox with specified name
        loTxt = CType(FindTextBox(Me, "txtField" & Format(Index, "00")), TextBox)

        Select Case Index
            Case 85, 83
                If IsDBNull(Value) Then
                    loTxt.Text = ""
                Else
                    loTxt.Text = Value
                End If
            Case 16, 18, 7
                loTxt.Text = Format(Value, "MMMM dd, yyyy")
            Case Else
                If IsDBNull(Value) Then
                    loTxt.Text = ""
                Else
                    loTxt.Text = Format(Value, xsDECIMAL)
                End If

                If Index = 19 Then
                    If IsDBNull(Value) Then
                        txtOther01.Text = "0.00"
                        txtOther02.Text = "0.00"
                    Else
                        txtOther01.Text = Format(p_oLoanMstr.Master("nIntTotal") / p_oLoanMstr.Master("nAcctTerm"), xsDECIMAL)
                        txtOther02.Text = Format(CDec(txtField19.Text) + CDec(txtOther01.Text), xsDECIMAL)
                    End If
                End If
        End Select
    End Sub

End Class