Imports ggcAppDriver
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text.RegularExpressions

Public Class frmReceipt
    Private p_oApp As GRider
    Private pnLoadx As Integer
    Private poControl As Control
    Private p_nButton As Integer
    Private p_bOnSeek As Boolean

    Private p_sBankIDxx As String
    Private p_nTotalAmt As Decimal
    Private p_cEntryTyp As String = "0"

    Private p_sClientID As String
    Private p_sTermIDxx As String

    'Property ShowMessage()
    Public WriteOnly Property AppDriver() As ggcAppDriver.GRider
        Set(ByVal value As ggcAppDriver.GRider)
            p_oApp = value
        End Set
    End Property

    Public WriteOnly Property TranTotal() As Decimal
        Set(ByVal value As Decimal)
            p_nTotalAmt = value
        End Set
    End Property

    Public Property BankID() As String
        Get
            Return p_sBankIDxx
        End Get
        Set(ByVal value As String)
            p_sBankIDxx = value
        End Set
    End Property

    Public Property Text_ORNo() As String
        Get
            Return txtField00.Text
        End Get
        Set(ByVal value As String)
            txtField00.Text = value
        End Set
    End Property

    Public WriteOnly Property Text_AcctNmbr() As String
        Set(ByVal value As String)
            txtField01.Text = value
        End Set
    End Property

    Public WriteOnly Property Text_ClientNm() As String
        Set(ByVal value As String)
            txtField80.Text = value
        End Set
    End Property

    Public WriteOnly Property Text_Addressx() As String
        Set(ByVal value As String)
            txtField81.Text = value
        End Set
    End Property

    Public Property Text_CashAmount() As String
        Get
            Return txtField02.Text
        End Get
        Set(ByVal value As String)
            txtField02.Text = Format(CDec(value), "#,##0.00")
        End Set
    End Property

    Public Property Text_PRNo() As String
        Get
            Return txtField03.Text
        End Get
        Set(ByVal value As String)
            txtField03.Text = value
        End Set
    End Property

    Public Property Text_BankName() As String
        Get
            Return txtField82.Text
        End Get
        Set(ByVal value As String)
            txtField82.Text = value
        End Set
    End Property

    Public Property Text_CheckNo() As String
        Get
            Return txtField04.Text
        End Get
        Set(ByVal value As String)
            txtField04.Text = value
        End Set
    End Property

    Public Property Text_BnkActNo() As String
        Get
            Return txtField05.Text
        End Get
        Set(ByVal value As String)
            txtField05.Text = value
        End Set
    End Property

    Public Property Text_CheckDate() As String
        Get
            Return txtField06.Text
        End Get
        Set(ByVal value As String)
            txtField06.Text = value
        End Set
    End Property

    Public Property Text_CheckAmount() As String
        Get
            Return CDec(txtField07.Text)
        End Get
        Set(ByVal value As String)
            txtField07.Text = Format(CDec(value), "#,##0.00")
        End Set
    End Property

    Public Property ClientID() As String
        Get
            Return p_sClientID
        End Get
        Set(ByVal value As String)
            p_sClientID = value
        End Set
    End Property

    Public Property TermCode() As String
        Get
            Return p_sTermIDxx
        End Get
        Set(ByVal value As String)
            p_sTermIDxx = value
        End Set
    End Property

    Public Property Text_EPClientNm() As String
        Get
            Return txtField51.Text
        End Get
        Set(ByVal value As String)
            txtField51.Text = value
        End Set
    End Property

    Public Property Text_EPReferNo() As String
        Get
            Return txtField52.Text
        End Get
        Set(ByVal value As String)
            txtField52.Text = value
        End Set
    End Property

    Public Property Text_EPTermNm() As String
        Get
            Return txtField53.Text
        End Get
        Set(ByVal value As String)
            txtField53.Text = value
        End Set
    End Property

    Public Property Text_EPRemarks() As String
        Get
            Return txtField54.Text
        End Get
        Set(ByVal value As String)
            txtField54.Text = value
        End Set
    End Property

    Public Property Text_EPAyAmt() As String
        Get
            Return txtField55.Text
        End Get
        Set(ByVal value As String)
            txtField55.Text = value
        End Set
    End Property

    Public ReadOnly Property Cancelled() As Boolean
        Get
            Return p_nButton <> 1
        End Get
    End Property

    Public Property EntryType() As String
        Get
            Return p_cEntryTyp
        End Get
        Set(ByVal value As String)
            p_cEntryTyp = value
        End Set
    End Property

    Private Sub frmReceipt_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Debug.Print("frmReceipt_Activated")
        If pnLoadx = 1 Then

            If p_cEntryTyp = "0" Then
                Label5.Text = "O.R. No"
            Else
                Label5.Text = "P.R. No"
            End If

            'Change display based on Reference Label
            If Label5.Text = "O.R. No" Then
                'Do not allow update of Check Group
                GroupBox1.Enabled = False
                GroupBox2.Enabled = True
                GroupBox3.Enabled = True
                'txtField03.ReadOnly = True
                'txtField04.ReadOnly = True
                'txtField05.ReadOnly = True
                'txtField06.ReadOnly = True
                'txtField07.ReadOnly = True
                'txtField82.ReadOnly = True

                txtField08.Text = txtField02.Text
            Else
                GroupBox1.Enabled = True
                GroupBox2.Enabled = False
                GroupBox3.Enabled = False

                Label15.Visible = False
                txtField03.Visible = False

                txtField07.Text = txtField02.Text
                txtField08.Text = txtField07.Text
                txtField02.Text = Format(0.0, "#,##0.00")
            End If

            txtField00.Focus()
            pnLoadx = 2
        End If
    End Sub

    Private Sub frmReceipt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Debug.Print("frmReceipt_Load")
        If pnLoadx = 0 Then

            'Set event Handler for txtField
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "GotFocus", AddressOf txtField_GotFocus)
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "LostFocus", AddressOf txtField_LostFocus)
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
                    Case 6
                        If Not IsDate(loTxt.Text) Then
                            loTxt.Text = p_oApp.getSysDate
                        End If
                        loTxt.Text = Format(loTxt.Text, "yyyy/MM/dd")
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
                Select Case loIndex
                    Case 6
                        If Not IsDate(loTxt.Text) Then
                            loTxt.Text = p_oApp.getSysDate
                        End If
                        loTxt.Text = Format(Convert.ToDateTime(loTxt.Text), xsDATE_LONG)
                    Case 2, 7, 55
                        Dim lnValue As Decimal
                        If Not IsNumeric(loTxt.Text) Then
                            loTxt.Text = 0
                        End If
                        lnValue = Convert.ToDecimal(loTxt.Text)
                        loTxt.Text = Format(lnValue, "#,##0.00")

                        If p_cEntryTyp = "0" Then
                            If loIndex = 2 Then
                                If lnValue = 0.0# Then
                                    txtField55.Text = Format(p_nTotalAmt, "#,##0.00")
                                End If
                            End If
                        End If

                        lnValue = 0
                        If IsNumeric(txtField02.Text) Then
                            lnValue = lnValue + Convert.ToDecimal(txtField02.Text)
                        End If
                        If IsNumeric(txtField07.Text) Then
                            lnValue = lnValue + Convert.ToDecimal(txtField07.Text)
                        End If
                        If IsNumeric(txtField55.Text) Then
                            lnValue = lnValue + Convert.ToDecimal(txtField55.Text)
                        End If
                        txtField08.Text = Format(lnValue, "#,##0.00")
                End Select

                loTxt.BackColor = SystemColors.Window
                'poControl = Nothing
            End If
        End If
    End Sub

    Private Sub txtField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)
            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtField" Then
                Select Case loIndex
                    Case 51
                        Call SearchClient(loTxt.Text & "%", True)
                    Case 53
                        Call SearchTerm(loTxt.Text & "%", True)
                    Case 82
                        Call SearchBank(loTxt.Text & "%", True)
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
                    p_nButton = 1
                    Me.Hide()
                End If
            Case 2 ' Cancel Update
                p_nButton = 2
                Me.Hide()
        End Select
    End Sub

    Private Sub SearchTerm(ByVal fsValue As String, ByVal fbIsSrch As Boolean)
        Dim lsSQL As String

        If fsValue = txtField53.Tag Then Exit Sub

        lsSQL = "SELECT" & _
                        "  sTermIDxx" & _
                        ", sTermName" & _
                " FROM Term" & _
                " WHERE cRecdStat = '1'"

        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                                , lsSQL _
                                                , True _
                                                , fsValue _
                                                , "sTermIDxx»sTermName" _
                                                , "ID»Name", _
                                                , "sTermIDxx»sTermName" _
                                                , 1)
            If IsNothing(loRow) Then
                txtField53.Text = ""
                txtField53.Tag = ""
                p_sTermIDxx = ""
            Else
                txtField53.Text = loRow("sTermName")
                txtField53.Tag = loRow("sTermName")
                p_sTermIDxx = loRow("sTermIDxx")
            End If

            Exit Sub
        End If
    End Sub

    Private Sub SearchClient(ByVal fsValue As String, ByVal fbIsSrch As Boolean)
        Dim lsSQL As String

        If fsValue = txtField51.Tag Then Exit Sub

        lsSQL = "SELECT" & _
                        "  a.sClientID" & _
                        ", a.sCompnyNm" & _
                " FROM Client_Master a" & _
                    ", Payment_Processor b" & _
                " WHERE a.sClientID = b.sClientID" & _
                    " AND b.cRecdStat = '1'"

        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                                , lsSQL _
                                                , True _
                                                , fsValue _
                                                , "sClientID»sCompnyNm" _
                                                , "ID»Name", _
                                                , "a.sClientID»a.sCompnyNm" _
                                                , 1)
            If IsNothing(loRow) Then
                txtField51.Text = ""
                txtField51.Tag = ""
                p_sClientID = ""
            Else
                txtField51.Text = loRow("sCompnyNm")
                txtField51.Tag = loRow("sCompnyNm")
                p_sClientID = loRow("sClientID")
            End If

            Exit Sub
        End If
    End Sub

    Private Sub SearchBank(ByVal fsValue As String, ByVal fbIsSrch As Boolean)
        Dim lsSQL As String

        If fsValue = txtField82.Tag Then Exit Sub

        lsSQL = "SELECT" & _
                       "  a.sBankIDxx" & _
                       ", a.sBankName" & _
              " FROM Banks a"

        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sBankIDxx»sBankName" _
                                             , "ID»Acct Name»Bank", _
                                             , "a.sBankIDxx»a.sBankName" _
                                             , 1)
            If IsNothing(loRow) Then
                txtField82.Text = ""
                txtField82.Tag = ""
                p_sBankIDxx = ""
            Else
                txtField82.Text = loRow("sBankName")
                txtField82.Tag = loRow("sBankName")
                p_sBankIDxx = loRow("sBankIDxx")
            End If

            Exit Sub
        End If
    End Sub

    Private Function isEntryOk() As Boolean
        Dim reGex As Regex = New Regex("[^a-zA-Z0-9]", RegexOptions.IgnoreCase)
        Dim lnCash As Decimal = 0
        Dim lnEPay As Decimal = 0

        If IsNumeric(txtField02.Text) Then
            lnCash = Convert.ToDecimal(txtField02.Text)
        End If

        If IsNumeric(txtField55.Text) Then
            lnEPay = Convert.ToDecimal(txtField55.Text)
        End If

        'Check cash payment part if cash payment is detected
        If lnCash + lnEPay > 0 Then
            If txtField00.Text = "" Then
                MsgBox("Invalid Receipt No detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
                Return False
            End If
        End If

        If reGex.IsMatch(txtField00.Text) Then
            MsgBox("Invalid Receipt No detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
            txtField00.Focus()
            Return False
        End If

        If lnEPay > 0 Then
            If p_sClientID = "" Then
                MsgBox("Invalid Payment Processor detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
                Return False
            End If

            If txtField52.Text = "" Then
                MsgBox("Invalid E-Pay Reference No. detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
                Return False
            End If
        End If

        Dim lnCheck As Decimal = 0
        If IsNumeric(txtField07.Text) Then
            lnCheck = Convert.ToDecimal(txtField07.Text)
        End If

        'Check check payment part if check payment is detected
        If lnCheck > 0 Then
            If p_sBankIDxx = "" Then
                MsgBox("Invalid Bank detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
                Return False
            End If

            If txtField04.Text = "" Then
                MsgBox("Invalid check no detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
                Return False
            End If

            If txtField05.Text = "" Then
                MsgBox("Invalid account no detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
                Return False
            End If

            If Not IsDate(txtField06.Text) Then
                MsgBox("Invalid check date detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
                Return False
            End If
        End If

        'Make sure that total payment collected is equivalent to transaction amount...
        If (lnCheck + lnCash + lnEPay) <> p_nTotalAmt Then
            MsgBox("Amount collected is different from transaction amount!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Payment Info")
            Return False
        End If

        Return True
    End Function
End Class