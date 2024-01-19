Imports ggcAppDriver
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmPaymentInfo
    Private p_oApp As GRider
    Private pnLoadx As Integer
    Private poControl As Control
    Private p_nButton As Integer
    Private p_bOnSeek As Boolean

    Private p_sBankIDxx As String

    'Property ShowMessage()
    Public WriteOnly Property AppDriver() As ggcAppDriver.GRider
        Set(ByVal value As ggcAppDriver.GRider)
            p_oApp = value
        End Set
    End Property

    Public WriteOnly Property Text_TransNox() As String
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

    Public Property Text_ReferNox() As String
        Get
            Return txtField02.Text
        End Get
        Set(ByVal value As String)
            txtField02.Text = value
        End Set
    End Property

    Public WriteOnly Property Text_TranAmtx() As String
        Set(ByVal value As String)
            txtField03.Text = value
        End Set
    End Property

    Public Property Text_BnkActNo() As String
        Get
            Return txtField04.Text
        End Get
        Set(ByVal value As String)
            txtField04.Text = value
        End Set
    End Property

    Public Property Text_BankName() As String
        Get
            Return txtField05.Text
        End Get
        Set(ByVal value As String)
            txtField05.Text = value
        End Set
    End Property

    Public Property Text_BankIDxx() As String
        Get
            Return p_sBankIDxx
        End Get
        Set(ByVal value As String)
            p_sBankIDxx = value
        End Set
    End Property

    Public Property Text_BankAddr() As String
        Get
            Return txtField06.Text
        End Get
        Set(ByVal value As String)
            txtField06.Text = value
        End Set
    End Property

    Public Property Text_CheckNox() As String
        Get
            Return txtField07.Text
        End Get
        Set(ByVal value As String)
            txtField07.Text = value
        End Set
    End Property

    Public Property Text_CheckDte() As String
        Get
            Return txtField08.Text
        End Get
        Set(ByVal value As String)
            txtField08.Text = value
        End Set
    End Property

    Public ReadOnly Property Cancelled() As Boolean
        Get
            Return p_nButton <> 1
        End Get
    End Property

    Private Sub frmPaymentInfo_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Debug.Print("frmPaymentInfo_Activated")
        If pnLoadx = 1 Then
            txtField02.Focus()
            pnLoadx = 2
        End If
    End Sub

    Private Sub frmPaymentInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Debug.Print("frmPaymentInfo_Load")
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
                    Case 8
                        If IsDate(loTxt.Text) Then
                            loTxt.Text = Format(loTxt.Text, "yyyy/MM/dd")
                        End If
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
                    Case 8
                        If Not IsDate(loTxt.Text) Then
                            loTxt.Text = p_oApp.getSysDate
                        End If
                        loTxt.Text = Format(CDate(loTxt.Text), xsDATE_LONG)
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
                    Case 5
                        Call SearchBank(loTxt.Text, True)
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

    Private Sub SearchBank(ByVal fsValue As String, ByVal fbIsSrch As Boolean)
        Dim lsSQL As String

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
                txtField05.Text = ""
                txtField05.Tag = ""
                p_sBankIDxx = ""
            Else
                txtField05.Text = loRow("sBankName")
                p_sBankIDxx = loRow("sBankIDxx")
            End If

            Exit Sub
        End If
    End Sub

    Private Function isEntryOk() As Boolean
        If txtField02.Text = "" Then
            MsgBox("Invalid Receipt No detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Release Info")
            Return False
        End If

        If Trim(txtField04.Text) <> "" Then
            If txtField05.Text = "" Then
                MsgBox("Invalid Bank detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Release Info")
                Return False
            End If

            If txtField06.Text = "" Then
                MsgBox("Invalid Branch of Bank detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Release Info")
                Return False
            End If

            If txtField07.Text = "" Then
                MsgBox("Invalid check no detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Release Info")
                Return False
            End If

            If txtField08.Text = "" Then
                MsgBox("Invalid check date detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Release Info")
                Return False
            End If
        End If

        Return True
    End Function
End Class