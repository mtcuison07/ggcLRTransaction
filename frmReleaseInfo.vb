Imports ggcAppDriver
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmReleaseInfo
    Private p_oApp As GRider
    Private pnLoadx As Integer
    Private poControl As Control
    Private p_nButton As Integer
    Private p_bOnSeek As Boolean

    Private p_sBnkActID As String

    'Property ShowMessage()
    Public WriteOnly Property AppDriver() As ggcAppDriver.GRider
        Set(ByVal value As ggcAppDriver.GRider)
            p_oApp = value
        End Set
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

    Public ReadOnly Property Cancelled() As Boolean
        Get
            Return p_nButton <> 1
        End Get
    End Property

    Private Sub frmReleaseInfo_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Debug.Print("frmReleaseInfo_Activated")
        If pnLoadx = 1 Then
            txtField01.Focus()
            pnLoadx = 2
        End If
    End Sub

    Private Sub txtField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then

            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)
            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If loTxt.Name = "txtField82" Then
                Call SearchAccount(loTxt.Text, True)
            End If

            'If Mid(loTxt.Name, 1, 8) = "txtField" Then
            '    Select Case loIndex
            '        Case 82
            '            Call SearchAccount(loTxt.Text, True)
            '    End Select
            'End If

            If TypeOf poControl Is TextBox Then
                SelectNextControl(loTxt, True, True, True, True)
            End If
        End If
    End Sub

    Private Sub frmReleaseInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Debug.Print("frmReleaseInfo_Load")
        If pnLoadx = 0 Then

            p_sBnkActID = ""

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
                    Case 4
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
                    Case 4
                        If Not IsDate(loTxt.Text) Then
                            loTxt.Text = p_oApp.getSysDate
                        End If
                        loTxt.Text = Format(CDate(loTxt.Text), "MMMM dd, yyyy")
                    Case 82 ' Account No
                        If Trim(loTxt.Text) <> "" Then
                            Call SearchAccount(loTxt.Text, False)
                        End If
                End Select

                loTxt.BackColor = SystemColors.Window
                'poControl = Nothing
            End If
        End If
    End Sub

    'Private Sub txtField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
    '        Dim loTxt As TextBox
    '        loTxt = CType(sender, System.Windows.Forms.TextBox)
    '        Dim loIndex As Integer
    '        loIndex = Val(Mid(loTxt.Name, 9))

    '        If Mid(loTxt.Name, 1, 8) = "txtField" Then
    '            Select Case loIndex
    '                Case 82
    '                    Call SearchAccount(loTxt.Text, True)
    '            End Select
    '        End If

    '        If TypeOf poControl Is TextBox Then
    '            SelectNextControl(loTxt, True, True, True, True)
    '        End If
    '    End If
    'End Sub

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

    Private Sub SearchAccount(ByVal fsValue As String, ByVal fbIsSrch As Boolean)
        Dim lsSQL As String

        lsSQL = "SELECT" & _
                       "  a.sActNumbr" & _
                       ", a.sActNamex" & _
                       ", b.sBankName" & _
                       ", a.sBnkActID" & _
              " FROM Bank_Account a" & _
                " LEFT JOIN Banks b ON a.sBankIDxx = b.sBankIDxx"

        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sActNumbr»sActNamex»sBankName" _
                                             , "Acct No»Acct Name»Bank", _
                                             , "a.sActNumbr»a.sActNamex»b.sBankName" _
                                             , 0)
            If IsNothing(loRow) Then
                txtField82.Text = ""
                txtField82.Tag = ""
                txtField83.Text = ""
                txtField84.Text = ""
                p_sBnkActID = ""
            Else
                txtField82.Text = loRow("sActNumbr")
                txtField82.Tag = loRow("sActNumbr")
                txtField83.Text = loRow("sActNamex")
                txtField84.Text = loRow("sBankName")
                p_sBnkActID = loRow("sBnkActID")
            End If

            Exit Sub
        End If
    End Sub

    Private Function isEntryOk() As Boolean
        If txtField01.Text = "" Then
            MsgBox("Invalid voucher no detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Release Info")
            Return False
        End If

        If p_sBnkActID <> "" Then
            If txtField03.Text = "" Then
                MsgBox("Invalid check no detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Release Info")
                Return False
            End If

            If txtField04.Text = "" Then
                MsgBox("Invalid check date detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Release Info")
                Return False
            End If
        End If

        Return True
    End Function

End Class