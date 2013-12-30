Public Class StocklessPNG

    '' church budget
    Private Const CBDW As String = "DoubleWide_CB"
    Private Const CBBK As String = "Booklet"
    Private Const CBUV As String = "BookletCovers"
    Private Const CBCN As String = "Cartons_CB"
    Private Const CBTN As String = "Kyocera_CB_Num10"
    Private Const CBKY As String = "Kyocera_CB"

    '' monthly mail
    Private Const MMDW As String = "DoubleWide_MM"
    Private Const MMCN As String = "Cartons_MM"
    Private Const MMTN As String = "Kyocera_MM_Num10"
    Private Const MMRT As String = "Kyocera_MM_Return"
    Private Const MMKY As String = "Kyocera_MM"


    Private Const PNGFONTPATH As String = "c:\pngfont\pngfont2.exe"


    Private Sub uiTxtFolder_KeyDown(sender As Object, e As KeyEventArgs) Handles uiTxtFolder.KeyDown

        '' only search if they pressed enter
        If Not (e.KeyCode = Keys.Return) Then
            Exit Sub
        End If

        Dim type As String = getTypeString()
        Dim folder As String = uiTxtFolder.Text.Trim

        '' open png font with the proper folder and type
        Dim cmd As String = PNGFONTPATH & " /o=" & folder.ToUpper & " /t=" & type
        Call Shell(cmd, AppWinStyle.MaximizedFocus)

    End Sub

    Private Function getTypeString() As String
        If uiRdioCBDW.Checked Then
            Return CBDW
        ElseIf uiRdioCBBK.Checked Then
            Return CBBK
        ElseIf uiRdioCBUV.Checked Then
            Return CBUV
        ElseIf uiRdioCBCN.Checked Then
            Return CBCN
        ElseIf uiRdioCBTN.Checked Then
            Return CBTN
        ElseIf uiRdioCBKY.Checked Then
            Return CBKY

        ElseIf uiRdioMMDW.Checked Then
            Return MMDW
        ElseIf uiRdioMMCN.Checked Then
            Return MMCN
        ElseIf uiRdioMMTN.Checked Then
            Return MMTN
        ElseIf uiRdioMMRT.Checked Then
            Return MMRT
        Else
            Return MMKY
        End If
    End Function
End Class
