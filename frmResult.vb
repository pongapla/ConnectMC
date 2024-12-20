Public Class frmResult

    Sub New(ByVal Result As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'Dim r() As String

        'r = Result.Split(" ")

        'For i As Integer = 0 To r.GetUpperBound(0) - 1
        '    Me.TextBox1.Text &= System.Convert.ToChar(Convert.ToUInt32(r(i)))
        'Next

        'For i As Integer = 0 To Result.Length - 1
        '    If i Mod 2 = 0 Then
        '        Me.TextBox1.Text &= System.Convert.ToChar(Convert.ToUInt32(Result.Substring(i, 2)))
        '    End If
        '    'Me.TextBox1.Text &= Chr(Asc(Result.Substring(i, 1))) & " "
        'Next

        Me.TextBox1.Text = Result
        'Me.TextBox1.Text = Data_Hex_Asc(Result)
    End Sub

    Public Function Data_Hex_Asc(ByRef Data As String) As String
        Dim Data1 As String = ""
        Dim sData As String = ""
        While Data.Length > 0
            'first take two hex value using substring.
            'then convert Hex value into ascii.
            'then convert ascii value into character.
            If Data <> " " Then
                If Data <> "" Then
                    Data1 = System.Convert.ToChar(System.Convert.ToUInt32(Data.Substring(0, 2), 16)).ToString()
                    sData = sData & Data1
                    Data = Data.Substring(2, Data.Length - 2)
                End If
                
            Else
                Data = ""
            End If
        End While
        Return sData
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmResult_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1.Focus()
    End Sub
End Class