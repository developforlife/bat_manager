Public Class Form1

    'Общие переменные
    Dim sAppPath As String = AppDomain.CurrentDomain.BaseDirectory ' Путь до самой программы
    Dim sAppPathINI As String = sAppPath & "\" & "config.ini" ' Путь до INI-файла

    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal ipAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer ' Чтение INI-файлов
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Section As String, ByVal Key As String, ByVal putStr As String, ByVal INIfile As String) As Integer ' Запись в INI-файл

    Private Function ReadINI(ByVal sSection As String, ByVal sKey As String, ByVal sIniFileName As String) ' Функция чтения INI-файла
        Dim nLength As Integer
        Dim sTemp As String
        Dim lsTemp As Integer
        sTemp = Space(255)
        lsTemp = Len(sTemp)
        nLength = GetPrivateProfileString(sSection, sKey, "", sTemp, lsTemp, sIniFileName)
        Return Microsoft.VisualBasic.Left(sTemp, nLength)

    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
Mark1:
        If IO.Directory.Exists(ReadINI("general_settings", "folder_in", sAppPathINI)) = False Then

                    Dim mess = MsgBox("Директория не доступна!" & vbCrLf & vbCrLf & "Проверьте доступность папки, и нажмите OK.", 16 + 1, "Ошибка! Директория не доступна.") 'значения цифр прилагаю
            If mess = DialogResult.OK Then
                'Process.Start(ReadINI("general_settings", "network_bat", sAppPathINI))
                GoTo Mark1
            Else
                End
            End If
        End If



        Dim iBatCount As Integer = ReadINI("general_settings", "bat_count", sAppPathINI)

        Dim iStep As Integer = 0
        For i As Integer = 0 To iBatCount
            Dim sPath As String = ReadINI("bat_" & i, "bat_path_" & i, sAppPathINI)
            Dim sText As String = ReadINI("bat_" & i, "button_text_" & i, sAppPathINI)
            Dim sToolTip As String = ReadINI("bat_" & i, "button_tool_tip_" & i, sAppPathINI)

            If sPath <> "" And sText <> "" And sToolTip <> "" Then
                Dim Btn As New Button With {.Name = "Button" & i, .Text = ReadINI("bat_" & i, "button_text_" & i, sAppPathINI), .Font = New Font("Microsoft Sans Serif", 10, FontStyle.Bold), .Size = New System.Drawing.Size(300, 40), .Location = New Point(8, 20 + iStep)}
                AddHandler Btn.Click, AddressOf Btn_Click 'Привязываем кнопку к событию
                AddHandler Btn.MouseDown, AddressOf Btn_MouseDown 'Привязываем кнопку к событию
                Me.Controls.Add(Btn)

                iStep += 60
                If i >= iBatCount Then
                    Me.Height += iStep
                End If

            Else
                Me.Height += iStep

            End If
        Next

        For Each cont In Me.Controls
            Dim contName() As String = Split(cont.Name, "n")
            Me.ToolTip1.SetToolTip(cont, ReadINI("bat_" & contName(1), "button_tool_tip_" & contName(1), sAppPathINI))
        Next

    End Sub

    Sub Btn_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)  ' По правому клику мыши по кнопке, открывает проводник с bat-файлом, и выделяет его курсором.
        Dim sSenderNameNumber() As String = Split(sender.Name, "n")
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Process.Start("explorer", "/select, """ & ReadINI("bat_" & sSenderNameNumber(1), "bat_path_" & sSenderNameNumber(1), sAppPathINI) & """")
        End If
    End Sub

    Sub Btn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSenderNameNumber() As String = Split(sender.Name, "n")

        Try
            Process.Start(ReadINI("bat_" & sSenderNameNumber(1), "bat_path_" & sSenderNameNumber(1), sAppPathINI))

            If ReadINI("bat_" & sSenderNameNumber(1), "button_msgbox_" & sSenderNameNumber(1), sAppPathINI) <> "" Then
                MsgBox(ReadINI("bat_" & sSenderNameNumber(1), "button_msgbox_" & sSenderNameNumber(1), sAppPathINI), MsgBoxStyle.Information)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

End Class
