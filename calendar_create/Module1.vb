Imports MySql.Data.MySqlClient

Module Module1

    Public CuDr As String       ''exeの動くカレントディレクトリを格納
    Public mysqlCon As New MySqlConnection
    Public sqlCommand As New MySqlCommand

    Sub sql_st()
        ''データベースに接続

        Dim Builder = New MySqlConnectionStringBuilder()
        ' データベースに接続するために必要な情報をBuilderに与える。データベース情報はGitに乗せないこと。
        Builder.Server = ""
        Builder.Port =
        Builder.UserID = ""
        Builder.Password = ""
        Builder.Database = ""
        Dim ConStr = Builder.ToString()

        mysqlCon.ConnectionString = ConStr
        mysqlCon.Open()

    End Sub

    Sub sql_cl()
        ' データベースの切断
        mysqlCon.Close()
    End Sub

    Function sql_result_return(ByVal query As String) As DataTable
        ''データセットを返すSELECT系のSQLを処理するコード

        Dim dt As New DataTable()

        Try
            ' 4.データ取得のためのアダプタの設定
            Dim Adapter = New MySqlDataAdapter(query, mysqlCon)

            ' 5.データを取得
            Dim Ds As New DataSet
            Adapter.Fill(dt)

            Return dt
        Catch ex As Exception

            Return dt
        End Try

    End Function

    Function sql_result_no(ByVal query As String)
        ''データセットを返さない、DELETE、UPDATE、INSERT系のSQLを処理するコード

        Try
            sqlCommand.Connection = mysqlCon
            sqlCommand.CommandText = query
            sqlCommand.ExecuteNonQuery()

            Return "Complete"
        Catch ex As Exception

            Return ex.Message
        End Try

    End Function

    Sub close_save()
        ''設定用ファイルの保存

        Dim dtx1 As String = ""

        For lp1 As Integer = 1 To 21
            Dim tbxn1 As String = "Cf_TextBox" & lp1.ToString

            Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
            If cs.Length > 0 Then
                dtx1 &= CType(cs(0), TextBox).Text
                dtx1 &= vbCrLf
            End If
        Next


        Dim stCurrentDir As String = System.IO.Directory.GetCurrentDirectory()
        CuDr = stCurrentDir

        Dim excsv1 As IO.StreamWriter
        excsv1 = New IO.StreamWriter(CuDr & "\config.ini", False, System.Text.Encoding.GetEncoding("shift_jis"))
        excsv1.Write(dtx1)
        excsv1.Close()
        excsv1.Dispose()

    End Sub

    Sub fil_ftp_up(ByVal upFile As String, ByVal upUrl As String, ByVal uid As String, ByVal ups As String)

        'アップロード先のURI
        Dim u As New Uri(upUrl)

        'FtpWebRequestの作成
        Dim ftpReq As System.Net.FtpWebRequest =
            CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
        'ログインユーザー名とパスワードを設定
        ftpReq.Credentials = New System.Net.NetworkCredential(uid, ups)
        'MethodにWebRequestMethods.Ftp.UploadFile("STOR")を設定
        ftpReq.Method = System.Net.WebRequestMethods.Ftp.UploadFile
        '要求の完了後に接続を閉じる
        ftpReq.KeepAlive = False
        'ASCIIモードで転送する
        ftpReq.UseBinary = False
        'PASVモードを無効にする
        ftpReq.UsePassive = True

        'ファイルをアップロードするためのStreamを取得
        Dim reqStrm As System.IO.Stream = ftpReq.GetRequestStream()
        'アップロードするファイルを開く
        Dim fs As New System.IO.FileStream(
            upFile, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        'アップロードStreamに書き込む
        Dim buffer(1023) As Byte
        While True
            Dim readSize As Integer = fs.Read(buffer, 0, buffer.Length)
            If readSize = 0 Then
                Exit While
            End If
            reqStrm.Write(buffer, 0, readSize)
        End While
        fs.Close()
        reqStrm.Close()

        'FtpWebResponseを取得
        Dim ftpRes As System.Net.FtpWebResponse =
            CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
        'FTPサーバーから送信されたステータスを表示
        Console.WriteLine("{0}: {1}", ftpRes.StatusCode, ftpRes.StatusDescription)
        '閉じる
        ftpRes.Close()
    End Sub

    Sub fil_ftps_up(ByVal upFile As String, ByVal upUrl As String, ByVal uid As String, ByVal ups As String)

        'アップロード先のURI
        Dim u As New Uri(upUrl)

        'FtpWebRequestの作成
        Dim ftpReq As System.Net.FtpWebRequest =
            CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
        'ログインユーザー名とパスワードを設定
        ftpReq.Credentials = New System.Net.NetworkCredential(uid, ups)
        'MethodにWebRequestMethods.Ftp.UploadFile("STOR")を設定
        ftpReq.Method = System.Net.WebRequestMethods.Ftp.UploadFile
        ftpReq.EnableSsl = True
        '要求の完了後に接続を閉じる
        ftpReq.KeepAlive = False
        'ASCIIモードで転送する
        ftpReq.UseBinary = False
        'PASVモードを無効にする
        ftpReq.UsePassive = True

        'ファイルをアップロードするためのStreamを取得
        Dim reqStrm As System.IO.Stream = ftpReq.GetRequestStream()
        'アップロードするファイルを開く
        Dim fs As New System.IO.FileStream(
            upFile, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        'アップロードStreamに書き込む
        Dim buffer(1023) As Byte
        While True
            Dim readSize As Integer = fs.Read(buffer, 0, buffer.Length)
            If readSize = 0 Then
                Exit While
            End If
            reqStrm.Write(buffer, 0, readSize)
        End While
        fs.Close()
        reqStrm.Close()

        'FtpWebResponseを取得
        Dim ftpRes As System.Net.FtpWebResponse =
            CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
        'FTPサーバーから送信されたステータスを表示
        Console.WriteLine("{0}: {1}", ftpRes.StatusCode, ftpRes.StatusDescription)
        '閉じる
        ftpRes.Close()
    End Sub

    Sub clc01()

        ''データベースと接続
        Call sql_st()

        ''表示の初期化
        For wn1 As Integer = 0 To 13
            For dn1 As Integer = 0 To 6
                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString

                Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    CType(cs(0), Label).Text = ""
                    CType(cs(0), Label).BackColor = Color.White
                    CType(cs(0), Label).Refresh()
                End If
            Next

        Next

        ''テキストボックスの年、月、1日で日付を合成。曜日を計算する
        Dim cday1 As New DateTime(Form1.TextBox_year1.Text, Form1.TextBox_month1.Text, 1)
        Dim cday2 As DateTime = cday1.AddMonths(1).AddDays(-1)
        Dim lped As Integer = cday2.Day

        Dim sql1 As String = "SELECT `日付`, `種別`"
        sql1 &= " FROM `no_ships_day`"
        sql1 &= " WHERE `日付` >= '" & cday1.Year.ToString("0000") & "-" & cday1.Month.ToString("00") & "-" & cday1.Day.ToString("00") & "'"
        sql1 &= " And"
        sql1 &= " `日付` <= '" & cday2.Year.ToString("0000") & "-" & cday2.Month.ToString("00") & "-" & cday2.Day.ToString("00") & "'"
        sql1 &= ";"
        Dim dTb1 As DataTable = sql_result_return(sql1)

        Dim ye1 As String = cday1.Year.ToString("0000")
        Dim mo1 As String = cday1.Month.ToString("00")

        Form1.TextBox_yearLeft.Text = cday1.Year.ToString
        Form1.TextBox_monthLeft.Text = cday1.Month.ToString

        Dim weekNumber As Integer = CInt(cday1.DayOfWeek)
        Dim dayn1 As Integer = 1

        For wn1 As Integer = 0 To 6

            For dn1 As Integer = 0 To 6

                If wn1 = 0 And dn1 = 0 Then
                    dn1 = weekNumber
                End If

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString

                Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    CType(cs(0), Label).Text = dayn1.ToString
                    Select Case dn1
                        Case 0
                            CType(cs(0), Label).BackColor = Color.Pink
                        Case 6
                            CType(cs(0), Label).BackColor = Color.Orange
                    End Select
                    CType(cs(0), Label).Refresh()
                End If

                Dim sql2 As String = "日付 = '" & ye1 & "/" & mo1 & "/" & dayn1.ToString("00") & "'"
                Dim DRow() As DataRow = dTb1.Select(sql2)

                If DRow.Count > 0 Then
                    ''この日付はデータベースに存在
                    Select Case DRow(0)(1)
                        Case 0
                            CType(cs(0), Label).BackColor = Color.Pink
                        Case 6
                            CType(cs(0), Label).BackColor = Color.Orange
                    End Select

                Else
                    ''この日付はデータベースに存在しない＝なにもしない
                End If

                dayn1 += 1

                If dayn1 > lped Then
                    GoTo loopend
                End If

            Next

        Next

loopend:

        cday1 = cday1.AddMonths(1)
        cday2 = cday1.AddMonths(1).AddDays(-1)
        lped = cday2.Day

        sql1 = "SELECT `日付`, `種別`"
        sql1 &= " FROM `no_ships_day`"
        sql1 &= " WHERE `日付` >= '" & cday1.Year.ToString("0000") & "-" & cday1.Month.ToString("00") & "-" & cday1.Day.ToString("00") & "'"
        sql1 &= " And"
        sql1 &= " `日付` <= '" & cday2.Year.ToString("0000") & "-" & cday2.Month.ToString("00") & "-" & cday2.Day.ToString("00") & "'"
        sql1 &= ";"
        dTb1 = sql_result_return(sql1)

        ye1 = cday1.Year.ToString("0000")
        mo1 = cday1.Month.ToString("00")

        Form1.TextBox_yearRight.Text = cday1.Year.ToString
        Form1.TextBox_monthRight.Text = cday1.Month.ToString

        weekNumber = CInt(cday1.DayOfWeek)
        dayn1 = 1

        For wn1 As Integer = 7 To 13

            For dn1 As Integer = 0 To 6

                If wn1 = 7 And dn1 = 0 Then
                    dn1 = weekNumber
                End If

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString

                Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    CType(cs(0), Label).Text = dayn1.ToString
                    Select Case dn1
                        Case 0
                            CType(cs(0), Label).BackColor = Color.Pink
                        Case 6
                            CType(cs(0), Label).BackColor = Color.Orange
                    End Select
                    CType(cs(0), Label).Refresh()
                End If

                Dim sql2 As String = "日付 = '" & ye1 & "/" & mo1 & "/" & dayn1.ToString("00") & "'"
                Dim DRow() As DataRow = dTb1.Select(sql2)

                If DRow.Count > 0 Then
                    ''この日付はデータベースに存在
                    Select Case DRow(0)(1)
                        Case 0
                            CType(cs(0), Label).BackColor = Color.Pink
                        Case 6
                            CType(cs(0), Label).BackColor = Color.Orange
                    End Select

                Else
                    ''この日付はデータベースに存在しない＝なにもしない
                End If

                dayn1 += 1

                If dayn1 > lped Then
                    GoTo loopend2
                End If

            Next

        Next

loopend2:

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub register_holiday()

        ''データベースと接続
        Call sql_st()

        Dim ye As String = ""
        Dim mo As String = ""
        Dim da As Integer = 0
        Dim sql1 As String = ""

        ''カレンダー上の休日をデータベースに登録
        For wn1 As Integer = 0 To 13

            If wn1 > 6 Then
                ye = Form1.TextBox_yearRight.Text
                mo = Form1.TextBox_monthRight.Text

            Else
                ye = Form1.TextBox_yearLeft.Text
                mo = Form1.TextBox_monthLeft.Text
            End If

            Dim check01 As Integer = 0
            For dn1 As Integer = 0 To 6

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString

                Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    Dim cv As String = CType(cs(0), Label).Text

                    If cv = "" Then

                    Else
                        Dim bc1 As String = CType(cs(0), Label).BackColor.ToString
                        da = Integer.Parse(CType(cs(0), Label).Text)
                        Select Case bc1
                            Case "Color [Orange]"
                                sql1 = "DELETE FROM `no_ships_day` WHERE `日付`='" & ye & "/" & mo & "/" & da.ToString("00") & "';"
                                Dim rs As String = sql_result_no(sql1)

                                If rs = "Complete" Then
                                Else
                                    MsgBox("失敗：" & sql1)

                                End If

                                sql1 = "INSERT INTO `no_ships_day`(`日付`, `種別`) VALUES ('" & ye & "/" & mo & "/" & da.ToString("00") & "',6);"
                                rs = sql_result_no(sql1)

                                If rs = "Complete" Then
                                Else
                                    MsgBox("失敗：" & sql1)
                                End If

                            Case "Color [Pink]"
                                sql1 = "DELETE FROM `no_ships_day` WHERE `日付`='" & ye & "/" & mo & "/" & da.ToString("00") & "';"
                                Dim rs As String = sql_result_no(sql1)

                                If rs = "Complete" Then
                                Else
                                    MsgBox("失敗：" & sql1)
                                End If

                                sql1 = "INSERT INTO `no_ships_day`(`日付`, `種別`) VALUES ('" & ye & "/" & mo & "/" & da.ToString("00") & "',0);"
                                rs = sql_result_no(sql1)

                                If rs = "Complete" Then
                                Else
                                    MsgBox("失敗：" & sql1)
                                End If


                            Case "Color [White]"


                        End Select


                    End If

                End If

            Next

        Next

        ''データベースを切断
        Call sql_cl()

    End Sub

    Function lineNumDecision01()

        Dim ln1 As Integer = 0

        For wn1 As Integer = 0 To 6

            Dim check01 As Integer = 0
            For dn1 As Integer = 0 To 6

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString

                Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    Dim cv As String = CType(cs(0), Label).Text

                    If cv = "" Then
                        check01 += 0
                    Else
                        check01 += Double.Parse(cv)
                    End If

                End If

            Next

            If check01 > 0 Then
                ln1 += 1
            End If

        Next

        Return ln1

    End Function

    Function lineNumDecision02()

        Dim ln1 As Integer = 0

        For wn1 As Integer = 7 To 13

            Dim check01 As Integer = 0
            For dn1 As Integer = 0 To 6

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString

                Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    Dim cv As String = CType(cs(0), Label).Text

                    If cv = "" Then
                        check01 += 0
                    Else
                        check01 += Double.Parse(cv)
                    End If

                End If

            Next

            If check01 > 0 Then
                ln1 += 1
            End If

        Next

        Return ln1

    End Function

End Module
