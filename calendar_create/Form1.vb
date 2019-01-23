Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Call clc01()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label_we00.Click,
        Label_we01.Click,
        Label_we02.Click,
        Label_we03.Click,
        Label_we04.Click,
        Label_we05.Click,
        Label_we06.Click,
        Label_we10.Click,
        Label_we11.Click,
        Label_we12.Click,
        Label_we13.Click,
        Label_we14.Click,
        Label_we15.Click,
        Label_we16.Click,
        Label_we20.Click,
        Label_we21.Click,
        Label_we22.Click,
        Label_we23.Click,
        Label_we24.Click,
        Label_we25.Click,
        Label_we26.Click,
        Label_we30.Click,
        Label_we31.Click,
        Label_we32.Click,
        Label_we33.Click,
        Label_we34.Click,
        Label_we35.Click,
        Label_we36.Click,
        Label_we40.Click,
        Label_we41.Click,
        Label_we42.Click,
        Label_we43.Click,
        Label_we44.Click,
        Label_we45.Click,
        Label_we46.Click,
        Label_we50.Click,
        Label_we51.Click,
        Label_we52.Click,
        Label_we53.Click,
        Label_we54.Click,
        Label_we55.Click,
        Label_we56.Click,
        Label_we60.Click,
        Label_we61.Click,
        Label_we62.Click,
        Label_we63.Click,
        Label_we64.Click,
        Label_we65.Click,
        Label_we66.Click,
        Label_we70.Click,
        Label_we71.Click,
        Label_we72.Click,
        Label_we73.Click,
        Label_we74.Click,
        Label_we75.Click,
        Label_we76.Click,
        Label_we80.Click,
        Label_we81.Click,
        Label_we82.Click,
        Label_we83.Click,
        Label_we84.Click,
        Label_we85.Click,
        Label_we86.Click,
        Label_we90.Click,
        Label_we91.Click,
        Label_we92.Click,
        Label_we93.Click,
        Label_we94.Click,
        Label_we95.Click,
        Label_we96.Click,
        Label_we100.Click,
        Label_we101.Click,
        Label_we102.Click,
        Label_we103.Click,
        Label_we104.Click,
        Label_we105.Click,
        Label_we106.Click,
        Label_we110.Click,
        Label_we111.Click,
        Label_we112.Click,
        Label_we113.Click,
        Label_we114.Click,
        Label_we115.Click,
        Label_we116.Click,
        Label_we120.Click,
        Label_we121.Click,
        Label_we122.Click,
        Label_we123.Click,
        Label_we124.Click,
        Label_we125.Click,
        Label_we126.Click,
        Label_we130.Click,
        Label_we131.Click,
        Label_we132.Click,
        Label_we133.Click,
        Label_we134.Click,
        Label_we135.Click,
        Label_we136.Click



        Dim bc1 As String = sender.BackColor.ToString

        Select Case bc1
            Case "Color [Orange]"
                sender.BackColor = Color.Pink

            Case "Color [Pink]"
                sender.BackColor = Color.White

            Case "Color [White]"
                sender.BackColor = Color.Orange

        End Select


    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''フォームが読み込まれたときの処理

        ''INIファイルを読み込む。
        Dim dbini As IO.StreamReader
        Dim stCurrentDir As String = System.IO.Directory.GetCurrentDirectory()
        CuDr = stCurrentDir

        If IO.File.Exists(CuDr & "\config.ini") = True Then
            dbini = New IO.StreamReader(CuDr & "\config.ini", System.Text.Encoding.Default)

            For lp1 As Integer = 1 To 21
                Dim tbxn1 As String = "Cf_TextBox" & lp1.ToString

                Dim cs As Control() = Me.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    CType(cs(0), TextBox).Text = dbini.ReadLine
                End If
            Next

            ''メイン作業タブへ
            Me.TabControl1.SelectedTab = TabPage1

            dbini.Close()
            dbini.Dispose()
        Else
            MessageBox.Show("設定ファイルが見つからないか壊れています。", "通知")
            Me.TabControl1.SelectedTab = TabPage2
        End If

        Dim cday1 As DateTime = DateTime.Now

        Me.TextBox_year1.Text = cday1.Year.ToString
        Me.TextBox_month1.Text = cday1.Month.ToString

        Call clc01()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ''左右カレンダーでループ数を揃えるためのチェック
        Dim lcl1 As Integer = lineNumDecision01()
        Dim rcl1 As Integer = lineNumDecision02()
        Dim lope As Integer = 0

        If rcl1 > lcl1 Then
            ''右カレンダーの行が多い
            lope = rcl1
        Else
            ''左カレンダーの行が多いか同数
            lope = lcl1
        End If

        lope -= 1

        Dim html1 As New System.Text.StringBuilder()
        Dim html2 As New System.Text.StringBuilder()
        Dim html3 As New System.Text.StringBuilder()

        html1.Append("<table class=""calendar"">").AppendLine()
        html1.Append("  <thead>").AppendLine()
        html1.Append("    <tr>").AppendLine()
        html1.Append("      <th colspan=""7"">" & Me.TextBox_yearLeft.Text & "年" & Me.TextBox_monthLeft.Text & "月</th>").AppendLine()
        html1.Append("    </tr>").AppendLine()
        html1.Append("  </thead>").AppendLine()
        html1.Append("  <tbody>").AppendLine()
        html1.Append("    <tr>").AppendLine()
        html1.Append("      <th>日</th>").AppendLine()
        html1.Append("      <th>月</th>").AppendLine()
        html1.Append("      <th>火</th>").AppendLine()
        html1.Append("      <th>水</th>").AppendLine()
        html1.Append("      <th>木</th>").AppendLine()
        html1.Append("      <th>金</th>").AppendLine()
        html1.Append("      <th>土</th>").AppendLine()
        html1.Append("    </tr>").AppendLine()

        html2.Append("<ul>").AppendLine()
        html2.Append("  <li>").AppendLine()
        html2.Append("    <div class=""calendar"">").AppendLine()
        html2.Append("      <p class=""curr_m"">" & Me.TextBox_yearLeft.Text & "年" & Me.TextBox_monthLeft.Text & "月</p>").AppendLine()
        html2.Append("      <table>").AppendLine()
        html2.Append("        <thead>").AppendLine()
        html2.Append("          <tr>").AppendLine()
        html2.Append("            <th>日</th>").AppendLine()
        html2.Append("            <th>月</th>").AppendLine()
        html2.Append("            <th>火</th>").AppendLine()
        html2.Append("            <th>水</th>").AppendLine()
        html2.Append("            <th>木</th>").AppendLine()
        html2.Append("            <th>金</th>").AppendLine()
        html2.Append("            <th>土</th>").AppendLine()
        html2.Append("          </tr>").AppendLine()
        html2.Append("        </thead>").AppendLine()
        html2.Append("        <tbody>").AppendLine()




        For wn1 As Integer = 0 To lope

            html1.Append("    <tr>").AppendLine()
            html2.Append("          <tr>").AppendLine()

            For dn1 As Integer = 0 To 6
                html1.Append("      <td")
                html2.Append("            <td")

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString

                Dim cs As Control() = Me.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    Dim bc1 As String = CType(cs(0), Label).BackColor.ToString

                    Select Case bc1
                        Case "Color [Orange]"
                            html1.Append(" class=""bg-orange"">")
                            html2.Append(" class=""brown"">")

                        Case "Color [Pink]"
                            html1.Append(" class=""bg-red"">")
                            html2.Append(" class=""red"">")

                        Case "Color [White]"
                            html1.Append(">")
                            html2.Append(">")

                        Case Else
                    End Select
                End If

                If CType(cs(0), Label).Text = "" Then
                Else
                    html1.Append(CType(cs(0), Label).Text)
                    html2.Append(CType(cs(0), Label).Text)
                End If

                html1.Append("</td>").AppendLine()
                html2.Append("</td>").AppendLine()
            Next

            html1.Append("    </tr>").AppendLine()
            html2.Append("          </tr>").AppendLine()

        Next

        html1.Append("  </tbody>").AppendLine()
        html1.Append("</table>").AppendLine()
        html1.Append("<table class=""calendar"">").AppendLine()
        html1.Append("  <thead>").AppendLine()
        html1.Append("    <tr>").AppendLine()
        html1.Append("      <th colspan=""7"">" & Me.TextBox_yearRight.Text & "年" & Me.TextBox_monthRight.Text & "月</th>").AppendLine()
        html1.Append("    </tr>").AppendLine()
        html1.Append("  </thead>").AppendLine()
        html1.Append("  <tbody>").AppendLine()
        html1.Append("    <tr>").AppendLine()
        html1.Append("      <th>日</th>").AppendLine()
        html1.Append("      <th>月</th>").AppendLine()
        html1.Append("      <th>火</th>").AppendLine()
        html1.Append("      <th>水</th>").AppendLine()
        html1.Append("      <th>木</th>").AppendLine()
        html1.Append("      <th>金</th>").AppendLine()
        html1.Append("      <th>土</th>").AppendLine()
        html1.Append("    </tr>").AppendLine()

        html2.Append("      </table>").AppendLine()
        html2.Append("    </div>").AppendLine()
        html2.Append("  </li>").AppendLine()
        html2.Append("  <li>").AppendLine()
        html2.Append("    <div class=""calendar"">").AppendLine()
        html2.Append("      <p class=""curr_m"">" & Me.TextBox_yearRight.Text & "年" & Me.TextBox_monthRight.Text & "月</p>").AppendLine()
        html2.Append("      <table>").AppendLine()
        html2.Append("        <thead>").AppendLine()
        html2.Append("          <tr>").AppendLine()
        html2.Append("            <th>日</th>").AppendLine()
        html2.Append("            <th>月</th>").AppendLine()
        html2.Append("            <th>火</th>").AppendLine()
        html2.Append("            <th>水</th>").AppendLine()
        html2.Append("            <th>木</th>").AppendLine()
        html2.Append("            <th>金</th>").AppendLine()
        html2.Append("            <th>土</th>").AppendLine()
        html2.Append("          </tr>").AppendLine()
        html2.Append("        </thead>").AppendLine()
        html2.Append("        <tbody>").AppendLine()


        lope += 7
        For wn1 As Integer = 7 To lope

            html1.Append("    <tr>").AppendLine()
            html2.Append("          <tr>").AppendLine()

            For dn1 As Integer = 0 To 6
                html1.Append("      <td")
                html2.Append("            <td")

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString

                Dim cs As Control() = Me.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    Dim bc1 As String = CType(cs(0), Label).BackColor.ToString

                    Select Case bc1
                        Case "Color [Orange]"
                            html1.Append(" class=""bg-orange"">")
                            html2.Append(" class=""brown"">")

                        Case "Color [Pink]"
                            html1.Append(" class=""bg-red"">")
                            html2.Append(" class=""red"">")

                        Case "Color [White]"
                            html1.Append(">")
                            html2.Append(">")

                        Case Else
                    End Select
                End If

                If CType(cs(0), Label).Text = "" Then
                Else
                    html1.Append(CType(cs(0), Label).Text)
                    html2.Append(CType(cs(0), Label).Text)
                End If

                html1.Append("</td>").AppendLine()
                html2.Append("</td>").AppendLine()
            Next

            html1.Append("    </tr>").AppendLine()
            html2.Append("          </tr>").AppendLine()
        Next


        html1.Append("  </tbody>").AppendLine()
        html1.Append("</table>").AppendLine()

        html2.Append("        </tbody>").AppendLine()
        html2.Append("      </table>").AppendLine()
        html2.Append("    </div>").AppendLine()
        html2.Append("  </li>").AppendLine()
        html2.Append("</ul>").AppendLine()

        '送信ファイル名
        Dim lofn1 As String = CuDr & "\pc_calendar.html"
        Dim lofn2 As String = CuDr & "\sp_calendar.html"

        '送信先URL
        Dim hofn01 As String = Me.Cf_TextBox1.Text & "/calendar.html"
        Dim hofn02 As String = Me.Cf_TextBox1.Text & "/sp/calendar.html"
        Dim hofn03 As String = Me.Cf_TextBox7.Text & ":16910/calendar.html"
        Dim hofn04 As String = Me.Cf_TextBox7.Text & ":16910/sp/calendar.html"
        Dim hofn05 As String = Me.Cf_TextBox4.Text & ":16910/calendar.html"
        Dim hofn06 As String = Me.Cf_TextBox4.Text & ":16910/sp/calendar.html"
        Dim hofn07 As String = Me.Cf_TextBox10.Text & "/calendar.html"
        Dim hofn08 As String = Me.Cf_TextBox10.Text & "/sp/calendar.html"
        Dim hofn09 As String = Me.Cf_TextBox13.Text & "/calendar.html"
        Dim hofn10 As String = Me.Cf_TextBox13.Text & "/sp/calendar.html"
        Dim hofn11 As String = Me.Cf_TextBox19.Text & "/calendar.html"
        Dim hofn12 As String = Me.Cf_TextBox19.Text & "/sp/calendar.html"

        ''UTF-8 でファイルに書き出す
        Dim sw1 As New System.IO.StreamWriter(lofn1)
        sw1.Write(html1)
        sw1.Close()

        Dim sw2 As New System.IO.StreamWriter(lofn2)
        sw2.Write(html2)
        sw2.Close()

        'FTPへ上記ファイルをアップ
        'ロイヤルYahoo!
        fil_ftp_up(lofn1, "ftp://" & hofn07, Me.Cf_TextBox11.Text, Me.Cf_TextBox12.Text)
        fil_ftp_up(lofn2, "ftp://" & hofn08, Me.Cf_TextBox11.Text, Me.Cf_TextBox12.Text)

        'マザープラスストア
        fil_ftp_up(lofn1, "ftp://" & hofn11, Me.Cf_TextBox20.Text, Me.Cf_TextBox21.Text)
        'fil_ftp_up(lofn2, "ftp://" & hofn12, Me.Cf_TextBox20.Text, Me.Cf_TextBox21.Text)

        '生活空間Yahoo!
        fil_ftp_up(lofn1, "ftp://" & hofn01, Me.Cf_TextBox2.Text, Me.Cf_TextBox3.Text)
        fil_ftp_up(lofn2, "ftp://" & hofn02, Me.Cf_TextBox2.Text, Me.Cf_TextBox3.Text)


        '生活空間楽天
        fil_ftp_up(lofn1, "ftp://" & hofn03, Me.Cf_TextBox8.Text, Me.Cf_TextBox9.Text)
        fil_ftp_up(lofn2, "ftp://" & hofn04, Me.Cf_TextBox8.Text, Me.Cf_TextBox9.Text)

        'ロイヤル楽天
        fil_ftp_up(lofn1, "ftp://" & hofn05, Me.Cf_TextBox5.Text, Me.Cf_TextBox6.Text)
        fil_ftp_up(lofn2, "ftp://" & hofn06, Me.Cf_TextBox5.Text, Me.Cf_TextBox6.Text)

        'ポンパレモール
        fil_ftps_up(lofn1, "ftp://" & hofn09, Me.Cf_TextBox14.Text, Me.Cf_TextBox15.Text)
        fil_ftps_up(lofn2, "ftp://" & hofn10, Me.Cf_TextBox14.Text, Me.Cf_TextBox15.Text)

        MessageBox.Show("各店FTPに、カレンダーをアップロードしました", "完了", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Call register_holiday()
        MessageBox.Show("データベース登録完了", "完了", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        close_save()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim html1 As String = ""

        html1 &= "<!doctype html>"
        html1 &= "<HTML>"
        html1 &= "<HEAD>"
        html1 &= "<meta charset=""utf-8"">"
        html1 &= "<TITLE>営業カレンダー</TITLE>"
        html1 &= "<STYLE type=""text/css"">"
        html1 &= "<!--"
        html1 &= "table#main2 {"
        html1 &= "	width: 234px;"
        html1 &= "}"
        html1 &= "table#cr1 {"
        html1 &= "	border-collapse: collapse;"
        html1 &= "	width: 112px;"
        html1 &= "	height: 128px;"
        html1 &= "}"
        html1 &= "table#cr2 {"
        html1 &= "	border-collapse: collapse;"
        html1 &= "	width: 112px;"
        html1 &= "	height: 128px;"
        html1 &= "}"
        html1 &= "TD {"
        html1 &= "	background-color: #00CCFF;"
        html1 &= "	border-bottom: 1px solid #000000;"
        html1 &= "	border-left: 1px solid #000000;"
        html1 &= "	border-right: 1px solid #000000;"
        html1 &= "	width: 16px;"
        html1 &= "	height: 14px;"
        html1 &= "	font-size: 12px;"
        html1 &= "	text-align: center;"
        html1 &= "	padding: 0px 1px 0px 1px;"
        html1 &= "}"
        html1 &= "TD#dwk {"
        html1 &= "	background-color: #66CC33;"
        html1 &= "}"
        html1 &= "TD#cgr {"
        html1 &= "	background-color: #33ffcc;"
        html1 &= "}"
        html1 &= "TD#cye {"
        html1 &= "	background-color: #FFFF99;"
        html1 &= "}"
        html1 &= "TD#mon {"
        html1 &= "	border-top: 1px solid #000000;"
        html1 &= "	font-weight: bold;"
        html1 &= "	background-color: #FFFFFF;"
        html1 &= "}"
        html1 &= "TD#nl1 {"
        html1 &= "	background-color: #FFFFFF;"
        html1 &= "	border-bottom: 1px solid #FFFFFF;"
        html1 &= "	border-top: 1px solid #FFFFFF;"
        html1 &= "	border-left: 1px solid #FFFFFF;"
        html1 &= "	border-right: 1px solid #FFFFFF;"
        html1 &= "	padding: 0px 2px 0px 2px;"
        html1 &= "}"
        html1 &= "-->"
        html1 &= "</STYLE>"
        html1 &= "</HEAD>"
        html1 &= ""
        html1 &= "<BODY>"
        html1 &= "<TABLE id=""main2"">"
        html1 &= "  <TBODY>"
        html1 &= "    <TR>"
        html1 &= "      <TD id=""nl1""><TABLE id=""cr1"">"
        html1 &= "          <TR>"
        html1 &= "            <TD colspan=""7"" id=""mon"">" & Me.TextBox_yearLeft.Text & "年" & Me.TextBox_monthLeft.Text & "月</TD>"
        html1 &= "          </TR>"
        html1 &= "          <TR>"
        html1 &= "            <TD id=""dwk"">日</TD>"
        html1 &= "            <TD id=""dwk"">月</TD>"
        html1 &= "            <TD id=""dwk"">火</TD>"
        html1 &= "            <TD id=""dwk"">水</TD>"
        html1 &= "            <TD id=""dwk"">木</TD>"
        html1 &= "            <TD id=""dwk"">金</TD>"
        html1 &= "            <TD id=""dwk"">土</TD>"
        html1 &= "          </TR>"

        For wn1 As Integer = 0 To 5

            html1 &= "          <TR>"

            For dn1 As Integer = 0 To 6
                html1 &= "            <TD"

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString
                Dim cs As Control() = Me.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    Dim bc1 As String = CType(cs(0), Label).BackColor.ToString

                    Select Case bc1
                        Case "Color [Orange]"
                            html1 &= " id=""cgr"">"
                            html1 &= CType(cs(0), Label).Text
                            html1 &= "</TD>"

                        Case "Color [Pink]"
                            html1 &= " id=""cye"">"
                            html1 &= CType(cs(0), Label).Text
                            html1 &= "</TD>"

                        Case "Color [White]"
                            If CType(cs(0), Label).Text = "" Then
                                html1 &= " id=""cye"">&nbsp;</TD>"

                            Else
                                html1 &= ">"
                                html1 &= CType(cs(0), Label).Text
                                html1 &= "</TD>"

                            End If
                        Case Else
                    End Select
                End If

            Next

            html1 &= "          </TR>"

        Next

        html1 &= "        </TABLE></TD>"
        html1 &= "      <TD id=""nl1""><TABLE id=""cr1"">"
        html1 &= "          <TR>"
        html1 &= "            <TD colspan=""7"" id=""mon"">" & Me.TextBox_yearRight.Text & "年" & Me.TextBox_monthRight.Text & "月</TD>"
        html1 &= "          </TR>"
        html1 &= "          <TR>"
        html1 &= "            <TD id=""dwk"">日</TD>"
        html1 &= "            <TD id=""dwk"">月</TD>"
        html1 &= "            <TD id=""dwk"">火</TD>"
        html1 &= "            <TD id=""dwk"">水</TD>"
        html1 &= "            <TD id=""dwk"">木</TD>"
        html1 &= "            <TD id=""dwk"">金</TD>"
        html1 &= "            <TD id=""dwk"">土</TD>"
        html1 &= "          </TR>"

        For wn1 As Integer = 7 To 12

            html1 &= "          <TR>"

            For dn1 As Integer = 0 To 6
                html1 &= "            <TD"

                Dim tbxn1 As String = "Label_we" & wn1.ToString & dn1.ToString
                Dim cs As Control() = Me.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    Dim bc1 As String = CType(cs(0), Label).BackColor.ToString

                    Select Case bc1
                        Case "Color [Orange]"
                            html1 &= " id=""cgr"">"
                            html1 &= CType(cs(0), Label).Text
                            html1 &= "</TD>"

                        Case "Color [Pink]"
                            html1 &= " id=""cye"">"
                            html1 &= CType(cs(0), Label).Text
                            html1 &= "</TD>"

                        Case "Color [White]"
                            If CType(cs(0), Label).Text = "" Then
                                html1 &= " id=""cye"">&nbsp;</TD>"

                            Else
                                html1 &= ">"
                                html1 &= CType(cs(0), Label).Text
                                html1 &= "</TD>"

                            End If
                        Case Else
                    End Select
                End If

            Next

            html1 &= "          </TR>"

        Next


        html1 &= "        </TABLE></TD>"
        html1 &= "    </TR>"
        html1 &= "  </TBODY>"
        html1 &= "</TABLE>"
        html1 &= "</BODY>"
        html1 &= "</HTML>"

        Dim lofn1 As String = CuDr & "\rm_calendar.html"
        Dim hofn1 As String = Me.Cf_TextBox16.Text & "/web/imaging_calendar.html"

        ''UTF-8 でファイルに書き出す
        Dim sw1 As New System.IO.StreamWriter(lofn1)
        sw1.Write(html1)
        sw1.Close()


        'FTPへ上記ファイルをアップ
        fil_ftps_up(lofn1, "ftp://" & hofn1, Me.Cf_TextBox17.Text, Me.Cf_TextBox18.Text)

        Dim result As DialogResult = MessageBox.Show("Amazon用カレンダーアップロード完了" & vbCrLf & "ブラウザで確認しますか？", "完了", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        '何が選択されたか調べる 
        If result = DialogResult.Yes Then
            '「はい」が選択された時 
            System.Diagnostics.Process.Start("http://www.royal-e.com/imaging_calendar.html")
        End If


    End Sub


End Class
