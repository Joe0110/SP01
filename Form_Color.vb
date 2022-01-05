Public Class Form_Color

    '--------------------
    '   目前無法使用 Color_List
    '--------------------
    'Public Const Color_List() As Color = {Color.White,
    '                                        Color.Red,
    '                                        Color.Green,
    '                                        Color.DarkOrange,
    '                                        Color.Black,
    '                                        Color.Yellow,
    '                                        Color.YellowGreen,
    '                                        Color.Purple,
    '                                        Color.Blue,
    '                                        Color.Pink,
    '                                        Color.DodgerBlue,
    '                                        Color.Aqua,
    '                                        Color.Magenta,
    '                                        Color.LimeGreen,
    '                                        Color.Lime,
    '                                        Color.LightSkyBlue,
    '                                        Color.LightGreen,
    '                                        Color.Gold,
    '                                        Color.Olive,
    '                                        Color.RoyalBlue,
    '                                        Color.Violet,
    '                                        Color.DarkGray,
    '                                        Color.LightSalmon,
    '                                        Color.DeepPink,
    '                                        Color.Navy,
    '                                        Color.DarkGreen
    '    }

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form_Spec_CHK_Ctl.CMB_Font_Color.SelectedIndex = Val(TB_Fcolor.Text)
    End Sub

    Private Sub Form_Color_Load(sender As Object, e As EventArgs) Handles Me.Load
        '[COLOR_02]
        Label0.ForeColor = Color.Blue
        Label1.ForeColor = Color.Red

        Label2.ForeColor = Color.Green
        Label3.ForeColor = Color.Black
        Label4.ForeColor = Color.DarkOrange

        Label5.ForeColor = Color.Yellow
        Label5.BackColor = Color.Black

        Label6.ForeColor = Color.YellowGreen
        Label7.ForeColor = Color.Purple
        Label8.ForeColor = Color.Magenta

        Label9.ForeColor = Color.Pink
        Label9.BackColor = Color.Black

        Label10.ForeColor = Color.DodgerBlue
        Label10.BackColor = Color.Black

        Label11.ForeColor = Color.Lime

        '-----------------------------
        Label12.ForeColor = Color.Gold
        Label13.ForeColor = Color.LightSalmon
        Label14.ForeColor = Color.DeepPink
        Label15.ForeColor = Color.Navy
        Label16.ForeColor = Color.DarkGray

        '===============================

        TextBox0.Text = "無"
        'TextBox0.BackColor = Color.Blue
        TextBox1.BackColor = Color.Red
        TextBox2.BackColor = Color.Green
        TextBox3.BackColor = Color.DarkOrange
        TextBox4.BackColor = Color.Black
        TextBox5.BackColor = Color.Yellow
        TextBox6.BackColor = Color.YellowGreen
        TextBox7.BackColor = Color.Purple
        TextBox8.BackColor = Color.Blue
        TextBox9.BackColor = Color.Pink

        TextBox10.BackColor = Color.DodgerBlue

        TextBox11.BackColor = Color.Aqua
        TextBox12.BackColor = Color.Magenta
        TextBox13.BackColor = Color.LimeGreen

        TextBox14.BackColor = Color.Lime

        TextBox15.BackColor = Color.LightSkyBlue
        TextBox16.BackColor = Color.LightGreen

        TextBox17.BackColor = Color.Gold

        TextBox18.BackColor = Color.Olive
        TextBox19.BackColor = Color.RoyalBlue
        TextBox20.BackColor = Color.Violet
        TextBox21.BackColor = Color.DarkGray
        TextBox22.BackColor = Color.LightSalmon
        TextBox23.BackColor = Color.DeepPink
        TextBox24.BackColor = Color.Navy
        TextBox25.BackColor = Color.DarkGreen



    End Sub

    Private Sub TextBox0_Click(sender As Object, e As EventArgs) Handles TextBox0.Click
        Change_Background_Color(0)
    End Sub

    Private Sub TextBox0_TextChanged(sender As Object, e As EventArgs) Handles TextBox0.TextChanged

    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        Change_Background_Color(1)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        Change_Background_Color(2)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
  
    End Sub

    Private Sub TextBox3_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        Change_Background_Color(3)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox4_Click(sender As Object, e As EventArgs) Handles TextBox4.Click
        Change_Background_Color(4)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox5_Click(sender As Object, e As EventArgs) Handles TextBox5.Click
        Change_Background_Color(5)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox6_Click(sender As Object, e As EventArgs) Handles TextBox6.Click
        Change_Background_Color(6)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox7_Click(sender As Object, e As EventArgs) Handles TextBox7.Click
        Change_Background_Color(7)
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox8_Click(sender As Object, e As EventArgs) Handles TextBox8.Click

        Change_Background_Color(8)
    End Sub
    Public Sub Change_Background_Color(ByVal AAA As Integer)
        '-----------------------------
        '   改變背景顏色
        '-----------------------------
        Form_Spec_CHK_Ctl.CMB_ItemType.SelectedIndex = 13
        Form_Spec_CHK_Ctl.TB_Modify_Value.Text = AAA
        TB_Bcolor.Text = Form_Spec_CHK_Ctl.TB_Modify_Value.Text
        Form_Spec_CHK_Ctl.Modify_BlockItem()

    End Sub
    Public Sub Change_Font_Color(ByVal AAA As Integer)
        '-----------------------------
        '   改變字型顏色
        '-----------------------------
        Form_Spec_CHK_Ctl.CMB_ItemType.SelectedIndex = 12
        Form_Spec_CHK_Ctl.TB_Modify_Value.Text = AAA
        TB_Fcolor.Text = Form_Spec_CHK_Ctl.TB_Modify_Value.Text
        Form_Spec_CHK_Ctl.Modify_BlockItem()

    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox9_Click(sender As Object, e As EventArgs) Handles TextBox9.Click
        Change_Background_Color(9)
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox10_Click(sender As Object, e As EventArgs) Handles TextBox10.Click
        Change_Background_Color(10)
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Change_Font_Color(1)
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Change_Font_Color(2)
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Change_Font_Color(3)
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        Change_Font_Color(4)
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        Change_Font_Color(5)
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        Change_Font_Color(6)
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        Change_Font_Color(7)
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        Change_Font_Color(8)
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        Change_Font_Color(9)
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        Change_Font_Color(10)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form_Spec_CHK_Ctl.CMB_BG_Color.SelectedIndex = Val(TB_Bcolor.Text)
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label0.Click
        Change_Font_Color(0)
    End Sub

    Private Sub TextBox11_Click(sender As Object, e As EventArgs) Handles TextBox11.Click


        Change_Background_Color(11)
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox12_Click(sender As Object, e As EventArgs) Handles TextBox12.Click

        Change_Background_Color(12)
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox13_Click(sender As Object, e As EventArgs) Handles TextBox13.Click

        Change_Background_Color(13)
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox14_Click(sender As Object, e As EventArgs) Handles TextBox14.Click

        Change_Background_Color(14)
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox15_Click(sender As Object, e As EventArgs) Handles TextBox15.Click

        Change_Background_Color(15)
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox16_Click(sender As Object, e As EventArgs) Handles TextBox16.Click

        Change_Background_Color(16)
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        Change_Font_Color(11)
    End Sub

    Private Sub TextBox17_Click(sender As Object, e As EventArgs) Handles TextBox17.Click
  '
        Change_Color_Action(TextBox17.Name)
    End Sub
    Public Sub Change_Color_Action(ByVal Str_xx As String)
        '---------------------------
        '   改變方塊顏色
        '---------------------------
        Form_Spec_CHK_Ctl.CMB_ItemType.SelectedIndex = 13
        'Debug.Print("ColorBTN1:" & Str_xx)
        Form_Spec_CHK_Ctl.TB_Modify_Value.Text = Str_xx.Replace("TextBox", "")
        TB_Fcolor.Text = Form_Spec_CHK_Ctl.TB_Modify_Value.Text
        'Debug.Print("ColorBTN2:" & Form_Spec_CHK_Ctl.TB_Modify_Value.Text)
        Form_Spec_CHK_Ctl.Modify_BlockItem()
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub TextBox18_Click(sender As Object, e As EventArgs) Handles TextBox18.Click

        Change_Color_Action(TextBox18.Name)
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged

    End Sub

    Private Sub TextBox19_Click(sender As Object, e As EventArgs) Handles TextBox19.Click
        Change_Color_Action(TextBox19.Name)
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub TextBox20_Click(sender As Object, e As EventArgs) Handles TextBox20.Click
        Change_Color_Action(TextBox20.Name)
    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub TextBox21_Click(sender As Object, e As EventArgs) Handles TextBox21.Click
        Change_Color_Action(TextBox21.Name)
    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged

    End Sub

    Private Sub TextBox22_Click(sender As Object, e As EventArgs) Handles TextBox22.Click
        Change_Color_Action(TextBox22.Name)
    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub TextBox23_Click(sender As Object, e As EventArgs) Handles TextBox23.Click
        Change_Color_Action(TextBox23.Name)
    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged

    End Sub

    Private Sub TextBox24_Click(sender As Object, e As EventArgs) Handles TextBox24.Click
        Change_Color_Action(TextBox24.Name)
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged

    End Sub

    Private Sub TextBox25_Click(sender As Object, e As EventArgs) Handles TextBox25.Click
        Change_Color_Action(TextBox25.Name)
    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

        Change_Font_Color(12)
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

        Change_Font_Color(13)
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click

        Change_Font_Color(14)
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        Change_Font_Color(15)
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        Change_Font_Color(16)
    End Sub
End Class