'Enum IO_Tpye
'    None = 0
'    Trivial = 1
'    Regular = 2
'    Important = 3
'    Critical = 4
'End Enum


Imports System.IO '新增此命名空間
Imports System.Math
Public Structure Spec_Item_Color
    Public XReady As Boolean

    Public Name() As String
    Public FColor() As String
    Public Color() As String
    Public Src_Path As String   '來源檔案

    Public Select_Trg As ComboBox
    Public Fcolor_Trg As ComboBox
    Public color_Trg As ComboBox

    'Public E_Sum As Single
    Sub init()
        XReady = False
        ReDim Name(50)
        ReDim FColor(50)
        ReDim Color(50)

    End Sub

    Sub set_Src_file(ByRef Dtrg As String)
        Src_Path = Dtrg
    End Sub
    Sub set_Select_item(ByRef AAA As ComboBox)
        Select_Trg = AAA
    End Sub
    Sub set_FColor_item(ByRef AAA As ComboBox)
        Fcolor_Trg = AAA
    End Sub
    Sub set_Color_item(ByRef AAA As ComboBox)
        color_Trg = AAA
    End Sub
    Sub Get_Color(ByVal AAA As Integer)
        If XReady = True Then
            '------------
            '   字的顏色
            '------------
            If IsNumeric(FColor(AAA)) = True Then
                Fcolor_Trg.SelectedIndex = Val(FColor(AAA))
            End If

            '------------
            '   背景的顏色
            '------------
            If IsNumeric(Color(AAA)) = True Then
                color_Trg.SelectedIndex = Val(Color(AAA))
            End If
        End If

    End Sub
    Sub Load_Src_file()
        'Src_Path = Dtrg
        Dim i As Integer = 0
        Dim temp_array() As String

        Dim Parameter_Line As String = ""






        Select_Trg.Items.Clear()

        '開啟證件
        Dim fs As FileStream = New FileStream(Src_Path, FileMode.Open)
        Using sr As StreamReader = New StreamReader(fs, _
                                                    System.Text.Encoding.Default)


            '讀取一行資料 略過標題



            While (sr.EndOfStream <> True)



                '送出相關設定
                Parameter_Line = sr.ReadLine()   ' & vbCrLf                                     '讀取一行資料

                temp_array = Parameter_Line.Split(",")
                Name(i) = temp_array(0)
                FColor(i) = temp_array(1)
                Color(i) = temp_array(2)
                Select_Trg.Items.Add(Name(i))
                'Debug.Print("ADD:" & Name(i))
                i = i + 1

            End While







            sr.Close()                                                          '關閉檔案




        End Using

        XReady = True

    End Sub















End Structure
Public Structure DWItem2
    Public LB_Name As Label         '表示項次用

    Public CKB_Name As CheckBox     '顯示Title

    Public TB_Vale As TextBox
    Public TB_Scale As TextBox
    Public TB_Offset As TextBox

    Public Enable As Boolean
    Public Type As Integer
    Public value As Single
    Public Scale As Single
    Public Offset As Single


End Structure


Public Structure ParaItem
    'Public NO As Integer
    Public NO As String             '前面加 "@" 時, 不進行判斷
    Public DispItem As String
    Public DispType As String
    Public AD_Type As Byte          '0: 不設定, 1:第1組, 2:第2組
    Public SPH1 As Integer
    Public SPL1 As Integer
    Public SPH2 As Integer
    Public SPL2 As Integer
    Public THI As Integer
    Public Setting_Type As String   '目前在設定PWM輸出時, 需要3個參數 (超過2000以上的視為PWM腳位輸出控制命令)

    Public Value_1 As Integer
    Public Value_1s As Single
    Public Value_1str As String

    Public Value_2 As Integer
    Public Value_3 As Integer

    'Public judge_type_ As Integer    '==> spec_type
    Public Spec_type As Integer

    Public Spec_value As String
    Public diff As Single
    Public delay As Integer
    Public trg_type As Integer          '20190318 New Add
    Public OK As Boolean
End Structure

Public Structure Simulate_Source
    Public SrcType As Byte

    Dim DATA As Single
    Dim Para As Single
    Public Value As Single
    Dim EN As Boolean

    Dim Src_DAT As TextBox
    Dim Src_DAT_LB As Label
    Dim Src_DAT_V As Single
    Dim Src_DAT_BAR As TrackBar
    ' Dim Src_Para As TextBox

    '設定來源
    '==================== 多型 ====================
    Sub Set_DAT_Source(ByRef Src As TextBox)
        Src_DAT = Src
        SrcType = 0
        EN = True
    End Sub
    Sub Set_DAT_Source(ByRef Src As Label)
        Src_DAT_LB = Src
        SrcType = 1
        EN = True
    End Sub
    Sub Set_DAT_Source(ByRef Src As Single)
        Src_DAT_V = Src
        SrcType = 2
        EN = True
    End Sub
    Sub Set_DAT_Source(ByRef Src As TrackBar)
        Src_DAT_BAR = Src
        SrcType = 3
        EN = True
    End Sub
    '==================== 多型 ====================

    'Sub Set_Para_Source(Src As TextBox)
    '    Src_Para = Src
    '    EN = True
    'End Sub


    Sub init()
        EN = False
        'Inst_EN = False
        SrcType = 255
        DATA = 0
        Para = 1
        'Is_Base = False
    End Sub

    '取得目前值
    Sub Get_Value()


        If EN = True Then
            Select Case (SrcType)
                Case 0
                    DATA = Val(Src_DAT.Text)
                Case 1
                    DATA = Val(Src_DAT_LB.Text)
                Case 2
                    DATA = Src_DAT_V
                Case 3
                    DATA = Src_DAT_BAR.Value
                Case Else
                    DATA = 0
            End Select


            'Para = Val(Src_Para.Text)
            Value = DATA
        Else
            DATA = 0
            Para = 0
            Value = 0
        End If
    End Sub

    

    '------------------------------

    '------------------------------





End Structure
Public Structure Simulate_Para_structure_element_Base
    Public Base_VAL As Simulate_Source
    Public Base_MUL As Simulate_Source
    Public Base_OFS As Simulate_Source

    Public Inst_VAL As Simulate_Source
    Public Inst_MUL As Simulate_Source
    Public Inst_OFS As Simulate_Source

    Public Para As Simulate_Source

    'Public SUM As Simulate_Source

    Public Sum_Value As Single
    Public Sum_Value_Trg As TextBox

    Public Result_Base_Value As Single
    Public Result_Inst_Value As Single
    'Public E_Sum As Single

    Sub set_sum(ByRef Dtrg As TextBox)
        Sum_Value_Trg = Dtrg
    End Sub
    Sub renew_sum()
        Sum_Value_Trg.Text = Sum_Value

    End Sub
    Sub cal()


        Result_Base_Value = Base_VAL.Value * Base_MUL.Value + Base_OFS.Value



        Result_Inst_Value = Inst_VAL.Value * Inst_MUL.Value + Inst_OFS.Value


        Sum_Value = (Result_Base_Value - Result_Inst_Value) * Para.Value

    End Sub

  


    Sub init()
        Base_VAL.init()
        Base_MUL.init()
        Base_OFS.init()

        Inst_VAL.init()
        Inst_MUL.init()
        Inst_OFS.init()

        Para.init()

    End Sub

    Sub Renew()
        Base_VAL.Get_Value()
        Base_MUL.Get_Value()
        Base_OFS.Get_Value()

        Inst_VAL.Get_Value()
        Inst_MUL.Get_Value()
        Inst_OFS.Get_Value()

        Para.Get_Value()

        cal()
        renew_sum()

    End Sub








End Structure
Public Structure Simulate_Para_structure_element_UI
    Dim SrcType As Byte

    Dim DATA As Single
    Dim Para As Single
    Dim EN As Boolean

    Dim Src_DAT As TextBox
    Dim Src_DAT_LB As Label
    Dim Src_DAT_V As Single
    Dim Src_Para As TextBox

    '==================== 多型 ====================
    Sub Set_DAT_Source(Src As TextBox)
        Src_DAT = Src
        SrcType = 0
        EN = True
    End Sub
    Sub Set_DAT_Source(Src As Label)
        Src_DAT_LB = Src
        SrcType = 1
        EN = True
    End Sub
    Sub Set_DAT_Source(Src As Single)
        Src_DAT_V = Src
        SrcType = 2
        EN = True
    End Sub
    '==================== 多型 ====================

    Sub Set_Para_Source(Src As TextBox)
        Src_Para = Src
        EN = True
    End Sub


    Sub init()
        EN = False


        DATA = 0
        Para = 0

    End Sub

    Sub Renew()

        If EN = True Then
            Select Case (SrcType)
                Case 0
                    DATA = Val(Src_DAT.Text)
                Case 1
                    DATA = Val(Src_DAT_LB.Text)
                Case 2
                    DATA = Src_DAT_V
                Case Else
                    DATA = 0
            End Select



            Para = Val(Src_Para.Text)
        Else
            DATA = 0
            Para = 0
        End If
    End Sub

    '------------------------------

    '------------------------------









End Structure
Public Structure Simulate_Para_structure_element

    Dim SrcType As Byte

    Dim DATA As Single
    Dim Para As Single
    Dim EN As Boolean

    Dim Src_DAT As TextBox
    Dim Src_DAT_LB As Label
    Dim Src_DAT_V As Single
    Dim Src_Para As TextBox

    '==================== 多型 ====================
    Sub Set_DAT_Source(Src As TextBox)
        Src_DAT = Src
        SrcType = 0
        EN = True
    End Sub
    Sub Set_DAT_Source(Src As Label)
        Src_DAT_LB = Src
        SrcType = 1
        EN = True
    End Sub
    Sub Set_DAT_Source(Src As Single)
        Src_DAT_V = Src
        SrcType = 2
        EN = True
    End Sub
    '==================== 多型 ====================

    Sub Set_Para(Src As Single)
        Para = Src
        'EN = True
    End Sub


    Sub init()
        EN = False


        DATA = 0
        Para = 0

    End Sub

    Sub Renew()

        If EN = True Then
            Select Case (SrcType)
                Case 0
                    DATA = Val(Src_DAT.Text)
                Case 1
                    DATA = Val(Src_DAT_LB.Text)
                Case 2
                    DATA = Src_DAT_V
                Case Else
                    DATA = 0
            End Select



            'Para = Val(Src_Para.Text)
        Else
            DATA = 0
            Para = 0
        End If
    End Sub

    '------------------------------

    '------------------------------









End Structure

Public Class CMB
    'Public 
End Class

Public Class SIM_Item
    Dim Use_Trg_bar As Boolean

    'Public Base As Simulate_Para_structure_element_Base
    Public ENV As Simulate_Para_structure_element_Base
    '上升

    Public U1 As Simulate_Para_structure_element_UI
    Public U2 As Simulate_Para_structure_element_UI
    Public U3 As Simulate_Para_structure_element_UI
    Public U4 As Simulate_Para_structure_element_UI

    '下降
    Public D1 As Simulate_Para_structure_element_UI
    Public D2 As Simulate_Para_structure_element_UI
    Public D3 As Simulate_Para_structure_element_UI
    Public D4 As Simulate_Para_structure_element_UI

    Public Trg_Bar As TrackBar


    '總和
    Dim Total_Cool As Single        '總制冷
    Dim Total_Hot As Single         '總制熱
    Dim Total_C_H_Result As Single  '總制冷熱

    '變化率
    Dim Simpling_HZ As Single
    Dim ChangeRate As Single
    Dim Total_C_H_Result_ACC As Single

    '質量
    Dim USE_compare_value As Integer
    Dim compare_value_U As Integer
    Dim compare_value_D As Integer

    Public compare_value_is_UP As Boolean

    Dim TBAR_OP_V As Integer

    Public compare_value_U_Array(16) As Integer
    Public compare_value_D_Array(16) As Integer


    Dim Src_Total_Cool_TB As TextBox
    Dim Src_Total_Hot_TB As TextBox
    Dim Src_Total_C_H_Result_TB As TextBox
    Dim Src_Total_ACC As TextBox                '總累加結果
    Dim Src_Use_Vsize_TB As TextBox


    Dim Src_Size_UP As TextBox
    Dim Src_Size_DN As TextBox

    '變化控制
    Dim plus_value As Single
    Dim total_plus_value As Single
    Dim total_plus_value_OP As Single
    Dim total_plus_value_Minus As Boolean

    Dim ParaID As Byte
    Public Para_AREA As Byte




    Dim TBarChange_V As Integer '總質量

    ' Sub ReNew_Use_Vsize()
    'Src_Use_Vsize_TB.Text = USE_compare_value
    'End Sub
    Sub cal_area_all(dat As Byte)
        Para_AREA = cal_Area2(dat)
    End Sub

    Sub Reset_ACC()
        Total_C_H_Result_ACC = 0
    End Sub
    Sub Vsize_Display()
        Src_Size_DN.Text = compare_value_D
        Src_Size_UP.Text = compare_value_U
        Src_Use_Vsize_TB.Text = USE_compare_value
    End Sub
    Sub ReNew_all()
        ENV.Renew()
        U1.Renew()
        U2.Renew()
        U3.Renew()
        U4.Renew()

        D1.Renew()
        D2.Renew()
        D3.Renew()
        D4.Renew()

    End Sub
    Sub Set_Trg_Bar(SrcTB As TrackBar)
        Trg_Bar = SrcTB
        Use_Trg_bar = True
    End Sub

    '=========================================
    '   設定熱量計算 顯示位置
    '=========================================
    Sub Set_Total_Cool_TB(SrcTB As TextBox)
        Src_Total_Cool_TB = SrcTB
    End Sub
    Sub Set_Use_Vsize_TB(SrcTB As TextBox)
        Src_Use_Vsize_TB = SrcTB
    End Sub
    Sub Set_Total_Hot_TB(SrcTB As TextBox)
        Src_Total_Hot_TB = SrcTB
    End Sub

    Sub Set_Total_H_L_Result_TB(SrcTB As TextBox)
        Src_Total_C_H_Result_TB = SrcTB
    End Sub
    Sub Set_Total_H_L_Result_ACC_TB(SrcTB As TextBox)
        Src_Total_ACC = SrcTB
    End Sub
    Sub Set_SIZE_UP_TB(SrcTB As TextBox)
        Src_Size_UP = SrcTB
    End Sub
    Sub Set_SIZE_DN_TB(SrcTB As TextBox)
        Src_Size_DN = SrcTB
    End Sub

    Sub Display_HL_Result()
        Src_Total_Cool_TB.Text = Total_Cool
        Src_Total_Hot_TB.Text = Total_Hot
        Src_Total_C_H_Result_TB.Text = Total_C_H_Result.ToString("#.##")
        Src_Total_ACC.Text = Total_C_H_Result_ACC

    End Sub
    Sub cal_base()
        'Base.Total_Base = Base.Base_Mul2
    End Sub

    '===========================
    '   計算 冷/熱量 總合
    '===========================
    Sub Cal_Total_C_H()
        Total_Cool = D1.DATA * D1.Para + D2.DATA * D2.Para + D3.DATA * D3.Para + D4.DATA * D4.Para
        Total_Hot = U1.DATA * U1.Para + U2.DATA * U2.Para + U3.DATA * U3.Para + U4.DATA * U4.Para

        '單次運算結果
        Total_C_H_Result = Total_Hot - Total_Cool + ENV.Sum_Value

        '累加結果
        Total_C_H_Result_ACC += Total_C_H_Result

        '設定 變化方向
        If Total_C_H_Result_ACC >= 0 Then
            compare_value_is_UP = True
        Else
            compare_value_is_UP = False
        End If


        Display_HL_Result()

    End Sub
    '=========================================
    '   更新 熱量計算 顯示結果
    '=========================================
    Sub Renew_HL_Result()                       '顯示熱量計算結果
        Src_Total_Cool_TB.Text = Total_Cool
        Src_Total_Hot_TB.Text = Total_Hot
        Src_Total_C_H_Result_TB.Text = Total_C_H_Result
    End Sub
    Sub Set_U_D_array(temp_U_array() As Integer, temp_D_array() As Integer)
        For i = 0 To 15
            compare_value_U_Array(i) = temp_U_array(i)
            compare_value_D_Array(i) = temp_D_array(i)
        Next i
    End Sub
    Sub Set_U_array(temp_U_array() As Integer)
        For i = 0 To 15
            compare_value_U_Array(i) = temp_U_array(i)
        Next i
    End Sub
    Sub Set_D_array(temp_D_array() As Integer)
        For i = 0 To 15
            compare_value_D_Array(i) = temp_D_array(i)
        Next i
    End Sub

    Dim obj_cube As Integer '總質量
    Sub GET_COMP_VALUE(tValue As Byte)
        Para_AREA = cal_Area2(tValue)
        Get_compare_value(Para_AREA)
        Vsize_Display()
    End Sub
    Sub Get_compare_value(Area_V As Integer)
        compare_value_U = compare_value_U_Array(Area_V)
        compare_value_D = compare_value_D_Array(Area_V)
        If compare_value_is_UP Then
            USE_compare_value = compare_value_U
        Else
            USE_compare_value = compare_value_D
        End If
    End Sub

    Sub Bar_Control()
        ' total_plus_value += Total_C_H_Result
        total_plus_value = Total_C_H_Result_ACC

        If total_plus_value >= 0 Then
            total_plus_value_Minus = False
            total_plus_value_OP = total_plus_value
        Else
            total_plus_value_Minus = True
            total_plus_value_OP = Abs(total_plus_value)
        End If

        If total_plus_value_Minus = True Then
            If (total_plus_value_OP > compare_value_D) Then
                TBAR_OP_V = total_plus_value_OP / compare_value_D
                OP_TBar_Renew_Minus(TBAR_OP_V)
                total_plus_value = 0
                Total_C_H_Result_ACC = 0

            End If
        Else
            If (total_plus_value_OP > compare_value_U) Then
                TBAR_OP_V = total_plus_value_OP / compare_value_U
                OP_TBar_Renew_Plus(TBAR_OP_V)
                total_plus_value = 0
                Total_C_H_Result_ACC = 0

            End If

        End If


    End Sub
    Sub Bar_Control_THI()
        ' total_plus_value += Total_C_H_Result
        total_plus_value = Total_C_H_Result_ACC

        If total_plus_value >= 0 Then
            total_plus_value_Minus = False
            total_plus_value_OP = total_plus_value
        Else
            total_plus_value_Minus = True
            total_plus_value_OP = Abs(total_plus_value)
        End If

        If total_plus_value_Minus = True Then
            If (total_plus_value_OP > compare_value_D) Then
                TBAR_OP_V = total_plus_value_OP / compare_value_D
                OP_TBar_Renew_Minus(TBAR_OP_V)
                'OP_TBar_Renew_Plus(TBAR_OP_V)
                total_plus_value = 0
                Total_C_H_Result_ACC = 0

            End If
        Else
            If (total_plus_value_OP > compare_value_U) Then
                TBAR_OP_V = total_plus_value_OP / compare_value_D
                OP_TBar_Renew_Plus(TBAR_OP_V)
                'OP_TBar_Renew_Minus(TBAR_OP_V)
                total_plus_value = 0
                Total_C_H_Result_ACC = 0

            End If

        End If


    End Sub
    Sub cal_Area(I_Value As Byte, Cp_array() As Byte)

    End Sub

    Function cal_Area2(B_Value As Byte)
        If B_Value < 17 Then
            Return 0
        ElseIf B_Value >= 17 And B_Value < 34 Then
            Return 1
        ElseIf B_Value >= 34 And B_Value < 51 Then
            Return 2
        ElseIf B_Value >= 51 And B_Value < 68 Then
            Return 3
        ElseIf B_Value >= 68 And B_Value < 85 Then
            Return 4
        ElseIf B_Value >= 85 And B_Value < 102 Then
            Return 5
        ElseIf B_Value >= 102 And B_Value < 119 Then
            Return 6
        ElseIf B_Value >= 119 And B_Value < 136 Then
            Return 7
        ElseIf B_Value >= 136 And B_Value < 153 Then
            Return 8
        ElseIf B_Value >= 153 And B_Value < 170 Then
            Return 9
        ElseIf B_Value >= 170 And B_Value < 187 Then
            Return 10
        ElseIf B_Value >= 187 And B_Value < 204 Then
            Return 11
        ElseIf B_Value >= 204 And B_Value < 221 Then
            Return 12
        ElseIf B_Value >= 221 And B_Value < 238 Then
            Return 13
        ElseIf B_Value >= 238 Then
            Return 14
        End If

        Return 14

    End Function
    Sub cal_TbarChange_v()
        ' TBarChange_V()
    End Sub



    Public OP_TBar As TrackBar
    Sub OP_TBar_Setup(ByRef TBar As TrackBar)
        OP_TBar = TBar
        Use_Trg_bar = True
    End Sub
    Sub OP_TBar_Renew(PM_value As Integer)
        OP_TBar.Value = OP_TBar.Value + PM_value
    End Sub
    Sub OP_TBar_Renew_Plus(PM_value As Integer)
        'OP_TBar.Value = OP_TBar.Value + PM_value

        If OP_TBar.Value + PM_value <= 255 Then
            OP_TBar.Value = OP_TBar.Value + PM_value
        Else
            OP_TBar.Value = 255
        End If

    End Sub
    Sub OP_TBar_Renew_Minus(PM_value As Integer)
        If OP_TBar.Value - PM_value >= 0 Then
            OP_TBar.Value = OP_TBar.Value - PM_value
        Else
            OP_TBar.Value = 0
        End If

    End Sub
    Sub Update()

        ReNew_all()
        Cal_Total_C_H()

        If Use_Trg_bar = True Then
            Bar_Control()
        Else
            For i = 1 To 5

            Next

        End If

    End Sub
    Sub Update_THI()

        ReNew_all()
        Cal_Total_C_H()

        If Use_Trg_bar = True Then
            Bar_Control_THI()
        End If

    End Sub

    Sub Initialize()
        ENV.init()
        U1.SrcType = 255
        U2.SrcType = 255
        U3.SrcType = 255
        U4.SrcType = 255

        D1.SrcType = 255
        D2.SrcType = 255
        D3.SrcType = 255
        D4.SrcType = 255

        ChangeRate = 0
        Total_C_H_Result_ACC = 0
        Use_Trg_bar = False

        ReDim compare_value_U_Array(16)
        ReDim compare_value_D_Array(16)
    End Sub

    Sub SetupSrc()

    End Sub
End Class
'================================================
'
'   [ DW__ ]
'================================================
Public Class DW_Item
    Private Name As String
    Private Max As Single
    Private Min As Single
    Private trun_range As Integer
    Private offset As Integer
    Private Zero_Point As Integer
    Private Src As TextBox
    'Private Src As String

    Private p1 As Point
    Private p2 As Point
    Private _Ori_buffer(1000) As Single
    Private _buffer(1000) As Single
    Private gp As Byte
    Public EN As Boolean = True

    Dim penQQ As New Pen(Color.Green, 1)


    Public Property UserName() As String

        Get
            ' Gets the property value.
            Return Name
        End Get
        Set(ByVal Value As String)
            ' Sets the property value.
            Name = Value
        End Set
    End Property
    Public Property SetSrc() As TextBox
        'Public Property SetSrc() As String
        Get
            ' Gets the property value.
            Return Src
        End Get
        Set(ByVal Value As TextBox)
            ' Sets the property value.
            Src = Value
        End Set
    End Property
    Public Property Graphic() As Byte
        Get
            '        ' Gets the property value.
            Return gp
        End Get
        Set(ByVal Value As Byte)
            '        ' Sets the property value.
            gp = Value
        End Set
    End Property
    'Public Property buf(ByVal ii As Integer) As Integer

    '    Get

    '        Return _buffer(ii)  ' Gets the property value.
    '    End Get

    '    Set(ByVal Value As Integer)

    '        _buffer(ii) = Value ' Sets the property value.
    '    End Set

    'End Property
    Public Property bufs(ByVal ii As Integer) As Single

        Get
            ' Gets the property value.
            Return _Ori_buffer(ii)
        End Get

        Set(ByVal Value As Single)
            ' Sets the property value.
            _Ori_buffer(ii) = Value
        End Set

    End Property

    Public Property g_p1() As Point

        Get
            ' Gets the property value.
            Return p1
        End Get
        Set(value As Point)

        End Set

    End Property
    Public Property g_p2() As Point

        Get
            ' Gets the property value.
            Return p2
        End Get
        Set(value As Point)

        End Set

    End Property

    Public Property XX_CVT() As Point

        Get
            ' Gets the property value.
            Return p2
        End Get

        Set(value As Point)

        End Set

    End Property

    Public Sub New(ByVal UserName As String)
        ' Set the property value.
        Me.UserName = UserName
    End Sub
    '===============================================
    '   設定最大值, 最大值, 範圍量
    '===============================================
    Public Sub SET_Range(ByVal data As Single, ByVal data2 As Single, ByVal data3 As Single)
        Max = data
        Min = data2
        trun_range = data3
        'Src = data4

        'offset = data4 - data3
        initial(0)
    End Sub
    'Public Sub SET_Source(ByRef tb As String)
    '====================================
    '   設定來源
    '====================================
    Public Sub SET_Source(ByRef tb As TextBox)
        Src = tb
    End Sub
    '====================================
    '   設定 位置, 顏色, Offset
    '====================================
    Public Sub SET_Local(ByRef LB As Label, ByRef PicBox As PictureBox, ByRef xColor As Color, ByVal xOffset As Integer)
        offset = LB.Location.Y - PicBox.Location.Y - trun_range + LB.Height + xOffset
        'PointTarget = yy
        Zero_Point = offset
        penQQ.Color = xColor
        'p1.Y = p2.Y
    End Sub
    '====================================
    '   設定 顏色
    '====================================
    Public Sub SET_Color(ByRef xColor As Color)

        penQQ.Color = xColor

    End Sub
    Public Sub Return_Zero()
        offset = Zero_Point
        initial(4)

        'p1.Y = p2.Y
    End Sub

    Public Sub initial(ByVal setval As Integer)
        p1.X = 0
        'p1.Y = setval
        p1.Y = trun_range + offset
        'EN = True
    End Sub

    Public Sub get_new_addr2(ByVal data1 As Integer)
        p2.X = data1
        'p2.Y = 240 - ((data2 + 0.4) / 3.8 * 80)
        'p2.Y = (trun_range - (((_buffer(data1) - Min) / (Max - Min)) * trun_range)) + offset
        If data1 <> 0 Then
            p2.Y = _buffer(data1)
        Else
            p2.Y = offset
        End If
    End Sub
    '=============================================
    '   畫線
    '=============================================
    Public Sub LINE_CVT(ByVal data1 As Integer, ByVal dataX As Single)
        'p2.X = data1
        'p2.Y = 240 - ((data2 + 0.4) / 3.8 * 80)
        _Ori_buffer(data1) = dataX
        _buffer(data1) = (trun_range - (((dataX - Min) / (Max - Min)) * trun_range)) + offset
    End Sub
    Public Sub Get_New_Value(ByVal data1 As Integer)
        Src.Refresh()

        LINE_CVT(data1, Val(Src.Text))
        'LINE_CVT(data1, Val(Src))
        'LINE_CVT(data1, Src)
    End Sub

    Public Sub DrawPoint(ByRef GP As Graphics, ByRef ii As Integer)
        '計算 數值對應的座標點位址
        get_new_addr2(ii)

        'PointTarget.DrawLine(penQQ, p1, p2)
        '在Graphics上畫圖
        GP.DrawLine(penQQ, p1, p2)

        '更新座標位址
        renew_coordinates()
    End Sub
    Public Sub No_DrawPoint(ByRef GP As Graphics, ByRef ii As Integer)
        '只更新座標, 不繪圖

        '計算 數值對應的座標點位址
        get_new_addr2(ii)

        'PointTarget.DrawLine(penQQ, p1, p2)
        '在Graphics上畫圖
        'GP.DrawLine(penQQ, p1, p2)

        '更新座標位址
        renew_coordinates()
    End Sub

    Public Sub renew_coordinates()
        p1.X = p2.X
        p1.Y = p2.Y
    End Sub



    'Public Sub set_pen(ByVal data As Color, ByVal d2 As Integer)
    '    Penx.Color = data
    '    Penx.Width = d2
    'End Sub
End Class
Public Class Diag_CHK
    'Dim CHK_NO() As Integer = New Integer(3) {1, 3, 5, 7}
    Dim CHK_NO() As Integer = New Integer(3) {10, 11, 13, 15}
    Dim count As Byte
    Dim last As Byte
    Dim ok_count As Integer
    Dim resultx As Boolean
    Public Sub init()
        resultx = False
        count = 0
        last = 0
        ok_count = 0
    End Sub



    Public WriteOnly Property Enter() As Integer

        Set(ByVal Value As Integer)
            ' Sets the property value.
            If (Value = CHK_NO(count)) Then
                count += 1
                last = Value
                If (count = 4) Then
                    count = 0
                    resultx = True
                    ok_count += 1
                End If
            Else
                If (Value <> last) Then
                    count = 0
                    resultx = 0

                End If

            End If
        End Set
    End Property
    Public ReadOnly Property Result() As Boolean
        Get
            Return resultx
        End Get
    End Property
    Public ReadOnly Property OK() As Integer
        Get
            Return ok_count
        End Get
    End Property

End Class

Public Class Diag_CHK_2

    Dim CHK_NO() As Integer = New Integer(3) {2, 4, 6, 8}
    Dim count As Byte
    Dim last As Byte
    Dim ok_count As Integer
    Dim resultx As Boolean
    Public Sub init()
        resultx = False
        count = 0
        last = 0
        ok_count = 0
    End Sub



    Public WriteOnly Property Enter() As Integer


        Set(ByVal Value As Integer)
            ' Sets the property value.
            If (Value = CHK_NO(count)) Then
                count += 1
                last = Value
                If (count = 4) Then
                    count = 0
                    resultx = True
                    ok_count += 1
                End If
            Else
                If (Value <> last) Then
                    count = 0
                    resultx = 0

                End If

            End If
        End Set
    End Property
    Public ReadOnly Property Result() As Boolean
        Get
            Return resultx
        End Get
    End Property
    Public ReadOnly Property OK() As Integer
        Get
            Return ok_count
        End Get
    End Property

End Class
Namespace JFILE
    Class AFile
        Public PCB_LOG_Path As String = "123"
    End Class

End Namespace


Public Class IO_status
    '=========================================================
    '類別: 
    ' 0: OUT 
    ' 1: IN
    ' 2: PWM
    ' 3: CAN
    ' 4: UI/RAM
    ' 5: AD_Value
    ' 6: +電
    ' 7: GND
    ' 8: None
    ' 9: 其它
    '10: ES34
    '11: LIN
    '==========================================================

    Enum IO_Type

        Out = 0
        In_ = 1
        PWM = 2
        'PWM_In = 3
        CAN = 3
        UI_RAM = 4
        AD_V = 5
        VP = 6
        GND = 7
        None = 8
        Another = 9
        ES34 = 10
        LIN = 11
    End Enum

    Public CN_to_Item_Trans(65) As Byte
    Public Type(65) As IO_Type


    'Public Type(64 + 1) As Byte

    '類別: 
    '0 :None

    ' 0: OUT        橙
    ' 1: IN         綠
    ' 2: PWM OUT    藍
    ' 3: PWM IN     藍
    ' 4: CAN        黃
    ' 5: UI/RAM     黃
    ' 6: AD_Value   粉紅
    ' 7: +電        紅
    ' : GND        黑
    ' 8: None       灰
    ' 9: 其它       棕色


    Public rev(64 + 1) As Boolean   '反相處理
    Public OC(64 + 1) As Byte       '外部控制: 0:OFF, 1:ON, 2:System
    Public TBarray(64) As TextBox
    Public Property Set_rev(ByVal Valuex As Boolean) As Boolean             '資料反相設定

        Get
            ' Gets the property value.
            Return rev(Valuex)
        End Get

        Set(ByVal Value As Boolean)
            ' Sets the property value.
            rev(Value) = True
        End Set
    End Property

    Public Sub Init_Reset(ByVal Nstr As String)
        'Debug.Print(Nstr & "IO_Initial_Reset:")
        Dim i As Byte
        For i = 0 To 64
            rev(i) = False
            Type(i) = IO_Type.None
        Next
    End Sub
    Public Sub ECU_Initial_V25()
        Init_Reset("ECU_V25")

        'Debug.Print("ECU_Initial_V25")

        rev(0) = True 'CN1

        'rev(5) = True 'CN6
        rev(6) = True 'CN7
        rev(7) = True 'CN8
        'rev(8) = True 'CN9
        'rev(9) = True 'CN10
        'rev(10) = True 'CN11
        'rev(17) = True 'CN18

        'rev(18) = True 'CN19
        'rev(21) = True 'CN22
        'rev(22) = True 'CN23
        'rev(23) = True 'CN24
        rev(29) = True 'CN30
        'rev(30) = True 'CN31
        'rev(31) = True 'CN32
        rev(37) = True 'CN38
        rev(40) = True 'CN41
        rev(53) = True 'CN54
        rev(62) = True 'CN63

        '--------------------------------------
        Type(0) = IO_Type.In_     'CN1
        Type(1) = IO_Type.Out    'CN2
        Type(2) = IO_Type.Out      'CN3
        Type(3) = IO_Type.PWM      'CN4
        Type(4) = IO_Type.PWM      'CN5
        Type(5) = IO_Type.Out      'CN6
        Type(6) = IO_Type.Out      'CN7
        Type(7) = IO_Type.Out      'CN8

        Type(8) = IO_Type.Out      'CN9
        Type(9) = IO_Type.Out      'CN10
        Type(10) = IO_Type.Out    'CN11
        Type(11) = IO_Type.ES34    'CN12
        Type(12) = IO_Type.PWM     'CN13
        Type(13) = IO_Type.Out     'CN14
        Type(14) = IO_Type.Out     'CN15
        Type(15) = IO_Type.ES34    'CN16

        Type(16) = IO_Type.ES34    'CN17
        Type(17) = IO_Type.Out     'CN18
        Type(18) = IO_Type.Out     'CN19
        Type(19) = IO_Type.VP   'CN20
        Type(20) = IO_Type.ES34     'CN21
        Type(21) = IO_Type.Out     'CN22
        Type(22) = IO_Type.Out    'CN23
        Type(23) = IO_Type.Out    'CN24

        Type(24) = IO_Type.PWM    'CN25
        Type(25) = IO_Type.PWM    'CN26
        Type(26) = IO_Type.CAN   'CN27
        Type(27) = IO_Type.LIN    'CN28
        Type(28) = IO_Type.UI_RAM    'CN29
        Type(29) = IO_Type.In_    'CN30
        Type(30) = IO_Type.Out    'CN31
        Type(31) = IO_Type.Out    'CN32

        Type(32) = IO_Type.PWM   'CN33
        Type(33) = IO_Type.ES34     'CN34
        Type(34) = IO_Type.CAN    'CN35
        Type(35) = IO_Type.GND    'CN36
        Type(36) = IO_Type.UI_RAM    'CN37
        Type(37) = IO_Type.In_     'CN38
        Type(38) = IO_Type.GND   'CN39
        Type(39) = IO_Type.GND   'CN40

        Type(40) = IO_Type.In_     'CN41
        Type(41) = IO_Type.In_     'CN42
        Type(42) = IO_Type.In_    'CN43
        Type(43) = IO_Type.AD_V   'CN44
        Type(44) = IO_Type.AD_V     'CN45
        Type(45) = IO_Type.AD_V    'CN46
        Type(46) = IO_Type.AD_V    'CN47
        Type(47) = IO_Type.VP    'CN48

        Type(48) = IO_Type.VP    'CN49
        Type(49) = IO_Type.In_     'CN50
        Type(50) = IO_Type.VP    'CN51
        Type(51) = IO_Type.GND   'CN52
        Type(52) = IO_Type.GND  'CN53
        Type(53) = IO_Type.In_    'CN54
        Type(54) = IO_Type.In_     'CN55
        Type(55) = IO_Type.In_     'CN56

        Type(56) = IO_Type.AD_V     'CN57
        Type(57) = IO_Type.AD_V   'CN58
        Type(58) = IO_Type.AD_V    'CN59
        Type(59) = IO_Type.AD_V    'CN60
        Type(60) = IO_Type.AD_V     'CN61
        Type(61) = IO_Type.ES34     'CN62
        Type(62) = IO_Type.In_     'CN63
        Type(63) = IO_Type.In_    'CN64



    End Sub
    Public Sub ARD_Initial_V25()
        Init_Reset("ARD_V25")

        ' Debug.Print("ARD_Initial_V25")

        rev(5) = True   'CN6
        rev(8) = True   'CN9
        rev(9) = True   'CN10
        rev(10) = True   'CN11
        rev(17) = True   'CN18
        rev(18) = True   'CN19
        rev(21) = True   'CN22
        rev(22) = True   'CN23
        rev(23) = True   'CN24
        rev(30) = True   'CN31
        rev(31) = True   'CN32
        rev(62) = True   'CN63


        '類別: 
        ' 0: OUT 
        ' 1: IN
        ' 2: PWM
        ' 3: CAN
        ' 4: UI/RAM
        ' 5: AD_Value
        ' 6: +電
        ' 7: GND


        Type(0) = IO_Type.Out     'CN1
        Type(1) = IO_Type.In_     'CN2
        Type(2) = IO_Type.In_     'CN3
        Type(3) = IO_Type.In_     'CN4
        Type(4) = IO_Type.In_     'CN5
        Type(5) = IO_Type.In_     'CN6
        Type(6) = IO_Type.In_     'CN7
        Type(7) = IO_Type.In_     'CN8

        Type(8) = IO_Type.In_     'CN9
        Type(9) = IO_Type.In_     'CN10
        Type(10) = IO_Type.In_    'CN11
        Type(11) = IO_Type.ES34   'CN12
        Type(12) = IO_Type.In_    'CN13
        Type(13) = IO_Type.In_    'CN14
        Type(14) = IO_Type.In_    'CN15
        Type(15) = IO_Type.ES34    'CN16

        Type(16) = IO_Type.ES34    'CN17
        Type(17) = IO_Type.In_    'CN18
        Type(18) = IO_Type.In_    'CN19
        Type(19) = 1    'CN20
        Type(20) = IO_Type.ES34    'CN21
        Type(21) = IO_Type.In_    'CN22
        Type(22) = IO_Type.In_    'CN23
        Type(23) = 1    'CN24

        Type(24) = IO_Type.PWM    'CN25
        Type(25) = IO_Type.PWM    'CN26
        Type(26) = IO_Type.CAN    'CN27
        Type(27) = IO_Type.LIN    'CN28
        Type(28) = IO_Type.UI_RAM   'CN29
        Type(29) = IO_Type.Out     'CN30
        Type(30) = IO_Type.In_    'CN31
        Type(31) = IO_Type.In_    'CN32

        Type(32) = 1    'CN33
        Type(33) = IO_Type.ES34    'CN34
        Type(34) = IO_Type.CAN    'CN35
        Type(35) = IO_Type.GND   'CN36
        Type(36) = IO_Type.UI_RAM    'CN37
        Type(37) = IO_Type.Out   'CN38
        Type(38) = IO_Type.GND    'CN39
        Type(39) = IO_Type.GND    'CN40

        Type(40) = IO_Type.Out    'CN41
        Type(41) = IO_Type.Out    'CN42
        Type(42) = IO_Type.Out    'CN43
        Type(43) = IO_Type.AD_V    'CN44
        Type(44) = IO_Type.AD_V    'CN45
        Type(45) = IO_Type.AD_V    'CN46
        Type(46) = IO_Type.AD_V    'CN47
        Type(47) = IO_Type.AD_V   'CN48

        Type(48) = IO_Type.AD_V    'CN49
        Type(49) = IO_Type.Out    'CN50
        Type(50) = IO_Type.Out   'CN51
        Type(51) = IO_Type.GND    'CN52
        Type(52) = IO_Type.GND    'CN53
        Type(53) = IO_Type.Out    'CN54
        Type(54) = IO_Type.Out    'CN55
        Type(55) = IO_Type.Out    'CN56

        Type(56) = IO_Type.AD_V     'CN57
        Type(57) = IO_Type.AD_V     'CN58
        Type(58) = IO_Type.AD_V     'CN59
        Type(59) = IO_Type.AD_V     'CN60
        Type(60) = IO_Type.AD_V     'CN61
        Type(61) = IO_Type.ES34    'CN62
        Type(62) = IO_Type.Out    'CN63
        Type(63) = IO_Type.Out    'CN64

        '------------------------------


    End Sub
    Public Sub ECU_Initial_V30()
        Init_Reset("ECU_V30")


        ' Debug.Print("ECU_Initial_V30")
        rev(0) = True   'CN1
        'rev(5) = True   'CN6

        'rev(6) = True   'CN7

        'rev(8) = True   'CN9
        'rev(9) = True   'CN10
        'rev(10) = True   'CN11
        'rev(17) = True   'CN18
        'rev(18) = True   'CN19
        'rev(21) = True   'CN22
        'rev(22) = True   'CN23
        'rev(23) = True   'CN24
        rev(29) = True   'CN30
        ' rev(30) = True   'CN31
        'rev(31) = True   'CN32
        rev(37) = True   'CN38

        '-----------------------------

        '類別: 
        ' 0: OUT 
        ' 1: IN
        ' 2: PWM
        ' 3: CAN
        ' 4: UI/RAM
        ' 5: AD_Value
        ' 6: +電
        ' 7: GND

        Type(0) = IO_Type.In_     'CN1
        Type(1) = IO_Type.None    'CN2
        Type(2) = IO_Type.None     'CN3
        Type(3) = IO_Type.None    'CN4
        Type(4) = IO_Type.None     'CN5
        Type(5) = IO_Type.Out     'CN6
        Type(6) = IO_Type.Out     'CN7
        Type(7) = IO_Type.None     'CN8

        Type(8) = IO_Type.Out     'CN9
        Type(9) = IO_Type.Out     'CN10
        Type(10) = IO_Type.Out    'CN11
        Type(11) = IO_Type.ES34   'CN12
        Type(12) = IO_Type.None    'CN13
        Type(13) = IO_Type.Out   'CN14
        Type(14) = IO_Type.Out    'CN15
        Type(15) = IO_Type.ES34    'CN16

        Type(16) = IO_Type.ES34    'CN17
        Type(17) = IO_Type.Out   'CN18
        Type(18) = IO_Type.Out    'CN19
        Type(19) = IO_Type.VP    'CN20
        Type(20) = IO_Type.ES34    'CN21
        Type(21) = IO_Type.Out    'CN22
        Type(22) = IO_Type.Out    'CN23
        Type(23) = IO_Type.Out   'CN24

        Type(24) = IO_Type.PWM    'CN25
        Type(25) = IO_Type.PWM    'CN26
        Type(26) = IO_Type.CAN    'CN27
        Type(27) = IO_Type.LIN    'CN28
        Type(28) = IO_Type.UI_RAM  'CN29
        Type(29) = IO_Type.In_     'CN30
        Type(30) = IO_Type.Out   'CN31
        Type(31) = IO_Type.Out    'CN32

        Type(32) = IO_Type.PWM    'CN33
        Type(33) = IO_Type.ES34    'CN34
        Type(34) = IO_Type.CAN    'CN35
        Type(35) = IO_Type.GND    'CN36
        Type(36) = IO_Type.UI_RAM    'CN37
        Type(37) = IO_Type.In_     'CN38
        Type(38) = IO_Type.GND    'CN39
        Type(39) = IO_Type.GND    'CN40

        Type(40) = IO_Type.None    'CN41
        Type(41) = IO_Type.None    'CN42
        Type(42) = IO_Type.None    'CN43
        Type(43) = IO_Type.AD_V    'CN44
        Type(44) = IO_Type.AD_V   'CN45
        Type(45) = IO_Type.AD_V    'CN46
        Type(46) = IO_Type.AD_V    'CN47
        Type(47) = IO_Type.VP   'CN48

        Type(48) = IO_Type.VP   'CN49
        Type(49) = IO_Type.None    'CN50
        Type(50) = IO_Type.VP    'CN51
        Type(51) = IO_Type.GND     'CN52
        Type(52) = IO_Type.GND    'CN53
        Type(53) = IO_Type.None    'CN54
        Type(54) = IO_Type.None    'CN55
        Type(55) = IO_Type.None    'CN56

        Type(56) = IO_Type.AD_V    'CN57
        Type(57) = IO_Type.AD_V    'CN58
        Type(58) = IO_Type.AD_V    'CN59
        Type(59) = IO_Type.AD_V    'CN60
        Type(60) = IO_Type.AD_V    'CN61
        Type(61) = IO_Type.ES34    'CN62
        Type(62) = IO_Type.In_    'CN63
        Type(63) = IO_Type.None   'CN64

        '------------------------------



    End Sub

    Public Sub ARD_Initial_V30()
        Init_Reset("ARD_V30")

        ' Debug.Print("ARD_Initial_V30")

        rev(5) = True   'CN6
        rev(6) = True   'CN7
        rev(8) = True   'CN9
        rev(9) = True   'CN10
        rev(10) = True   'CN11
        rev(17) = True   'CN18
        rev(18) = True   'CN19
        rev(21) = True   'CN22
        rev(22) = True   'CN23
        rev(23) = True   'CN24
        rev(30) = True   'CN31
        rev(31) = True   'CN32
        rev(62) = True   'CN63

        '---------------------------

        '類別: 
        ' 0: OUT 
        ' 1: IN
        ' 2: PWM
        ' 3: CAN
        ' 4: UI/RAM
        ' 5: AD_Value
        ' 6: +電
        ' 7: GND

        Type(0) = IO_Type.Out     'CN1
        Type(1) = IO_Type.In_    'CN2
        Type(2) = IO_Type.In_     'CN3
        Type(3) = IO_Type.In_    'CN4
        Type(4) = IO_Type.In_     'CN5
        Type(5) = IO_Type.In_     'CN6
        Type(6) = IO_Type.In_     'CN7
        Type(7) = IO_Type.In_     'CN8

        Type(8) = IO_Type.In_     'CN9
        Type(9) = IO_Type.In_     'CN10
        Type(10) = IO_Type.In_    'CN11
        Type(11) = IO_Type.ES34   'CN12
        Type(12) = IO_Type.In_    'CN13
        Type(13) = IO_Type.In_   'CN14
        Type(14) = IO_Type.In_    'CN15
        Type(15) = IO_Type.ES34    'CN16

        Type(16) = IO_Type.ES34    'CN17
        Type(17) = IO_Type.In_   'CN18
        Type(18) = IO_Type.In_    'CN19
        Type(19) = IO_Type.VP    'CN20
        Type(20) = IO_Type.ES34    'CN21
        Type(21) = IO_Type.In_    'CN22
        Type(22) = IO_Type.In_    'CN23
        Type(23) = IO_Type.In_   'CN24

        Type(24) = IO_Type.PWM    'CN25
        Type(25) = IO_Type.PWM    'CN26
        Type(26) = IO_Type.CAN    'CN27
        Type(27) = IO_Type.LIN    'CN28
        Type(28) = IO_Type.UI_RAM  'CN29
        Type(29) = IO_Type.Out     'CN30
        Type(30) = IO_Type.In_   'CN31
        Type(31) = IO_Type.In_    'CN32

        Type(32) = IO_Type.PWM    'CN33
        Type(33) = IO_Type.ES34    'CN34
        Type(34) = IO_Type.CAN    'CN35
        Type(35) = IO_Type.GND    'CN36
        Type(36) = IO_Type.UI_RAM    'CN37
        Type(37) = IO_Type.Out     'CN38
        Type(38) = IO_Type.GND    'CN39
        Type(39) = IO_Type.GND    'CN40

        Type(40) = IO_Type.Out    'CN41
        Type(41) = IO_Type.Out    'CN42
        Type(42) = IO_Type.Out    'CN43
        Type(43) = IO_Type.AD_V    'CN44
        Type(44) = IO_Type.AD_V   'CN45
        Type(45) = IO_Type.AD_V    'CN46
        Type(46) = IO_Type.AD_V    'CN47
        Type(47) = IO_Type.VP   'CN48

        Type(48) = IO_Type.VP   'CN49
        Type(49) = IO_Type.Out    'CN50
        Type(50) = IO_Type.VP    'CN51
        Type(51) = IO_Type.GND     'CN52
        Type(52) = IO_Type.GND    'CN53
        Type(53) = IO_Type.Out    'CN54
        Type(54) = IO_Type.Out    'CN55
        Type(55) = IO_Type.Out    'CN56

        Type(56) = IO_Type.AD_V    'CN57
        Type(57) = IO_Type.AD_V    'CN58
        Type(58) = IO_Type.AD_V    'CN59
        Type(59) = IO_Type.AD_V    'CN60
        Type(60) = IO_Type.AD_V    'CN61
        Type(61) = IO_Type.ES34    'CN62
        Type(62) = IO_Type.Out    'CN63
        Type(63) = IO_Type.Out   'CN64

        '------------------------------



    End Sub

    Public Sub ECU_Initial_V37()
        Init_Reset("ECU_V37")


        'Debug.Print("ECU_Initial_V37")
        rev(0) = True   'CN1


        'rev(5) = True   'CN6

        'rev(6) = True   'CN7

        'rev(8) = True   'CN9
        'rev(9) = True   'CN10
        'rev(10) = True   'CN11
        'rev(17) = True   'CN18
        'rev(18) = True   'CN19
        'rev(21) = True   'CN22
        rev(19) = True   'CN20

        rev(22) = True   'CN23  SREF
        'rev(23) = True   'CN24
        rev(29) = True   'CN30
        ' rev(30) = True   'CN31
        'rev(31) = True   'CN32
        rev(37) = True   'CN38

        rev(50) = True   'CN51  - HW_EN
        rev(53) = True   'CN54  - HW_EN

        rev(61) = True   'CN62  - HW_EN

        rev(62) = True   'CN63


        '-----------------------------

        '類別: 
        ' 0: OUT 
        ' 1: IN
        ' 2: PWM
        ' 3: CAN
        ' 4: UI/RAM
        ' 5: AD_Value
        ' 6: +電
        ' 7: GND

        Type(0) = IO_Type.In_     'CN1
        Type(1) = IO_Type.None    'CN2
        Type(2) = IO_Type.None     'CN3 HVIL_VOUT
        Type(3) = IO_Type.None    'CN4  HVIL_I
        Type(4) = IO_Type.None     'CN5   
        Type(5) = IO_Type.None     'CN6   HVIL_VIN
        Type(6) = IO_Type.None     'CN7
        Type(7) = IO_Type.Out     'CN8

        Type(8) = IO_Type.Out     'CN9
        Type(9) = IO_Type.Out     'CN10
        Type(10) = IO_Type.Out    'CN11
        Type(11) = IO_Type.Out   'CN12
        Type(12) = IO_Type.None    'CN13
        Type(13) = IO_Type.Out   'CN14
        Type(14) = IO_Type.Out    'CN15
        Type(15) = IO_Type.None    'CN16

        Type(16) = IO_Type.ES34    'CN17
        Type(17) = IO_Type.ES34   'CN18
        Type(18) = IO_Type.Out    'CN19
        Type(19) = IO_Type.ES34    'CN20
        Type(20) = IO_Type.ES34    'CN21
        Type(21) = IO_Type.None    'CN22    LIN_1
        Type(22) = IO_Type.In_    'CN23
        Type(23) = IO_Type.Out   'CN24

        Type(24) = IO_Type.PWM    'CN25
        Type(25) = IO_Type.PWM    'CN26
        Type(26) = IO_Type.CAN    'CN27
        Type(27) = IO_Type.LIN    'CN28
        Type(28) = IO_Type.UI_RAM  'CN29
        Type(29) = IO_Type.In_     'CN30
        Type(30) = IO_Type.Out   'CN31
        Type(31) = IO_Type.Out    'CN32

        Type(32) = IO_Type.Out    'CN33
        Type(33) = IO_Type.Out    'CN34
        Type(34) = IO_Type.None    'CN35
        Type(35) = IO_Type.CAN    'CN36
        Type(36) = IO_Type.UI_RAM    'CN37
        Type(37) = IO_Type.In_     'CN38
        Type(38) = IO_Type.GND    'CN39
        Type(39) = IO_Type.GND    'CN40

        Type(40) = IO_Type.PWM    'CN41
        Type(41) = IO_Type.AD_V    'CN42
        Type(42) = IO_Type.AD_V    'CN43
        Type(43) = IO_Type.AD_V    'CN44
        Type(44) = IO_Type.AD_V   'CN45
        Type(45) = IO_Type.AD_V    'CN46
        Type(46) = IO_Type.AD_V    'CN47
        Type(47) = IO_Type.VP   'CN48

        Type(48) = IO_Type.VP   'CN49
        Type(49) = IO_Type.None    'CN50
        Type(50) = IO_Type.VP    'CN51
        Type(51) = IO_Type.ES34     'CN52
        Type(52) = IO_Type.Out    'CN53
        Type(53) = IO_Type.ES34    'CN54
        Type(54) = IO_Type.VP    'CN55
        Type(55) = IO_Type.None    'CN56 LIN_2

        Type(56) = IO_Type.AD_V    'CN57
        Type(57) = IO_Type.AD_V    'CN58
        Type(58) = IO_Type.AD_V    'CN59
        Type(59) = IO_Type.AD_V    'CN60
        Type(60) = IO_Type.None    'CN61
        Type(61) = IO_Type.In_    'CN62
        Type(62) = IO_Type.In_    'CN63
        Type(63) = IO_Type.Out   'CN64

        '------------------------------



    End Sub
    Public Sub ARD_Initial_V37()
        Init_Reset("ARD_V37")

        'Debug.Print("ARD_Initial_V37")

        rev(5) = True   'CN6
        rev(6) = True   'CN7

        rev(7) = True   'CN8
        rev(8) = True   'CN9
        rev(9) = True   'CN10
        rev(10) = True   'CN11

        rev(11) = True   'CN12
        rev(17) = True   'CN18
        rev(18) = True   'CN19
        rev(21) = True   'CN22
        rev(22) = True   'CN23
        rev(23) = True   'CN24
        rev(30) = True   'CN31
        rev(31) = True   'CN32
        rev(32) = True   'CN33
        rev(33) = True   'CN34
        rev(52) = True   'CN53

        rev(61) = True   'CN62 - HW_EN

        rev(62) = True   'CN63
        rev(63) = True   'CN64


        '---------------------------

        '類別: 
        ' 0: OUT 
        ' 1: IN
        ' 2: PWM
        ' 3: CAN
        ' 4: UI/RAM
        ' 5: AD_Value
        ' 6: +電
        ' 7: GND


        'Type(0) = IO_Type.In_     'CN1
        'Type(1) = IO_Type.None    'CN2
        'Type(2) = IO_Type.None     'CN3 HVIL_VOUT
        'Type(3) = IO_Type.None    'CN4  HVIL_I
        'Type(4) = IO_Type.None     'CN5   
        'Type(5) = IO_Type.None     'CN6   HVIL_VIN
        'Type(6) = IO_Type.None     'CN7
        'Type(7) = IO_Type.Out     'CN8

        'Type(8) = IO_Type.Out     'CN9
        'Type(9) = IO_Type.Out     'CN10
        'Type(10) = IO_Type.Out    'CN11
        'Type(11) = IO_Type.Out   'CN12
        'Type(12) = IO_Type.None    'CN13
        'Type(13) = IO_Type.Out   'CN14
        'Type(14) = IO_Type.Out    'CN15
        'Type(15) = IO_Type.None    'CN16


        Type(0) = IO_Type.Out     'CN1
        Type(1) = IO_Type.None    'CN2
        Type(2) = IO_Type.None     'CN3
        Type(3) = IO_Type.None    'CN4
        Type(4) = IO_Type.None     'CN5
        Type(5) = IO_Type.None     'CN6
        Type(6) = IO_Type.None     'CN7
        Type(7) = IO_Type.In_     'CN8

        Type(8) = IO_Type.In_     'CN9
        Type(9) = IO_Type.In_     'CN10
        Type(10) = IO_Type.In_    'CN11
        Type(11) = IO_Type.In_   'CN12
        Type(12) = IO_Type.None    'CN13
        Type(13) = IO_Type.In_   'CN14
        Type(14) = IO_Type.In_    'CN15
        Type(15) = IO_Type.None    'CN16

        'Type(16) = IO_Type.ES34    'CN17
        'Type(17) = IO_Type.ES34   'CN18
        'Type(18) = IO_Type.Out    'CN19
        'Type(19) = IO_Type.ES34    'CN20
        'Type(20) = IO_Type.ES34    'CN21
        'Type(21) = IO_Type.None    'CN22    LIN_1
        'Type(22) = IO_Type.In_    'CN23
        'Type(23) = IO_Type.Out   'CN24

        'Type(24) = IO_Type.PWM    'CN25
        'Type(25) = IO_Type.PWM    'CN26
        'Type(26) = IO_Type.CAN    'CN27
        'Type(27) = IO_Type.LIN    'CN28
        'Type(28) = IO_Type.UI_RAM  'CN29
        'Type(29) = IO_Type.In_     'CN30
        'Type(30) = IO_Type.Out   'CN31
        'Type(31) = IO_Type.Out    'CN32

        Type(16) = IO_Type.ES34    'CN17
        Type(17) = IO_Type.ES34   'CN18
        Type(18) = IO_Type.In_    'CN19
        Type(19) = IO_Type.ES34    'CN20
        Type(20) = IO_Type.ES34    'CN21
        Type(21) = IO_Type.None    'CN22
        Type(22) = IO_Type.Out    'CN23
        Type(23) = IO_Type.In_   'CN24

        Type(24) = IO_Type.PWM    'CN25
        Type(25) = IO_Type.PWM    'CN26
        Type(26) = IO_Type.CAN    'CN27
        Type(27) = IO_Type.LIN    'CN28
        Type(28) = IO_Type.UI_RAM  'CN29
        Type(29) = IO_Type.Out     'CN30
        Type(30) = IO_Type.In_   'CN31
        Type(31) = IO_Type.In_    'CN32

        'Type(32) = IO_Type.Out    'CN33
        'Type(33) = IO_Type.Out    'CN34
        'Type(34) = IO_Type.None    'CN35
        'Type(35) = IO_Type.CAN    'CN36
        'Type(36) = IO_Type.UI_RAM    'CN37
        'Type(37) = IO_Type.In_     'CN38
        'Type(38) = IO_Type.GND    'CN39
        'Type(39) = IO_Type.GND    'CN40

        'Type(40) = IO_Type.PWM    'CN41
        'Type(41) = IO_Type.AD_V    'CN42
        'Type(42) = IO_Type.AD_V    'CN43
        'Type(43) = IO_Type.AD_V    'CN44
        'Type(44) = IO_Type.AD_V   'CN45
        'Type(45) = IO_Type.AD_V    'CN46
        'Type(46) = IO_Type.AD_V    'CN47
        'Type(47) = IO_Type.VP   'CN48

        Type(32) = IO_Type.In_    'CN33
        Type(33) = IO_Type.In_    'CN34
        Type(34) = IO_Type.None    'CN35
        Type(35) = IO_Type.CAN    'CN36
        Type(36) = IO_Type.UI_RAM    'CN37
        Type(37) = IO_Type.Out     'CN38
        Type(38) = IO_Type.GND    'CN39
        Type(39) = IO_Type.GND    'CN40

        Type(40) = IO_Type.PWM    'CN41
        Type(41) = IO_Type.AD_V    'CN42
        Type(42) = IO_Type.AD_V    'CN43
        Type(43) = IO_Type.AD_V    'CN44
        Type(44) = IO_Type.AD_V   'CN45
        Type(45) = IO_Type.AD_V    'CN46
        Type(46) = IO_Type.AD_V    'CN47
        Type(47) = IO_Type.VP   'CN48


        'Type(48) = IO_Type.VP   'CN49
        'Type(49) = IO_Type.None    'CN50
        'Type(50) = IO_Type.VP    'CN51
        'Type(51) = IO_Type.ES34     'CN52
        'Type(52) = IO_Type.Out    'CN53
        'Type(53) = IO_Type.ES34    'CN54
        'Type(54) = IO_Type.VP    'CN55
        'Type(55) = IO_Type.None    'CN56 LIN_2

        'Type(56) = IO_Type.AD_V    'CN57
        'Type(57) = IO_Type.AD_V    'CN58
        'Type(58) = IO_Type.AD_V    'CN59
        'Type(59) = IO_Type.AD_V    'CN60
        'Type(60) = IO_Type.None    'CN61
        'Type(61) = IO_Type.In_    'CN62
        'Type(62) = IO_Type.In_    'CN63
        'Type(63) = IO_Type.Out   'CN64

        Type(48) = IO_Type.VP   'CN49
        Type(49) = IO_Type.None    'CN50
        Type(50) = IO_Type.VP    'CN51
        Type(51) = IO_Type.ES34     'CN52
        Type(52) = IO_Type.In_    'CN53
        Type(53) = IO_Type.ES34    'CN54
        Type(54) = IO_Type.VP    'CN55
        Type(55) = IO_Type.None    'CN56

        Type(56) = IO_Type.AD_V    'CN57
        Type(57) = IO_Type.AD_V    'CN58
        Type(58) = IO_Type.AD_V    'CN59
        Type(59) = IO_Type.AD_V    'CN60
        Type(60) = IO_Type.None    'CN61
        Type(61) = IO_Type.Out    'CN62
        Type(62) = IO_Type.Out    'CN63
        Type(63) = IO_Type.In_   'CN64

        '------------------------------



    End Sub
    Public Sub ECU_Initial_V50()
        Init_Reset("ECU_V50")


        'Debug.Print("ECU_Initial_V50")


        rev(0) = True 'CN1

        'rev(5) = True 'CN6
        'rev(6) = True 'CN7
        rev(7) = True 'CN8
        'rev(8) = True 'CN9
        'rev(9) = True 'CN10
        'rev(10) = True 'CN11
        'rev(17) = True 'CN18

        'rev(18) = True 'CN19
        'rev(21) = True 'CN22
        'rev(22) = True 'CN23
        'rev(23) = True 'CN24
        rev(29) = True 'CN30
        'rev(30) = True 'CN31
        'rev(31) = True 'CN32
        rev(37) = True 'CN38
        rev(40) = True 'CN41
        rev(49) = True 'CN50    'SOA
        rev(53) = True 'CN54
        rev(62) = True 'CN63

        '--------------------------------------
        Type(0) = IO_Type.In_     'CN1
        Type(1) = IO_Type.Out     'CN2
        Type(2) = IO_Type.Out      'CN3
        Type(3) = IO_Type.PWM      'CN4
        Type(4) = IO_Type.PWM      'CN5

        Type(5) = IO_Type.None      'CN6   空

        Type(6) = IO_Type.Out      'CN7
        Type(7) = IO_Type.Out      'CN8

        Type(8) = IO_Type.Out      'CN9

        Type(9) = IO_Type.Out      'CN10
        Type(10) = IO_Type.Out    'CN11

        Type(11) = IO_Type.None    'CN12

        Type(12) = IO_Type.Out     'CN13    EVBP

        Type(13) = IO_Type.Out     'CN14
        Type(14) = IO_Type.Out     'CN15

        Type(15) = IO_Type.In_    'CN16

        Type(16) = IO_Type.In_    'CN17

        Type(17) = IO_Type.Out     'CN18

        Type(18) = IO_Type.None     'CN19
        Type(19) = IO_Type.None   'CN20

        Type(20) = IO_Type.In_     'CN21

        Type(21) = IO_Type.Out     'CN22
        Type(22) = IO_Type.Out    'CN23
        Type(23) = IO_Type.Out    'CN24

        Type(24) = IO_Type.PWM    'CN25
        Type(25) = IO_Type.PWM    'CN26

        Type(26) = IO_Type.CAN   'CN27
        Type(27) = IO_Type.LIN    'CN28
        Type(28) = IO_Type.UI_RAM    'CN29
        Type(29) = IO_Type.In_    'CN30

        Type(30) = IO_Type.Out    'CN31
        Type(31) = IO_Type.Out    'CN32

        Type(32) = IO_Type.PWM   'CN33
        Type(33) = IO_Type.AD_V     'CN34
        Type(34) = IO_Type.CAN    'CN35
        Type(35) = IO_Type.GND    'CN36
        Type(36) = IO_Type.UI_RAM    'CN37
        Type(37) = IO_Type.In_     'CN38
        Type(38) = IO_Type.GND   'CN39

        Type(39) = IO_Type.GND   'CN40

        Type(40) = IO_Type.None     'CN41

        Type(41) = IO_Type.None     'CN42
        Type(42) = IO_Type.In_    'CN43

        Type(43) = IO_Type.AD_V   'CN44
        Type(44) = IO_Type.AD_V     'CN45

        Type(45) = IO_Type.AD_V    'CN46
        Type(46) = IO_Type.AD_V    'CN47

        Type(47) = IO_Type.VP    'CN48

        Type(48) = IO_Type.VP    'CN49

        Type(49) = IO_Type.In_     'CN50

        Type(50) = IO_Type.VP    'CN51

        Type(51) = IO_Type.GND   'CN52
        Type(52) = IO_Type.GND  'CN53

        Type(53) = IO_Type.AD_V    'CN54

        Type(54) = IO_Type.None     'CN55

        Type(55) = IO_Type.In_     'CN56

        Type(56) = IO_Type.AD_V     'CN57
        Type(57) = IO_Type.AD_V   'CN58
        Type(58) = IO_Type.AD_V    'CN59
        Type(59) = IO_Type.AD_V    'CN60
        Type(60) = IO_Type.AD_V     'CN61
        Type(61) = IO_Type.None     'CN62
        Type(62) = IO_Type.In_     'CN63
        Type(63) = IO_Type.In_    'CN64


    End Sub
    Public Sub ARD_Initial_V50()
        Init_Reset("ARD_V50")

        ' Debug.Print("ARD_Initial_V50")

        'Init_Reset("ARD_V25")

        'Debug.Print("ARD_Initial_V25")

        rev(5) = True   'CN6
        rev(6) = True   'CN7
        rev(8) = True   'CN9
        rev(9) = True   'CN10
        rev(10) = True   'CN11
        rev(17) = True   'CN18
        rev(18) = True   'CN19
        rev(21) = True   'CN22
        rev(22) = True   'CN23
        rev(23) = True   'CN24
        rev(30) = True   'CN31
        rev(49) = True   'CN50 'SOA
        rev(31) = True   'CN32
        rev(62) = True   'CN63


        '類別: 
        ' 0: OUT 
        ' 1: IN
        ' 2: PWM
        ' 3: CAN
        ' 4: UI/RAM
        ' 5: AD_Value
        ' 6: +電
        ' 7: GND


        Type(0) = IO_Type.Out     'CN1
        Type(1) = IO_Type.In_     'CN2
        Type(2) = IO_Type.In_     'CN3

        Type(3) = IO_Type.PWM     'CN4
        Type(4) = IO_Type.PWM     'CN5

        Type(5) = IO_Type.None     'CN6

        Type(6) = IO_Type.In_     'CN7
        Type(7) = IO_Type.In_     'CN8

        Type(8) = IO_Type.In_     'CN9
        Type(9) = IO_Type.In_     'CN10
        Type(10) = IO_Type.In_    'CN11

        Type(11) = IO_Type.None   'CN12

        Type(12) = IO_Type.In_    'CN13
        Type(13) = IO_Type.In_    'CN14
        Type(14) = IO_Type.In_    'CN15
        Type(15) = IO_Type.Out    'CN16

        Type(16) = IO_Type.Out    'CN17
        Type(17) = IO_Type.In_    'CN18

        Type(18) = IO_Type.None    'CN19
        Type(19) = IO_Type.None    'CN20

        Type(20) = IO_Type.Out    'CN21
        Type(21) = IO_Type.In_    'CN22
        Type(22) = IO_Type.In_    'CN23
        Type(23) = IO_Type.In_    'CN24

        Type(24) = IO_Type.PWM    'CN25
        Type(25) = IO_Type.PWM    'CN26

        Type(26) = IO_Type.CAN    'CN27
        Type(27) = IO_Type.LIN    'CN28

        Type(28) = IO_Type.UI_RAM   'CN29
        Type(29) = IO_Type.Out     'CN30

        Type(30) = IO_Type.In_    'CN31
        Type(31) = IO_Type.In_    'CN32

        Type(32) = IO_Type.PWM    'CN33

        Type(33) = IO_Type.AD_V    'CN34
        Type(34) = IO_Type.CAN    'CN35

        Type(35) = IO_Type.GND   'CN36
        Type(36) = IO_Type.UI_RAM    'CN37
        Type(37) = IO_Type.Out   'CN38

        Type(38) = IO_Type.GND    'CN39
        Type(39) = IO_Type.GND    'CN40

        Type(40) = IO_Type.None    'CN41

        Type(41) = IO_Type.None    'CN42

        Type(42) = IO_Type.Out    'CN43
        Type(43) = IO_Type.AD_V    'CN44
        Type(44) = IO_Type.AD_V    'CN45
        Type(45) = IO_Type.AD_V    'CN46
        Type(46) = IO_Type.AD_V    'CN47

        Type(47) = IO_Type.AD_V   'CN48

        Type(48) = IO_Type.AD_V    'CN49

        Type(49) = IO_Type.Out    'CN50
        Type(50) = IO_Type.Out   'CN51

        Type(51) = IO_Type.GND    'CN52
        Type(52) = IO_Type.GND    'CN53

        Type(53) = IO_Type.AD_V    'CN54
        Type(54) = IO_Type.None    'CN55

        Type(55) = IO_Type.Out    'CN56

        Type(56) = IO_Type.AD_V     'CN57
        Type(57) = IO_Type.AD_V     'CN58
        Type(58) = IO_Type.AD_V     'CN59
        Type(59) = IO_Type.AD_V     'CN60
        Type(60) = IO_Type.AD_V     'CN61

        Type(61) = IO_Type.None    'CN62
        Type(62) = IO_Type.Out    'CN63
        Type(63) = IO_Type.Out    'CN64

        '------------------------------



    End Sub
End Class
'Public Class Comp_CVT
'    Public CVT_TC12 As Char
'    Public CVT_FRO1 As Char
'    Public CVT_FRO2 As Char
'    Public CVT_PL1 As Char
'    Public CVT_PL2 As Char
'    Public CVT_PH1 As Char
'    Public CVT_PH2 As Char
'    Public CVT_FIN1 As Char
'    Public CVT_FIN2 As Char

'    Public ECC1 As Char
'    Public ECC2 As Char
'End Class
'Public Class ECU_Item
'    Public RLCDH1 As Boolean
'    Public RLCDH2 As Boolean
'    Public RLCDH3 As Boolean
'    Public RLCDL1 As Boolean
'    Public RLCDL2 As Boolean
'    Public SPH1 As PressV
'    Public SPL1 As PressV
'    Public SPH2 As PressV
'    Public SPL2 As PressV
'    Public THI As Single
'    Public TSET As Single
'    Public Y As String
'    Public AC_OP As Comp_CVT
'    Public ECC1_times As String
'    Public ECC1_CNR As String
'    Public ECC2_times As String
'    Public ECC2_CNR As String
'    Public CDSH_times As String
'    Public CDSL_times As String

'End Class