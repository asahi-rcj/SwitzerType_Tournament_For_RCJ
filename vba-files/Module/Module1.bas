Attribute VB_Name = "Module1"

Sub Initialize_TeamForm()

    result = MsgBox("現在の情報をすべて初期化します。よろしいですか？", Buttons:=vbYesNo)

    If result = vbYes Then

        Team_Count = WorkSheets("Home").Cells(5, "E").Value
        Head_Text = WorkSheets("Home").Cells(7, "E").Value

        Base_Cell_Index = 4

        For i = 1 To 100 
            
            WorkSheets("Teams").Cells(Base_Cell_Index + (i - 1), "B").Value = ""
            WorkSheets("Teams").Cells(Base_Cell_Index + (i - 1), "C").Value = ""

        Next i

        For i = 1 To Team_Count 

            If i < 10 Then
                WorkSheets("Teams").Cells(Base_Cell_Index + (i - 1), "B").Value = Head_Text + "00" + CStr(i)
            ElseIf i < 100 Then
                WorkSheets("Teams").Cells(Base_Cell_Index + (i - 1), "B").Value = Head_Text + "0" + CStr(i)
            Else
                WorkSheets("Teams").Cells(Base_Cell_Index + (i - 1), "B").Value = Head_Text + CStr(i)
            End If

        Next i

        WorkSheets("ProgramData").Cells(4, "B").Value = 0

    End If

End Sub

Sub Initialize_RankingForm()
    
    Team_Count = WorkSheets("Home").Cells(5, "E").Value

    Base_Cell_Index_Teams = 4
    Base_Cell_Index_Ranking = 4

    For i = 1 To 100

        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "B").Value = ""
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "C").Value = ""
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "D").Value = ""
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "E").Value = ""
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "F").Value = ""
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "G").Value = ""
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "H").Value = ""

        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "J").Value = ""
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "K").Value = ""
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "L").Value = ""

    Next i

    For i = 1 To Team_Count

        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "B").Value = WorkSheets("Teams").Cells(Base_Cell_Index_Teams + (i - 1), "B").Value
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "C").Value = WorkSheets("Teams").Cells(Base_Cell_Index_Teams + (i - 1), "C").Value

        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "D").Value = 0
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "E").Value = 0
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "F").Value = 0
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "G").Value = 0
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "H").Value = 0
        
        WorkSheets("Ranking").Cells(Base_Cell_Index_Ranking + (i - 1), "J").Value = i

    Next i

End Sub

Sub Generate_Match()

    Team_Count = WorkSheets("Home").Cells(5, "E").Value
    Court_Count = WorkSheets("Home").Cells(6, "E").Value
    Game_Count = Application.WorksheetFunction.RoundUp(Team_Count / 2.0, 0)

    Now_Game_Count = Val(WorkSheets("ProgramData").Cells(4, "C").Value)

    Base_Cell_Index_Games = 4
    Now_Base_Cells_Games = Base_Cell_Index_Games + (1 + 1 + Game_Count + 1) * Now_Game_Count
    
    '**********************************
    '      ゲーム生成アルゴリズム
    '**********************************

    Dim Team_List As New collection
    Dim Fight_List As New collection

    For i = 1 To Team_Count
        Team_List.Add WorkSheets("Teams").Cells(4 + i - 1, "C").Value
    Next i

    Debug.Print(Team_List.Count)

    ' 初回のゲーム生成
    If Now_Game_Count = 0 Then
        For i = 1 To Team_Count
            Fight_List.Add Team_List(i)
        Next i
    Else
        For i = 1 To Team_Count
            Fight_List.Add Team_List(i)
        Next i
    End If


    '**********************************
    '         シートへの情報反映
    '**********************************
    
    WorkSheets("Games").Cells(Now_Base_Cells_Games, "B").Font.Bold = True
    WorkSheets("Games").Cells(Now_Base_Cells_Games, "B").Value = "ROUND" + CStr(Now_Game_Count + 1)
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "B").Font.Bold = True
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "B").Value = "GAME_ID"
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "C").Font.Bold = True
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "C").Value = "COURT"
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "D").Font.Bold = True
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "D").Value = "TEAM_A"
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "E").Font.Bold = True
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "E").Value = "TEAM_B"
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "F").Font.Bold = True
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "F").Value = "SCORE_A"
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "G").Font.Bold = True
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "G").Value = "SCORE_B"
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "H").Font.Bold = True
    WorkSheets("Games").Cells(Now_Base_Cells_Games + 1, "H").Value = "RESULT"

    For i = 1 To Game_Count
        WorkSheets("Games").Cells(Now_Base_Cells_Games + 1 + i, "B").Value = "GAME_" + Get_AlphaName(i)
        WorkSheets("Games").Cells(Now_Base_Cells_Games + 1 + i, "C").Value = Get_AlphaName(((i - 1) Mod Court_Count) + 1)
        WorkSheets("Games").Cells(Now_Base_Cells_Games + 1 + i, "D").Value = Fight_List(i * 2 - 1)
        WorkSheets("Games").Cells(Now_Base_Cells_Games + 1 + i, "E").Value = Fight_List(i * 2)
        WorkSheets("Games").Cells(Now_Base_Cells_Games + 1 + i, "F").Value = 0
        WorkSheets("Games").Cells(Now_Base_Cells_Games + 1 + i, "G").Value = 0
        WorkSheets("Games").Cells(Now_Base_Cells_Games + 1 + i, "H").FormulaR1C1 = "=IF(RC[-2]=RC[-1], ""DRAW"", IF(RC[-2]>RC[-1], RC[-4], RC[-3]))"
    Next i

    WorkSheets("ProgramData").Cells(4, "C").Value = Now_Game_Count + 1

End Sub

Sub Update_Ranking()
    'ランキングを更新する。

    Base_Cell_Index_Ranking = 4
    Team_Count = WorkSheets("Home").Cells(5, "E").Value
    Game_Count = Application.WorksheetFunction.RoundUp(Team_Count / 2.0, 0)

    Now_Game_Count = Val(WorkSheets("ProgramData").Cells(4, "C").Value)

    Base_Cell_Index_Games = 4

    Dim Points() As Variant
    Dim GF() As Variant
    Dim GA() As Variant
    Dim GD() As Variant
    Dim Team_Rank As Variant
    ReDim Points(Team_Count)
    ReDim GF(Team_Count)
    ReDim GA(Team_Count)
    ReDim GD(Team_Count)
    ReDim Team_Rank(Team_Count)
    
    For i = 0 To Team_Count - 1
        Points(i) = 0
        Team_Rank(i) = 1
    Next i

    If Now_Game_Count >= 1 Then
        
        For i = 0 To (Now_Game_Count - 1)

            For j = 1 To Game_Count

                Checking_Index = Base_Cell_Index_Games + (1 + 1 + Game_Count + 1) * (i) + 2 + (j - 1)

                '各対戦の結果を見て各チームに必要な値を加算
                Team_A = WorkSheets("Games").Cells(Checking_Index, "D").Value
                Team_B = WorkSheets("Games").Cells(Checking_Index, "E").Value
                Result = WorkSheets("Games").Cells(Checking_Index, "H").Value

                GF(GetTeamIndexFromName(Team_A)) = GF(GetTeamIndexFromName(Team_A)) + Val(WorkSheets("Games").Cells(Checking_Index, "F").Value)
                GF(GetTeamIndexFromName(Team_B)) = GF(GetTeamIndexFromName(Team_B)) + Val(WorkSheets("Games").Cells(Checking_Index, "G").Value)
                GA(GetTeamIndexFromName(Team_A)) = GA(GetTeamIndexFromName(Team_A)) + Val(WorkSheets("Games").Cells(Checking_Index, "G").Value)
                GA(GetTeamIndexFromName(Team_B)) = GA(GetTeamIndexFromName(Team_B)) + Val(WorkSheets("Games").Cells(Checking_Index, "F").Value)
                GD(GetTeamIndexFromName(Team_A)) = GF(GetTeamIndexFromName(Team_A)) - GA(GetTeamIndexFromName(Team_A))
                GD(GetTeamIndexFromName(Team_B)) = GF(GetTeamIndexFromName(Team_B)) - GA(GetTeamIndexFromName(Team_B))

                If Result = "DRAW" Then

                    Points(GetTeamIndexFromName(Team_A)) = Points(GetTeamIndexFromName(Team_A)) + 1
                    Points(GetTeamIndexFromName(Team_B)) = Points(GetTeamIndexFromName(Team_B)) + 1

                ElseIf Result = Team_A Then

                    Points(GetTeamIndexFromName(Team_A)) = Points(GetTeamIndexFromName(Team_A)) + 3

                ElseIf Result = Team_B Then

                    Points(GetTeamIndexFromName(Team_B)) = Points(GetTeamIndexFromName(Team_B)) + 3

                EndIf 

            Next j

        Next i

    End If

    '次に仮順位を決定する
    'これは勝ち点のみを考慮した順位

    Team_Rank = GetRankFromArray(Points(), Team_Count, 0)

    'あとはセルに反映させる

    For i = 0 To Team_Count - 1
        WorkSheets("Ranking").Cells(4 + (i), "D").Value = Points(i)
        WorkSheets("Ranking").Cells(4 + (i), "E").Value = GF(i)
        WorkSheets("Ranking").Cells(4 + (i), "F").Value = GA(i)
        WorkSheets("Ranking").Cells(4 + (i), "G").Value = GD(i)
        WorkSheets("Ranking").Cells(4 + (i), "H").Value = Team_Rank(i)
    Next i


    'ここからは本順位を計算してランク順に並べなおす作業...。

    Dim Team_Finally_Rank() As Variant
    ReDim Team_Finally_Rank(Team_Count)

    For i = 0 To Team_Count - 1

        Team_Finally_Rank(i) = -1

    Next i

    For i = 0 To Team_Count - 1

        If ArrayCountIF(Team_Rank, i + 1) = 1 Then

            Team_Finally_Rank(i) = IndexOf(Team_Rank, i + 1)

        ElseIf ArrayCountIF(Team_Rank, i + 1) > 1 Then

            CountOfSameRank = ArrayCountIF(Team_Rank, i + 1)
            Dim Team_SameRank_Index_ As collection
            Dim Team_SameRank_GF_ As collection
            Set Team_SameRank_Index_ = New collection
            Set Team_SameRank_GF_ = New collection

            Dim Team_SameRank_GF_Rank() As Variant
            Dim Team_SameRank_Index() As Variant
            Dim Team_SameRank_GF() As Variant
            ReDim Team_SameRank_GF_Rank(CountOfSameRank)
            ReDim Team_SameRank_Index(CountOfSameRank)
            ReDim Team_SameRank_GF(CountOfSameRank)

            For j = 0 To Team_Count - 1

                If Team_Rank(j) = i + 1 Then

                    Team_SameRank_Index_.Add j
                    Team_SameRank_GF_.Add GF(j)

                EndIf

            Next j

            For j = 0 To CountOfSameRank - 1

                Team_SameRank_Index(j) = Team_SameRank_Index_(j + 1)
                Team_SameRank_GF(j) = Team_SameRank_GF_(j + 1)

            Next j

            Team_SameRank_GF_Rank = GetRankFromArray(Team_SameRank_GF, CountOfSameRank, 0)

            For j = 0 To CountOfSameRank - 1

                If ArrayCountIF(Team_SameRank_GF_Rank, j + 1) = 1 Then

                    Team_Finally_Rank(i + j) = Team_SameRank_Index(IndexOf(Team_SameRank_GF_Rank, j + 1))

                ElseIf ArrayCountIF(Team_SameRank_GF_Rank, j + 1) > 1 Then

                    'ここは明日またやる
                    Team_Finally_Rank(i + j) = -1

                End If

            Next j
            
        EndIf

    Next i

    For i = 0 To Team_Count - 1

        WorkSheets("Ranking").Cells(4 + i, "L").Value = WorkSheets("Ranking").Cells(4 + Team_Finally_Rank(i), "C").Value 

    Next i

End Sub




Function IndexOf(TargetArray, element)
    For i = 0 To UBound(TargetArray)
        If TargetArray(i) = element Then Exit For
    Next
    IndexOf = i
End Function

Function ArrayCountIF(source_array, search_criteria) As Long
    Dim Counter As Long
    Dim a As Variant
    For Each a In source_array
        If a = search_criteria Then
            Counter = Counter + 1
        End If
    Next
        
    ArrayCountIF = Counter
End Function

Function GetRankFromArray(Check_Array, Array_Count, Check_Mode) As Variant

    Dim Team_Rank() As Variant
    ReDim Team_Rank(Array_Count)

    For i = 0 To Array_Count - 1
        Team_Rank(i) = 1
    Next i

    For i = 0 To Array_Count - 1
        
        Team_A_Point = Check_Array(i)

        For j = i + 1 To Array_Count - 1

            Team_B_Point = Check_Array(j)

            If Check_Mode = 0 Then
                If Team_A_Point < Team_B_Point Then
                    Team_Rank(i) = Team_Rank(i) + 1
                ElseIf Team_A_Point > Team_B_Point Then
                    Team_Rank(j) = Team_Rank(j) + 1
                EndIf
            Else
                If Team_A_Point > Team_B_Point Then
                    Team_Rank(i) = Team_Rank(i) + 1
                ElseIf Team_A_Point < Team_B_Point Then
                    Team_Rank(j) = Team_Rank(j) + 1
                EndIf
            EndIf

        Next j

    Next i
    
    GetRankFromArray = Team_Rank

End Function

Public Function GetTeamIndexFromName(TeamName) As Integer

    Base_Cell_Index_Teams = 4
    Team_Count = WorkSheets("Home").Cells(5, "E").Value

    For i = 0 To Team_Count - 1
        If Cells(Base_Cell_Index_Teams + i, "C") = TeamName Then
            GetTeamIndexFromName = i
            Exit For
        End If
    Next i
End Function







'概要：列ナンバーをアルファベット列名に変換
'第一引数：数字
'返却値：対応するアルファベット
'例)1→A
Public Function Get_AlphaName(ColomnNum) As String
    
    '変数定義
    Dim ColomnName As String
    
    '列位置→アルファベット列名変換
    ColomnName = Cells(1, ColomnNum).Address(True, False)
    ColomnName = Left(ColomnName, InStr(2, ColomnName, "$") - 1)
    
    '返却値
    Get_AlphaName = ColomnName
    
End Function

'概要：アルファベット列名から列数値に変換
'第一引数：アルファベット
'返却値：対応する数値
'例)A→1
Public Function Get_Alpha_ColumnNum(AlphaName As String) As Integer
    
    '返却値
    Get_Alpha_ColumnNum = Range(AlphaName & "1").Column
    
End Function