Attribute VB_Name = "MdlAudioTreat"
Option Explicit

Enum operationType
    AddOperation
    ChangeOperation
    RemoveOperation
    DeleteOperation
End Enum

Sub AddAudioToSlides()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If ActivePresentation.Path = "" Then
        MsgBox "プレゼンテーションが保存されていません。一度ファイルを保存してから実行してください。", vbExclamation, "保存確認"
        Exit Sub
    End If

    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        On Error Resume Next ' ▼ エラーが発生してもプログラムを止めずに次へ進む
        
        ' 1. まずは現在の選択状態（サムネイル、図形、テキストなど）からスライドの取得を試みる
        Set slds = ActiveWindow.Selection.SlideRange
        
        ' 2. リボン操作中などで上記が失敗(Nothing)した場合、現在画面に表示されているスライドを取得
        If slds Is Nothing Then
            Set slds = ActivePresentation.Slides.Range(ActiveWindow.View.Slide.SlideIndex)
        End If
        
        On Error GoTo 0 ' ▲ エラー無視の設定を解除（以降の予期せぬエラーは通常通り表示する）
        
        ' 3. 万が一、何らかの理由でどうしても取得できなかった場合の安全装置
        If slds Is Nothing Then
            MsgBox "スライドを特定できませんでした。スライド画面を一度クリックしてから再実行してください。", vbExclamation, "スライド特定エラー"
            Exit Sub
        End If
    End If
    
    For Each sld In slds
        AddAudioToSlide sld
        AddAutoTransitToSlide sld
        TreattransitOnSlide sld, AddOperation
    Next sld
End Sub


Sub RemoveAudioFromSlides()
    Dim sld As Slide

    ' 各スライドをループ処理
    For Each sld In ActivePresentation.Slides
        RemoveAudioFromSlide sld
        RemoveAudioFromSlideRegacy sld
    Next sld
End Sub

Sub MoveAudioInSlides()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If ActivePresentation.Path = "" Then
        MsgBox "プレゼンテーションが保存されていません。一度ファイルを保存してから実行してください。", vbExclamation, "保存確認"
        Exit Sub
    End If

    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        On Error Resume Next ' ▼ エラーが発生してもプログラムを止めずに次へ進む
        
        ' 1. まずは現在の選択状態（サムネイル、図形、テキストなど）からスライドの取得を試みる
        Set slds = ActiveWindow.Selection.SlideRange
        
        ' 2. リボン操作中などで上記が失敗(Nothing)した場合、現在画面に表示されているスライドを取得
        If slds Is Nothing Then
            Set slds = ActivePresentation.Slides.Range(ActiveWindow.View.Slide.SlideIndex)
        End If
        
        On Error GoTo 0 ' ▲ エラー無視の設定を解除（以降の予期せぬエラーは通常通り表示する）
        
        ' 3. 万が一、何らかの理由でどうしても取得できなかった場合の安全装置
        If slds Is Nothing Then
            MsgBox "スライドを特定できませんでした。スライド画面を一度クリックしてから再実行してください。", vbExclamation, "スライド特定エラー"
            Exit Sub
        End If
    End If
    
    For Each sld In slds
        MoveAudioInSlide sld
        AddAutoTransitToSlide sld
        TreattransitOnSlide sld, AddOperation
    Next sld
End Sub


' スライドへの音声配置処理
Sub AddAudioToSlide(sld As Slide)
    Dim shp As Shape
    Dim filePath As String
    Dim audioFile As String
    Dim slideNumber As Long
    Dim effect As effect
    Dim presentationName As String

    slideNumber = sld.slideNumber

    If useAudioFolder Then
        filePath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\audio\"
    Else
        ' PowerPointファイルの名前（拡張子なし）を取得
        presentationName = Left(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1)
        filePath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\" & presentationName & "\"
    End If

    If Dir(filePath, vbDirectory) = "" Then
        Exit Sub ' フォルダが無ければ音声ファイルもないので処理をスキップ
    End If

    audioFile = ""
    ' 音声ファイルのパスを定義する
    If Dir(filePath & slideNumber & ".wav") <> "" Then
        audioFile = filePath & slideNumber & ".wav"
    ElseIf Dir(filePath & slideNumber & ".mp3") <> "" Then
        audioFile = filePath & slideNumber & ".mp3"
    End If

    If audioFile <> "" Then
        If doOverride Then
            RemoveAudioFromSlide sld
        End If
        ' スライドに音声ファイルを追加する
        Set shp = sld.Shapes.AddMediaObject2(audioFile, msoFalse, msoTrue, sld.Master.Width + audioXPosition, sld.Master.Height - 50)

        ' **ここに音声オブジェクトにタグを設定するコードを追加**
        shp.Tags.Add Name:="AudioObject", Value:="True" ' 例：AudioObjectという名前でTrueの値を設定

        ' 音声ファイルにアニメーション効果を追加する
        Set effect = sld.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectMediaPlay, Trigger:=msoAnimTriggerWithPrevious)
        ' アニメーションの開始を遅らせる
        effect.Timing.TriggerDelayTime = startDelay
    End If
End Sub

' スライドを3秒後に自動的に進めるように設定する
Sub AddAutoTransitToSlide(sld As Slide)
    sld.SlideShowTransition.AdvanceOnTime = msoTrue
    sld.SlideShowTransition.AdvanceTime = transitTime
End Sub

Sub RemoveAudioFromSlide(sld As Slide)
    Dim shp As Shape
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(i)

        ' 音声オブジェクトの場合
        If shp.Type = msoMedia Then
            If shp.Tags.Item("AudioObject") <> "" Then
                shp.Delete
                GoTo NextShape ' 削除したら次のシェイプへ
            End If
        End If

        ' 楕円の場合
        If shp.AutoShapeType = msoShapeOval Then
            If shp.Tags.Item("AudioControl") <> "" Then
                shp.Delete
                GoTo NextShape ' 削除したら次のシェイプへ
            End If
        End If
NextShape:
    Next i
End Sub

Sub MoveAudioInSlide(sld As Slide)
    ' スライドに配置されたすべてのシェイプをループする
    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Type = msoMedia Then
            If shp.Left = sld.Master.Width + 50 And shp.Top = sld.Master.Height - 50 Then
                GoTo MoveProcess
            ElseIf shp.Left = sld.Master.Width - 50 And shp.Top = sld.Master.Height - 50 Then
                GoTo MoveProcess
            ElseIf shp.Left = sld.Master.Width - 100 And shp.Top = sld.Master.Height - 50 Then
                GoTo MoveProcess
            ElseIf shp.Left = sld.Master.Width - 150 And shp.Top = sld.Master.Height - 50 Then
                GoTo MoveProcess
            ElseIf shp.Left = sld.Master.Width - 200 And shp.Top = sld.Master.Height - 50 Then
                GoTo MoveProcess
            ElseIf shp.Left = sld.Master.Width - 250 And shp.Top = sld.Master.Height - 50 Then
                GoTo MoveProcess
            End If
        End If
        GoTo NextShape
MoveProcess:
        shp.Left = sld.Master.Width + audioXPosition
        shp.Top = sld.Master.Height - 50
NextShape:
    Next shp
End Sub


Sub MoveAudioPosition(x As Integer, y As Integer)

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        ' スライドに配置されたすべてのシェイプをループする
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                If shp.Left = sld.Master.Width + 50 And shp.Top = sld.Master.Height - 50 Then
                    GoTo MoveProcess
                ElseIf shp.Left = sld.Master.Width - 50 And shp.Top = sld.Master.Height - 50 Then
                    GoTo MoveProcess
                ElseIf shp.Left = sld.Master.Width - 100 And shp.Top = sld.Master.Height - 50 Then
                    GoTo MoveProcess
                ElseIf shp.Left = sld.Master.Width - 150 And shp.Top = sld.Master.Height - 50 Then
                    GoTo MoveProcess
                ElseIf shp.Left = sld.Master.Width - 200 And shp.Top = sld.Master.Height - 50 Then
                    GoTo MoveProcess
                ElseIf shp.Left = sld.Master.Width - 250 And shp.Top = sld.Master.Height - 50 Then
                    GoTo MoveProcess
                End If
            End If
            GoTo NextShape
MoveProcess:
            shp.Left = sld.Master.Width + x
            shp.Top = sld.Master.Height + y
NextShape:
        Next shp
    Next sld
End Sub

Sub MakeVideoTransparent(shp As Shape)
    ' 動画オブジェクトを透明にするサブルーチン
    On Error Resume Next
    shp.Fill.Transparency = 1 ' 透明度を100%に設定する
    shp.line.Transparency = 1 ' 線の透明度を100%に設定する
    On Error GoTo 0
End Sub

Sub MakeAllVideosTransparent()
    ' 各スライドから動画オブジェクトを透明にするサブルーチン
    Dim sld As Slide
    Dim shp As Shape

    On Error GoTo ErrorHandler ' エラーハンドラの定義

    ' プレゼンテーションの各スライドをループする
    For Each sld In ActivePresentation.Slides
        ' スライドに配置されたすべてのシェイプをループする
        For Each shp In sld.Shapes
            If shp.Type = msoMedia Then
                ' 特定の位置にあるシェイプのみを対象とする場合
                If shp.Left = sld.Master.Width - 50 And shp.Top = sld.Master.Height - 50 Then
                    shp.Fill.Transparency = 1
                End If
            End If
        Next shp
    Next sld

    Exit Sub ' 正常終了

ErrorHandler:
    MsgBox "エラーが発生しました。エラー番号: " & Err.Number & " エラーの内容: " & Err.Description, vbCritical
End Sub


Sub MakeAudioTransparent(shp As Shape, Optional transparencyLevel As Single = 1)
    ' 音声オブジェクトを透明にするサブルーチン

    On Error Resume Next
    If shp.Type = msoMedia Then
        If transparencyLevel = 1 Then
            shp.Fill.Visible = msoFalse ' 透明度100%の場合、オブジェクトを非表示にする
        Else
            shp.Fill.Transparency = transparencyLevel ' 指定された透明度を設定する
            shp.line.Transparency = transparencyLevel ' 指定された透明度を設定する
        End If
    End If
    On Error GoTo 0
End Sub


Sub TreattransitOnSlide(sld As Slide, optype As operationType)
    Dim shpcnt As Integer
    Dim shp As Shape
    Dim eff As effect
    shpcnt = 0
    
    For Each shp In sld.Shapes
        If shp.AutoShapeType = msoShapeOval Then
            ' AudioControlタグが付いている図形なら無条件でアニメーション処理へ
            If shp.Tags.Item("AudioControl") = "True" Then
                GoTo AnimationProcess
            End If
        End If
        GoTo NextShape
AnimationProcess:
        Dim i As Integer
        Dim effect As effect
        Select Case optype
            Case AddOperation, ChangeOperation
                If shp.AnimationSettings.Animate = msoTrue Then
                    For i = sld.TimeLine.MainSequence.Count To 1 Step -1
                        Set effect = sld.TimeLine.MainSequence(i)
                        If effect.Shape.Name = shp.Name Then
                            sld.TimeLine.MainSequence(i).Delete
                        End If
                    Next i
                End If
                
                ' アニメーションを追加
                Set eff = sld.TimeLine.MainSequence.AddEffect(Shape:=shp, effectId:=msoAnimEffectSplit)
                With eff.Timing
                    .Duration = endDelay
                    .TriggerType = msoAnimTriggerAfterPrevious ' 前のアニメーションの後に開始
                End With
                
                shpcnt = shpcnt + 1
            Case RemoveOperation
                If shp.AnimationSettings.Animate = msoTrue Then
                    For i = sld.TimeLine.MainSequence.Count To 1 Step -1
                        Set effect = sld.TimeLine.MainSequence(i)
                        If effect.Shape.Name = shp.Name Then
                            sld.TimeLine.MainSequence(i).Delete
                        End If
                    Next i
                End If
            Case DeleteOperation
                shp.Delete
        End Select
NextShape:
    Next shp
    
    If (optype = AddOperation) And (shpcnt = 0) Then
        Dim posX As Single
        Dim posY As Single
        ' 固定値の 50 ではなく、リボンで設定した circleXPosition を使う
        posX = sld.Master.Width + circleXPosition
        posY = sld.Master.Height - 50
    
        Set shp = sld.Shapes.AddShape(msoShapeOval, posX, posY, 50, 50)
        
        ' **ここにタグを設定するコードを追加**
        shp.Tags.Add Name:="AudioControl", Value:="True"
    
        shp.Fill.Transparency = 1
        shp.line.Transparency = 1

    
        If shp.AnimationSettings.Animate = msoTrue Then
            shp.AnimationSettings.Animate = msoFalse
        End If
        
        ' アニメーションを追加
        Set eff = sld.TimeLine.MainSequence.AddEffect(Shape:=shp, effectId:=msoAnimEffectSplit)
        With eff.Timing
            .Duration = endDelay
            .TriggerType = msoAnimTriggerAfterPrevious ' 前のアニメーションの後に開始
        End With
    End If
    
    Select Case optype
        Case AddOperation
            sld.SlideShowTransition.AdvanceOnTime = msoTrue
            sld.SlideShowTransition.AdvanceTime = transitTime
        Case ChangeOperation
            sld.SlideShowTransition.AdvanceOnTime = msoTrue
            sld.SlideShowTransition.AdvanceTime = transitTime
        Case RemoveOperation
            sld.SlideShowTransition.AdvanceOnTime = msoFalse
        Case DeleteOperation
            sld.SlideShowTransition.AdvanceOnTime = msoFalse
    End Select
End Sub
