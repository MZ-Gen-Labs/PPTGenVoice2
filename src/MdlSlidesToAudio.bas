Attribute VB_Name = "MdlSlidesToAudio"

Sub ExportNoteToAudiosAll()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        If Not ActiveWindow.Selection.SlideRange Is Nothing Then
            Set slds = ActiveWindow.Selection.SlideRange
        End If
    End If
    
    If Not slds Is Nothing Then
        UserForm1.Show vbModeless
        PlaceFormAtCenter UserForm1
                
        For Each sld In slds
            DoEvents ' 他のイベント処理を許可
            UserForm1.Label_Status.Caption = "処理中: " & sld.slideNumber & " / " & slds.Count
            UserForm1.Repaint ' 表示を更新
        
            If ExportNoteToAudio(sld) Then
                AddAudioToSlide sld
                AddAutoTransitToSlide sld
                TreattransitOnSlide sld, AddOperation
            Else
                Exit For
            End If
        Next sld
        
        Unload UserForm1
    End If
End Sub

Sub ExportNoteToAudios()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        If Not ActiveWindow.Selection.SlideRange Is Nothing Then
            Set slds = ActiveWindow.Selection.SlideRange
        End If
    End If
    
    If Not slds Is Nothing Then
        UserForm1.Show vbModeless
        PlaceFormAtCenter UserForm1
                
        For Each sld In slds
            DoEvents ' 他のイベント処理を許可
            UserForm1.Label_Status.Caption = "処理中: " & sld.slideNumber & " / " & slds.Count
            UserForm1.Repaint ' 表示を更新
        
            If ExportNoteToAudio(sld) Then
            Else
                Exit For
            End If
        Next sld
        
        Unload UserForm1
    End If
End Sub

Function ExportNoteToAudio(sld As Slide) As Boolean
    ExportNoteToAudio = False
    
    Dim baseUrl As String
    Dim speakerID As Long
    baseUrl = "http://localhost:" & CStr(voicePort)
    speakerID = voiceId
    VoiceInitialize baseUrl, speakerID

    Dim textFilePath As String
    Dim textBasepath As String
    Dim textFldrpath As String
    textFilePath = GetTextFullpath()
    textBasepath = GetTextBasepath()
    textFldrpath = GetTextFldrpath()
    Dim audioFldrpath As String
    If useAudioFolder Then
        audioFldrpath = textFldrpath & "\" & "audio"
    Else
        audioFldrpath = textFldrpath & "\" & textBasepath
    End If

    ' ノートのテキストを抽出する
    notesText = sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text
    If notesText <> "" Then
        ' スライド番号を取得する
        Dim slideNumber As Long
        baseFileName = sld.slideNumber
        
        saveFolder = audioFldrpath
    
        If GenerateVoiceFile(notesText, saveFolder, baseFileName) Then
            ExportNoteToAudio = True
        Else
            ' MsgBox "音声ファイルの生成に失敗しました。"
        End If
    End If
End Function


Sub PlaceFormAtCenter(frm As Object)
    ' アクティブなプレゼンテーションのウィンドウ情報を取得
    Dim pptLeft As Long
    Dim pptTop As Long
    Dim pptWidth As Long
        Dim pptHeight As Long
    With Application.ActiveWindow
        pptLeft = .Left
        pptTop = .Top
        pptWidth = .Width
        pptHeight = .Height
    End With

    ' ユーザーフォームのサイズを取得
    Dim frmWidth As Long
    Dim frmHeight As Long
    frmWidth = frm.Width
    frmHeight = frm.Height

    ' ユーザーフォームの表示位置を計算 (ウィンドウの中央)
    Dim leftPos As Long
    Dim topPos As Long
    leftPos = pptLeft + (pptWidth - frmWidth) / 2
    topPos = pptTop + (pptHeight - frmHeight) / 2

    ' ユーザーフォームの位置を設定
    frm.Left = leftPos
    frm.Top = topPos
End Sub
