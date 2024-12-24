Attribute VB_Name = "MdlSlidesToAudio"

Sub TestCurrent()
    Dim sld As Slide
    Dim slds As SlideRange
    
    If doAllSlides Then
        Set slds = ActivePresentation.Slides.Range
    Else
        ' 選択が有効かを確認
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
            ' スライド範囲を取得
            If Not ActiveWindow.Selection.SlideRange Is Nothing Then
                Set slds = ActiveWindow.Selection.SlideRange
            End If
        End If
    End If
    
    For Each sld In slds
        ExportNoteToText sld
        AddAudioToSlide sld
        AddAutoTransitToSlide sld
        TreattransitOnSlide sld, AddOperation
    Next sld
End Sub

Sub ExportNoteToText(sld As Slide)
    Dim synthesizer As VoicevoxSynthesizer
    Set synthesizer = New VoicevoxSynthesizer
    Dim baseUrl As String
    Dim speakerID As Long
    baseUrl = "http://localhost:" & CStr(voicePort)
    speakerID = voiceId
    synthesizer.Initialize baseUrl, speakerID

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

    ' スライド番号を取得する
    Dim slideNumber As Long
    baseFileName = sld.slideNumber
    
    saveFolder = audioFldrpath


    If synthesizer.GenerateVoiceFile(notesText, saveFolder, baseFileName) Then
    Else
        ' MsgBox "音声ファイルの生成に失敗しました。"
    End If


End Sub
