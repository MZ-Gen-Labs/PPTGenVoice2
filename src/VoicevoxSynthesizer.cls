VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VoicevoxSynthesizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Windows API宣言
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As Long

Private baseUrl As String
Private speakerID As Long

' 初期化メソッド
Public Sub Initialize(ByVal apiUrl As String, ByVal speaker As Long)
    baseUrl = apiUrl
    speakerID = speaker
End Sub

' 音声ファイルを生成するメソッド
Public Function GenerateVoiceFile(ByVal inputText As String, ByVal saveFolder As String, ByVal baseFileName As String) As Boolean
    On Error GoTo ErrorHandler

    ' フォルダが存在しない場合は作成
    If Dir(saveFolder, vbDirectory) = "" Then
        MkDir saveFolder
    End If

    ' Audio Queryの作成
    Dim audioQueryJson As String
    audioQueryJson = CreateAudioQuery(inputText)
    If audioQueryJson = "" Then Exit Function

    ' 音声合成
    Dim audioData() As Byte
    audioData = SynthesizeAudio(audioQueryJson)
    If IsEmpty(audioData) Then Exit Function

    ' 保存先フォルダとファイルパスの生成
    If Right(saveFolder, 1) <> "\" Then
        saveFolder = saveFolder & "\"
    End If
    Dim wavFilePath As String
    wavFilePath = saveFolder & baseFileName & ".wav"

    ' WAVファイルの保存
    If SaveToFile(wavFilePath, audioData) Then
        ' WAVをMP3に変換
        Dim mp3FilePath As String
        mp3FilePath = ConvertWavToMp3(wavFilePath)

        If mp3FilePath <> "" Then
            GenerateVoiceFile = True
        Else
            MsgBox "MP3への変換に失敗しました。", vbExclamation
            GenerateVoiceFile = False
        End If
    Else
        MsgBox "WAVファイルの保存に失敗しました。", vbExclamation
        GenerateVoiceFile = False
    End If

    Exit Function

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    GenerateVoiceFile = False
End Function

' 音声合成のためのクエリを作成する
Private Function CreateAudioQuery(ByVal text As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim url As String
    url = baseUrl & "/audio_query?text=" & URLEncodeUTF8(text) & "&speaker=" & speakerID

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send ""

    If http.Status = 200 Then
        CreateAudioQuery = http.responseText
    Else
        MsgBox "Audio Queryに失敗しました。Status: " & http.Status & vbCrLf & http.responseText, vbExclamation
        CreateAudioQuery = ""
    End If
End Function

' 音声を合成する
Private Function SynthesizeAudio(ByVal audioQueryJson As String) As Byte()
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim url As String
    url = baseUrl & "/synthesis?speaker=" & speakerID

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send audioQueryJson

    If http.Status = 200 Then
        SynthesizeAudio = http.responseBody
    Else
        MsgBox "Synthesisに失敗しました。Status: " & http.Status, vbExclamation
        SynthesizeAudio = Empty
    End If
End Function

' ファイルにデータを保存する
Private Function SaveToFile(ByVal filePath As String, ByRef data() As Byte) As Boolean
    On Error GoTo ErrorHandler
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' バイナリデータ
    stream.Open
    stream.Write data
    stream.SaveToFile filePath, 2 ' 2は上書き保存
    stream.Close
    SaveToFile = True
    Exit Function

ErrorHandler:
    MsgBox "ファイル保存に失敗しました: " & Err.Description, vbCritical
    SaveToFile = False
End Function

Private Function URLEncodeUTF8(ByVal text As String) As String
    Dim bytes() As Byte
    Dim i As Long
    Dim result As String

    ' UTF-8バイト配列を取得
    bytes = StringToUtf8Bytes(text)

    ' バイト配列をループしてURLエンコード処理を実行
    For i = LBound(bytes) To UBound(bytes)
        Select Case bytes(i)
            Case 9, 10 ' Tab, LFをスキップ
                ' スキップする処理（何もしない）
            Case Else
                ' URLエンコードの条件に一致する場合、そのまま追加
                If (bytes(i) >= 48 And bytes(i) <= 57) Or _
                   (bytes(i) >= 65 And bytes(i) <= 90) Or _
                   (bytes(i) >= 97 And bytes(i) <= 122) Or _
                   InStr("~-_.", Chr(bytes(i))) > 0 Then          ' 特殊文字
                    result = result & Chr(bytes(i))
                Else
                    ' それ以外は%xx形式でエンコード
                    result = result & "%" & Right("0" & Hex(bytes(i)), 2)
                End If
        End Select
    Next i

    URLEncodeUTF8 = result
End Function


' UTF-8バイト配列を生成
Private Function StringToUtf8Bytes(ByVal text As String) As Byte()
    Dim byteCount As Long
    Dim utf8Bytes() As Byte

    ' 必要なバイト数を取得
    byteCount = WideCharToMultiByte(65001, 0, StrPtr(text), -1, 0, 0, 0, 0)

    ' バイト配列を確保
    ReDim utf8Bytes(0 To byteCount - 2) ' -2 は終端のNullバイトを除くため

    ' UTF-8エンコード
    WideCharToMultiByte 65001, 0, StrPtr(text), -1, VarPtr(utf8Bytes(0)), byteCount - 1, 0, 0

    StringToUtf8Bytes = utf8Bytes
End Function

' WAVファイルをMP3に変換する
Private Function ConvertWavToMp3(ByVal wavFilePath As String) As String
    On Error GoTo ErrorHandler

    ' FFmpegのパス（適宜修正してください）
    Dim ffmpegPath As String
    ffmpegPath = "C:\tool\ffmpeg\bin\ffmpeg.exe" ' FFmpegの実行ファイルのフルパス

    ' MP3ファイルのパスを生成
    Dim mp3FilePath As String
    mp3FilePath = Left(wavFilePath, InStrRev(wavFilePath, ".")) & "mp3"

    ' コマンドを構築
    Dim cmd As String
    cmd = Chr(34) & ffmpegPath & Chr(34) & " -y -i " & Chr(34) & wavFilePath & Chr(34) & " " & Chr(34) & mp3FilePath & Chr(34)

    ' FFmpegを実行
    Shell cmd, vbHide

    ' MP3ファイルのパスを返す
    ConvertWavToMp3 = mp3FilePath
    Exit Function

ErrorHandler:
    MsgBox "MP3変換中にエラーが発生しました: " & Err.Description, vbCritical
    ConvertWavToMp3 = ""
End Function




