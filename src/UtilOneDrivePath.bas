Attribute VB_Name = "UtilOneDrivePath"
Option Explicit

' Onedriveフォルダ取得関数

' https://kuroihako.com/vba/onedriveurltolocalpath/
' パワーポイント用に以下のみ修正
'        PathSeparator = "/"
'        ' パワーポイントでは以下の処理がないためハードコード
'        ' PathSeparator = Application.PathSeparator


' [VBA]OneDriveで同期しているファイルまたはフォルダのURLをローカルパスに変換する関数
' Copyright (c) 2020-2023  黒箱
' This software is released under the GPLv3.
' このソフトウェアはGNU GPLv3の下でリリースされています。

'* @fn Public Function OneDriveUrlToLocalPath(ByRef Url As String) As String
'* @brief OneDriveのファイルURL又はフォルダURLをローカルパスに変換します。
'* @param[in] Url OneDrive内に保存されたのファイル又はフォルダのURL
'* @return Variant ローカルパスを返します。引数Urlにローカルパスに"https://"以外から始まる文字列を指定した場合、引数Urlを返します。
'* @details OneDriveのファイルURL又はフォルダURLをローカルパスに変換します。本関数は、ExcelブックがOneDrive内に格納されている場合に、Workbook.Path又はWorkbook.FullNameがURLを返す問題を解決するためのものです。
'*
Public Function OneDriveUrlToLocalPath(ByRef url As String) As String
Const OneDriveCommercialUrlPattern As String = "*my.sharepoint.com*" '法人向けOneDriveのURLか否かを判定するためのLike右辺値

    '引数がURLでない場合、引数はローカルパスと判断してそのまま返す。
    If Not (url Like "https://*") Then
        OneDriveUrlToLocalPath = url
        Exit Function
    End If
    
    'OneDriveのパスを取得しておく(パフォーマンス優先)。
    Static PathSeparator As String
    Static OneDriveCommercialPath As String
    Static OneDriveConsumerPath As String
    
    If (PathSeparator = "") Then
        PathSeparator = "/"
        ' パワーポイントでは以下の処理がないためハードコード
        ' PathSeparator = Application.PathSeparator
        
        '法人向けOneDrive(OneDrive for Business)のパス
        OneDriveCommercialPath = Environ("OneDriveCommercial")
        If (OneDriveCommercialPath = "") Then OneDriveCommercialPath = Environ("OneDrive")
        
        '個人向けOneDriveのパス
        OneDriveConsumerPath = Environ("OneDriveConsumer")
        If (OneDriveConsumerPath = "") Then OneDriveConsumerPath = Environ("OneDrive")

    End If
    
    '法人向けOneDrive：URL＝"https://会社名-my.sharepoint.com/personal/ユーザー名_domain_com/Documentsファイルパス")
    Dim FilePathPos As Long
    If (url Like OneDriveCommercialUrlPattern) Then
        FilePathPos = InStr(1, url, "/Documents") + 10 '10 = Len("/Documents")
        OneDriveUrlToLocalPath = OneDriveCommercialPath & Replace(Mid(url, FilePathPos), "/", PathSeparator)
        
    '個人向けOneDrive：URL＝"https://d.docs.live.net/CID番号/ファイルパス"
    Else
        FilePathPos = InStr(9, url, "/") '9 == Len("https://") + 1
        FilePathPos = InStr(FilePathPos + 1, url, "/")

        If (FilePathPos = 0) Then
            OneDriveUrlToLocalPath = OneDriveConsumerPath
        Else
            OneDriveUrlToLocalPath = OneDriveConsumerPath & Replace(Mid(url, FilePathPos), "/", PathSeparator)
        End If
    End If

End Function



