Attribute VB_Name = "xls2xlsx_or_xlsm"
Sub ConvertToXlsx()
    Dim strPath As String
    Dim strFile As String
    Dim strNewFile As String
    Dim fileOpened As Boolean
    Dim marcoFile As Boolean
    Dim intNewFileType As Integer
    Dim savedNewFile As Boolean
    Dim wbk As Workbook
    ' Path must end in trailing backslash
    strPath = "C:\"
    
    Dim colFiles As New Collection
    Debug.Print "Collecting files start"
    RecursiveDir colFiles, strPath, "*.xls", True
    Debug.Print "Collecting files Finished"
    
    For Each vFile In colFiles
        strFile = vFile
        If Right(strFile, 3) = "xls" Then
            fileOpened = True
            On Error Resume Next
            Set wbk = Workbooks.Open(Filename:=strFile)
            marcoFile = wbk.HasVBProject
            If Err Then
                fileOpened = False
            End If
            On Error GoTo 0
            
            If fileOpened = False Then
                Debug.Print "Unable to open: " & strFile
            Else
                Debug.Print "Processing: " & strFile
            
                If marcoFile Then
                    strNewFile = strFile & "m"
                    intNewFileType = xlOpenXMLWorkbookMacroEnabled
                Else
                    strNewFile = strFile & "x"
                    intNewFileType = xlOpenXMLWorkbook
                End If
                
                savedNewFile = True
                On Error Resume Next
                wbk.SaveAs Filename:=strNewFile & "x", FileFormat:=intNewFileType
                If Err Then
                    savedNewFile = False
                    Debug.Print "Cannot save file as: " & strNewFile
                End If
                On Error GoTo 0
                wbk.Close SaveChanges:=False
        
                'rename processed file
                If savedNewFile = True Then
                    Name strFile As strFile & "-"
                End If
            End If
        End If
    Next vFile
    
    Debug.Print "Finished updating excel files"
End Sub

' http://www.ammara.com/access_image_faq/recursive_folder_search.html

Public Function RecursiveDir(colFiles As Collection, _
                             strFolder As String, _
                             strFileSpec As String, _
                             bIncludeSubfolders As Boolean)

    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant
    Dim skipFolder As Integer

    'Add files in strFolder matching strFileSpec to colFiles
    strFolder = TrailingSlash(strFolder)
    On Error Resume Next
    strTemp = Dir(strFolder & strFileSpec)
    If Err Then
        strTemp = ""
    End If
    On Error GoTo 0
    
    Do While strTemp <> vbNullString
        colFiles.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Fill colFolders with list of subdirectories of strFolder
        On Error Resume Next
        strTemp = Dir(strFolder, vbDirectory)
        If Err Then
            'No access rights
            a = Err.Number
            b = Err.Description
            strTemp = ""
        End If
        On Error GoTo 0
    
        Do While strTemp <> vbNullString

            If (strTemp <> ".") And (strTemp <> "..") Then
                On Error Resume Next
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                    skipFolder = InStr(strTemp, "Windows") + InStr(strTemp, "Program Files") + InStr(strTemp, "PerfLogs")
                    If skipFolder > 0 Then
                        Debug.Print "Skipping folder: " & strTemp
                    Else
                        colFolders.Add strTemp
                    End If
                End If
                If Err Then
                    Debug.Print "Cannot acces: " & strTemp & ", in folder: " & strFolder
                End If
                On Error GoTo 0
            End If
            strTemp = Dir
        Loop

        'Call RecursiveDir for each subfolder in colFolders
        For Each vFolderName In colFolders
            Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
        Next vFolderName
    End If

End Function


Public Function TrailingSlash(strFolder As String) As String
    If Len(strFolder) > 0 Then
        If Right(strFolder, 1) = "\" Then
            TrailingSlash = strFolder
        Else
            TrailingSlash = strFolder & "\"
        End If
    End If
End Function
