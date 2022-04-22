'This is the top level method to create an Access or Excel file and/or import all the elements such as code modules, 
'forms, references, properties, tables, queries, etc...

Public Sub ImportAll(Optional Publishing As Boolean = False)
        ImportAll = False
        'On Error Resume Next
        If Publishing Then
            If MsgBox("Are you sure you want to Publish your project?" & vbCrLf & "This will overwrite the selected elements in your PUBLISH FILE?", 4) = 7 Then
                Exit Sub
            End If
        Else

            If MsgBox("Are you sure you want to OVERWRITE" & vbCrLf & "the selected elements in your DEV FILE?", 4) = 7 Then
                Exit Sub
            End If
        End If
        myMainWindow.importingAll = True

        myMainWindow.InitializeStandardObjects()
        If Publishing Then

            myMainWindow.filePath = myUtilities.ReplaceEnvironVars(myMainWindow.txtPublishFolder.Text) & "\" & myMainWindow.AppName.Text & IIf(myMainWindow.TestMode, "Test", "")
        Else
            myMainWindow.filePath = myUtilities.ReplaceEnvironVars(myMainWindow.txtsourceFile.Text & IIf(myMainWindow.TestMode, "Test", ""))
        End If
        If InStr(1, myMainWindow.txtsourceFile.Text, ".xls") <> 0 Then
            xl = True
            If Publishing Then
                myMainWindow.filePath += ".xlsm"
            End If
        ElseIf InStr(1, myMainWindow.txtsourceFile.Text, ".accdb") <> 0 Or InStr(1, myMainWindow.txtsourceFile.Text, ".mdb") <> 0 Then
            xl = False
            If Publishing Then
                myMainWindow.filePath += ".accdb"
            End If
        Else
            MsgBox("The file type is incorrect. Please select an excel or access file and try again.")
        End If

        If xl Then
            If myMainWindow.chkExcelFile.IsChecked Then
                CreateExcelFile(myMainWindow.filePath)
                myUtilities.Delay(5)
                newfile = True
            End If
        Else
            If Not myUtilities.FileExists(myMainWindow.filePath) Then ' ,  myMainWindow.AppName & ".accdb") Then
                CreateAccessFile(myMainWindow.filePath)
                newfile = True
            End If
        End If
        Dim myFile As String
        If Publishing Then
            myFile = myMainWindow.filePath
            'myFile = myUtilities.ReplaceEnvironVars(myMainWindow.PublishFolder & myMainWindow.AppName.Text & IIf(myMainWindow.xl, ".xlsm", ".accdb"))
        Else
            myFile = myUtilities.ReplaceEnvironVars(myMainWindow.txtsourceFile.Text)
        End If
        If (IsNothing(accApp) And IsNothing(xlApp)) Or Publishing Then
            If Publishing And Not IsNothing(accApp) And Not IsNothing(xlApp) Then
                Dim ans = MsgBox("We have to close your projcet file. Would you like to save it?", vbYesNo)
                ReleaseMSFile(True, ans)
            End If

            If Not OpenMSFile(myFile,, Publishing) Then
                Exit Sub
            End If
        End If

        If Publishing Then
            ImportReferenceFromGUID("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
        Else

        End If

        If xl Then
        Else
            If myMainWindow.chkForms.IsChecked Then
                ImportAccessForms()
            End If

            If myMainWindow.chkReports.IsChecked Then
                If myUtilities.FolderExists(myMainWindow.txtProjectFolder.Text & "\Reports\") Then
                    ImportAccessReports()
                End If
            End If

            If myMainWindow.chkMacros.IsChecked Then
                If myUtilities.FolderExists(myMainWindow.txtProjectFolder.Text & "\Macros\") Then
                    ImportMacros()
                End If
            End If

            If myMainWindow.chkQueries.IsChecked Then
                If myUtilities.FolderExists(myMainWindow.txtProjectFolder.Text & "\Queries\") Then
                    ImportQueries()
                End If
            End If

            If myMainWindow.chkTables.IsChecked Then
                If myUtilities.FolderExists(myMainWindow.txtProjectFolder.Text & "\Tables\") Then
                    ImportTables()
                End If
            End If

            If myMainWindow.chkRelationships.IsChecked Then
                ImportRelationships()
            End If

            If myMainWindow.chkSpecs.IsChecked Then
                ImportSpecs()
            End If

            If myMainWindow.chkProperties.IsChecked Or Publishing Then
                ImportDatabaseProperties(Publishing)
            End If

        End If
        If myMainWindow.chkProperties.IsChecked Then
              ImportVBProjectName()
        End If
        If myMainWindow.chkVBProject().IsChecked Or Publishing Then
            ImportVBProject(Publishing)
        Else
            If myMainWindow.chkReferences.IsChecked Or (myMainWindow.chkExcelFile.IsChecked And Publishing) Then
                ImportReferences()
            End If
        End If

        If Not Publishing And Not newfile Then
            myMainWindow.ExitGracefully(CloseWhenDone, , True)
        End If
        ImportAll = True
        myMainWindow.importingAll = False

    End Sub





'This is a method to compare the VB Project to the code repository. I creates dictionaries containing hash codes created from the content of the code modules,
'and compares the dictionaries to alert the user that there are  discrepencies between the project file and the repository.

Public Sub CompareHashCodes(Optional NewDict As Dictionary(Of String, String()) = Nothing, Optional CompareGit As Boolean = False, Optional CompareProject As Boolean = False)
        On Error GoTo ErrorHandler
        Dim vbObj As VBComponent
        Dim toImport As String = ""
        Dim toImportForms As String = ""
        myMainWindow.CodeModule.Text = ""
        Dim sfx As String
        Dim regex = New Regex("((" & vbCr & ")(?!" & vbLf & ")|(?<!" & vbCr & ")(" & vbLf & "))")
     
   For Each vbObj In vbComps
            If vbObj.Name <> "ClickOnce" Then
                Dim objName As String
                Dim projDiffs As Boolean = False
                Dim gitDiffs As Boolean = False
                Dim sbfolder As String = ""
                Select Case vbObj.Type
                    Case 2
                        sfx = ".cls"
                        objName = vbObj.Name
                        sbfolder = "Classes"
                    Case 100
                        If xl Then
                            sfx = ".cls"
                            objName = vbObj.Name
                            sbfolder = "Worksheets"
                        Else
                            sfx = ".txt"
                            objName = Right(vbObj.Name, Len(vbObj.Name) - 5)
                            sbfolder = "Forms"
                        End If
                    Case 1
                        sfx = ".bas"
                        objName = vbObj.Name
                        sbfolder = "Modules"
                    Case 3
                        If xl Then
                            sfx = ".frm"
                            objName = vbObj.Name
                            sbfolder = "Forms"
                        Else
                            sfx = ".frm"
                            objName = vbObj.Name
                            sbfolder = "Forms"
                        End If
                    Case Else
                        sfx = ""
                End Select
                If vbObj.Name Like "*CallGuide*" Then
                    Dim a As String = ""
                End If
                If CompareProject And Not CompareGit Then
                    If Not IsNothing(ProjectHash) Then
                        If If(IsNothing(NewDict), ProjectHash, NewDict).ContainsKey(objName) Then
                            If NewDict(vbObj.Name)(0) <> ProjectHash(objName)(0) Then
                                projDiffs = True
                            End If
                        Else
                            projDiffs = True
                        End If
                    End If
                End If
                If CompareProject And CompareGit Then
                    If (Not IsNothing(ProjectHash) Or Not IsNothing(NewDict)) And Not IsNothing(myMainWindow.gitHash) Then
                        If IIf(IsNothing(NewDict), ProjectHash, NewDict).ContainsKey(objName) Then
                            If myMainWindow.gitHash.ContainsKey(objName) Then
                                If myMainWindow.gitHash(objName)(0) <> IIf(IsNothing(NewDict), ProjectHash, NewDict)(objName)(0) Then
                                    Dim i As Int16 = 0
                                    Do
                                        projDiffs = False
                                        myMainWindow.gitHash(objName)(0) = UCase(myUtilities.TrimTrailingCRLF(myUtilities.ReadTextFile(myMainWindow.ProjectFolder & "\" & sbfolder & "\" & objName & sfx))).GetHashCode
                                        IIf(IsNothing(NewDict), ProjectHash, NewDict)(objName)(0) = UCase(myUtilities.TrimTrailingCRLF(myUtilities.ReadTextFile(myUtilities.ReplaceEnvironVars("%LOCALAPPDATA%\Temp\") & objName & sfx))).GetHashCode
                                        If myMainWindow.gitHash(objName)(0) <> IIf(IsNothing(NewDict), ProjectHash, NewDict)(objName)(0) Then
                                            projDiffs = True
                                        End If
                                        i += 1
                                    Loop Until projDiffs = False Or i = 3
                                End If
                            End If
                        Else
                            projDiffs = True
                        End If
                    End If
                End If
                If CompareGit And Not CompareProject Then
                    If Not myMainWindow.gitHash.ContainsKey(objName) Then
                        gitDiffs = True
                    Else
                        gitDiffs = UCase(myUtilities.TrimTrailingCRLF(myUtilities.ReadTextFile(myUtilities.ReplaceEnvironVars("%LOCALAPPDATA%\Temp\" & objName & sfx)))).GetHashCode() <>
                            UCase(myUtilities.TrimTrailingCRLF(myUtilities.ReadTextFile(myUtilities.ReplaceEnvironVars(myMainWindow.txtProjectFolder.Text & "\" & sbfolder & "\" & objName & sfx)))).GetHashCode()
                        If gitDiffs Then
                            gitDiffs = UCase(myUtilities.TrimTrailingCRLF(myUtilities.ReadTextFile(myUtilities.ReplaceEnvironVars("%LOCALAPPDATA%\Temp\" & objName & sfx)))).GetHashCode() <>
                            UCase(myUtilities.TrimTrailingCRLF(myUtilities.ReadTextFile(myUtilities.ReplaceEnvironVars(myMainWindow.txtProjectFolder.Text & "\" & sbfolder & "\" & objName & sfx)))).GetHashCode()
                        End If
                    End If
                End If
                If CompareGit And CompareProject Then

                End If
                If projDiffs Or gitDiffs Then
                    If If(IsNothing(NewDict), ProjectHash, NewDict)(objName)(1) = ".frm" And Not xl Then
                        If myMainWindow.Form.Text = vbNullString Then
                            myMainWindow.Form.Text = objName
                        Else
                            myMainWindow.Form.Text = myMainWindow.Form.Text & "," & objName
                        End If
                    Else
                        If myMainWindow.CodeModule.Text = vbNullString Then
                            myMainWindow.CodeModule.Text = objName
                        Else
                            myMainWindow.CodeModule.Text = myMainWindow.CodeModule.Text & "," & objName
                        End If
                    End If
                End If
            End If
        Next

        If CompareGit And CompareProject Then
            For Each kvp As KeyValuePair(Of String, String()) In IIf(IsNothing(NewDict), ProjectHash, NewDict)
                Dim addMod As Boolean
                If Not myMainWindow.gitHash.ContainsKey(kvp.Key) Then
                    addMod = True
                Else
                    addMod = IIf(IsNothing(NewDict), ProjectHash, NewDict)(kvp.Key)(0).GetHashCode <> kvp.Value(0).GetHashCode
                End If

                If addMod Then
                    If vbObj.Type = 100 And Not xl Then
                        If myMainWindow.Form.Text = vbNullString Then
                            myMainWindow.Form.Text = kvp.Key
                        Else
                            myMainWindow.Form.Text = myMainWindow.Form.Text & "," & kvp.Key
                        End If
                    Else
                        If myMainWindow.CodeModule.Text = vbNullString Then
                            myMainWindow.CodeModule.Text = kvp.Key
                        Else
                            myMainWindow.CodeModule.Text = myMainWindow.CodeModule.Text & "," & kvp.Key
                        End If
                    End If
                End If
            Next
        End If
        If Not IsNothing(NewDict) Or Not IsNothing(ProjectHash) Then
            For Each kvp As KeyValuePair(Of String, String()) In myMainWindow.gitHash
                If Not If(IsNothing(NewDict), ProjectHash, NewDict).ContainsKey(kvp.Key) Then
                    If myMainWindow.gitHash(kvp.Key)(1) = ".frm" And Not xl Then
                        If toImportForms = vbNullString Then
                            toImportForms = kvp.Key
                        Else
                            toImportForms += "," & kvp.Key
                        End If
                    Else
                        If toImport = vbNullString Then
                            toImport = kvp.Key
                        Else
                            toImport += "," & kvp.Key
                        End If
                    End If
                End If
            Next
        End If
        If Not toImport = vbNullString Or Not toImportForms = vbNullString Then
            MsgBox("There are modules in your repo, that aren't in your project file. " & vbCrLf _
                    & "You should import code before proceeding")
            myMainWindow.CodeModule.Text = toImport
            myMainWindow.Form.Text = toImportForms
        End If
        If Not IsNothing(ProjectHash) Then
            ProjectHash.Clear()
        End If

ExitHandler:
        vbObj = Nothing
        Exit Sub
ErrorHandler:
        MsgBox("Error in GetProjectHashCodes " & Err.Number & ", " & Err.Description & " at line number " & Erl())
        GoTo ExitHandler
        Resume Next
    End Sub
      
      
      
      'Thsi method will create or update a version of the project for deployment by doing such things as compacting and repairing, disabling the nav pane
      ', ribbon, and special keys, and saving as an accde file. It also creates a setup and uninstaller file.
      
     Private Sub PublishNow()
        Try
            Dim addUninstaller As Boolean
            Dim gitFolder As String
            Dim pubFolder As String = myUtilities.ReplaceEnvironVars(PublishFolder)
            gitFolder = IIf(Right(pubFolder, Len(pubFolder) - InStrRev(pubFolder, "\")) = cmbRepositoryName.Text, "", "/" & Right(pubFolder, Len(pubFolder) - InStrRev(pubFolder, "\")))

            SetTestMode()

            CopyClickOnceToRepo()
            If Not myMSFile.ImportAll(True) Then
                Exit Sub
            End If
            CopyEmailerToRepo()
            CreateANSIPCConverter()
            CreateVersionFile()
            addUninstaller = False
            If MsgBox("Do you want to add the uninstaller in the start menu?", MsgBoxStyle.YesNo, "Create Uninstaller") = MessageBoxResult.Yes Then
                addUninstaller = True
            End If
            CreateSetupFile(addUninstaller)
            CreateUninstaller()

            myUtilities.ZipFolder(pubFolder & "\" & "Setup", pubFolder & "\" & "Setup.zip")
            '
            'myMSFile.OpenMSFile(txtPublishFolder.Text & "\" & AppName.Text & ".accdb")

            If myMSFile.xl Then
                myMSFile.wb.Save()
                myMSFile.ReleaseMSFile(True, False)
                'myUtilities.KillTask("EXCEL.EXE")
            Else

                If chkHideNav.IsChecked Then
                    myMSFile.accApp.DoCmd.ShowToolbar("Ribbon", Access.AcShowToolbar.acToolbarNo)
                    myMSFile.accApp.DoCmd.NavigateTo("acNavigationCategoryObjectType")
                    myMSFile.accApp.DoCmd.RunCommand(Access.AcCommand.acCmdWindowHide)
                    myMSFile.accApp.DoCmd.NavigateTo("acNavigationCategoryObjectType")
                    myMSFile.accApp.DoCmd.RunCommand(Access.AcCommand.acCmdWindowHide)
                    myMSFile.SetProp(myMSFile.accApp.CurrentDb, "AllowShortcutMenus", False, 1)
                End If
                If chkSpecialKeys.IsChecked Then
                    myMSFile.DisableShift()
                End If
                myMSFile.accApp.Quit(Access.AcQuitOption.acQuitSaveAll)
                myMSFile.ReleaseMSFile(True, True)

                myUtilities.KillTask("MSACCESS.EXE")
            End If
            Dim myAppName As String
            If Not myMSFile.xl Then
                Dim newAccApp As New Access.Application
                myAppName = myUtilities.ReplaceEnvironVars(PublishFolder) & "\" & AppName.Text & IIf(TestMode, "Test", "") & ".accdb"
                Dim TempName As String = myUtilities.ReplaceEnvironVars(PublishFolder) & "\" & IIf(TestMode, "Test", "") & AppName.Text & "1.accdb"
                myUtilities.DeleteFile(myUtilities.ReplaceEnvironVars(PublishFolder) & "\" & IIf(TestMode, "Test", "") & AppName.Text & ".laccdb")
                myUtilities.DeleteFile(myUtilities.ReplaceEnvironVars(PublishFolder) & "\" & IIf(TestMode, "Test", "") & AppName.Text & "1.laccdb")
                myUtilities.DeleteFile(TempName)

                newAccApp.DBEngine.CompactDatabase(myAppName, TempName)
                myUtilities.RenameFile(TempName, myAppName)
                newAccApp.Quit(Access.AcQuitOption.acQuitSaveAll)
                newAccApp = Nothing

                myUtilities.KillTask("MSACCESS.EXE")
                If chkAccde.IsChecked Then
                    newAccApp = New Access.Application
                    newAccApp.AutomationSecurity = 1
                    newAccApp.UserControl = True
                    myUtilities.DeleteFile(txtPublishFolder.Text & "\" & AppName.Text & ".accde")
                    newAccApp.SysCmd(603, myAppName, myUtilities.ReplaceEnvironVars(PublishFolder) & "\" & AppName.Text & IIf(TestMode, "Test", "") & ".accde")
                    newAccApp.Quit(Access.AcQuitOption.acQuitSaveNone)
                    newAccApp = Nothing
                End If
            End If
            If chkDeploy.IsChecked Then
                If Not myMSFile.xl Then
                    myAppName = myUtilities.ReplaceEnvironVars(PublishFolder) & "\" & AppName.Text & ".accdb"
                Else
                    myAppName = myUtilities.ReplaceEnvironVars(PublishFolder) & "\" & AppName.Text & ".xlsm"
                End If
                If InStr(txtDeployFolder.Text, ";") Then
                    Dim deployArr() As String = Split(txtDeployFolder.Text, ";")
                    For i = 0 To UBound(deployArr)
                        myUtilities.DeleteFile(myUtilities.ReplaceEnvironVars(Trim(Replace(deployArr(i), vbCrLf, ""))))
                        objFSO.CopyFile(myAppName, myUtilities.ReplaceEnvironVars(Trim(Replace(deployArr(i), vbCrLf, ""))))
                    Next
                Else
                    myUtilities.DeleteFile(myUtilities.ReplaceEnvironVars(txtDeployFolder.Text))
                    objFSO.CopyFile(myAppName, myUtilities.ReplaceEnvironVars(txtDeployFolder.Text))
                End If
            End If
            ExitGracefully(True, "Project Published. Commit and push to GitHub and share the install link with users.")

ErrorHandler:
        Catch e As Exception
            MsgBox("Error " & e.Message & " In " & System.Reflection.MethodBase.GetCurrentMethod.Name & " at line " & e.StackTrace)
        End Try
        myMSFile.myUtilities.SetStatus("Ready")
        Exit Sub

    End Sub

        
        
        'This is a method will build a batch file, using the values saved in the apps config file, that will download and install the 
        'application from github.
   Public Sub CreateSetupFile(MakeUninstaller As Boolean)
        Try
            If Not myUtilities.FolderExists(PublishFolder & "\Setup") Then
                myUtilities.CreateFolder(PublishFolder & "\Setup")
            End If
            myUtilities.CreateNewTextFile(PublishFolder & "\Setup\Setup.bat")
            Dim myText As String
            myText = "@echo off" & vbCrLf _
                    & "Title Download Setup file from github And run it." & vbCrLf & vbCrLf _
                    & "If Not " & Chr(34) & "%minimized%" & Chr(34) & "==" & Chr(34) & Chr(34) & " GoTo :minimized" & vbCrLf _
                    & "Set minimized=True" & vbCrLf _
                    & "start /min cmd /C" & Chr(34) & "%~dpnx0" & Chr(34) & vbCrLf _
                    & "GoTo :EOF" & vbCrLf _
                    & ":minimized" & vbCrLf _
                   & "If Not exist " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & Chr(34) & " mkdir " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & Chr(34) & vbCrLf _
                    & "If Not exist " & Chr(34) & "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\" & If(txtStartMenuFolder.Text = "", AppName.Text, txtStartMenuFolder.Text) & IIf(TestMode, "Test", "") & "\" & Chr(34) & " mkdir " & Chr(34) & "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\" & IIf(txtStartMenuFolder.Text <> vbNullString, txtStartMenuFolder.Text, AppName.Text) & IIf(TestMode, "Test", "") & Chr(34) & vbCrLf & vbCrLf _
                    & "Call :CreateConverter" & vbCrLf _
                    & ":CheckForFile" & vbCrLf _
                    & "IF EXIST " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test", "") & "\converttoAnsiPC.VBS" & Chr(34) & " GOTO FoundIt" & vbCrLf _
                    & "GOTO CheckForFile" & vbCrLf _
                    & ":FoundIt" & vbCrLf _
                    & "If Not exist " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test", "") & "\converttoAnsiPC.VBS" & Chr(34) & " (" & vbCrLf _
                    & "Pause" & vbCrLf _
                    & ")" & vbCrLf _
                    & vbCrLf
            Dim gitFolder As String
            gitFolder = IIf(Right(PublishFolder, Len(PublishFolder) - InStrRev(PublishFolder, "\")) = cmbRepositoryName.Text, "", Right(PublishFolder, Len(PublishFolder) - InStrRev(PublishFolder, "\")) & "/")

            myText += AddFileToInstaller(gitPubURL, "SendEmailToAdmin.vbs", True)

            myText += "wscript " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & "SendEmailToAdmin.vbs" & Chr(34) & " " & Chr(34) & "DoNotReply@uhc.com" & Chr(34) & " " _
                    & Chr(34) & txtAdminEmail.Text & Chr(34) & " " _
                    & Chr(34) & "Installation of: " & AppName.Text & " " _
                    & Replace(Replace(Trim(txtVersion.Text), vbCr, ""), vbLf, "") & Chr(34) & " " _
                    & Chr(34) & "Installed on: %COMPUTERNAME% By: %USERNAME%" & Chr(34) & " " _
                    & Chr(34) & "" & Chr(34) & " " _
                    & Chr(34) & "" & Chr(34) & " " _
                    & Chr(34) & "mailo2.uhc.com" & Chr(34) & " " _
                    & "25" & vbCrLf
            myText += AddFileToInstaller(gitPubURL, AppName.Text & IIf(TestMode, "Test", "") & IIf(myMSFile.xl, ".xlsm", IIf(chkAccde.IsChecked, ".accde", ".accdb")))
            If Not chkAccde.IsChecked And Not myMSFile.xl Then
                myText += AddFileToInstaller(gitPubURL, AppName.Text & IIf(TestMode, "Test", "") & ".accdb")
            End If
            myText += AddFileToInstaller(gitPubURL, "Version.txt")
            myText += AddFileToInstaller(gitPubURL, "Uninstall " & AppName.Text & ".bat", True)
            myText += "set SCRIPT=" & Chr(34) & "%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs" & Chr(34) & vbCrLf
            myText += vbCrLf
            myText += "echo Set oWS = WScript.CreateObject(" & Chr(34) & "WScript.Shell" & Chr(34) & ") >> %SCRIPT%" & vbCrLf
            myText += "echo sLinkFile = " & Chr(34)
            myText += "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\"
            myText += If(txtStartMenuFolder.Text = "", AppName.Text, txtStartMenuFolder.Text)
            myText += IIf(TestMode, "Test", "") & "\" & AppName.Text & IIf(TestMode, "Test", "") & ".lnk"
            myText += Chr(34) & " >> %SCRIPT% " & vbCrLf
            myText += "echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%" & vbCrLf
            myText += "echo oLink.TargetPath = " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & AppName.Text
            myText += IIf(TestMode, "Test", "") & IIf(myMSFile.xl, ".xlsm", IIf(chkAccde.IsChecked, ".accde", ".accdb"))
            myText += Chr(34) & " >> %SCRIPT%" & vbCrLf
            myText += "echo oLink.Save >> %SCRIPT%" & vbCrLf
            myText += vbCrLf
            myText += "wscript /nologo %SCRIPT%" & vbCrLf
            myText += "del %SCRIPT%" & vbCrLf
            myText += vbCrLf
            If MakeUninstaller Then
                myText += "set SCRIPT=" & Chr(34) & "%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs" & Chr(34) & vbCrLf
                myText += vbCrLf
                myText += "echo Set oWS = WScript.CreateObject(" & Chr(34) & "WScript.Shell" & Chr(34) & ") >> %SCRIPT%" & vbCrLf
                myText += "echo sLinkFile = " & Chr(34) & "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\" & If(txtStartMenuFolder.Text = "", AppName.Text, txtStartMenuFolder.Text) & IIf(TestMode, "Test", "") & "\Uninstall " & AppName.Text & ".lnk" & Chr(34) & " >> %SCRIPT%" & vbCrLf
                myText += "echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%" & vbCrLf
                myText += "echo oLink.TargetPath = " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & "Uninstall " & AppName.Text & ".bat" & Chr(34) & " >> %SCRIPT%" & vbCrLf
                myText += "echo oLink.Save >> %SCRIPT%" & vbCrLf
                myText += vbCrLf
                myText += "wscript /nologo %SCRIPT%" & vbCrLf
                myText += "del %SCRIPT%" & vbCrLf
                myText += vbCrLf
            End If
            myText += "Call :OpenSetup" & vbCrLf
            myText += "timeout 2 > nul" & vbCrLf
            myText += "@Echo off" & vbCrLf
            myText += vbCrLf
            myText += "Exit /b" & vbCrLf
            myText += "ECHO" & vbCrLf
            myText += "::*********************************************************************************" & vbCrLf
            myText += ":CreateConverter" & vbCrLf
            myText += "@echo off" & vbCrLf
            myText += "DEL " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & "converttoAnsiPC.VBS" & Chr(34) & vbCrLf
            myText += "echo Do Until WScript.StdIn.AtEndOfStream>> " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & "converttoAnsiPC.VBS" & Chr(34) & vbCrLf
            myText += "echo WScript.StdOut.WriteLine WScript.StdIn.ReadLine>> " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & "converttoAnsiPC.VBS" & Chr(34) & vbCrLf
            myText += "echo Loop>> " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & "converttoAnsiPC.VBS" & Chr(34) & vbCrLf
            myText += "" & vbCrLf
            myText += "Exit /b" & vbCrLf
            myText += ":Download <url> <File>" & vbCrLf
            myText += "@ECHO off" & vbCrLf
            myText += "Powershell.exe -command " & Chr(34) & "(New-Object System.Net.WebClient).DownloadFile('%1','%2')" & Chr(34) & vbCrLf
            myText += "exit /b" & vbCrLf
            myText += ":DeleteSetup" & vbCrLf
            myText += "@ECHO off" & vbCrLf
            myText += "del " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & AppName.Text & IIf(myMSFile.xl, ".xlsm", IIf(chkAccde.IsChecked, ".accde", ".accdb")) & Chr(34) & " /s /f /q" & vbCrLf
            myText += "exit /b" & vbCrLf
            myText += ":OpenSetup" & vbCrLf
            myText += "@ECHO off" & vbCrLf
            myText += "Start " & Chr(34) & Chr(34) & " " & Chr(34) & txtPCInstallFolder.Text & IIf(InStr(txtPCInstallFolder.Text, "\") > 0, "", "\" & cmbRepositoryName.Text) & IIf(TestMode, "Test\", "\") & AppName.Text & IIf(TestMode, "Test", "") & IIf(myMSFile.xl, ".xlsm", IIf(chkAccde.IsChecked, ".accde", ".accdb")) & Chr(34) & vbCrLf
            myText += "exit /b" & vbCrLf

            myUtilities.WriteTextFile(txtPublishFolder.Text & "\Setup\Setup.bat", myText)
            myUtilities.WriteTextFile(txtPublishFolder.Text & "\Setup.bat", myText)
        Catch e As Exception
            MsgBox("Error " & e.Message & " In " & System.Reflection.MethodBase.GetCurrentMethod.Name & " at line " & e.StackTrace)
        End Try
    End Sub  
