Imports System.IO
Imports System.Net.WebRequestMethods
Imports EdmLib
Imports SolidWorks.Interop
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swdocumentmgr

Public Class BatchPDM
    Implements IEdmAddIn5

    Const RESTARTSWCOUNT As Integer = 3
    Const LOGPATH As String = "C:\Temp\BatchPDM_Log_"
    Public Const MATERIALDBPATH As String = "C:\PDM\00_Admin\Materiales\MATERIALES_PDM.sldmat"

    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo

        Try
            poInfo.mbsAddInName = "SerratAutomation"
            poInfo.mbsCompany = "Written by Lee Priest www.cadinnovations.ca"

            'Specify information to display in the add-in's Properties dialog box
            poInfo.mbsDescription = "Custom PDM functionality"
            poInfo.mlAddInVersion = 1.0
            poInfo.mlRequiredVersionMajor = 8
            poInfo.mlRequiredVersionMinor = 0

            'poCmdMgr.AddCmd(1, "Batch Set Assy Filename", EdmMenuFlags.EdmMenu_MustHaveSelection + EdmMenuFlags.EdmMenu_OnlyFolders)
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_PostAdd)
        Catch
        End Try

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd

        If poCmd.meCmdType = EdmCmdType.EdmCmd_PostAdd Then

            PopulateCode(poCmd, ppoData)

        End If

        'If poCmd.meCmdType = EdmCmdType.EdmCmd_Menu Then

        '    If poCmd.mlCmdID = 1 Then

        '        Dim eVault As EdmVault5 = poCmd.mpoVault

        '        Dim eUserMgr As IEdmUserMgr5 = eVault.CreateUtility(EdmUtility.EdmUtil_UserMgr)
        '        Dim eUser As IEdmUser5 = eUserMgr.GetLoggedInUser()

        '        If eUser.Name.ToLower() = "admin" Then

        '            Dim confirmList As String = ""
        '            For Each folderData As EdmCmdData In ppoData
        '                confirmList += folderData.mbsStrData1 + vbNewLine
        '            Next

        '            Dim firstFolderData As EdmCmdData = ppoData(0)
        '            Dim folderLetter As String = Strings.Left(firstFolderData.mbsStrData1, 1)

        '            If MsgBox(confirmList, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
        '                FindFiles(poCmd, ppoData, eVault, folderLetter)
        '            End If
        '        Else

        '            MsgBox("Adming login required to run this function", MsgBoxStyle.Exclamation, "BatchPDM")

        '        End If

        '    End If

        'End If

    End Sub

    Private Sub PopulateCode(ByRef poCmd As EdmCmd, ByRef ppoData As Array)

        Dim eVault As EdmVault5 = poCmd.mpoVault

        For Each fileData As EdmCmdData In ppoData

            If Strings.Right(fileData.mbsStrData1, 6).ToLower() = "sldasm" Then

                Dim eFile As IEdmFile5 = eVault.GetFileFromPath(fileData.mbsStrData1)
                Dim eFileConfigs As EdmStrLst5 = eFile.GetConfigurations()

                Dim splitName() As String = eFile.Name.ToString.Split(New String() {" "}, StringSplitOptions.None)

                Dim eFileCard As IEdmEnumeratorVariable5 = eFile.GetEnumeratorVariable()
                eFileCard.SetVar("Codigo", "@", splitName(0))

                Dim pos As IEdmPos5
                pos = eFileConfigs.GetHeadPosition

                While Not pos.IsNull
                    eFileCard.SetVar("Codigo", eFileConfigs.GetNext(pos), splitName(0))
                End While

                eFileCard.Flush()

            End If

        Next

    End Sub

    Private Sub FindFiles(poCmd As EdmCmd, ByRef ppoData As System.Array, eVault As EdmVault5, folderLetter As String)

        Dim eFolder As IEdmFolder6 = Nothing

        Try
            Dim processedList As New List(Of String)
            Dim count As Integer = 0
            Dim success As Boolean = True
            Dim swApp As SldWorks = StartSW(count)

            Dim docMgrKey As String = "CONSTRUCCIONESMECANICASALCAYSL:swdocmgr_general-11785-02051-00064-17409-08723-34307-00007-06120-12153-28675-47147-36320-07780-58580-20483-13007-16485-58752-40693-63371-17264-24369-15628-19769-18769-03413-09485-14653-19733-05429-01293-09529-01293-01357-03377-25861-12621-14337-27236-56922-59590-25690-25696-1026"
            Dim classFactory As SwDMClassFactory = TryCast(Activator.CreateInstance(Type.GetTypeFromProgID("SwDocumentMgr.SwDMClassFactory")), SwDMClassFactory)
            Dim swDmApp As SwDMApplication4 = classFactory.GetApplication(docMgrKey)

            'If swApp IsNot Nothing Then
            If swDmApp IsNot Nothing Then

                For Each folderData In ppoData

                    If processedList.Contains(folderData.mlObjectID2) = False Then

                        processedList.Add(folderData.mlObjectID2)

                        eFolder = eVault.GetObject(EdmObjectType.EdmObject_Folder, folderData.mlObjectID2)

                        If eFolder IsNot Nothing Then
                            'TraverseFolderForParts(swApp, count, eFolder)
                            TraverseFolderForAssemblies(swDmApp, count, eFolder, folderLetter)
                        Else
                            WriteToLog(True, $"Unable to get folder object with ID: {folderData.mlObjectID2}", folderLetter)
                        End If

                        eFolder = Nothing
                    End If
                Next

                CloseSW(swApp)

                MsgBox($"Successfully processed {count} files", MsgBoxStyle.Information, "BatchPDM")
                WriteToLog(False, $"Job complete - Successfully processed {count} files", folderLetter)

            End If

        Catch ex As System.Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)

            MsgBox($"The following error occurred:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})", MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        End Try

    End Sub

    Sub TraverseFolderForAssemblies(ByRef swDmApp As SwDMApplication4, ByRef count As Integer, eFolder As IEdmFolder5, folderLetter As String)

        If eFolder IsNot Nothing Then

            Dim pdmFilePos As IEdmPos5
            pdmFilePos = eFolder.GetFirstFilePosition()

            While pdmFilePos.IsNull = False
                Dim eFile As IEdmFile5
                eFile = eFolder.GetNextFile(pdmFilePos)

                If eFile IsNot Nothing Then

                    If Strings.Right(eFile.Name, 6).ToLower() = "sldasm" Then

                        If eFile.IsLocked = True Then
                            If eFile.LockedByUser.Name.ToLower() = "admin" Then

                                'SetAssyProps(swApp, eFile.LockPath, folderLetter)
                                SetFilenameProperty(eFile, folderLetter)

                                Dim splitName() As String = eFile.Name.ToString.Split(New String() {" "}, StringSplitOptions.None)

                                If splitName.GetUpperBound(0) > 0 Then

                                    Dim result As SwDmDocumentOpenError
                                    Dim swDoc As SwDMDocument10 = swDmApp.GetDocument(eFile.LockPath, SwDmDocumentType.swDmDocumentAssembly, False, result)

                                    If result <> SwDmDocumentOpenError.swDmDocumentOpenErrorNone Then
                                        WriteToLog(True, $"Error opening file: {result.ToString} ({eFile.Name})", folderLetter)
                                    Else
                                        WriteToLog(False, $"Success setting filename property: {eFile.Name}", folderLetter)
                                    End If

                                    Dim configNames As Object = swDoc.ConfigurationManager.GetConfigurationNames

                                    For Each configName In configNames
                                        Dim swConfig As SwDMConfiguration10 = swDoc.ConfigurationManager.GetConfigurationByName(configName)

                                        Dim propNames As Object = swConfig.GetCustomPropertyNames

                                        For Each propName In propNames
                                            Dim propExists As Boolean = False
                                            If propName = "Código" Then propExists = True

                                            If propExists = True Then
                                                swConfig.SetCustomProperty("Código", splitName(0))
                                            Else
                                                swConfig.AddCustomProperty("Código", SwDmCustomInfoType.swDmCustomInfoText, splitName(0))
                                            End If

                                        Next

                                    Next

                                    swDoc.Save()
                                    swDoc.CloseDoc()

                                End If

                                count += 1

                                'If count Mod RESTARTSWCOUNT = 0 Then

                                '    WriteToLog(False, $"Restart solidworks: {count} files processed", folderLetter)
                                '    CloseSW(swApp)
                                '    swApp = StartSW(count)

                                '    If swApp Is Nothing Then WriteToLog(True, $"Batch did not complete successfully", folderLetter)

                                'End If

                            Else
                                WriteToLog(True, $"Not checked out to admin: {eFile.Name}", folderLetter)
                            End If
                        End If
                    End If
                End If

            End While

            Dim pdmSubFolderPos As IEdmPos5
            pdmSubFolderPos = eFolder.GetFirstSubFolderPosition()

            While Not pdmSubFolderPos.IsNull
                Dim pdmSubFolder As IEdmFolder5
                pdmSubFolder = eFolder.GetNextSubFolder(pdmSubFolderPos)

                TraverseFolderForAssemblies(swDmApp, count, pdmSubFolder, folderLetter)
            End While

        End If

    End Sub

    Sub SetFilenameProperty(eFile As IEdmFile5, folderLetter As String)

        Dim eFileCard As IEdmEnumeratorVariable5 = eFile.GetEnumeratorVariable()
        eFileCard.SetVar("Filename", "@", eFile.Name)
        eFileCard.Flush()

    End Sub

    Sub SetAssyProps(swApp As SldWorks, filePath As String, folderLetter As String)

        'swApp.OpenDoc(filePath, swDocumentTypes_e.swDocASSEMBLY)
        Dim errors As Integer
        Dim warnings As Integer
        'swApp.OpenDocSilent(filePath, swDocumentTypes_e.swDocASSEMBLY, errors)
        swApp.OpenDoc6(filePath, swDocumentTypes_e.swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight + swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)
        If errors <> 0 Then
            WriteToLog(True, $"Open error {errors}: {filePath}", folderLetter)
        End If

        Dim swFile As ModelDoc2
        swFile = swApp.ActiveDoc

        Dim boolstatus As Boolean

        boolstatus = swFile.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms)


        Dim swFilePath As String
        swFilePath = swFile.GetPathName

        Dim swFileName As String
        swFileName = Strings.Right(swFilePath, Len(swFilePath) - InStrRev(swFilePath, "\"))

        Dim swFileExt As String
        swFileExt = Strings.Right(swFileName, 6)

        Dim swPropMgr As CustomPropertyManager
        swPropMgr = swFile.Extension.CustomPropertyManager("")


        swPropMgr.Add3("Material", swCustomInfoType_e.swCustomInfoText, """" & "SW-Material@" & swFileName & """", 1)
        swPropMgr.Add3("Weight", swCustomInfoType_e.swCustomInfoText, """" & "SW-Masa@" & swFileName & """", 1)
        swPropMgr.Add3("Código", swCustomInfoType_e.swCustomInfoText, "$PRP:" & """" & "SW-Nombre del archivo(File Name)" & """", 1)


        Dim swProp As CustomPropertyManager
        swProp = swFile.Extension.CustomPropertyManager("")

        Dim propNames As Object
        propNames = swProp.GetNames

        Dim swConfigs As Object
        swConfigs = swFile.GetConfigurationNames

        For Each swConfig In swConfigs

            Dim swPropConfig As CustomPropertyManager
            swPropConfig = swFile.Extension.CustomPropertyManager(swConfig)

            On Error Resume Next

            For i = 0 To UBound(propNames)
                Dim propName As String
                propName = propNames(i)

                Dim propVal As String
                Dim propValRes As String
                swProp.Get6(propName, False, propVal, propValRes, True, False)

                swPropConfig.Add3(propName, swCustomInfoType_e.swCustomInfoText, propVal, 1)
            Next

        Next

        swFile.ViewZoomtofit()

        swFile.Save2(True)

        swApp.QuitDoc(swFile.GetPathName)

    End Sub

    Sub TraverseFolderForParts(ByRef swApp As SldWorks, ByRef count As Integer, eFolder As IEdmFolder5, folderLetter As String)

        If eFolder IsNot Nothing Then

            Dim pdmFilePos As IEdmPos5
            pdmFilePos = eFolder.GetFirstFilePosition()

            While pdmFilePos.IsNull = False
                Dim eFile As IEdmFile5
                eFile = eFolder.GetNextFile(pdmFilePos)

                If eFile IsNot Nothing Then

                    If Strings.Right(eFile.Name, 6).ToLower() = "sldprt" Then

                        Dim eFileCard As IEdmEnumeratorVariable8 = eFile.GetEnumeratorVariable()
                        Dim variableValue As Object = Nothing
                        If eFileCard IsNot Nothing Then
                            eFileCard.GetVar("Laser", "@", variableValue)
                        Else
                            WriteToLog(True, $"Unable to read 'Laser' property: {eFile.Name}", folderLetter)
                        End If

                        If variableValue IsNot Nothing Then
                            If variableValue.ToString().ToLower() = "x" Then
                                If eFile.IsLocked = True Then
                                    If eFile.LockedByUser.Name.ToLower() = "admin" Then
                                        SetMaterial(swApp, eFile.LockPath)

                                        count += 1

                                        If count Mod RESTARTSWCOUNT = 0 Then

                                            WriteToLog(False, $"Restart solidworks: {count} files processed", folderLetter)
                                            CloseSW(swApp)
                                            swApp = StartSW(count)

                                            If swApp Is Nothing Then WriteToLog(True, $"Batch did not complete successfully", folderLetter)

                                        End If

                                        WriteToLog(False, $"Success setting material: {eFile.Name}", folderLetter)
                                    Else
                                        WriteToLog(True, $"Not checked out to admin: {eFile.Name}", folderLetter)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            End While

            Dim pdmSubFolderPos As IEdmPos5
            pdmSubFolderPos = eFolder.GetFirstSubFolderPosition()

            While Not pdmSubFolderPos.IsNull
                Dim pdmSubFolder As IEdmFolder5
                pdmSubFolder = eFolder.GetNextSubFolder(pdmSubFolderPos)

                TraverseFolderForParts(swApp, count, pdmSubFolder, folderLetter)
            End While

        End If

    End Sub

    Sub SetMaterial(swApp As SldWorks, filePath As String)

        'Dim swApp As SldWorks = CreateObject("SldWorks.Application")
        swApp.OpenDoc(filePath, swDocumentTypes_e.swDocPART)

        Dim swFile As ModelDoc2
        swFile = swApp.ActiveDoc

        Dim swPart As PartDoc
        swPart = swFile

        Dim vMat As Object
        vMat = swFile.Extension.GetMaterialPropertyValues(swInConfigurationOpts_e.swAllConfiguration, "")
        swPart.SetMaterialPropertyName2("", MATERIALDBPATH, "S235JR-1.0037-ST37")
        swFile.Extension.RemoveMaterialProperty(swInConfigurationOpts_e.swAllConfiguration, "")
        swFile.Extension.SetMaterialPropertyValues(vMat, swInConfigurationOpts_e.swAllConfiguration, "")
        swFile.EditRebuild3()

        swFile.Save()

        swApp.QuitDoc(swFile.GetPathName)

    End Sub

    Private Function StartSW(folderLetter As String) As SldWorks

        Dim swApp As SldWorks = Nothing
        Try
            swApp = CreateObject("SldWorks.Application")
        Catch ex As Exception
            WriteToLog(True, ex.Message, folderLetter)
        End Try

        If swApp Is Nothing Then WriteToLog(True, "swApp object is null", folderLetter)

        Return swApp

    End Function

    Private Sub CloseSW(swApp As SldWorks)
        Dim swProcess() As Process = Process.GetProcessesByName("SLDWORKS")

        swApp.ExitApp()
        swProcess(0).Kill()
        swProcess(0).WaitForExit()
    End Sub

    Private Sub WriteToLog(logError As Boolean, message As String, folderLetter As String)

        Dim messageType As String = " [INFO]"
        If logError = True Then
            messageType = "[ERROR]"

            Dim streamWriter_Error As New StreamWriter($"{LOGPATH}_ERROR_{Strings.Format(DateTime.Now, "yyMMdd")}.txt", True)
            streamWriter_Error.WriteLine($"{messageType} {Strings.Format(DateTime.Now, "hhmmss")}: {message}")
            streamWriter_Error.Close()

        End If

        Dim streamWriter As New StreamWriter($"{LOGPATH}{Strings.Format(DateTime.Now, "yyMMdd")}_{folderLetter}.txt", True)
        streamWriter.WriteLine($"{messageType} {Strings.Format(DateTime.Now, "hhmmss")}: {message}")
        streamWriter.Close()

    End Sub

End Class
