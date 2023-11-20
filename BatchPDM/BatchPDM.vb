Imports System.IO
Imports System.Net.WebRequestMethods
Imports System.Windows.Forms
Imports EdmLib
Imports SolidWorks.Interop
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swdocumentmgr

Public Class BatchPDM
    Implements IEdmAddIn5

    Const RESTARTSWCOUNT As Integer = 500
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

            'poCmdMgr.AddCmd(1, "Batch Check Assy Material", EdmMenuFlags.EdmMenu_MustHaveSelection + EdmMenuFlags.EdmMenu_OnlyFolders)
            'poCmdMgr.AddCmd(2, "Batch Remove Assy Material", EdmMenuFlags.EdmMenu_MustHaveSelection + EdmMenuFlags.EdmMenu_OnlyFolders)
            'poCmdMgr.AddCmd(3, "Process IN Freeze Bar", EdmMenuFlags.EdmMenu_ShowInMenuBarAction)
            'poCmdMgr.AddCmd(4, "Check out Files with Property Missing", EdmMenuFlags.EdmMenu_ShowInMenuBarAction)
            'poCmdMgr.AddCmd(5, "Set Part Properties", EdmMenuFlags.EdmMenu_ShowInMenuBarAction)
            'poCmdMgr.AddCmd(6, "Increment Part Revisions", EdmMenuFlags.EdmMenu_MustHaveSelection + EdmMenuFlags.EdmMenu_OnlyFolders)
            poCmdMgr.AddCmd(7, "Find Macro Runner", EdmMenuFlags.EdmMenu_MustHaveSelection + EdmMenuFlags.EdmMenu_OnlyFolders)
            poCmdMgr.AddCmd(8, "Get Latest In Selected Folder(s)", EdmMenuFlags.EdmMenu_MustHaveSelection + EdmMenuFlags.EdmMenu_OnlyFolders)
            poCmdMgr.AddCmd(9, "Find grams", EdmMenuFlags.EdmMenu_MustHaveSelection + EdmMenuFlags.EdmMenu_OnlyFolders)
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_PostAdd)
        Catch
        End Try

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd

        If poCmd.meCmdType = EdmCmdType.EdmCmd_PostAdd Then

            PopulateCode(poCmd, ppoData)

        End If

        Dim eVault As EdmVault5 = poCmd.mpoVault

        Dim eUserMgr As IEdmUserMgr5 = eVault.CreateUtility(EdmUtility.EdmUtil_UserMgr)
        Dim eUser As IEdmUser5 = eUserMgr.GetLoggedInUser()

        If eUser.Name.ToLower() = "admin" Then

            If poCmd.meCmdType = EdmCmdType.EdmCmd_Menu Then

                If poCmd.mlCmdID = 1 Then

                    Dim confirmList As String = ""
                    For Each folderData As EdmCmdData In ppoData
                        confirmList += folderData.mbsStrData1 + vbNewLine
                    Next

                    Dim firstFolderData As EdmCmdData = ppoData(0)
                    Dim folderLetter As String = Strings.Left(firstFolderData.mbsStrData1, 1)

                    If MsgBox(confirmList, MsgBoxStyle.OkCancel, "Check Material") = MsgBoxResult.Ok Then
                        FindFiles(poCmd, ppoData, eVault, folderLetter, True)
                    End If

                ElseIf poCmd.mlCmdID = 2 Then

                    Dim confirmList As String = ""
                    For Each folderData As EdmCmdData In ppoData
                        confirmList += folderData.mbsStrData1 + vbNewLine
                    Next

                    Dim firstFolderData As EdmCmdData = ppoData(0)
                    Dim folderLetter As String = Strings.Left(firstFolderData.mbsStrData1, 1)

                    If MsgBox(confirmList, MsgBoxStyle.OkCancel, "Remove Material") = MsgBoxResult.Ok Then
                        FindFiles(poCmd, ppoData, eVault, folderLetter, False)
                    End If

                ElseIf poCmd.mlCmdID = 3 Then

                    Dim csvPath As String = "C:\Users\administrador\Desktop\Macros\IN_Type.csv"
                    SetFreezeBar(poCmd, ppoData, csvPath)

                ElseIf poCmd.mlCmdID = 4 Then

                    Dim csvPath As String = "C:\Users\administrador\Desktop\Macros\Codigo_Missing.txt"
                    CheckOutFiles(poCmd, ppoData, csvPath)

                ElseIf poCmd.mlCmdID = 5 Then

                    Dim csvPath As String = "C:\Users\administrador\Desktop\Macros\Codigo_Missing.txt"
                    SetPartProperties(poCmd, ppoData, csvPath)

                ElseIf poCmd.mlCmdID = 6 Then

                    Dim confirmList As String = ""
                    For Each folderData As EdmCmdData In ppoData
                        confirmList += folderData.mbsStrData1 + vbNewLine
                    Next

                    If MsgBox(confirmList, MsgBoxStyle.OkCancel, "Set Part Revisions") = MsgBoxResult.Ok Then
                        SetPartRevisions(poCmd, ppoData, eVault)
                    End If

                ElseIf poCmd.mlCmdID = 7 Then

                    FindMacroRunner(poCmd, ppoData, eVault)

                ElseIf poCmd.mlCmdID = 8 Then

                    GetLatestInFolders(poCmd, ppoData, eVault)

                ElseIf poCmd.mlCmdID = 9 Then

                    FindGrams(poCmd, ppoData, eVault)

                End If

            End If
        Else

            MsgBox("Adming login required to run this function", MsgBoxStyle.Exclamation, "Serrat Automation")

        End If

    End Sub

    Private Sub GetLatestInFolders(poCmd As EdmCmd, ByRef ppoData As System.Array, eVault As EdmVault5)

        Dim eFolder As IEdmFolder6 = Nothing

        Try
            Dim processedList As New List(Of String)
            Dim success As Boolean = True

            For Each folderData In ppoData

                If processedList.Contains(folderData.mlObjectID2) = False Then

                    processedList.Add(folderData.mlObjectID2)

                    eFolder = eVault.GetObject(EdmObjectType.EdmObject_Folder, folderData.mlObjectID2)

                    If eFolder IsNot Nothing Then
                        TraverseFolderForAssys_GetLatest(eFolder)
                    Else
                        WriteToLog(True, $"Unable to get folder object with ID: {folderData.mlObjectID2}")
                    End If

                    eFolder = Nothing
                End If
            Next

            MsgBox($"Get latest completed", MsgBoxStyle.Information, "BatchPDM")

        Catch ex As System.Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)

            MsgBox($"The following error occurred:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})", MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        End Try

    End Sub

    Sub TraverseFolderForAssys_GetLatest(eFolder As IEdmFolder5)

        If eFolder IsNot Nothing Then

            Dim splitFolderName() As String = Split(eFolder.Name, "-")

            If splitFolderName(0).Length = 3 Then

                Dim pdmFilePos As IEdmPos5
                pdmFilePos = eFolder.GetFirstFilePosition()

                While pdmFilePos.IsNull = False
                    Try

                        Dim eFile As IEdmFile5
                        eFile = eFolder.GetNextFile(pdmFilePos)

                        If eFile IsNot Nothing Then

                            eFile.GetFileCopy(0, lEdmGetFlags:=EdmGetFlag.EdmGet_RefsOnlyMissing)

                        End If
                    Catch ex As Exception
                        Dim st As New StackTrace(True)
                        st = New StackTrace(ex, True)

                        WriteToLog(True, $"The following error occurred getting latest:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})")

                    End Try

                End While

                Dim pdmSubFolderPos As IEdmPos5
                pdmSubFolderPos = eFolder.GetFirstSubFolderPosition()

                While Not pdmSubFolderPos.IsNull
                    Dim pdmSubFolder As IEdmFolder5
                    pdmSubFolder = eFolder.GetNextSubFolder(pdmSubFolderPos)

                    TraverseFolderForAssys_GetLatest(pdmSubFolder)
                End While

            End If

        End If


    End Sub

    Private Sub FindGrams(poCmd As EdmCmd, ByRef ppoData As System.Array, eVault As EdmVault5)

        Dim eFolder As IEdmFolder6 = Nothing

        Dim swApp As SldWorks = StartSW(True)

        Dim processedList As New List(Of String)
        Dim count As Integer = 0
        Dim success As Boolean = True

        For Each folderData In ppoData

            If processedList.Contains(folderData.mlObjectID2) = False Then

                processedList.Add(folderData.mlObjectID2)

                eFolder = eVault.GetObject(EdmObjectType.EdmObject_Folder, folderData.mlObjectID2)

                If eFolder IsNot Nothing Then
                    TraverseFolderForParts_Grams(swApp, count, eFolder)
                Else
                    WriteToLog(True, $"Unable to get folder object with ID: {folderData.mlObjectID2}")
                End If

                eFolder = Nothing
            End If
        Next

        CloseSW(swApp)

        MsgBox($"Completed mass unit check on {count} files", MsgBoxStyle.Information, "BatchPDM")

        'Try
        'Catch ex As System.Exception
        '    Dim st As New StackTrace(True)
        '    st = New StackTrace(ex, True)

        '    MsgBox($"The following error occurred:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})", MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        'End Try

    End Sub

    Sub TraverseFolderForParts_Grams(ByRef swApp As SldWorks, ByRef count As Integer, eFolder As IEdmFolder5)

        If eFolder IsNot Nothing Then

            Dim splitFolderName() As String = Split(eFolder.Name, "-")

            If splitFolderName(0).Length = 3 Then

                Dim pdmFilePos As IEdmPos5
                pdmFilePos = eFolder.GetFirstFilePosition()

                While pdmFilePos.IsNull = False

                    Try
                        Dim eFile As IEdmFile5
                        eFile = eFolder.GetNextFile(pdmFilePos)

                        If eFile IsNot Nothing Then

                            If Strings.Right(eFile.Name, 6).ToLower() = "sldprt" Then

                                eFile.GetFileCopy(0)

                                Dim filePath As String = eFile.GetLocalPath(eFolder.ID)

                                Dim errors As Integer
                                Dim warnings As Integer

                                If swApp Is Nothing Then
                                    swApp = StartSW(True)
                                    WriteToLog(False, $"Restart solidworks due to crash")
                                End If

                                Dim swFile As ModelDoc2
                                Try
                                    swFile = swApp.OpenDoc6(filePath, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)
                                Catch
                                End Try

                                If errors = 0 And swFile IsNot Nothing Then

                                    Dim massUnits As Integer = swFile.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified)

                                    If massUnits <> 3 Then
                                        WriteToLog(False, $"Units not set to kg: {filePath}")
                                    End If

                                    swApp.QuitDoc(swFile.GetPathName)

                                    count += 1

                                    'If count Mod RESTARTSWCOUNT = 0 Then

                                    '    WriteToLog(False, $"Restart solidworks: {count} files processed")
                                    '    CloseSW(swApp)
                                    '    swApp = StartSW(True)

                                    '    If swApp Is Nothing Then WriteToLog(True, $"Batch did not complete successfully")

                                    'End If
                                Else
                                    WriteToLog(True, $"Open error {errors}: {filePath}")
                                End If

                            End If
                        End If

                    Catch ex As Exception
                        Dim st As New StackTrace(True)
                        st = New StackTrace(ex, True)

                        WriteToLog(True, $"The following error occurred checking mass units:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})")

                    End Try

                End While

                Dim pdmSubFolderPos As IEdmPos5
                pdmSubFolderPos = eFolder.GetFirstSubFolderPosition()

                While Not pdmSubFolderPos.IsNull
                    Dim pdmSubFolder As IEdmFolder5
                    pdmSubFolder = eFolder.GetNextSubFolder(pdmSubFolderPos)

                    TraverseFolderForParts_Grams(swApp, count, pdmSubFolder)
                End While

            End If

        End If


    End Sub

    Private Sub FindMacroRunner(poCmd As EdmCmd, ByRef ppoData As System.Array, eVault As EdmVault5)

        Dim eFolder As IEdmFolder6 = Nothing

        Dim swApp As SldWorks = StartSW(True)

        Dim processedList As New List(Of String)
        Dim count As Integer = 0
        Dim success As Boolean = True

        For Each folderData In ppoData

            If processedList.Contains(folderData.mlObjectID2) = False Then

                processedList.Add(folderData.mlObjectID2)

                eFolder = eVault.GetObject(EdmObjectType.EdmObject_Folder, folderData.mlObjectID2)

                If eFolder IsNot Nothing Then
                    TraverseFolderForAssys_Macro(swApp, count, eFolder)
                Else
                    WriteToLog(True, $"Unable to get folder object with ID: {folderData.mlObjectID2}")
                End If

                eFolder = Nothing
            End If
        Next

        CloseSW(swApp)

        MsgBox($"Completed macro runner check on {count} files", MsgBoxStyle.Information, "BatchPDM")

        'Try
        'Catch ex As System.Exception
        '    Dim st As New StackTrace(True)
        '    st = New StackTrace(ex, True)

        '    MsgBox($"The following error occurred:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})", MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        'End Try

    End Sub

    Sub TraverseFolderForAssys_Macro(ByRef swApp As SldWorks, ByRef count As Integer, eFolder As IEdmFolder5)

        If eFolder IsNot Nothing Then

            Dim splitFolderName() As String = Split(eFolder.Name, "-")

            If splitFolderName(0).Length = 3 Then

                Dim pdmFilePos As IEdmPos5
                pdmFilePos = eFolder.GetFirstFilePosition()

                While pdmFilePos.IsNull = False

                    Try
                        Dim eFile As IEdmFile5
                        eFile = eFolder.GetNextFile(pdmFilePos)

                        If eFile IsNot Nothing Then

                            If Strings.Right(eFile.Name, 6).ToLower() = "sldasm" Then

                                eFile.GetFileCopy(0)

                                Dim filePath As String = eFile.GetLocalPath(eFolder.ID)
                                'WriteToLog(False, $"check file: {filePath}")

                                Dim errors As Integer
                                Dim warnings As Integer

                                Dim swFile As ModelDoc2 = swApp.OpenDoc6(filePath, swDocumentTypes_e.swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent + swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", errors, warnings)

                                If errors = 0 Then

                                    Dim swFeat As Feature
                                    swFeat = swFile.FirstFeature

                                    While Not swFeat Is Nothing
                                        If InStr(1, LCase(swFeat.Name), "macrorunner") > 0 Then
                                            WriteToLog(False, $"Contains macro runner, {filePath}")
                                            Exit While
                                        End If

                                        swFeat = swFeat.GetNextFeature
                                    End While

                                    swApp.QuitDoc(swFile.GetPathName)

                                    count += 1

                                    If count Mod RESTARTSWCOUNT = 0 Then

                                        WriteToLog(False, $"Restart solidworks: {count} files processed")
                                        CloseSW(swApp)
                                        swApp = StartSW(True)
                                        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUserEnableFreezeBar, True)

                                        If swApp Is Nothing Then WriteToLog(True, $"Batch did not complete successfully")

                                    End If
                                Else
                                    WriteToLog(True, $"Open error {errors}: {filePath}")
                                End If

                                'If eFile.CurrentRevision = "" Then
                                '    eFile.IncrementRevision(0, eFolder.ID, "REVISION A")

                                '    count += 1

                                '    WriteToLog(False, $"Set Revision: {eFile.Name}")
                                'Else
                                '    WriteToLog(False, $"Exisisting revision {eFile.Name}: {eFile.CurrentRevision}")
                                'End If

                            End If
                        End If

                    Catch ex As Exception
                        Dim st As New StackTrace(True)
                        st = New StackTrace(ex, True)

                        WriteToLog(True, $"The following error occurred removing macro:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})")

                    End Try

                End While

                Dim pdmSubFolderPos As IEdmPos5
                pdmSubFolderPos = eFolder.GetFirstSubFolderPosition()

                While Not pdmSubFolderPos.IsNull
                    Dim pdmSubFolder As IEdmFolder5
                    pdmSubFolder = eFolder.GetNextSubFolder(pdmSubFolderPos)

                    TraverseFolderForAssys_Macro(swApp, count, pdmSubFolder)
                End While

            End If

        End If


    End Sub

    Private Sub SetPartRevisions(poCmd As EdmCmd, ByRef ppoData As System.Array, eVault As EdmVault5)

        Dim eFolder As IEdmFolder6 = Nothing

        Try
            Dim processedList As New List(Of String)
            Dim count As Integer = 0
            Dim success As Boolean = True

            For Each folderData In ppoData

                If processedList.Contains(folderData.mlObjectID2) = False Then

                    processedList.Add(folderData.mlObjectID2)

                    eFolder = eVault.GetObject(EdmObjectType.EdmObject_Folder, folderData.mlObjectID2)

                    If eFolder IsNot Nothing Then
                        TraverseFolderForParts_Rev(count, eFolder)
                    Else
                        WriteToLog(True, $"Unable to get folder object with ID: {folderData.mlObjectID2}")
                    End If

                    eFolder = Nothing
                End If
            Next

            MsgBox($"Successfully incremented {count} file revisions", MsgBoxStyle.Information, "BatchPDM")

        Catch ex As System.Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)

            MsgBox($"The following error occurred:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})", MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        End Try

    End Sub

    Sub TraverseFolderForParts_Rev(ByRef count As Integer, eFolder As IEdmFolder5)

        If eFolder IsNot Nothing Then

            Dim splitFolderName() As String = Split(eFolder.Name, "-")

            If splitFolderName(0).Length = 3 Then

                Dim pdmFilePos As IEdmPos5
                pdmFilePos = eFolder.GetFirstFilePosition()

                While pdmFilePos.IsNull = False

                    Dim eFile As IEdmFile5
                    eFile = eFolder.GetNextFile(pdmFilePos)

                    If eFile IsNot Nothing Then

                        If Strings.Right(eFile.Name, 6).ToLower() = "sldprt" Then

                            If eFile.CurrentRevision = "" Then
                                eFile.IncrementRevision(0, eFolder.ID, "REVISION A")

                                count += 1

                                WriteToLog(False, $"Set Revision: {eFile.Name}")
                            Else
                                WriteToLog(False, $"Exisisting revision {eFile.Name}: {eFile.CurrentRevision}")
                            End If

                        End If
                    End If

                End While

                Dim pdmSubFolderPos As IEdmPos5
                pdmSubFolderPos = eFolder.GetFirstSubFolderPosition()

                While Not pdmSubFolderPos.IsNull
                    Dim pdmSubFolder As IEdmFolder5
                    pdmSubFolder = eFolder.GetNextSubFolder(pdmSubFolderPos)

                    TraverseFolderForParts_Rev(count, pdmSubFolder)
                End While

            Else
                WriteToLog(False, $"Skipping folder: {eFolder.Name}")
            End If

        End If


    End Sub

    Private Sub SetPartProperties(ByRef poCmd As EdmCmd, ByRef ppoData As Array, csvPath As String)

        Try
            Dim swApp As SldWorks = StartSW()
            swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUserEnableFreezeBar, True)

            If swApp IsNot Nothing Then

                Dim csvList As List(Of String) = ReadCSV(csvPath)

                WriteToLog(False, $"Read CSV file: {csvPath} ({csvList.Count} files)")

                Dim errors As Integer
                Dim warnings As Integer

                Dim count As Integer = 0

                For Each strFileInfo In csvList

                    Dim fileInfo() As String = strFileInfo.Split(",")
                    Dim filePath As String = IO.Path.Combine(fileInfo(6), $"{fileInfo(0)}.sldprt")

                    Dim swFile As ModelDoc2 = swApp.OpenDoc6(filePath, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)

                    If errors = 0 Then

                        SetUnits(swFile)
                        If Strings.Right(filePath, 6) = "sldprt" Then RunSetMaterial(swFile)
                        AddSpecialProperties(swFile)
                        CopyProps(swFile)

                        swFile.Save2(True)

                        swApp.QuitDoc(swFile.GetPathName)

                        WriteToLog(False, $"Set All Properties: {fileInfo(0)}")

                        count += 1

                        If count Mod RESTARTSWCOUNT = 0 Then

                            WriteToLog(False, $"Restart solidworks: {count} files processed")
                            CloseSW(swApp)
                            swApp = StartSW()
                            swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUserEnableFreezeBar, True)

                            If swApp Is Nothing Then WriteToLog(True, $"Batch did not complete successfully")

                        End If
                    Else
                        WriteToLog(True, $"Open error {errors}: {filePath}")
                    End If

                Next

            End If

            CloseSW(swApp)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

        MsgBox("Done Setting Properties", MsgBoxStyle.Information)

    End Sub

    Private Sub SetUnits(swFile As ModelDoc2)

        Dim boolstatus As Boolean

        boolstatus = swFile.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms)
        boolstatus = swFile.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropDecimalPlaces, swUserPreferenceOption_e.swDetailingNoOptionSpecified, 2)

    End Sub

    Private Sub RunSetMaterial(swFile As ModelDoc2)

        Dim swFilePath As String
        swFilePath = swFile.GetPathName

        Dim swFileName As String
        swFileName = Strings.Right(swFilePath, Len(swFilePath) - InStrRev(swFilePath, "\"))

        Dim swFileExt As String
        swFileExt = Strings.Right(swFileName, 6)

        Dim swPropMgr As CustomPropertyManager
        swPropMgr = swFile.Extension.CustomPropertyManager("")

        If LCase(swFileExt) = "sldprt" Then

            Dim swPart As PartDoc
            swPart = swFile

            Dim propLaser As String
            propLaser = swPropMgr.Get("Laser")

            Dim setMaterialSuccess As Boolean

            If propLaser <> "" Then

                If Strings.Left(propLaser, 2) = "AN" Then

                    setMaterialSuccess = SetMaterial(swFile, swPart, "HARDOX400-ANTIDESGASTE")

                ElseIf Strings.Left(propLaser, 2) = "PA" Then

                    setMaterialSuccess = SetMaterial(swFile, swPart, "S700MC-1.8974-PAS700")

                ElseIf Strings.Left(propLaser, 2) = "CR" Then

                    setMaterialSuccess = SetMaterial(swFile, swPart, "13CrMo4-5 1.7335")

                ElseIf Strings.Left(propLaser, 2) = "AL" Then

                    setMaterialSuccess = SetMaterial(swFile, swPart, "ALUMINIO 6061-T6")

                End If

            End If

            If setMaterialSuccess = False Then

                Dim materialListPath As String = "C:\Users\administrador\Desktop\Macros\MaterialsList.txt"

                Dim materialListReader As New StreamReader(materialListPath)

                Do While Not materialListReader.EndOfStream

                    Dim materialLineText As String
                    materialLineText = materialListReader.ReadLine

                    Dim materialSplit() As String

                    materialSplit = Split(materialLineText, ",")

                    If LCase(Strings.Left(swFileName, 3)) = LCase(materialSplit(0)) Then

                        SetMaterial(swFile, swPart, materialSplit(1))

                    End If
                Loop

                materialListReader.Close()

            End If

        End If
    End Sub

    Private Sub AddSpecialProperties(swFile As ModelDoc2)

        Dim swFilePath As String
        swFilePath = swFile.GetPathName

        Dim swFileName As String
        swFileName = Strings.Right(swFilePath, Len(swFilePath) - InStrRev(swFilePath, "\"))

        Dim swPropMgr As CustomPropertyManager
        swPropMgr = swFile.Extension.CustomPropertyManager("")

        swPropMgr.Add3("Material", swCustomInfoType_e.swCustomInfoText, """" & "SW-Material@" & swFileName & """", 1)
        swPropMgr.Add3("Weight", swCustomInfoType_e.swCustomInfoText, """" & "SW-Masa@" & swFileName & """", 1)
        swPropMgr.Add3("Código", swCustomInfoType_e.swCustomInfoText, "$PRP:" & """" & "SW-Nombre del archivo(File Name)" & """", 1)

        Try
            swPropMgr.Delete("Description")
        Catch
        End Try

    End Sub

    Private Sub CopyProps(swFile As ModelDoc2)

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
    End Sub


    Private Sub SetFreezeBar(ByRef poCmd As EdmCmd, ByRef ppoData As Array, csvPath As String)

        Dim swApp As SldWorks = StartSW()
        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUserEnableFreezeBar, True)

        If swApp IsNot Nothing Then

            Dim csvList As List(Of String) = ReadCSV(csvPath)

            WriteToLog(False, $"Read CSV file: {csvPath} ({csvList.Count} files)")

            Dim errors As Integer
            Dim warnings As Integer

            Dim count As Integer = 0

            For Each strFileInfo In csvList

                Dim fileInfo() As String = strFileInfo.Split(",")
                Dim filePath As String = IO.Path.Combine(fileInfo(6), $"{fileInfo(0)}.sldprt")

                Dim swFile As ModelDoc2 = swApp.OpenDoc6(filePath, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)

                If errors = 0 Then

                    Dim featMgr As FeatureManager = swFile.FeatureManager

                    featMgr.EditFreeze2(swconst.swMoveFreezeBarTo_e.swMoveFreezeBarToEnd, "", True, True)

                    swFile.Save2(True)

                    swApp.QuitDoc(swFile.GetPathName)

                    WriteToLog(False, $"Freeze Bar Updated: {fileInfo(0)}")

                    count += 1

                    If count Mod RESTARTSWCOUNT = 0 Then

                        WriteToLog(False, $"Restart solidworks: {count} files processed")
                        CloseSW(swApp)
                        swApp = StartSW()
                        swApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUserEnableFreezeBar, True)

                        If swApp Is Nothing Then WriteToLog(True, $"Batch did not complete successfully")

                    End If
                Else
                    WriteToLog(True, $"Open error {errors}: {filePath}")
                End If

            Next

        End If

        CloseSW(swApp)

        MsgBox("Done Setting Freeze Bar", MsgBoxStyle.Information)

    End Sub

    Private Sub CheckOutFiles(ByRef poCmd As EdmCmd, ByRef ppoData As Array, csvPath As String)


        Dim csvList As List(Of String) = ReadCSV(csvPath)

        Dim eVault As EdmVault5 = poCmd.mpoVault

        For Each strFileInfo In csvList

            Dim fileInfo() As String = strFileInfo.Split(",")
            Dim filePath As String = IO.Path.Combine(fileInfo(6), $"{fileInfo(0)}.sldprt")

            Try

                Dim eFile As IEdmFile5 = eVault.GetFileFromPath(filePath)
                eFile.GetFileCopy(0)

                Dim eFolder As IEdmFolder5 = eVault.GetFolderFromPath(fileInfo(6))
                eFile.LockFile(eFolder.ID, 0)

            Catch ex As Exception
                WriteToLog(True, $"Checkout error {ex.Message}: {filePath}")
            End Try

        Next
    End Sub

    Private Function ReadCSV(csvPath As String) As List(Of String)

        Dim streamReader As New StreamReader(csvPath)
        Dim csvList As New List(Of String)

        Do While Not streamReader.EndOfStream
            csvList.Add(streamReader.ReadLine)
        Loop

        streamReader.Close()

        Return csvList

    End Function

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

    Private Sub FindFiles(poCmd As EdmCmd, ByRef ppoData As System.Array, eVault As EdmVault5, folderLetter As String, findReadOnly As Boolean)

        Dim eFolder As IEdmFolder6 = Nothing

        Try
            Dim processedList As New List(Of String)
            Dim count As Integer = 0
            Dim success As Boolean = True
            Dim swApp As SldWorks = StartSW(folderLetter:=count)

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
                            If findReadOnly = False Then
                                TraverseFolderForAssemblies(swDmApp, count, eFolder, folderLetter)
                            Else
                                TraverseFolderForAssembliesReadOnly(swDmApp, count, eFolder)
                            End If
                        Else
                            WriteToLog(True, $"Unable to get folder object with ID: {folderData.mlObjectID2}")
                        End If

                        eFolder = Nothing
                    End If
                Next

                CloseSW(swApp)

                MsgBox($"Successfully processed {count} files", MsgBoxStyle.Information, "BatchPDM")
                'WriteToLog(False, $"Job complete - Successfully processed {count} files", folderLetter)

            End If

        Catch ex As System.Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)

            MsgBox($"The following error occurred:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})", MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        End Try

    End Sub

    Sub TraverseFolderForAssembliesReadOnly(ByRef swDmApp As SwDMApplication4, ByRef count As Integer, eFolder As IEdmFolder5)

        If eFolder IsNot Nothing Then

            Dim pdmFilePos As IEdmPos5
            pdmFilePos = eFolder.GetFirstFilePosition()

            While pdmFilePos.IsNull = False
                Dim eFile As IEdmFile5
                eFile = eFolder.GetNextFile(pdmFilePos)

                If eFile IsNot Nothing Then

                    If Strings.Right(eFile.Name, 6).ToLower() = "sldasm" Then

                        Dim result As SwDmDocumentOpenError
                        Dim swDoc As SwDMDocument10 = swDmApp.GetDocument(eFile.GetLocalPath(eFolder.ID), SwDmDocumentType.swDmDocumentAssembly, True, result)

                        If result <> SwDmDocumentOpenError.swDmDocumentOpenErrorNone Then
                            WriteToLog(True, $"Error opening file: {result.ToString} ({eFile.Name})")
                            'Else
                            '    WriteToLog(False, $"Success setting filename property: {eFile.Name}")
                        End If

                        Dim propExists As Boolean = False

                        Dim configNames As Object = swDoc.ConfigurationManager.GetConfigurationNames

                        For Each configName In configNames
                            Dim swConfig As SwDMConfiguration10 = swDoc.ConfigurationManager.GetConfigurationByName(configName)

                            Dim propNames As Object = swConfig.GetCustomPropertyNames

                            For Each propName In propNames

                                If propName = "Material" Then propExists = True

                            Next

                        Next

                        swDoc.CloseDoc()

                        If propExists = True Then WriteToLog(False, $"Material property exists: {eFile.Name}")

                        count += 1

                    End If
                End If

            End While

            Dim pdmSubFolderPos As IEdmPos5
            pdmSubFolderPos = eFolder.GetFirstSubFolderPosition()

            While Not pdmSubFolderPos.IsNull
                Dim pdmSubFolder As IEdmFolder5
                pdmSubFolder = eFolder.GetNextSubFolder(pdmSubFolderPos)

                TraverseFolderForAssembliesReadOnly(swDmApp, count, pdmSubFolder)
            End While

        End If

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
                                'SetFilenameProperty(eFile, folderLetter)

                                'Dim splitName() As String = eFile.Name.ToString.Split(New String() {" "}, StringSplitOptions.None)

                                'If splitName.GetUpperBound(0) > 0 Then

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
                                        If propName = "Material" Then
                                            swConfig.DeleteCustomProperty("Material")
                                        End If

                                        'If propExists = True Then
                                        '    swConfig.SetCustomProperty("Código", splitName(0))
                                        'Else
                                        '    swConfig.AddCustomProperty("Código", SwDmCustomInfoType.swDmCustomInfoText, splitName(0))
                                        'End If

                                    Next

                                Next

                                swDoc.Save()
                                swDoc.CloseDoc()

                                'End If

                                count += 1

                                'If count Mod RESTARTSWCOUNT = 0 Then

                                '    WriteToLog(False, $"Restart solidworks: {count} files processed", folderLetter)
                                '    CloseSW(swApp)
                                '    swApp = StartSW(folderLetter:=count)

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
                                        'SetMaterial(swApp, eFile.LockPath)

                                        count += 1

                                        If count Mod RESTARTSWCOUNT = 0 Then

                                            WriteToLog(False, $"Restart solidworks: {count} files processed", folderLetter)
                                            CloseSW(swApp)
                                            swApp = StartSW(folderLetter:=count)

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

    Function SetMaterial(swFile As ModelDoc2, swPart As PartDoc, materialName As String)

        Dim vMat As Object
        vMat = swFile.Extension.GetMaterialPropertyValues(swInConfigurationOpts_e.swThisConfiguration, "")
        swPart.SetMaterialPropertyName2("", MATERIALDBPATH, materialName)
        swFile.Extension.RemoveMaterialProperty(swInConfigurationOpts_e.swThisConfiguration, "")
        swFile.Extension.SetMaterialPropertyValues(vMat, swInConfigurationOpts_e.swThisConfiguration, "")
        swFile.EditRebuild3()

        SetMaterial = True

    End Function

    Private Function StartSW(Optional background As Boolean = False, Optional folderLetter As String = "") As SldWorks

        Dim swApp As SldWorks = Nothing
        Try
            swApp = CreateObject("SldWorks.Application")
            If background = False Then swApp.Visible = True
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

    Private Sub WriteToLog(logError As Boolean, message As String, Optional folderLetter As String = "")

        Dim messageType As String = " [INFO]"
        If logError = True Then
            messageType = "[ERROR]"

            Dim streamWriter_Error As New StreamWriter($"{LOGPATH}_ERROR_{Strings.Format(DateTime.Now, "yyMMdd")}.txt", True)
            streamWriter_Error.WriteLine($"{messageType} {Strings.Format(DateTime.Now, "hhmmss")}: {message}")
            streamWriter_Error.Close()
        Else

            Dim messageLogPath As String = $"{LOGPATH}{Strings.Format(DateTime.Now, "yyMMdd")}"

            If folderLetter <> "" Then messageLogPath += $"_{folderLetter}"
            messageLogPath += ".txt"

            Dim streamWriter As New StreamWriter(messageLogPath, True)
            streamWriter.WriteLine($"{messageType} {Strings.Format(DateTime.Now, "hhmmss")}: {message}")
            streamWriter.Close()

        End If

    End Sub

End Class
