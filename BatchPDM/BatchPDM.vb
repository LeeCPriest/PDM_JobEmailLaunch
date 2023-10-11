Imports System.IO
Imports EdmLib
Imports SolidWorks.Interop
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Public Class BatchPDM
    Implements IEdmAddIn5

    Const RESTARTSWCOUNT As Integer = 250
    Const LOGPATH As String = "C:\Temp\BatchPDM_Log_"
    Public Const MATERIALDBPATH As String = "C:\PDM\00_Admin\Materiales\MATERIALES_PDM.sldmat"

    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo

        Try
            poInfo.mbsAddInName = "BatchPDM"
            poInfo.mbsCompany = "Written by CAD Innovations www.cadinnovations.ca"

            'Specify information to display in the add-in's Properties dialog box
            poInfo.mbsDescription = "Batch processing of PDM files"
            poInfo.mlAddInVersion = 1.0
            poInfo.mlRequiredVersionMajor = 8
            poInfo.mlRequiredVersionMinor = 0

            poCmdMgr.AddCmd(1, "Batch Set Laser Material", EdmMenuFlags.EdmMenu_MustHaveSelection + EdmMenuFlags.EdmMenu_OnlyFolders)
        Catch
        End Try

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd

        If poCmd.meCmdType = EdmCmdType.EdmCmd_Menu Then

            If poCmd.mlCmdID = 1 Then

                Dim eVault As EdmVault5 = poCmd.mpoVault

                Dim eUser As IEdmUser5 = eVault.CreateUtility(EdmUtility.EdmUtil_UserMgr)

                If eUser.Name.ToLower() = "admin" Then

                    Dim confirmList As String = ""
                    For Each folderData As EdmCmdData In ppoData
                        confirmList += folderData.mbsStrData1 + vbNewLine
                    Next

                    If MsgBox(confirmList, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        FindFiles(poCmd, ppoData, eVault)
                    End If
                Else

                    MsgBox("Adming login required to run this function", MsgBoxStyle.Exclamation, "BatchPDM")

                End If

            End If

        End If

    End Sub

    Private Sub FindFiles(poCmd As EdmCmd, ByRef ppoData As System.Array, eVault As EdmVault5)

        Dim eFolder As IEdmFolder6 = Nothing

        Try
            Dim processedList As New List(Of String)

            Dim swApp As SldWorks = StartSW()

            Dim count As Integer = 0
            Dim success As Boolean = True

            If swApp IsNot Nothing Then

                For Each folderData In ppoData

                    If processedList.Contains(folderData.mlObjectID2) = False Then

                        WriteToLog(False, $"Folder ID: {folderData.mlObjectID2}")

                        processedList.Add(folderData.mlObjectID2)

                        eFolder = eVault.GetObject(EdmObjectType.EdmObject_Folder, folderData.mlObjectID2)

                        If eFolder IsNot Nothing Then
                            TraverseFolder(swApp, count, eFolder)
                        Else
                            WriteToLog(True, $"Unable to get folder object with ID: {folderData.mlObjectID2}")
                        End If

                        eFolder = Nothing
                    End If
                Next

                CloseSW(swApp)


                If success = True Then
                    MsgBox($"Successfully processed {count} files")

                    WriteToLog(False, $"Job complete - Successfully processed {count} files")
                End If

            End If

        Catch ex As System.Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)

            MsgBox($"The following error occurred:{vbNewLine}{vbNewLine}{ex.Message} (Line: {st.GetFrame(0).GetFileLineNumber()})", MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        End Try

    End Sub

    Sub TraverseFolder(ByRef swApp As SldWorks, ByRef count As Integer, eFolder As IEdmFolder5)

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
                            WriteToLog(True, $"Unable to read 'Laser' property: {eFile.Name}")
                        End If

                        If variableValue IsNot Nothing Then
                            If variableValue.ToString().ToLower() = "x" Then
                                If eFile.IsLocked = True Then
                                    If eFile.LockedByUser.Name.ToLower() = "admin" Then
                                        SetMaterial(swApp, eFile.LockPath)

                                        count += 1

                                        If count Mod RESTARTSWCOUNT = 0 Then

                                            WriteToLog(False, $"Restart solidworks: {count} files processed")
                                            CloseSW(swApp)
                                            swApp = StartSW()

                                            If swApp Is Nothing Then WriteToLog(True, $"Batch did not complete successfully")

                                        End If

                                        WriteToLog(False, $"Success setting material: {eFile.Name}")
                                    Else
                                        WriteToLog(True, $"Not checked out to admin: {eFile.Name}")
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

                TraverseFolder(swApp, count, pdmSubFolder)
            End While

        End If

    End Sub

    Sub SetMaterial(swApp As SldWorks, filePath As String)

        'Dim swApp As SldWorks = CreateObject("SldWorks.Application")
        swApp.OpenDoc(filePath, swconst.swDocumentTypes_e.swDocPART)

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

    Private Function StartSW() As SldWorks

        Dim swApp As SldWorks = Nothing
        Try
            swApp = CreateObject("SldWorks.Application")
        Catch ex As Exception
            WriteToLog(True, ex.Message)
        End Try

        If swApp Is Nothing Then WriteToLog(True, "swApp object is null")

        Return swApp

    End Function

    Private Sub CloseSW(swApp As SldWorks)
        Dim swProcess() As Process = Process.GetProcessesByName("SLDWORKS")

        swApp.ExitApp()
        swProcess(0).Kill()
        swProcess(0).WaitForExit()
    End Sub

    Private Sub WriteToLog(logError As Boolean, message As String)

        Dim messageType As String = " [INFO]"
        If logError = True Then messageType = "[ERROR]"

        Dim streamWriter As New StreamWriter($"{LOGPATH}{Strings.Format(DateTime.Now, "yyMMdd")}.txt", True)
        streamWriter.WriteLine($"{messageType} {Strings.Format(DateTime.Now, "hhmmss")}: {message}")
        streamWriter.Close()

    End Sub

End Class
