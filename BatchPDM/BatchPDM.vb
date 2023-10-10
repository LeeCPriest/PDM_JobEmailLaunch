Imports System.IO
Imports EdmLib
Imports SolidWorks.Interop
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Public Class BatchPDM
    Implements IEdmAddIn5

    Const LOGPATH As String = "C:\Temp\BatchPDM_Log.txt"
    Const LOGPATH_ERROR As String = "C:\Temp\BatchPDM_ERRORS_Log.txt"
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

            'Notify the add-in after a file state change event occurs
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_Menu)
        Catch
        End Try

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd

        If poCmd.meCmdType = EdmCmdType.EdmCmd_Menu Then 'run the following function when a PostState event occurs (i.e. after the transition has occurred)

            Dim confirmList As String = ""
            For Each folderData As EdmCmdData In ppoData
                confirmList += folderData.mbsStrData1 + vbNewLine
            Next

            If MsgBox(confirmList, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                FindFiles(poCmd, ppoData)
            End If

        End If

    End Sub

    Private Sub FindFiles(ByVal poCmd As EdmCmd, ByRef ppoData As System.Array)

        Dim eVault As EdmVault5 = Nothing
        Dim eFolder As IEdmFolder6 = Nothing

        Try
            eVault = poCmd.mpoVault

            Dim processedList As New List(Of String)

            Dim swApp As SldWorks = CreateObject("SldWorks.Application")

            For Each folderData As EdmCmdData In ppoData
                processedList.Add(folderData.mbsStrData1)

                eFolder = eVault.GetFolderFromPath(folderData.mbsStrData1)

                TraverseFolder(swApp, eFolder)
            Next

            swApp.ExitApp()

        Catch ex As System.Exception

            MsgBox("The following error occurred:" & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        End Try

    End Sub

    Sub TraverseFolder(swApp As SldWorks, folder As IEdmFolder5)

        Dim pdmFilePos As IEdmPos5
        pdmFilePos = folder.GetFirstFilePosition()

        While pdmFilePos.IsNull = False
            Dim eFile As IEdmFile5
            eFile = folder.GetNextFile(pdmFilePos)

            If eFile.Name.Substring(Len(eFile.Name) - 6, 6) = "sldprt" Then

                Dim eFileCard As IEdmEnumeratorVariable8 = eFile.GetEnumeratorVariable()
                Dim variableValue As Object
                eFileCard.GetVar("Laser", "@", variableValue)

                If variableValue.ToString().ToLower() = "x" Then
                    If eFile.IsLocked = True Then
                        If eFile.LockedByUser.Name.ToLower() = "admin" Then
                            SetMaterial(swApp, eFile.LockPath)

                            Dim streamWriter As New StreamWriter(LOGPATH)
                            streamWriter.WriteLine($"Sucess setting material: {eFile.Name}")
                            streamWriter.Close()
                        End If
                    Else
                        Dim streamWriter As New StreamWriter(LOGPATH_ERROR)
                        streamWriter.WriteLine($"Not checked out to admin: {eFile.Name}")
                        streamWriter.Close()
                    End If
                End If
            End If

        End While

        Dim pdmSubFolderPos As IEdmPos5
        pdmSubFolderPos = folder.GetFirstSubFolderPosition()

        While Not pdmSubFolderPos.IsNull
            Dim pdmSubFolder As IEdmFolder5
            pdmSubFolder = folder.GetNextSubFolder(pdmSubFolderPos)
            TraverseFolder(swApp, pdmSubFolder)
        End While

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


End Class
