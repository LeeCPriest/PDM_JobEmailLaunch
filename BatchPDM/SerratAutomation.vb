Imports EdmLib

Public Class SerratAutomation
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo

        Try
            poInfo.mbsAddInName = "SerratAutomation"
            poInfo.mbsCompany = "Written by Lee Priest www.cadinnovations.ca"

            'Specify information to display in the add-in's Properties dialog box
            poInfo.mbsDescription = "Custom PDM functionality"
            poInfo.mlAddInVersion = 1.0
            poInfo.mlRequiredVersionMajor = 8
            poInfo.mlRequiredVersionMinor = 0

            poCmdMgr.AddHook(EdmCmdType.EdmCmd_PostAdd)
        Catch
        End Try

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd

        If poCmd.meCmdType = EdmCmdType.EdmCmd_PostAdd Then

            PopulateCode(poCmd, ppoData)

        End If

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

End Class
