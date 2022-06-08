Imports EdmLib

Public Class JobEmailLaunch
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo

        Try
            poInfo.mbsAddInName = "Job Email Launch"
            poInfo.mbsCompany = "Written by Lee Priest - leeclarkepriest@gmail.com"

            'Specify information to display in the add-in's Properties dialog box
            poInfo.mbsDescription = "Add-in to automatically open custom email on workflow transition"
            poInfo.mlAddInVersion = 1.0
            poInfo.mlRequiredVersionMajor = 8
            poInfo.mlRequiredVersionMinor = 0

            'Notify the add-in after a file state change event occurs
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_PostState)
        Catch
        End Try

    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd

        If poCmd.meCmdType = EdmCmdType.EdmCmd_PostState Then 'run the following function when a PostState event occurs (i.e. after the transition has occurred)

            OpenEmail(poCmd, ppoData)

        End If

    End Sub

    Private Sub OpenEmail(ByVal poCmd As EdmCmd, ByRef ppoData As System.Array)

        Dim eVault As EdmVault5 = Nothing
        Dim eFile As IEdmFile8 = Nothing
        Dim eFolder As IEdmFolder6 = Nothing
        Dim eFileCard As IEdmEnumeratorVariable8 = Nothing

        Dim strFilePath As String

        Dim dictVariable1_Value As New Dictionary(Of String, String)
        Dim dictVariable2_Value As New Dictionary(Of String, String)
        Dim dictVariable3_Value As New Dictionary(Of String, String)

        Try
            eVault = poCmd.mpoVault

            Dim eData_First As EdmCmdData = ppoData.GetValue(0) 'get the transition info from the first file in the array

            If eData_First.mbsStrData2 = "Destination State Name" Then 'if the user is transitioning files to the specified destination state

                For Each eCmdData As EdmCmdData In ppoData 'loop through all files that are being transitioned

                    strFilePath = eCmdData.mbsStrData1 'get the path of the file being transitioned

                    If dictVariable1_Value.ContainsKey(strFilePath) = False Then 'for efficiency, it is only necessary to proces each file once

                        eFile = eVault.GetFileFromPath(strFilePath) 'get the PDM file object
                        eFolder = eVault.GetObject(EdmObjectType.EdmObject_Folder, eCmdData.mlObjectID2) 'get the PDM parent folder

                        Dim strVariable1_Value As String = ""
                        Dim strVariable2_Value As String = ""
                        Dim strVariable3_Value As String = ""

                        eFileCard = eFile.GetEnumeratorVariable() 'get the filecard object
                        eFileCard.GetVarFromDb("Variable 1 Name", "@", strVariable1_Value) 'get the variable value
                        eFileCard.GetVarFromDb("Variable 2 Name", "@", strVariable2_Value) 'get the variable value
                        eFileCard.GetVarFromDb("Variable 3 Name", "@", strVariable3_Value) 'get the variable value
                        eFileCard.CloseFile(False) 'close the filecard (can flush be false!?!?)

                        dictVariable1_Value.Add(strFilePath, strVariable1_Value) 'store the variable value in a dictionary with the PDM filepath as the key value
                        dictVariable2_Value.Add(strFilePath, strVariable2_Value) 'store the variable value in a dictionary with the PDM filepath as the key value
                        dictVariable3_Value.Add(strFilePath, strVariable3_Value) 'store the variable value in a dictionary with the PDM filepath as the key value

                    End If

                Next


                Dim emailSubject As String = "Example email subject"
                Dim emailBody = "Email body text goes here" & "%0D" 'note: %0D is the new line character for the 'mailto' protocol

                For i = 0 To dictVariable1_Value.Count - 1 'loop through dictionary items, adding to email body

                    Dim eData As EdmCmdData = ppoData.GetValue(i)

                    emailBody &= "Variable 1 Value: " & dictVariable1_Value(eData.mbsStrData1) & "%0D"
                    emailBody &= "Variable 2 Value: " & dictVariable2_Value(eData.mbsStrData1) & "%0D"
                    emailBody &= "Variable 3 Value: " & dictVariable3_Value(eData.mbsStrData1) & "%0D"

                Next


                'open mail message
                System.Diagnostics.Process.Start("mailto:user@company.com" &
                    "?Subject=" & emailSubject &
                    "&Body=" & emailBody)
            End If

        Catch ex As System.Exception

            MsgBox("The following error occurred:" & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Exclamation, My.Application.Info.AssemblyName)

        End Try

    End Sub

End Class
