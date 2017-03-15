Module TOOLS
    Private _SBOApplication As SAPbouiCOM.Application
    Public Property SBOApplication() As SAPbouiCOM.Application
        Get
            Return _SBOApplication
        End Get
        Set(ByVal value As SAPbouiCOM.Application)
            _SBOApplication = value
        End Set
    End Property

    Private _Company As SAPbobsCOM.Company
    Public docEntry As String
    Public code As String = ""

    Public Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
        Set(ByVal value As SAPbobsCOM.Company)
            _Company = value
        End Set
    End Property
    Public Sub userField(ByVal oCompany As SAPbobsCOM.Company, ByVal tableName As String, ByVal Descripcion As String, ByVal size As Integer, ByVal namefield As String, ByVal type As SAPbobsCOM.BoFieldTypes, ByVal validation As Boolean, ByVal SBO_app As SAPbouiCOM.Application)
        Dim err As String = ""
        Dim num As Integer = 0
        Dim row As Integer = -1
        Try
            If fieldExist(oCompany, tableName, namefield) = False Then
                Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
                oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUserFieldsMD.TableName = tableName
                oUserFieldsMD.Name = namefield   '"DOCUMENTO"
                oUserFieldsMD.Description = Descripcion  '"DOCUMENTO"
                oUserFieldsMD.Type = type
                If type = 0 Then
                    oUserFieldsMD.EditSize = size
                End If

                If validation = True Then
                    oUserFieldsMD.ValidValues.Value = "1"
                    oUserFieldsMD.ValidValues.Description = "INICIO"
                    oUserFieldsMD.ValidValues.Add()
                End If
                If oUserFieldsMD.Add() <> 0 Then
                    oCompany.GetLastError(num, err)
                    SBO_app.SetStatusBarMessage(num & " " & err, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            End If

            GC.Collect()
        Catch ex As Exception
            SBO_app.SetStatusBarMessage(ex.Message & "  " & num & " " & err, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try


    End Sub
    Private Function fieldExist(oCompany As SAPbobsCOM.Company, tableName As String, namefield As String) As Boolean

        Dim existe As Boolean = False
        Dim record As SAPbobsCOM.Recordset

        record = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        record.DoQuery("CALL AliasID ('" & tableName & "','" & namefield & "')")
        If record.RecordCount > 0 Then
            existe = True
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(record)
        record = Nothing
        GC.Collect()
        Return existe
    End Function
End Module
