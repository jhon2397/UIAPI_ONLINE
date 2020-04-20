Public Class clsFormAcceso
    Private WithEvents _SBOA As SAPbouiCOM.Application
    Private _SBOC As SAPbobsCOM.Company

    Public Sub New(ByVal oApplication As SAPbouiCOM.Application, ByVal oCompany As SAPbobsCOM.Company)
        Me._SBOA = oApplication
        Me._SBOC = oCompany
    End Sub

#Region "Eventos de la Clase"

    Public Function FormLoad(ByRef oEvent As SAPbouiCOM.MenuEvent) As Boolean

        REM Establecer ruta de archivo srf
        Dim rutaform As String = "C:\Users\Admin\Google Drive\Documentacion SAP\Capacitaciones\UI-API\Elementos\Formularios\frm_user.srf"

        Dim creaciondepaquetes As SAPbouiCOM.FormCreationParams = _SBOA.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)

        Dim formSAP As SAPbouiCOM.Form = _SBOA.Forms.AddEx(creaciondepaquetes)

        REM Cargar el formulario .srf utilizando metedo de .NET.
        Dim xmlForm As New Xml.XmlDocument
        xmlForm.Load(rutaform)

        REM esta validacion chequea la existencia de un formulario abierto, con el mismo nombre
        Dim contardorform As Integer = 0
        For index As Integer = 0 To _SBOA.Forms.Count - 1

            REM Si ya existe una instancia abierta, contabilizarla
            If _SBOA.Forms.Item(index).TypeEx = "frm_user" Then contardorform += 1
        Next
        contardorform += 1

        creaciondepaquetes.XmlData = xmlForm.InnerXml
        creaciondepaquetes.UniqueID = "frm_user" & "_" & contardorform.ToString() REM crear una nueva instancia
        formSAP = _SBOA.Forms.AddEx(creaciondepaquetes)

        formSAP.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        formSAP.Visible = True

    End Function
#End Region


End Class
