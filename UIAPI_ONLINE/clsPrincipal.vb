Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class clsPrincipal

#Region "Variable uso local"
    REM Variables de conexion a SAP BO
    Private WithEvents SBOA As Application
    Private SBOC As SAPbobsCOM.Company

    Private SBOGUI As SboGuiApi

    REM Variables de Menu SAP BO
    Public SBOMENU As SAPbouiCOM.Menus
    Public SBOMENUITEM As SAPbouiCOM.MenuItem

    REM Variables de Filtros
    Dim SBOFiltro As SAPbouiCOM.EventFilter


#End Region

#Region "Secuencia Inicial"
    Public Sub New()
        REM Paso 1: conexion UI-API
        ConectaSBOA()
        REM Paso 2: conexion DI-API
        ConectaSBOC()
        REM Paso 3: Test conexion
        'TesConexion()
        REM Paso 4: Carga de Menu
        CargaMenu()
        REM Paso 5: Aplicar filtros
        AplicarFiltro()

    End Sub

#End Region

#Region "Funciones Locales"

    Private Sub ConectaSBOA()
        Dim CadenaConexion As String = Environment.GetCommandLineArgs.GetValue(1)

        SBOGUI = New SAPbouiCOM.SboGuiApi

        SBOGUI.Connect(CadenaConexion)
        SBOA = SBOGUI.GetApplication()

    End Sub

    Private Sub ConectaSBOC()
        Try
            SBOC = SBOA.Company.GetDICompany
        Catch ex As Exception
            SBOC.MessageBox("Error en conexion: " & ex.ToString, 1, "Ok")
        End Try
    End Sub

    Private Sub TesConexion()
        SBOA.MessageBox("Hola gente")

        Dim TotalUser As SAPbobsCOM.Recordset = SBOC.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim Total As String = String.Empty

        TotalUser.DoQuery("select count(*) as Total From OUSR")
        Total = TotalUser.Fields.Item("Total").Value

        SBOA.MessageBox("Acutalmente, para la compañia " & SBOC.CompanyName & ", Tengo " & Total & " Usuarios")
    End Sub

    Private Sub CargaMenu()
        REM Congelar el Menu
        SBOA.Forms.GetFormByTypeAndCount(169, 1).Freeze(True)

        Dim menu As New Xml.XmlDocument
        SBOMENU = SBOA.Menus

        Try
            menu.Load("C:\Users\Admin\Google Drive\Documentacion SAP\Capacitaciones\UI-API\Elementos\MenuUno.xml")
            SBOA.LoadBatchActions(menu.InnerXml)

        Catch ex As System.IO.FileNotFoundException
            SBOA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Catch ex2 As System.Exception
            SBOA.MessageBox(ex2.Message + ex2.ToString)
        Finally
            REM escongelar el menu Principal
            SBOA.Forms.GetFormByTypeAndCount(169, 1).Freeze(False)
            SBOA.Forms.GetFormByTypeAndCount(169, 1).Update()
        End Try
    End Sub

    Public Sub AplicarFiltro()
        Dim XmlFilters As New Xml.XmlDocument
        Try

            Dim nroevent = BoEventTypes.et_CLICK


            SBOFiltro = New EventFilter

            XmlFilters.Load("C:\Users\Admin\Google Drive\Documentacion SAP\Capacitaciones\UI-API\Elementos\filtro.xml")
            SBOFiltro.LoadFromXML(XmlFilters.InnerXml)

            SBOA.SetFilter(SBOFiltro)

        Catch ex As Exception
            SBOA.MessageBox(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Manejo de Eventos de SAP"

    Private Sub ManejoEventosSBO(ByVal EventType As SAPbouiCOM.BoEventTypes) Handles SBOA.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                SBOA.MessageBox("AddOn Desconectado") ', BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                SBOC.Disconnect()
                End
            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                SBOA.StatusBar.SetText("Desconectando", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                SBOC.Disconnect()
                'End

        End Select
    End Sub

    Private Sub ManejoEventosItem(ByVal FormUID As String, ByRef oEvent As SAPbouiCOM.ItemEvent, ByRef bBubbleEvent As Boolean) Handles SBOA.ItemEvent
        Try
            REM Seleccionar segun el ID unico del documento
            Select Case oEvent.FormTypeEx
                Case Is = "150"

                    Dim frmArticulo As clsArticulo = New clsArticulo()
                    bBubbleEvent = frmArticulo.ManejaEventoForm(FormUID, oEvent, SBOA, SBOC)
                Case Else

            End Select
        Catch ex As Exception
            SBOA.StatusBar.SetText("Evento: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Private Sub ManejoEventosMenu(ByRef oEvent As SAPbouiCOM.MenuEvent, ByRef bBubbleEvent As Boolean) Handles SBOA.MenuEvent
        Try
            If Not oEvent.BeforeAction Then
                REM Seleccionar segun el ID unico del documento
                Select Case oEvent.MenuUID
                    Case Is = "mnuSUNO"

                        Dim frmFormAcceso As clsFormAcceso = New clsFormAcceso(SBOA, SBOC)
                        bBubbleEvent = frmFormAcceso.FormLoad(oEvent)
                    Case Else

                End Select
            End If

        Catch ex As Exception
            SBOA.StatusBar.SetText("Evento: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            bBubbleEvent = True
        Finally

        End Try
    End Sub
#End Region

End Class
