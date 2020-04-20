Public Class clsArticulo
#Region "Control de Eventos en SAP BO"

    Public Function ManejaEventoForm(ByVal FormUID As String, ByVal oEvent As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByVal oCompany As SAPbobsCOM.Company) As Boolean
        Dim bBubbleEvent As Boolean = True
        Select Case oEvent.EventType
            Case Is = SAPbouiCOM.BoEventTypes.et_CLICK
                bBubbleEvent = PinchaBoton(FormUID, oEvent, oApplication, oCompany)
            Case Is = SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                'oApplication.MessageBox("No puede ingresar a este moulo")
                bBubbleEvent = CerrarVentana(FormUID, oEvent, oApplication, oCompany)

        End Select
        Return bBubbleEvent

    End Function

    Private Function CargaFormArticulo(ByVal FormUID As String, ByVal oEvent As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByVal oCompany As SAPbobsCOM.Company) As Boolean
        Dim bBubbleEvent As Boolean = True
        If Not oEvent.BeforeAction Then
            Try
                Dim objeto As SAPbouiCOM.Item
                Dim objetoref As SAPbouiCOM.Item
                Dim formulario As SAPbouiCOM.Form

                Dim botonSAP As SAPbouiCOM.Button

                formulario = oApplication.Forms.Item(FormUID)
                'objeto = formulario.Items.Item("2")
                objeto = formulario.Items.Add("btnActItem", SAPbouiCOM.BoFormItemTypes.it_BUTTON)

                objetoref = formulario.Items.Item("2")

                objeto.Top = objetoref.Top
                objeto.Width = objetoref.Width
                objeto.Height = objetoref.Height
                objeto.Left = objetoref.Left + 100

                botonSAP = objeto.Specific
                botonSAP.Caption = "Prueba"
            Catch ex As Exception
                oApplication.MessageBox(ex.Message)
            End Try
        End If
        Return bBubbleEvent


    End Function

    Private Function PinchaBoton(ByVal FormUID As String, ByVal oEvent As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByVal oCompany As SAPbobsCOM.Company) As Boolean
        Dim bBubbleEvent As Boolean = True
        Select Case oEvent.ItemUID
            Case Is = "btnActItem"
                If Not oEvent.BeforeAction Then
                    bBubbleEvent = ProcesaDocumento(FormUID, oEvent, oApplication, oCompany)
                End If

        End Select
        Return bBubbleEvent

    End Function

    Private Function CerrarVentana(ByVal FormUID As String, ByVal oEvent As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByVal oCompany As SAPbobsCOM.Company) As Boolean
        Dim bBubbleEvent As Boolean = True
        If Not oEvent.BeforeAction Then
            Try
                'Dim cerrar As SAPbouiCOM.Form
                Dim usuario As String

                usuario = oCompany.UserName
                oApplication.MessageBox(usuario & "No puede ingresar a este modulo")

                Dim formulario As SAPbouiCOM.Form
                formulario = oApplication.Forms.Item(FormUID)
                formulario.Close()

            Catch ex As Exception
                oApplication.MessageBox(ex.Message)
            End Try
        End If
        Return bBubbleEvent


    End Function
#End Region

#Region "Funciones locales de clase"

    Private Function ProcesaDocumento(ByVal FormUID As String, ByVal oEvent As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByVal oCompany As SAPbobsCOM.Company) As Boolean

        Try
            If Not oEvent.BeforeAction Then
                Dim formulario As SAPbouiCOM.Form

                Dim textoform As SAPbouiCOM.EditText

                formulario = oApplication.Forms.Item(FormUID)

                textoform = formulario.Items.Item("5").Specific

                Dim codart As String = textoform.Value

                oApplication.MessageBox(codart)

            End If

        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try

    End Function

#End Region
End Class
