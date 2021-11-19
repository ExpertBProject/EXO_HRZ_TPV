Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class EXO_ODLN
    Private objGlobal As EXO_UIAPI.EXO_UIAPI

    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            'Apaño por un error que da EXO_Basic.dll al consultar infoEvento.FormTypeEx
            Try
                If infoEvento.FormTypeEx <> "" Then

                End If
            Catch ex As Exception
                Return False
            End Try

            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "140"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "140"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "140"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "140"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                End If
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Form_Load(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Visible = False

            'Buscar XML de update
            objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Campos identificación del cobro"
            oItem = oForm.Items.Add("lblCobro", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("89").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.LinkTo = "222"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Referencia Cobro: "
            oItem.TextStyle = 1


            oItem = oForm.Items.Add("txtCDEntry", BoFormItemTypes.it_EXTEDIT)
            oItem.Top = oForm.Items.Item("103").Top
            oItem.Left = oForm.Items.Item("222").Left
            oItem.Height = oForm.Items.Item("222").Height
            oItem.Width = oForm.Items.Item("222").Width
            oItem.LinkTo = "lblCobro"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "ODLN", "U_EXO_CDOCENTRY")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oItem = oForm.Items.Add("IrPago", BoFormItemTypes.it_LINKED_BUTTON)
            oItem.Top = oForm.Items.Item("txtCDEntry").Top  'Incidencia
            oItem.Left = oForm.Items.Item("229").Left
            oItem.Height = oForm.Items.Item("87").Height
            oItem.Width = oForm.Items.Item("87").Width
            oItem.LinkTo = "txtCDEntry"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.LinkedButton).LinkedObjectType = CType(BoObjectTypes.oIncomingPayments, String)
            CType(oItem.Specific, SAPbouiCOM.LinkedButton).Item.LinkTo = "txtCDEntry"


            oItem = oForm.Items.Add("lblCDEntry", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("txtCDEntry").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.LinkTo = "txtCDEntry"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Nº Interno"

            oItem = oForm.Items.Add("txtCDNum", BoFormItemTypes.it_EDIT)
            oItem.Top = oForm.Items.Item("27").Top
            oItem.Left = oForm.Items.Item("222").Left
            oItem.Height = oForm.Items.Item("222").Height
            oItem.Width = oForm.Items.Item("222").Width
            oItem.LinkTo = "lblCobro"
            oItem.FromPane = 0
            oItem.ToPane = 0
            oItem.Enabled = False
            CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "ODLN", "U_EXO_CDOCNUM")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oItem = oForm.Items.Add("lblCDNum", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("txtCDNum").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.LinkTo = "txtCDNum"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Nº Cobro"

            oItem = oForm.Items.Add("txtCTipo", BoFormItemTypes.it_COMBO_BOX)
            oItem.Top = oForm.Items.Item("29").Top
            oItem.Left = oForm.Items.Item("222").Left
            oItem.Height = oForm.Items.Item("222").Height
            oItem.Width = oForm.Items.Item("222").Width
            oItem.LinkTo = "lblCobro"
            oItem.FromPane = 0
            oItem.ToPane = 0
            oItem.DisplayDesc = True
            oItem.Enabled = False
            CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "ODLN", "U_EXO_CTIPO")
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oItem = oForm.Items.Add("lblCTipo", BoFormItemTypes.it_STATIC)
            oItem.Top = oForm.Items.Item("txtCTipo").Top
            oItem.Left = oForm.Items.Item("230").Left
            oItem.Height = oForm.Items.Item("230").Height
            oItem.Width = oForm.Items.Item("230").Width
            oItem.LinkTo = "txtCTipo"
            oItem.FromPane = 0
            oItem.ToPane = 0
            CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Tipo"
#End Region
#Region "Botones"
            oItem = oForm.Items.Add("btnCOBROT", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("10000330").Left - (oForm.Items.Item("10000330").Width * 2) - 5
            oItem.Width = oForm.Items.Item("10000330").Width
            oItem.Top = oForm.Items.Item("46").Top + 25
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            Dim oBtnAct As SAPbouiCOM.Button
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "PAGO TOTAL"
            oItem.TextStyle = 1
            oItem.LinkTo = "46"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oItem = oForm.Items.Add("btnCOBROC", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("btnCOBROT").Left + oForm.Items.Item("btnCOBROT").Width + 2
            oItem.Width = (oForm.Items.Item("btnCOBROT").Width * 2) - 30
            oItem.Top = oForm.Items.Item("btnCOBROT").Top
            oItem.Height = oForm.Items.Item("btnCOBROT").Height
            oItem.Enabled = False
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Cancelar Pago Asociado"
            oItem.TextStyle = 1
            oItem.LinkTo = "btnCOBROT"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region
            oForm.Visible = True

            EventHandler_Form_Load = True

        Catch ex As Exception
            oForm.Visible = True
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnCOBROT"
                    If pVal.ActionSuccess = True Then

                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
End Class
