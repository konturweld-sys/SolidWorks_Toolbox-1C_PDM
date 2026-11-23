Option Explicit

Private Const SERVER_NAME As String = "192.168.50.173"
Private Const INFOBASE_NAME As String = "1C_pdm"
Private Const INFOBASE_USER As String = "Администратор"
Private Const COMMON_MODULE_NAME As String = "ПДМ_ИнтеграцияВызов"
Private Const SERVER_FUNCTION_NAME As String = "СоздатьСтандартноеИзделиеИзCAD"
Private Const CLIENT_LOG_PATH As String = "D:\1C Log\sw_integration_log.txt"

Public Sub main()

    Dim swApp As Object
    Dim swModel As Object
    Dim swConfig As Object
    Dim connector As Object
    Dim connection As Object
    Dim commonModule As Object

    Dim configurationName As String
    Dim hexConfigurationName As String
    Dim password As String
    Dim connectionString As String
    Dim result As Variant

    Dim stage As String
    Dim errorText As String

    On Error GoTo ErrorHandler

    stage = "START"
    LogMessage "STAGE=" & stage

    stage = "GET_SOLIDWORKS"
    LogMessage "STAGE=" & stage

    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Err.Raise vbObjectError + 1000, , _
            "Не удалось получить объект SolidWorks."
    End If

    stage = "GET_ACTIVE_MODEL"
    LogMessage "STAGE=" & stage

    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        Err.Raise vbObjectError + 1001, , _
            "В SolidWorks нет активного документа."
    End If

    LogMessage "DOCUMENT_TITLE=" & swModel.GetTitle
    LogMessage "DOCUMENT_PATH=" & swModel.GetPathName

    stage = "GET_ACTIVE_CONFIGURATION"
    LogMessage "STAGE=" & stage

    Set swConfig = swModel.ConfigurationManager.ActiveConfiguration

    If swConfig Is Nothing Then
        Err.Raise vbObjectError + 1002, , _
            "Не удалось получить активную конфигурацию."
    End If

    configurationName = CStr(swConfig.Name)

    If Len(Trim$(configurationName)) = 0 Then
        Err.Raise vbObjectError + 1003, , _
            "Имя активной конфигурации пустое."
    End If

    LogMessage "ACTIVE_CONFIGURATION=" & configurationName

    stage = "ENCODE_CONFIGURATION"
    LogMessage "STAGE=" & stage

    hexConfigurationName = EncodeUtf16Hex(configurationName)

    LogMessage "HEX_LENGTH=" & CStr(Len(hexConfigurationName))
    LogMessage "HEX_VALUE=" & hexConfigurationName

    stage = "GET_PASSWORD"
    LogMessage "STAGE=" & stage

    password = InputBox( _
        "Введите пароль пользователя 1С: " & INFOBASE_USER, _
        "Подключение к 1С:PDM")

    If Len(password) = 0 Then
        Err.Raise vbObjectError + 1004, , _
            "Подключение отменено: пароль не введён."
    End If

    stage = "CREATE_CONNECTOR"
    LogMessage "STAGE=" & stage

    Set connector = CreateObject("V83.COMConnector")

    If connector Is Nothing Then
        Err.Raise vbObjectError + 1005, , _
            "Не удалось создать V83.COMConnector."
    End If

    stage = "CONNECT"
    LogMessage "STAGE=" & stage

    connectionString = _
        "Srvr=""" & SERVER_NAME & """;" & _
        "Ref=""" & INFOBASE_NAME & """;" & _
        "Usr=""" & INFOBASE_USER & """;" & _
        "Pwd=""" & password & """;"

    Set connection = connector.Connect(connectionString)

    If connection Is Nothing Then
        Err.Raise vbObjectError + 1006, , _
            "COMConnector не вернул соединение с информационной базой."
    End If

    LogMessage "CONNECT_OK"

    stage = "GET_COMMON_MODULE"
    LogMessage "STAGE=" & stage

    Set commonModule = CallByName( _
        connection, COMMON_MODULE_NAME, VbGet)

    If commonModule Is Nothing Then
        Err.Raise vbObjectError + 1007, , _
            "Не удалось получить общий модуль " & COMMON_MODULE_NAME
    End If

    LogMessage "COMMON_MODULE=" & COMMON_MODULE_NAME
    LogMessage "MODULE_ACCESS_OK"

    stage = "CALL_SERVER_FUNCTION"
    LogMessage "STAGE=" & stage
    LogMessage "FUNCTION=" & SERVER_FUNCTION_NAME
    LogMessage "ARGUMENT_HEX=" & hexConfigurationName

    result = CallByName( _
        commonModule, _
        SERVER_FUNCTION_NAME, _
        VbMethod, _
        hexConfigurationName)

    LogMessage "SERVER_RESULT=" & CStr(result)
    LogMessage "STAGE=FINISHED"

    MsgBox _
        "Вызов 1С завершён." & vbCrLf & _
        "Активная конфигурация: " & configurationName & vbCrLf & _
        "Результат: " & CStr(result), _
        vbInformation, _
        "SolidWorks > 1C:PDM"

CleanExit:
    On Error Resume Next
    Set commonModule = Nothing
    Set connection = Nothing
    Set connector = Nothing
    Set swConfig = Nothing
    Set swModel = Nothing
    Set swApp = Nothing
    Exit Sub

ErrorHandler:
    errorText = Err.Description

    LogMessage "FAILED_STAGE=" & stage
    LogMessage "FAILED_NUMBER=" & CStr(Err.Number)
    LogMessage "FAILED_ERROR=" & errorText

    MsgBox _
        "Ошибка на этапе: " & stage & vbCrLf & _
        "Номер: " & CStr(Err.Number) & vbCrLf & _
        "Описание: " & errorText & vbCrLf & vbCrLf & _
        "Подробности: " & CLIENT_LOG_PATH, _
        vbCritical, _
        "SolidWorks > 1C:PDM"

    Resume CleanExit

End Sub

Private Function EncodeUtf16Hex(ByVal sourceText As String) As String

    Dim i As Long
    Dim charCode As Long
    Dim result As String

    result = ""

    For i = 1 To Len(sourceText)
        charCode = AscW(Mid$(sourceText, i, 1))

        If charCode < 0 Then
            charCode = charCode + 65536
        End If

        result = result & Right$("0000" & Hex$(charCode), 4)
    Next i

    EncodeUtf16Hex = result

End Function

Private Sub LogMessage(ByVal messageText As String)

    Dim fileNumber As Integer

    On Error Resume Next

    fileNumber = FreeFile

    Open CLIENT_LOG_PATH For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
        " | " & messageText
    Close #fileNumber

    On Error GoTo 0

End Sub
