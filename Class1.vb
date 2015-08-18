Public Class Class1
    Implements StatConnectorCommonLib.IStatConnector

    Public Sub AddGraphicsDevice(ByVal bstrName As String, ByVal pDevice As StatConnectorCommonLib.ISGFX) Implements StatConnectorCommonLib.IStatConnector.AddGraphicsDevice

    End Sub

    Public Sub Close() Implements StatConnectorCommonLib.IStatConnector.Close

    End Sub

    Public Function Evaluate(ByVal bstrExpression As String) As Object Implements StatConnectorCommonLib.IStatConnector.Evaluate

    End Function

    Public Sub EvaluateNoReturn(ByVal bstrExpression As String) Implements StatConnectorCommonLib.IStatConnector.EvaluateNoReturn

    End Sub

    Public Function GetConnectorInformation(ByVal lInformationType As StatConnectorCommonLib.InformationType) As String Implements StatConnectorCommonLib.IStatConnector.GetConnectorInformation

    End Function

    Public Function GetErrorId() As Integer Implements StatConnectorCommonLib.IStatConnector.GetErrorId

    End Function

    Public Function GetErrorText() As String Implements StatConnectorCommonLib.IStatConnector.GetErrorText

    End Function

    Public Function GetInterpreterInformation(ByVal lInformationType As StatConnectorCommonLib.InformationType) As String Implements StatConnectorCommonLib.IStatConnector.GetInterpreterInformation

    End Function

    Public Function GetServerInformation(ByVal lInformationType As StatConnectorCommonLib.InformationType) As String Implements StatConnectorCommonLib.IStatConnector.GetServerInformation

    End Function

    Public Sub GetSupportedTypes(ByRef pulTypeMask As Integer) Implements StatConnectorCommonLib.IStatConnector.GetSupportedTypes

    End Sub

    Public Function GetSymbol(ByVal bstrSymbolName As String) As Object Implements StatConnectorCommonLib.IStatConnector.GetSymbol

    End Function

    Public Sub Init(ByVal bstrConnectorName As String) Implements StatConnectorCommonLib.IStatConnector.Init

    End Sub

    Public Sub RemoveGraphicsDevice(ByVal bstrName As String) Implements StatConnectorCommonLib.IStatConnector.RemoveGraphicsDevice

    End Sub

    Public Sub SetCharacterOutputDevice(ByVal pCharDevice As StatConnectorCommonLib.IStatConnectorCharacterDevice) Implements StatConnectorCommonLib.IStatConnector.SetCharacterOutputDevice

    End Sub

    Public Sub SetErrorDevice(ByVal pCharDevice As StatConnectorCommonLib.IStatConnectorCharacterDevice) Implements StatConnectorCommonLib.IStatConnector.SetErrorDevice

    End Sub

    Public Sub SetSymbol(ByVal bstrSymbolName As String, ByVal vData As Object) Implements StatConnectorCommonLib.IStatConnector.SetSymbol

    End Sub

    Public Sub SetTracingDevice(ByVal pCharDevice As StatConnectorCommonLib.IStatConnectorCharacterDevice) Implements StatConnectorCommonLib.IStatConnector.SetTracingDevice

    End Sub

    Public Sub SetUserInterfaceAgent(ByVal pUIAgent As StatConnectorCommonLib.IStatConnectorUIAgent) Implements StatConnectorCommonLib.IStatConnector.SetUserInterfaceAgent

    End Sub
End Class
