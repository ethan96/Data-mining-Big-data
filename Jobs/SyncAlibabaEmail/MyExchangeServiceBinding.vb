Imports SyncInvalidEmail.AEUOWA
Imports System.Web.Services.Protocols
Imports System.Xml

Public Class MyExchangeServiceBinding
    Inherits ExchangeServiceBinding

    Protected Overrides Function GetReaderForMessage(ByVal message As SoapClientMessage, ByVal bufferSize As Integer) As XmlReader
        Dim retval As XmlReader = MyBase.GetReaderForMessage(message, bufferSize)
        Dim xrt As XmlTextReader = CType(retval, XmlTextReader)
        If (Not (xrt) Is Nothing) Then
            xrt.Normalization = False
        End If
        Return retval
    End Function
End Class