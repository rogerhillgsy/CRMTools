' =====================================================================
'
'  This file is part of the Microsoft Dynamics CRM SDK code samples.
'
'  Copyright (C) Microsoft Corporation.  All rights reserved.
'
'  This source code is intended only as a supplement to Microsoft
'  Development Tools and/or on-line documentation.  See these other
'  materials for detailed information regarding Microsoft code samples.
'
'  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
'  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
'  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'  PARTICULAR PURPOSE.
'
' =====================================================================
Imports System.Globalization
Imports System.IdentityModel.Tokens
Imports System.IO
Imports System.Net
Imports System.ServiceModel
Imports System.ServiceModel.Description
Imports System.ServiceModel.Security
Imports System.Text
Imports System.Xml

Imports Microsoft.IdentityModel.Protocols.WSTrust

Imports Microsoft.Crm.Services.Utility

Namespace Microsoft.Crm.Sdk.Samples
    ''' <summary>
    ''' Utility to authenticate Microsoft account and Microsoft Office 365 (i.e. OSDP / OrgId) users 
    ''' without using the classes exposed in Microsoft.Xrm.Sdk.dll
    ''' </summary>
    Public NotInheritable Class WsdlTokenManager
        Private Shared FederationMetadataUrlFormat As String = "https://nexus.passport{0}.com/federationmetadata/2007-06/FederationMetaData.xml"

        Private Const DeviceTokenResponseXPath As String = "S:Envelope/S:Body/wst:RequestSecurityTokenResponse/wst:RequestedSecurityToken"
        Private Const UserTokenResponseXPath As String = "S:Envelope/S:Body/wst:RequestSecurityTokenResponse/wst:RequestedSecurityToken"

#Region "Templates"
        Private Const DeviceTokenTemplate As String = _
            "<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf & _
            "        <s:Envelope " & ControlChars.CrLf & _
            "          xmlns:s=""http://www.w3.org/2003/05/soap-envelope"" " & ControlChars.CrLf & _
            "          xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" " & ControlChars.CrLf & _
            "          xmlns:wsp=""http://schemas.xmlsoap.org/ws/2004/09/policy"" " & ControlChars.CrLf & _
            "          xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"" " & ControlChars.CrLf & _
            "          xmlns:wsa=""http://www.w3.org/2005/08/addressing"" " & ControlChars.CrLf & _
            "          xmlns:wst=""http://schemas.xmlsoap.org/ws/2005/02/trust"">" & ControlChars.CrLf & _
            "           <s:Header>" & ControlChars.CrLf & _
            "            <wsa:Action s:mustUnderstand=""1"">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action>" & ControlChars.CrLf & _
            "            <wsa:To s:mustUnderstand=""1"">http://Passport.NET/tb</wsa:To>    " & ControlChars.CrLf & _
            "            <wsse:Security>" & ControlChars.CrLf & _
            "              <wsse:UsernameToken wsu:Id=""devicesoftware"">" & ControlChars.CrLf & _
            "                <wsse:Username>{0:deviceName}</wsse:Username>" & ControlChars.CrLf & _
            "                <wsse:Password>{1:password}</wsse:Password>" & ControlChars.CrLf & _
            "              </wsse:UsernameToken>" & ControlChars.CrLf & _
            "            </wsse:Security>" & ControlChars.CrLf & _
            "          </s:Header>" & ControlChars.CrLf & _
            "          <s:Body>" & ControlChars.CrLf & _
            "            <wst:RequestSecurityToken Id=""RST0"">" & ControlChars.CrLf & _
            "                 <wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType>" & ControlChars.CrLf & _
            "                 <wsp:AppliesTo>" & ControlChars.CrLf & _
            "                    <wsa:EndpointReference>" & ControlChars.CrLf & _
            "                       <wsa:Address>http://Passport.NET/tb</wsa:Address>" & ControlChars.CrLf & _
            "                    </wsa:EndpointReference>" & ControlChars.CrLf & _
            "                 </wsp:AppliesTo>" & ControlChars.CrLf & _
            "              </wst:RequestSecurityToken>" & ControlChars.CrLf & _
            "          </s:Body>" & ControlChars.CrLf & _
            "        </s:Envelope>" & ControlChars.CrLf & _
            "        "

        Private Const UserTokenTemplate As String = _
            "<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf & _
            "    <s:Envelope " & ControlChars.CrLf & _
            "      xmlns:s=""http://www.w3.org/2003/05/soap-envelope"" " & ControlChars.CrLf & _
            "      xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" " & ControlChars.CrLf & _
            "      xmlns:wsp=""http://schemas.xmlsoap.org/ws/2004/09/policy"" " & ControlChars.CrLf & _
            "      xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"" " & ControlChars.CrLf & _
            "      xmlns:wsa=""http://www.w3.org/2005/08/addressing"" " & ControlChars.CrLf & _
            "      xmlns:wst=""http://schemas.xmlsoap.org/ws/2005/02/trust"">" & ControlChars.CrLf & _
            "       <s:Header>" & ControlChars.CrLf & _
            "        <wsa:Action s:mustUnderstand=""1"">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action>" & ControlChars.CrLf & _
            "        <wsa:To s:mustUnderstand=""1"">http://Passport.NET/tb</wsa:To>    " & ControlChars.CrLf & _
            "       <ps:AuthInfo Id=""PPAuthInfo"" xmlns:ps=""http://schemas.microsoft.com/LiveID/SoapServices/v1"">" & ControlChars.CrLf & _
            "             <ps:HostingApp>{0:clientId}</ps:HostingApp>" & ControlChars.CrLf & "          </ps:AuthInfo>" & ControlChars.CrLf & _
            "          <wsse:Security>" & ControlChars.CrLf & _
            "             <wsse:UsernameToken wsu:Id=""user"">" & ControlChars.CrLf & _
            "                <wsse:Username>{1:userName}</wsse:Username>" & ControlChars.CrLf & _
            "                <wsse:Password>{2:password}</wsse:Password>" & ControlChars.CrLf & _
            "             </wsse:UsernameToken>" & ControlChars.CrLf & _
            "             {3:binarySecurityToken}" & ControlChars.CrLf & _
            "          </wsse:Security>" & ControlChars.CrLf & _
            "      </s:Header>" & ControlChars.CrLf & _
            "      <s:Body>" & ControlChars.CrLf & _
            "        <wst:RequestSecurityToken Id=""RST0"">" & ControlChars.CrLf & _
            "             <wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType>" & ControlChars.CrLf & _
            "             <wsp:AppliesTo>" & ControlChars.CrLf & _
            "                <wsa:EndpointReference>" & ControlChars.CrLf & _
            "                   <wsa:Address>{4:partnerUrl}</wsa:Address>" & ControlChars.CrLf & _
            "                </wsa:EndpointReference>" & ControlChars.CrLf & _
            "             </wsp:AppliesTo>" & ControlChars.CrLf & _
            "             <wsp:PolicyReference URI=""{5:policy}""/>" & ControlChars.CrLf & _
            "          </wst:RequestSecurityToken>" & ControlChars.CrLf & _
            "      </s:Body>" & ControlChars.CrLf & _
            "    </s:Envelope>"
        Private Const BinarySecurityToken As String = _
            "<wsse:BinarySecurityToken ValueType=""urn:liveid:device"">" & ControlChars.CrLf & _
            "                   {0:deviceTokenValue}" & ControlChars.CrLf & _
            "</wsse:BinarySecurityToken>"
#End Region

        ''' <summary>
        ''' This shows the method to retrieve the security ticket for the Microsoft account user or OrgId user
        ''' without using any certificate for authentication.     
        ''' </summary>
        ''' <param name="credentials">User credentials that should be used to connect to the server</param>
        ''' <param name="appliesTo">Indicates the AppliesTo that is required for the token</param>
        ''' <param name="policy">Policy that should be used when communicating with the server</param>
        Private Sub New()
        End Sub
        Public Shared Function Authenticate(ByVal credentials As ClientCredentials, ByVal appliesTo As String, ByVal policy As String) _
               As SecurityToken
            Return Authenticate(credentials, appliesTo, policy, Nothing)
        End Function

        ''' <summary>
        ''' This shows the method to retrieve the security ticket for the Microsoft account user or OrgId user
        ''' without using any certificate for authentication.     
        ''' </summary>
        ''' <param name="credentials">User credentials that should be used to connect to the server</param>
        ''' <param name="appliesTo">Indicates the AppliesTo that is required for the token</param>
        ''' <param name="policy">Policy that should be used when communicating with the server</param>
        ''' <param name="issuerUri">URL for the current token issuer</param> 
        Public Shared Function Authenticate(ByVal credentials As ClientCredentials, ByVal appliesTo As String, ByVal policy As String, _
                                            ByVal issuerUri As Uri) As SecurityToken
            Dim serviceUrl As String = issuerUri.ToString()
            ' if serviceUrl starts with "https://login.live.com", it means Microsoft account authentication is needed otherwise OSDP authentication.
            If (Not String.IsNullOrEmpty(serviceUrl)) AndAlso serviceUrl.StartsWith("https://login.live.com") Then
                serviceUrl = GetServiceEndpoint(issuerUri)

                'Authenticate the device
                Dim deviceCredentials As ClientCredentials = DeviceIdManager.LoadOrRegisterDevice(issuerUri)
                Dim deviceToken As String = IssueDeviceToken(serviceUrl, deviceCredentials)
                'Use the device token to authenticate the user
                Return Issue(serviceUrl, credentials, appliesTo, policy, Guid.NewGuid(), deviceToken)
            End If
            ' Default to OSDP authentication.
            Return Issue(serviceUrl, credentials, appliesTo, policy, Guid.NewGuid(), Nothing)
        End Function

#Region "Private Methods"
        Private Shared Function Issue(ByVal serviceUrl As String, ByVal credentials As ClientCredentials, ByVal partner As String, _
                                      ByVal policy As String, ByVal clientId As Guid, ByVal deviceToken As String) As SecurityToken
            Dim soapEnvelope As String

            If Nothing IsNot deviceToken Then
                soapEnvelope = String.Format(CultureInfo.InvariantCulture, UserTokenTemplate, clientId.ToString(), credentials.UserName.UserName, credentials.UserName.Password, String.Format(CultureInfo.InvariantCulture, BinarySecurityToken, deviceToken), partner, policy)
            Else
                soapEnvelope = String.Format(CultureInfo.InvariantCulture, UserTokenTemplate, clientId.ToString(), credentials.UserName.UserName, credentials.UserName.Password, String.Empty, partner, policy)
            End If
            Dim doc As XmlDocument = CallOnlineSoapServices(serviceUrl, "POST", soapEnvelope)

            Dim namespaceManager As XmlNamespaceManager = CreateNamespaceManager(doc.NameTable)

            Dim serializedTokenNode As XmlNode = SelectNode(doc, namespaceManager, UserTokenResponseXPath)
            If Nothing Is serializedTokenNode Then
                Throw New InvalidOperationException("Unable to Issue User Token due to error" & Environment.NewLine & FormatXml(doc))
            End If
            Return ConvertToToken(serializedTokenNode.InnerXml)
        End Function

        Private Shared Function IssueDeviceToken(ByVal serviceUrl As String, ByVal deviceCredentials As ClientCredentials) As String
            Dim soapEnvelope As String = String.Format(CultureInfo.InvariantCulture, DeviceTokenTemplate, _
                                                       deviceCredentials.UserName.UserName, deviceCredentials.UserName.Password)
            Dim doc As XmlDocument = CallOnlineSoapServices(serviceUrl, "POST", soapEnvelope)

            Dim namespaceManager As XmlNamespaceManager = CreateNamespaceManager(doc.NameTable)
            Dim tokenNode As XmlNode = SelectNode(doc, namespaceManager, DeviceTokenResponseXPath)
            If Nothing Is tokenNode Then
                Throw New InvalidOperationException("Unable to Issue Device Token due to error" & Environment.NewLine & FormatXml(doc))
            End If

            Return tokenNode.InnerXml
        End Function

        Private Shared Function FormatXml(ByVal doc As XmlDocument) As String
            'Create the writer settings
            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.OmitXmlDeclaration = True

            'Write the data
            Using stringWriter_Renamed As New StringWriter(CultureInfo.InvariantCulture)
                Using writer As XmlWriter = XmlWriter.Create(stringWriter_Renamed, settings)
                    doc.Save(writer)
                End Using

                'Return the writer's contents
                Return stringWriter_Renamed.ToString()
            End Using
        End Function

        Private Shared Function GetServiceEndpoint(ByVal issuerUri As Uri) As String
            Dim passportEnvironment As String = DeviceIdManager.DiscoverEnvironment(issuerUri)
            Dim federationMetadataUrl As String = String.Format(CultureInfo.InvariantCulture, FederationMetadataUrlFormat, _
                                                    If(String.IsNullOrEmpty(passportEnvironment), Nothing, "-" & passportEnvironment))

            Dim doc As XmlDocument = CallOnlineSoapServices(federationMetadataUrl, "GET", Nothing)

            Dim namespaceManager As New XmlNamespaceManager(doc.NameTable)
            namespaceManager.AddNamespace("fed", "http://docs.oasis-open.org/wsfed/federation/200706")
            namespaceManager.AddNamespace("wsa", "http://www.w3.org/2005/08/addressing")
            namespaceManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")
            namespaceManager.AddNamespace("core", "urn:oasis:names:tc:SAML:2.0:metadata")

            Return SelectNode(doc, namespaceManager, "//core:EntityDescriptor/core:RoleDescriptor[@xsi:type='fed:ApplicationServiceType']" & _
                              "/fed:ApplicationServiceEndpoint/wsa:EndpointReference/wsa:Address").InnerText.Trim()
        End Function

        Private Shared Function CallOnlineSoapServices(ByVal serviceUrl As String, ByVal method As String, ByVal soapMessageEnvelope As String) As XmlDocument
            ' Buid the web request
            Dim url As String = serviceUrl
            Dim request As WebRequest = WebRequest.Create(url)
            request.Method = method
            request.Timeout = 180000
            If method = "POST" Then
                ' If we are "posting" then this is always a SOAP message
                request.ContentType = "application/soap+xml; charset=UTF-8"
            End If

            ' If a SOAP envelope is supplied, then we need to write to the request stream
            ' If there isn't a SOAP message supplied then continue onto just process the raw XML
            If Not String.IsNullOrEmpty(soapMessageEnvelope) Then
                Dim bytes() As Byte = Encoding.UTF8.GetBytes(soapMessageEnvelope)
                Using str As Stream = request.GetRequestStream()
                    str.Write(bytes, 0, bytes.Length)
                    str.Close()
                End Using
            End If

            ' Read the response into an XmlDocument and return that doc
            Dim xml As String
            Using response As WebResponse = request.GetResponse()
                Using reader As New StreamReader(response.GetResponseStream())
                    xml = reader.ReadToEnd()
                End Using
            End Using

            Dim document As New XmlDocument()
            document.LoadXml(xml)
            Return document
        End Function

        Private Shared Function ConvertToToken(ByVal xml As String) As SecurityToken
            Dim binding As New WS2007FederationHttpBinding(WSFederationHttpSecurityMode.TransportWithMessageCredential, False)
            Dim factory As New Microsoft.IdentityModel.Protocols.WSTrust.WSTrustChannelFactory(binding, New EndpointAddress("https://null-EndPoint"))
            factory.TrustVersion = TrustVersion.WSTrustFeb2005

            Dim trustChannel As Microsoft.IdentityModel.Protocols.WSTrust.WSTrustChannel = CType(factory.CreateChannel(), Microsoft.IdentityModel.Protocols.WSTrust.WSTrustChannel)

            Dim response As RequestSecurityTokenResponse = trustChannel.WSTrustResponseSerializer.CreateInstance()
            response.RequestedSecurityToken = New RequestedSecurityToken(LoadXml(xml).DocumentElement)
            response.IsFinal = True

            Dim requestToken As New RequestSecurityToken(WSTrustFeb2005Constants.RequestTypes.Issue)
            requestToken.KeyType = WSTrustFeb2005Constants.KeyTypes.Symmetric

            Return trustChannel.GetTokenFromResponse(requestToken, response)
        End Function

        Private Shared Function LoadXml(ByVal xml As String) As XmlDocument
            Dim settings As New XmlReaderSettings()
            settings.IgnoreWhitespace = True
            settings.ConformanceLevel = ConformanceLevel.Fragment

            Using memoryReader As New StringReader(xml)
                Using reader As XmlReader = XmlReader.Create(memoryReader, settings)
                    'Create the xml document
                    Dim doc As New XmlDocument()
                    doc.XmlResolver = Nothing

                    'Load the data from the reader
                    doc.Load(reader)

                    Return doc
                End Using
            End Using
        End Function

        Private Shared Function SelectNode(ByVal document As XmlDocument, ByVal namespaceManager As XmlNamespaceManager, _
                                           ByVal xPathToNode As String) As XmlNode
            Dim nodes As XmlNodeList = document.SelectNodes(xPathToNode, namespaceManager)
            If nodes IsNot Nothing AndAlso nodes.Count > 0 AndAlso nodes(0) IsNot Nothing Then
                Return nodes(0)
            End If
            Return Nothing
        End Function

        Private Shared Function CreateNamespaceManager(ByVal nameTable As XmlNameTable) As XmlNamespaceManager
            Dim namespaceManager As New XmlNamespaceManager(nameTable)
            namespaceManager.AddNamespace("wst", "http://schemas.xmlsoap.org/ws/2005/02/trust")
            namespaceManager.AddNamespace("S", "http://www.w3.org/2003/05/soap-envelope")
            namespaceManager.AddNamespace("wsu", "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd")
            Return namespaceManager
        End Function
#End Region
    End Class
End Namespace