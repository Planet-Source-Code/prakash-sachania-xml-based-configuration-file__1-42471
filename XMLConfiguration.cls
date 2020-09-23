VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* File name   : XMLConfiguration.cls
'*
'* Copyright (c) 2002 by House of Code, Inc. All Rights Reserved.
'*
'* This software is the proprietary information of House of Code, Inc.
'* Use is subject to license terms.
'*
'*
'* @author  Prakash Sachania (prakash.sachania@in.houseofcode.com)
'* @version 1.00
'* @date    January, 2003
'*
'* Known Bugs and/or side effects:Nil
'*************************************************************************************************

'************************  Modification Log ******************************************************
'Date         Modified By     Reason                             tag
'*************************************************************************************************

'*************************************************************************************************
' Module/Form name :    XMLConfiguration
' Abstract :            This class handles reading configuration information from an XML file. It
'                       can also write a new configuration file.
'                       Also, there are functions defined related to encryption and decryption but
'                       are not used in the program. They are meant in case if the configuration
'                       file has to be stored/retrieved in encrypted format
'*************************************************************************************************
' Examples :
'*************************************************************************************************
'
'    '0. create object
'    'set clsConfig = new XMLConfiguration
'
'    '1. Encryption to be used or not (Default is not use)
'    'clsConfig.Encryption = True
'
'    '2. open configuration file. creates if it does not exist
'    clsConfig.LoadConfigurationFile App.Path & "\ej2.xml"
'
'    '3. read a value
'    gsDriverName = clsConfig.ReadValue("system", "driver")
'
'    '4. write a value in the config file and save immediately
'    clsConfig.WriteValue "system", "driver", "new driver", True
'
'    '4.1 write a value in the config file but do not save immediately
'    clsConfig.WriteValue "system", "login", "new login"
'
'    '4.2 write a value in the config file but do not save immediately
'    clsConfig.WriteValue "database", "server", "new server"
'
'    '5. explicitly save configuration file
'    clsConfig.SaveConfigurationFile
'*************************************************************************************************

'---------------
'Error messages
'---------------
Private Const ERR_3010 = "Error occurred while opening configuration file."
Private Const ERR_3020 = "Error occurred while reading a value."
Private Const ERR_3030 = "Error occurred while writing a value."
Private Const ERR_3040 = "Error occurred while saving configuration file."

Private gsXMLFile                           As String                   'xml file location
Private gstrKey                             As String                   'Encryption key
Private blnEncrypt                          As Boolean                  'Determines if config file should be encrypted or decrypted - default is false i.e. no encryption
Private docXML                              As DOMDocument              'XML document object where configuration is loaded

Option Explicit

'*************************************************************************************************
' Property name :       Encryption
' Abstract :            This will get if encryption is used for read/write.
'*************************************************************************************************
Public Property Get Encryption() As Boolean
    Encryption = blnEncrypt
End Property

'*************************************************************************************************
' Property name :       Encryption
' Abstract :            This will set if encryption should be used for read/write.
'*************************************************************************************************
Public Property Let Encryption(ByVal blnEncryption As Boolean)
    blnEncrypt = blnEncryption
End Property

'*************************************************************************************************
' Procedure name :      LoadConfigurationFile
' Abstract :            This method will load a configuration file.
' Input Parameters :    File name - file to be opened
' Output Parameters :   Returns true if file is created
'
' Example :
'
'    'OPTION 1: open configuration file. creates if it does not exist
'    clsConfig.LoadConfigurationFile App.Path & "\ej2.xml"
'
'    'OPTION 2: open configuration file. and check if it is created or not
'    if (clsConfig.LoadConfigurationFile App.Path & "\ej2.xml") then
'       'already exists
'    else
'       'not created
'    end if
'
'    'OPTION 3: open configuration file. but do not create if it does not exist
'    if (clsConfig.LoadConfigurationFile App.Path & "\ej2.xml", false) then
'       'already exists
'    else
'       'not created
'    end if
'
'*************************************************************************************************
Public Function LoadConfigurationFile(strFilename As String, Optional blnCreate As Boolean = True) As Boolean
    
    LoadConfigurationFile = False       'config file not created
    
    gsXMLFile = strFilename
    
    If XMLReadConfiguration(blnCreate) Then
        LoadConfigurationFile = True    'config file created
    End If
    
End Function

'*************************************************************************************************
' Procedure name :      ReadValue
' Abstract :            This method will read value from the config file
' Input Parameters :    Section - section name
'                       Name - name of the item/subsection
' Output Parameters :   Value found in the config file. Return empty string if not found
'*************************************************************************************************
Public Function ReadValue(strSection As String, strName As String) As String
    Dim node                As IXMLDOMNode
    Dim nodeValue           As IXMLDOMNode
    Dim strSearch           As String
    
    On Error GoTo ErrHandler
    
    ReadValue = ""
    
    If docXML Is Nothing Then
        Exit Function
    End If
    
    If docXML.documentElement Is Nothing Then
        Exit Function
    End If
    
    strSearch = "//configuration/section[@name=""" & strSection & """]/item[@name=""" & strName & """]"
    Set node = docXML.documentElement.selectSingleNode(strSearch)
    
    If Not node Is Nothing Then
        Set nodeValue = node.Attributes.getNamedItem("value")
        If Not nodeValue Is Nothing Then
            ReadValue = nodeValue.Text
        End If
    End If

    Exit Function

ErrHandler:
    Err.Raise 3020, "XMLReadConfiguration", ConstructMessage(ERR_3020)
    Exit Function
    
End Function

'*************************************************************************************************
' Procedure name :      WriteValue
' Abstract :            This method will write to the config file
' Input Parameters :    Section - section to be written to
'                       Name - item/subsection name to be written to
'                       Value - value to be written
'                       optional Save flag - whether to save immediately or save will be called
'                       explicitly
' Output Parameters :   Return true if writing is successful else false
'
' Example :
'
'    'OPTION 1. write a value in the config file but do not save immediately
'    clsConfig.WriteValue "system", "login", "new login"
'
'    'OPTION 2. write a value in the config file and save immediately
'    clsConfig.WriteValue "system", "driver", "new driver", True
'
' Hint: In order to improve performance, delay saving until all the changes are made.
'*************************************************************************************************
Public Function WriteValue(strSection As String, _
                           strName As String, _
                           strValue As String, _
                           Optional blnSave As Boolean = False) As Boolean
    Dim nodeSection         As IXMLDOMNode
    Dim nodeItem            As IXMLDOMNode
    Dim nodeValue           As IXMLDOMNode
    Dim nodeName            As IXMLDOMNode
    Dim strSearch           As String
    Dim blnSectionCreated   As Boolean
    
    On Error GoTo ErrHandler
    
    WriteValue = False
    
    blnSectionCreated = False
    
    If docXML Is Nothing Then
        Exit Function
    End If
    
    If docXML.documentElement Is Nothing Then
        Exit Function
    End If
    
    'search for SECTION node
    strSearch = "//configuration/section[@name=""" & strSection & """]"
    Set nodeSection = docXML.documentElement.selectSingleNode(strSearch)
    If nodeSection Is Nothing Then
        'formatting
        docXML.documentElement.appendChild docXML.createTextNode(vbCrLf & vbTab)
        
        'create SECTION node
        Set nodeSection = docXML.createNode(NODE_ELEMENT, "section", "")
        Set nodeName = docXML.createNode(NODE_ATTRIBUTE, "name", "")
        nodeName.Text = strSection
        nodeSection.Attributes.setNamedItem nodeName
        docXML.documentElement.appendChild nodeSection
        
        'formatting
        nodeSection.appendChild docXML.createTextNode(vbCrLf)
        
        'formatting
        docXML.documentElement.appendChild docXML.createTextNode(vbCrLf)
        
        blnSectionCreated = True
    End If
    
    'search for ITEM node
    strSearch = "item[@name=""" & strName & """]"
    Set nodeItem = nodeSection.selectSingleNode(strSearch)
    If nodeItem Is Nothing Then
        'formatting
        If blnSectionCreated Then
            nodeSection.appendChild docXML.createTextNode(vbTab & vbTab)
        ElseIf nodeSection.selectNodes("item").length > 0 Then
            nodeSection.appendChild docXML.createTextNode(vbTab)
        Else
            nodeSection.appendChild docXML.createTextNode(vbCrLf & vbTab & vbTab)
        End If
        
        'create ITEM node
        Set nodeItem = docXML.createNode(NODE_ELEMENT, "item", "")
        Set nodeName = docXML.createNode(NODE_ATTRIBUTE, "name", "")
        nodeName.Text = strName
        nodeItem.Attributes.setNamedItem nodeName
        nodeSection.appendChild nodeItem
        
        'formatting
        nodeSection.appendChild docXML.createTextNode(vbCrLf & vbTab)
    
    End If
    
    'search for VALUE attribute
    Set nodeValue = nodeItem.Attributes.getNamedItem("value")
    If nodeValue Is Nothing Then
        Set nodeValue = docXML.createNode(NODE_ATTRIBUTE, "value", "")
        nodeItem.Attributes.setNamedItem nodeValue
    End If
    
    nodeValue.Text = strValue

    If blnSave Then
        SaveConfigurationFile
    End If

    WriteValue = True
    
    Exit Function

ErrHandler:
    Err.Raise 3030, "XMLReadConfiguration", ConstructMessage(ERR_3030)
    Exit Function
    
End Function

'*************************************************************************************************
' Procedure name :      SaveConfigurationFile
' Abstract :            This method will save changes to the configuration file
' Input Parameters :    None
' Output Parameters :   None
'*************************************************************************************************
Public Sub SaveConfigurationFile()
    
    On Error GoTo ErrHandler
    
    If docXML Is Nothing Then
        Exit Sub
    End If
    
    If docXML.documentElement Is Nothing Then
        Exit Sub
    End If

    SaveEncryptedXMLFile gsXMLFile, docXML.xml
    
    Exit Sub
    
ErrHandler:
    Err.Raise 3040, "XMLReadConfiguration", ConstructMessage(ERR_3040)
    Exit Sub
    
End Sub



'*************************************************************************************************
' Procedure name :      XMLReadConfiguration
' Abstract :            This method will read config file. If not found, it can create.
'                       Set 'docXML' global variable which hold reference to the DOM
' Input Parameters :    blnCreate - whether to create file if does not exist
' Output Parameters :   returns true if config is created
'*************************************************************************************************
Private Function XMLReadConfiguration(blnCreate As Boolean) As Boolean
    Dim bLoaded                     As Boolean
    Dim bLoading                    As Boolean
    Dim iFileNumber                 As Integer
    Dim nodeRoot                    As IXMLDOMNode
    Dim nodeName                    As IXMLDOMNode
    Dim piVersion                   As IXMLDOMProcessingInstruction

    On Error GoTo ErrHandler
    
    XMLReadConfiguration = False        'config file is not created
    
    If Len(Trim(Dir(gsXMLFile))) = 0 Then
        If blnCreate Then
            Set docXML = New DOMDocument
            Set nodeRoot = docXML.createNode(NODE_ELEMENT, "configuration", "")
            docXML.appendChild nodeRoot
            Set piVersion = docXML.createProcessingInstruction("xml", "version=""1.0""")
            docXML.insertBefore piVersion, docXML.childNodes(0)
            SaveConfigurationFile
            
            XMLReadConfiguration = True
        Else
            Exit Function
        End If
    End If
    
    Set docXML = New DOMDocument
    
    docXML.async = True
    bLoaded = docXML.loadXML(OpenEncryptedXMLFile(gsXMLFile))
    bLoading = True
    Do While Not docXML.parsed
        DoEvents
    Loop
    bLoading = False
    
    If Not bLoaded Or docXML.parseError <> 0 Then
        On Error GoTo 0
        Err.Raise 3010, "XMLReadConfiguration", ConstructMessage(ERR_3010)
        Exit Function
    End If
    
    Exit Function

ErrHandler:
    Err.Raise 3010, "XMLReadConfiguration", ConstructMessage(ERR_3010)
    Exit Function
    
End Function

'*************************************************************************************************
' Procedure name :      Encrypt
' Abstract :            This method will encrypt input data
' Input Parameters :    Data - to be encrypted
' Output Parameters :   Data - encrypted
'*************************************************************************************************
Private Function Encrypt(strData As String) As String
    Dim iKeyIndex               As Integer
    Dim iKeyChar                As Integer
    Dim iDataChar               As Integer
    Dim iCryptChar              As Integer
    Dim strEncrypted            As String
    Dim i                       As Integer

    If Not blnEncrypt Then
        Encrypt = strData
        Exit Function
    End If
    
    For i = 1 To Len(strData)
        iKeyIndex = (i - 1) Mod Len(gstrKey) + 1
        iKeyChar = Asc(Mid(gstrKey, iKeyIndex, 1))
        iDataChar = Asc(Mid(strData, i, 1))
        iCryptChar = iKeyChar Xor iDataChar
        strEncrypted = strEncrypted & Chr(iCryptChar)
    Next

    Encrypt = strEncrypted
    
End Function

'*************************************************************************************************
' Procedure name :      Decrypt
' Abstract :            This method will decrypt input data
' Input Parameters :    Data - to be decrypted
' Output Parameters :   Data - decrypted
'*************************************************************************************************
Private Function Decrypt(strData As String) As String
    Dim iKeyIndex               As Integer
    Dim iKeyChar                As Integer
    Dim iDataChar               As Integer
    Dim iDecryptChar            As Integer
    Dim strDecrypted            As String
    Dim i                       As Integer

    If Not blnEncrypt Then
        Decrypt = strData
        Exit Function
    End If
    
    For i = 1 To Len(strData)
        iKeyIndex = (i - 1) Mod Len(gstrKey) + 1
        iKeyChar = Asc(Mid(gstrKey, iKeyIndex, 1))
        iDataChar = Asc(Mid(strData, i, 1))
        iDecryptChar = iKeyChar Xor iDataChar
        strDecrypted = strDecrypted & Chr(iDecryptChar)
    Next

    Decrypt = strDecrypted
    
End Function

'*************************************************************************************************
' Procedure name :      GenerateKey
' Abstract :            This method will generate key to be used for encryption/decryption
' Input Parameters :    None
' Output Parameters :   Returns key
'*************************************************************************************************
Private Function GenerateKey() As String
    GenerateKey = "JN128SJHDJ23JHBCSNE23328JJJASJPUOIOQWUJXZNHJ273U54JBCJYH37YYGARQOOEPMASCXNHASGUYSJCXASJH23JHJHSDQWPIOEUDFKJSGM237846JSBAJHDG88R"
End Function

'*************************************************************************************************
' Procedure name :      ConstructMessage
' Abstract :            This procedure replace % with passed paramters.The Email message and the
'                       error message is constructed here.
'                       Err.number & err.description will be added to the message
'                       XML.parseError.errorCode & XML.parseError.reason will be added to the message
' Input Parameters :    Message string      - Message to be displayed/written
'                       4 optional parameters to be substituted
' Output Parameters :   Message
'*************************************************************************************************
Private Function ConstructMessage(ByRef sMessage As String, _
                         Optional sParam1 As String, _
                         Optional sParam2 As String, _
                         Optional sParam3 As String, _
                         Optional sParam4 As String) As String
    Dim iStart              As Integer
    Dim iEnd                As Integer
    Dim iParamNumber        As Integer
    Dim sNewMessage         As String
    Dim sCurrentParam       As String
    
    Dim lErrNumber              As Long
    Dim sErrDesc                As String
    Dim lXMLErrNumber           As Long
    Dim sXMLErrDesc             As String
    
    With Err
        lErrNumber = .Number
        sErrDesc = .Description
    End With
    
    With docXML
        lXMLErrNumber = .parseError.errorCode
        sXMLErrDesc = .parseError.reason
    End With
    
    iStart = 1
    iParamNumber = 1
    
    Do While (InStr(iStart, sMessage, "%", vbTextCompare) > 0) And (iStart <= Len(sMessage))
        iEnd = InStr(iStart, sMessage, "%", vbTextCompare)
        sNewMessage = sNewMessage & Mid$(sMessage, iStart, iEnd - iStart)
        Select Case iParamNumber
        Case 1
            If IsMissing(sParam1) Then
                Exit Do
            End If
            sCurrentParam = sParam1
        Case 2
            If IsMissing(sParam2) Then
                Exit Do
            End If
            sCurrentParam = sParam2
        Case 3
            If IsMissing(sParam3) Then
                Exit Do
            End If
            sCurrentParam = sParam3
        Case 4
            If IsMissing(sParam4) Then
                Exit Do
            End If
            sCurrentParam = sParam4
        Case Else
            Exit Do
        End Select
        sNewMessage = sNewMessage & sCurrentParam
        iStart = iEnd + 1
        iParamNumber = iParamNumber + 1
    Loop

    If iStart = 1 Then
        sNewMessage = sMessage
    Else
        sNewMessage = sNewMessage & Right(sMessage, Len(sMessage) - iStart + 1)
    End If
        
    If lErrNumber <> 0 Then
        sNewMessage = sNewMessage & vbCrLf & _
                        lErrNumber & vbCrLf & _
                        sErrDesc
    End If
    
    If lXMLErrNumber <> 0 Then
        sNewMessage = sNewMessage & vbCrLf & _
                        lXMLErrNumber & vbCrLf & _
                        sXMLErrDesc
    End If
    
    ConstructMessage = sNewMessage
    
End Function

'*************************************************************************************************
' Procedure name :      Class_Initialize
' Abstract :            This procedure will initialize class variables. In this case,
'                       Key is generated for encryption and decryption
' Input Parameters :    None
' Output Parameters :   None
'*************************************************************************************************
Private Sub Class_Initialize()
    blnEncrypt = False
    gstrKey = GenerateKey
End Sub

'*************************************************************************************************
' Procedure name :      SaveEncryptedXMLFile
' Abstract :            This procedure
' Input Parameters :    None
' Output Parameters :   None
'*************************************************************************************************
Private Sub SaveEncryptedXMLFile(strFilename As String, strData As String)
    Dim strEncrypt                  As String
    Dim intFilenum                  As Integer
    Dim strDegbugInfo               As String
    
    strEncrypt = Encrypt(strData)
    
    'File is created
    intFilenum = FreeFile
    
    If Len(Trim(Dir(strFilename))) Then
        Kill strFilename
    End If
    
    Open strFilename For Binary Access Write As #intFilenum
        
    Put #intFilenum, , strEncrypt
    
    Close #intFilenum
    
End Sub

'*************************************************************************************************
' Procedure name :      OpenEncryptedXMLFile
' Abstract :            This procedure
' Input Parameters :    None
' Output Parameters :   None
'*************************************************************************************************
Private Function OpenEncryptedXMLFile(strFilename As String) As String
    Dim strData                     As String
    Dim strLine                     As String
    Dim intFilenum                  As Integer

    'open file
    intFilenum = FreeFile
    
    Open strFilename For Binary Access Read As #intFilenum
    
    strData = Input(LOF(intFilenum), intFilenum)
    
    Close intFilenum
    
    OpenEncryptedXMLFile = Decrypt(strData)

End Function

'*************************************************************************************************
'* File name   : XMLConfiguration.cls
'*
'* Copyright (c) 2002 by House of Code, Inc. All Rights Reserved.
'*
'* This software is the proprietary information of House of Code, Inc.
'* Use is subject to license terms.
'*
'*
'* @author  Prakash Sachania (prakash.sachania@in.houseofcode.com)
'* @version 1.00
'* @date    January, 2003
'*
'* Known Bugs and/or side effects:Nil
'*************************************************************************************************


