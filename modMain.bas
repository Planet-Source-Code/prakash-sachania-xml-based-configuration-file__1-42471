Attribute VB_Name = "modMain"
'*************************************************************************************************
'* File name   : modMain.bas
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
' Module/Form name :    Main module (modMain)
' Abstract :            This module has "main" method defined which is called when the program
'                       starts.
'*************************************************************************************************

'-----------------'
' Public constants'
'-----------------'

'--------------------------------------------------------'
'Other global variables
'--------------------------------------------------------'
Public gsDriverName                         As String                   'sql server driver

Public clsConfig                            As New XMLConfiguration     'XML Configuration class

Option Explicit


'*************************************************************************************************
' Procedure name :      Main
' Abstract :            This method makes connections to ejems2, fa2 and act.Then it calls
'                       transfer module to transfer customer's data from ejems2 and fa2 to ACT.
' '                     Then it calls the XferLog record to to write into the log file.
' Input Parameters : None
' Output Parameters : None
'*************************************************************************************************
Public Sub Main()
   
    On Error GoTo MainErr

    GlobalInitialize             'initializes global variables
    
    Exit Sub
    
MainErr:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
    End
    
End Sub

'*************************************************************************************************
' Procedure name :      GlobalInitialize
' Abstract :            This method gets driver from the XML file
' Input Parameters :    None
' Output Parameters :   None
'*************************************************************************************************
Private Sub GlobalInitialize()
    
    'Encryption to be used
    'clsConfig.Encryption = True
    
    'open configuration file. creates if it does not exist
    clsConfig.LoadConfigurationFile App.Path & "\ej2.xml"
    
    'read a value
    gsDriverName = clsConfig.ReadValue("system", "driver")
    
    'write a value in the config file and save immediately
    clsConfig.WriteValue "system", "driver", "new driver", True
    
    'write a value in the config file but do not save immediately
    clsConfig.WriteValue "system", "login", "new login"
    
    'write a value in the config file but do not save immediately
    clsConfig.WriteValue "system", "database", "new database"
    
    'explicitly save configuration file
    clsConfig.SaveConfigurationFile
    
End Sub

'*************************************************************************************************
'* File name   : modMain.bas
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

