﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:2.0.50727.832
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On



<Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
 Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "8.0.0.0")>  _
Partial Friend NotInheritable Class TableNamesSettings
    Inherits Global.System.Configuration.ApplicationSettingsBase
    
    Private Shared defaultInstance As TableNamesSettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New TableNamesSettings),TableNamesSettings)
    
    Public Shared ReadOnly Property [Default]() As TableNamesSettings
        Get
            Return defaultInstance
        End Get
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("MapIndex")>  _
    Public Property MapIndexFC() As String
        Get
            Return CType(Me("MapIndexFC"),String)
        End Get
        Set
            Me("MapIndexFC") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("TaxCode")>  _
    Public Property TaxCodeFC() As String
        Get
            Return CType(Me("TaxCodeFC"),String)
        End Get
        Set
            Me("TaxCodeFC") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("TaxLot")>  _
    Public Property TaxLotFC() As String
        Get
            Return CType(Me("TaxLotFC"),String)
        End Get
        Set
            Me("TaxLotFC") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("TaxLotLines")>  _
    Public Property TaxLotLinesFC() As String
        Get
            Return CType(Me("TaxLotLinesFC"),String)
        End Get
        Set
            Me("TaxLotLinesFC") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("Plats")>  _
    Public Property PlatsFC() As String
        Get
            Return CType(Me("PlatsFC"),String)
        End Get
        Set
            Me("PlatsFC") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("CartographicLines")>  _
    Public Property CartographicLinesFC() As String
        Get
            Return CType(Me("CartographicLinesFC"),String)
        End Get
        Set
            Me("CartographicLinesFC") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("ReferenceLines")>  _
    Public Property ReferenceLinesFC() As String
        Get
            Return CType(Me("ReferenceLinesFC"),String)
        End Get
        Set
            Me("ReferenceLinesFC") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("CancelledNumbers")>  _
    Public Property CancelledNumbersTable() As String
        Get
            Return CType(Me("CancelledNumbersTable"),String)
        End Get
        Set
            Me("CancelledNumbersTable") = value
        End Set
    End Property
End Class
