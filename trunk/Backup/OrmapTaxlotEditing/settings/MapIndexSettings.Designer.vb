﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:2.0.50727.4918
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On



<Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
 Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "8.0.0.0")>  _
Partial Friend NotInheritable Class MapIndexSettings
    Inherits Global.System.Configuration.ApplicationSettingsBase
    
    Private Shared defaultInstance As MapIndexSettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MapIndexSettings),MapIndexSettings)
    
    Public Shared ReadOnly Property [Default]() As MapIndexSettings
        Get
            Return defaultInstance
        End Get
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("County")>  _
    Public Property CountyField() As String
        Get
            Return CType(Me("CountyField"),String)
        End Get
        Set
            Me("CountyField") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("MapNumber")>  _
    Public Property MapNumberField() As String
        Get
            Return CType(Me("MapNumberField"),String)
        End Get
        Set
            Me("MapNumberField") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("ORMapNum")>  _
    Public Property OrmapMapNumberField() As String
        Get
            Return CType(Me("OrmapMapNumberField"),String)
        End Get
        Set
            Me("OrmapMapNumberField") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("MapScale")>  _
    Public Property MapScaleField() As String
        Get
            Return CType(Me("MapScaleField"),String)
        End Get
        Set
            Me("MapScaleField") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("MapSuffixType")>  _
    Public Property MapSuffixTypeField() As String
        Get
            Return CType(Me("MapSuffixTypeField"),String)
        End Get
        Set
            Me("MapSuffixTypeField") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("MapSuffixNum")>  _
    Public Property MapSuffixNumberField() As String
        Get
            Return CType(Me("MapSuffixNumberField"),String)
        End Get
        Set
            Me("MapSuffixNumberField") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("PageNumber")>  _
    Public Property PageNumberField() As String
        Get
            Return CType(Me("PageNumberField"),String)
        End Get
        Set
            Me("PageNumberField") = value
        End Set
    End Property
    
    <Global.System.Configuration.UserScopedSettingAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Configuration.DefaultSettingValueAttribute("ReliaCode")>  _
    Public Property ReliabilityCodeField() As String
        Get
            Return CType(Me("ReliabilityCodeField"),String)
        End Get
        Set
            Me("ReliabilityCodeField") = value
        End Set
    End Property
End Class
