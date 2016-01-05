﻿'------------------------------------------------------------------------------
' <autogenerated>
'     This code was generated by a tool.
'     Runtime Version: 1.1.4322.2032
'
'     Changes to this file may cause incorrect behavior and will be lost if 
'     the code is regenerated.
' </autogenerated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.Data
Imports System.Runtime.Serialization
Imports System.Xml


<Serializable(),  _
 System.ComponentModel.DesignerCategoryAttribute("code"),  _
 System.Diagnostics.DebuggerStepThrough(),  _
 System.ComponentModel.ToolboxItem(true)>  _
Public Class Dataset1
    Inherits DataSet
    
    Private tablesample As sampleDataTable
    
    Public Sub New()
        MyBase.New
        Me.InitClass
        Dim schemaChangedHandler As System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
        AddHandler Me.Tables.CollectionChanged, schemaChangedHandler
        AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
    End Sub
    
    Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.New
        Dim strSchema As String = CType(info.GetValue("XmlSchema", GetType(System.String)),String)
        If (Not (strSchema) Is Nothing) Then
            Dim ds As DataSet = New DataSet
            ds.ReadXmlSchema(New XmlTextReader(New System.IO.StringReader(strSchema)))
            If (Not (ds.Tables("sample")) Is Nothing) Then
                Me.Tables.Add(New sampleDataTable(ds.Tables("sample")))
            End If
            Me.DataSetName = ds.DataSetName
            Me.Prefix = ds.Prefix
            Me.Namespace = ds.Namespace
            Me.Locale = ds.Locale
            Me.CaseSensitive = ds.CaseSensitive
            Me.EnforceConstraints = ds.EnforceConstraints
            Me.Merge(ds, false, System.Data.MissingSchemaAction.Add)
            Me.InitVars
        Else
            Me.InitClass
        End If
        Me.GetSerializationData(info, context)
        Dim schemaChangedHandler As System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
        AddHandler Me.Tables.CollectionChanged, schemaChangedHandler
        AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
    End Sub
    
    <System.ComponentModel.Browsable(false),  _
     System.ComponentModel.DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Content)>  _
    Public ReadOnly Property sample As sampleDataTable
        Get
            Return Me.tablesample
        End Get
    End Property
    
    Public Overrides Function Clone() As DataSet
        Dim cln As Dataset1 = CType(MyBase.Clone,Dataset1)
        cln.InitVars
        Return cln
    End Function
    
    Protected Overrides Function ShouldSerializeTables() As Boolean
        Return false
    End Function
    
    Protected Overrides Function ShouldSerializeRelations() As Boolean
        Return false
    End Function
    
    Protected Overrides Sub ReadXmlSerializable(ByVal reader As XmlReader)
        Me.Reset
        Dim ds As DataSet = New DataSet
        ds.ReadXml(reader)
        If (Not (ds.Tables("sample")) Is Nothing) Then
            Me.Tables.Add(New sampleDataTable(ds.Tables("sample")))
        End If
        Me.DataSetName = ds.DataSetName
        Me.Prefix = ds.Prefix
        Me.Namespace = ds.Namespace
        Me.Locale = ds.Locale
        Me.CaseSensitive = ds.CaseSensitive
        Me.EnforceConstraints = ds.EnforceConstraints
        Me.Merge(ds, false, System.Data.MissingSchemaAction.Add)
        Me.InitVars
    End Sub
    
    Protected Overrides Function GetSchemaSerializable() As System.Xml.Schema.XmlSchema
        Dim stream As System.IO.MemoryStream = New System.IO.MemoryStream
        Me.WriteXmlSchema(New XmlTextWriter(stream, Nothing))
        stream.Position = 0
        Return System.Xml.Schema.XmlSchema.Read(New XmlTextReader(stream), Nothing)
    End Function
    
    Friend Sub InitVars()
        Me.tablesample = CType(Me.Tables("sample"),sampleDataTable)
        If (Not (Me.tablesample) Is Nothing) Then
            Me.tablesample.InitVars
        End If
    End Sub
    
    Private Sub InitClass()
        Me.DataSetName = "Dataset1"
        Me.Prefix = ""
        Me.Namespace = "http://tempuri.org/Dataset1.xsd"
        Me.Locale = New System.Globalization.CultureInfo("en-US")
        Me.CaseSensitive = false
        Me.EnforceConstraints = true
        Me.tablesample = New sampleDataTable
        Me.Tables.Add(Me.tablesample)
    End Sub
    
    Private Function ShouldSerializesample() As Boolean
        Return false
    End Function
    
    Private Sub SchemaChanged(ByVal sender As Object, ByVal e As System.ComponentModel.CollectionChangeEventArgs)
        If (e.Action = System.ComponentModel.CollectionChangeAction.Remove) Then
            Me.InitVars
        End If
    End Sub
    
    Public Delegate Sub sampleRowChangeEventHandler(ByVal sender As Object, ByVal e As sampleRowChangeEvent)
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class sampleDataTable
        Inherits DataTable
        Implements System.Collections.IEnumerable
        
        Private columnNo As DataColumn
        
        Private columnName_ As DataColumn
        
        Friend Sub New()
            MyBase.New("sample")
            Me.InitClass
        End Sub
        
        Friend Sub New(ByVal table As DataTable)
            MyBase.New(table.TableName)
            If (table.CaseSensitive <> table.DataSet.CaseSensitive) Then
                Me.CaseSensitive = table.CaseSensitive
            End If
            If (table.Locale.ToString <> table.DataSet.Locale.ToString) Then
                Me.Locale = table.Locale
            End If
            If (table.Namespace <> table.DataSet.Namespace) Then
                Me.Namespace = table.Namespace
            End If
            Me.Prefix = table.Prefix
            Me.MinimumCapacity = table.MinimumCapacity
            Me.DisplayExpression = table.DisplayExpression
        End Sub
        
        <System.ComponentModel.Browsable(false)>  _
        Public ReadOnly Property Count As Integer
            Get
                Return Me.Rows.Count
            End Get
        End Property
        
        Friend ReadOnly Property NoColumn As DataColumn
            Get
                Return Me.columnNo
            End Get
        End Property
        
        Friend ReadOnly Property Name_Column As DataColumn
            Get
                Return Me.columnName_
            End Get
        End Property
        
        Public Default ReadOnly Property Item(ByVal index As Integer) As sampleRow
            Get
                Return CType(Me.Rows(index),sampleRow)
            End Get
        End Property
        
        Public Event sampleRowChanged As sampleRowChangeEventHandler
        
        Public Event sampleRowChanging As sampleRowChangeEventHandler
        
        Public Event sampleRowDeleted As sampleRowChangeEventHandler
        
        Public Event sampleRowDeleting As sampleRowChangeEventHandler
        
        Public Overloads Sub AddsampleRow(ByVal row As sampleRow)
            Me.Rows.Add(row)
        End Sub
        
        Public Overloads Function AddsampleRow(ByVal No As String, ByVal Name_ As String) As sampleRow
            Dim rowsampleRow As sampleRow = CType(Me.NewRow,sampleRow)
            rowsampleRow.ItemArray = New Object() {No, Name_}
            Me.Rows.Add(rowsampleRow)
            Return rowsampleRow
        End Function
        
        Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return Me.Rows.GetEnumerator
        End Function
        
        Public Overrides Function Clone() As DataTable
            Dim cln As sampleDataTable = CType(MyBase.Clone,sampleDataTable)
            cln.InitVars
            Return cln
        End Function
        
        Protected Overrides Function CreateInstance() As DataTable
            Return New sampleDataTable
        End Function
        
        Friend Sub InitVars()
            Me.columnNo = Me.Columns("No")
            Me.columnName_ = Me.Columns("Name ")
        End Sub
        
        Private Sub InitClass()
            Me.columnNo = New DataColumn("No", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnNo)
            Me.columnName_ = New DataColumn("Name ", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnName_)
        End Sub
        
        Public Function NewsampleRow() As sampleRow
            Return CType(Me.NewRow,sampleRow)
        End Function
        
        Protected Overrides Function NewRowFromBuilder(ByVal builder As DataRowBuilder) As DataRow
            Return New sampleRow(builder)
        End Function
        
        Protected Overrides Function GetRowType() As System.Type
            Return GetType(sampleRow)
        End Function
        
        Protected Overrides Sub OnRowChanged(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowChanged(e)
            If (Not (Me.sampleRowChangedEvent) Is Nothing) Then
                RaiseEvent sampleRowChanged(Me, New sampleRowChangeEvent(CType(e.Row,sampleRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowChanging(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowChanging(e)
            If (Not (Me.sampleRowChangingEvent) Is Nothing) Then
                RaiseEvent sampleRowChanging(Me, New sampleRowChangeEvent(CType(e.Row,sampleRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowDeleted(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowDeleted(e)
            If (Not (Me.sampleRowDeletedEvent) Is Nothing) Then
                RaiseEvent sampleRowDeleted(Me, New sampleRowChangeEvent(CType(e.Row,sampleRow), e.Action))
            End If
        End Sub
        
        Protected Overrides Sub OnRowDeleting(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowDeleting(e)
            If (Not (Me.sampleRowDeletingEvent) Is Nothing) Then
                RaiseEvent sampleRowDeleting(Me, New sampleRowChangeEvent(CType(e.Row,sampleRow), e.Action))
            End If
        End Sub
        
        Public Sub RemovesampleRow(ByVal row As sampleRow)
            Me.Rows.Remove(row)
        End Sub
    End Class
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class sampleRow
        Inherits DataRow
        
        Private tablesample As sampleDataTable
        
        Friend Sub New(ByVal rb As DataRowBuilder)
            MyBase.New(rb)
            Me.tablesample = CType(Me.Table,sampleDataTable)
        End Sub
        
        Public Property No As String
            Get
                Try 
                    Return CType(Me(Me.tablesample.NoColumn),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tablesample.NoColumn) = value
            End Set
        End Property
        
        Public Property Name_ As String
            Get
                Try 
                    Return CType(Me(Me.tablesample.Name_Column),String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set
                Me(Me.tablesample.Name_Column) = value
            End Set
        End Property
        
        Public Function IsNoNull() As Boolean
            Return Me.IsNull(Me.tablesample.NoColumn)
        End Function
        
        Public Sub SetNoNull()
            Me(Me.tablesample.NoColumn) = System.Convert.DBNull
        End Sub
        
        Public Function IsName_Null() As Boolean
            Return Me.IsNull(Me.tablesample.Name_Column)
        End Function
        
        Public Sub SetName_Null()
            Me(Me.tablesample.Name_Column) = System.Convert.DBNull
        End Sub
    End Class
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class sampleRowChangeEvent
        Inherits EventArgs
        
        Private eventRow As sampleRow
        
        Private eventAction As DataRowAction
        
        Public Sub New(ByVal row As sampleRow, ByVal action As DataRowAction)
            MyBase.New
            Me.eventRow = row
            Me.eventAction = action
        End Sub
        
        Public ReadOnly Property Row As sampleRow
            Get
                Return Me.eventRow
            End Get
        End Property
        
        Public ReadOnly Property Action As DataRowAction
            Get
                Return Me.eventAction
            End Get
        End Property
    End Class
End Class
