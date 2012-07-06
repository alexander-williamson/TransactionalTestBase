Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.EnterpriseServices
Imports System.Transactions

''' <summary>
''' A base class that automatically puts each test case in a Serviced Transaction and performs Rollback after each TestMethod.
''' Use with Microsoft.VisualStudio.TestTools.UnitTesting test classes. Simply inherit to gain the transactions.
''' Adapted from http://stackoverflow.com/questions/4927814/mstest-transactionscope-not-rolling-back
''' </summary>
''' <remarks>Last updated 15th September 2011 by Alexander Williamson</remarks>
<TestClass()> _
<System.Diagnostics.DebuggerStepThrough()> _
Public Class TransactionalTestBase
    Inherits ServicedComponent

    Private Shared ReadOnly RandomGenerator As New Random()
    Private Const RandomChars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-"
    ''' <summary>
    ''' Return a random String. Chars will include "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-" and include space " " if Include Space is true.
    ''' </summary>
    ''' <param name="Length">Length of desired Random String</param>
    ''' <param name="IncludeSpace">Include the " " (space) char</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function RandomString(ByVal Length As Integer, Optional ByVal IncludeSpace As Boolean = False) As String
        Dim CharRange As String = RandomChars
        If IncludeSpace Then CharRange = CharRange & " "
        Dim CharRangeLength As Integer = CharRange.Length
        Dim ReturnString As New System.Text.StringBuilder
        For i As Integer = 1 To Length
            ReturnString.Append(RandomChars(RandomGenerator.Next(CharRangeLength)))
        Next
        Return ReturnString.ToString
    End Function

    Public Shared Function GenerateString(ByVal Length As Integer) As String
        Dim CharRange As String = "ABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJABCDEFGHIJ" '260
        If Length > CharRange.Length Then Throw New ArgumentOutOfRangeException("Length", "Must be between 1 and " & CharRange.Length)
        Return CharRange.Substring(0, Length)
    End Function

    Public Shared Sub DumpBrokenRules(ByVal Rules As Csla.Validation.BrokenRulesCollection)
        For Each element As Csla.Validation.BrokenRule In Rules
            With element
                System.Diagnostics.Trace.WriteLine("Property: " & .Property & " Description: " & .Description)
                System.Diagnostics.Trace.WriteLine("Rule Name: " & .RuleName)
                System.Diagnostics.Trace.WriteLine(String.Empty)
            End With
        Next
    End Sub

#Region "Properties and Methods"

    ''' <summary>
    ''' The Current TransactionScope Object
    ''' </summary>
    ''' <remarks></remarks>
    Private CurrentTransactionScope As TransactionScope = Nothing

    Private CanSet_UseTransactions As Boolean = False
    ''' <summary>
    ''' True if the class should use Transactions. Throws ArgumentException if not set in PreTransactionBegin.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UseTransactions() As Boolean
        Get
            Return mUseTransactions
        End Get
        Set(ByVal value As Boolean)
            If Not CanSet_UseTransactions Then Throw New ArgumentException("Cannot set UseTransactions when a Transaction has already begun")
            mUseTransactions = value
        End Set
    End Property

    ''' <summary>
    ''' True if a Transaction should wrap the TestMethod
    ''' </summary>
    ''' <remarks></remarks>
    Private mUseTransactions As Boolean = True

#End Region

#Region "Transaction Begin and Rollback"

    ''' <summary>
    ''' Begins a COM+ 1.5 transaction for the test (marked as TestInitialize). Calls Overridable PreTransactionBegin before starting the transaction. Calls Overridable PostTransactionBegin after.
    ''' </summary>
    ''' <remarks></remarks>
    <TestInitialize()> _
    Public Sub TransactionBegin()
        PreTransactionBegin()
        If UseTransactions Then
            CurrentTransactionScope = New TransactionScope(TransactionScopeOption.RequiresNew)
        End If
        PostTransactionBegin()
    End Sub

    ''' <summary>
    ''' Rolls back the COM+ 1.5 transaction (marked as TestCleanup). Calls Overridable PreRollback before starting the transaction. Calls Overridable PostRollback after.
    ''' </summary>
    ''' <remarks></remarks>
    <TestCleanup()> _
    Public Sub TransactionRollback()
        PreRollback()
        If UseTransactions Then
            Transaction.Current.Rollback()
            If CurrentTransactionScope IsNot Nothing Then CurrentTransactionScope.Dispose()
            CurrentTransactionScope = Nothing
        End If
        PostRollback()
    End Sub

#End Region

#Region "Overridable Helper Subs"

    ''' <summary>
    ''' Called before the Transaction is started
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub PreTransactionBegin()
    End Sub

    ''' <summary>
    ''' Called after the Transaction has started
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub PostTransactionBegin()
    End Sub

    ''' <summary>
    ''' Called before the Transaction is rolled back
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub PreRollback()
    End Sub

    ''' <summary>
    ''' Called after the Transaction rollback
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub PostRollback()
    End Sub

    Public Sub New()
        'Public Constructor
    End Sub

#End Region

End Class