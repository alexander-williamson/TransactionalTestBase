TransactionalTestBase
=====================

A Visual Basic .NET (VB.NET) Unit Test base class which automatically rolls back any ADO.NET or LINQ transactions.

Uses ServicedComponent with System.Transactions and provides event hooks for ease of use.

Minimum-Code Example
====================

This example is the least amount of code required to setup a working Unit Test. There are more examples further down this README.

	Imports Microsoft.VisualStudio.TestTools.UnitTesting

	<TestClass()> _
	Public Class Example_Tests
		Inherits TransactionalTestBase
		
		' Example barebones UnitTest with automatic transactioning and roll-back
		' Author: Alexander Williamson
		' Website: http://www.alexw.co.uk
		
		#Region "Unit Tests"
		
			<TestMethod()> _
			Public Sub ExampleUnitTest
				Assert.IsTrue(True) ' This test will pass
			End Sub
		
		#End Region
		
	End Class

Usage
=====

Create a new Unit Test class. Include any (required) Microsoft Unit Testing references:

	Imports Microsoft.VisualStudio.TestTools.UnitTesting

Add the TransactionalTestBase file wherever your test are. Subclass child tests off this class.

	<TestClass()> _
	Public Class ExampleUnitTest
		Inherits TransactionalTestBase
		
	End Class

Each Sub that is decorated with the TestMethod() Attribute will automatically be wrapped in a transaction and rolled back when finished (or when an exception is encountered).

    <TestMethod()> _
    Public Sub ExampleUnitTest()
		Assert.Pass("Hello world!")
	End Sub

Hooks
=====

Set-up Sequence
--------------

Init Unit Test -> PreTransactionBegin -> (Transactions are created) -> PostTransactionBegin -> Unit Test Runs

Tear-down Sequence
------------------

Unit Test Ends -> PreRollback -> (Transactions rolled back) -> PostRollback -> End

Overridable functions
---------------------

You can optionally override the following functions in your Unit test

PreTransactionBegin
-------------------

Called before the transaction is created.

    Public Overridable Sub PreTransactionBegin()
	
PostTransactionBegin
--------------------

Called after the transaction is created, just before the Unit test is executed.

    Public Overridable Sub PostTransactionBegin()
	
PreRollback
-----------

Called after the unit test has run, failed or caused an exception; before the transaction is rolled back.

    Public Overridable Sub PreRollback()
	
PostRollback
------------

Called after the transaction has been rolled back:

    Public Overridable Sub PostRollback()

Recommended Example
==================

Normally only overriding PostTransactionBegin to add Set-Up code is required. All transactions will roll back when each Unit Test finishes.

	Imports Microsoft.VisualStudio.TestTools.UnitTesting

	<TestClass()> _
	Public Class Example_Tests
		Inherits TransactionalTestBase
		
		' Example barebones UnitTest with automatic transactioning and roll-back
		' Author: Alexander Williamson
		' Website: http://www.alexw.co.uk
		
		#Region "Set-Up and Tear-Down"
		
			Public Overrides Sub PostTransactionBegin()
				' Called after the TransactionalTestBase transactions are created.
				' We are in a transaction at this point.
				' This is an optional override
				' Put your Set-Up Code in here. This is normally the only section you will need.
			End Sub
		
		#End Region
		
		#Region "Unit Tests"
		
			<TestMethod()> _
			Public Sub ExampleUnitTest
				Assert.IsTrue(True) ' This test will pass
			End Sub
		
		#End Region
		
	End Class
	
Full Unit Test Example
======================

	Imports Microsoft.VisualStudio.TestTools.UnitTesting

	<TestClass()> _
	Public Class Example_Tests
		Inherits TransactionalTestBase
		
		' Example UnitTest with automatic transactioning and roll-back
		' Author: Alexander Williamson
		' Website: http://www.alexw.co.uk
		
		#Region "Set-Up and Tear-Down"
		
			Public Overrides Sub PreTransactionBegin()
				' Called before the TransactionalTestBase transactions are created.
				' We are not in a transaction at this point.
				' This is an optional override.
			End Sub
			
			Public Overrides Sub PostTransactionBegin()
				' Called after the TransactionalTestBase transactions are created.
				' We are in a transaction at this point.
				' This is an optional override
				' Put your Set-Up Code in here. This is normally the only section you will need.
			End Sub
			
			Public Overrides Sub PreRollback()
				' Called just before a transactional roll-back.
				' Put your Tear-Down code here (although you probably don't need anything in here). 
				' We are still in a transaction at this point.
				' This is an optional override.
			End Sub

			Public Overrides Sub PostRollback()
				' Called after the TransactionalTestBase transactions are created.
				' We are not in a transaction at this point.
				' This is an optional override.
			End Sub
		
		#End Region
		
		#Region "Unit Tests"
		
			<TestMethod()> _
			Public Sub ExampleUnitTest
				Assert.IsTrue(True) ' This test will pass
			End Sub
		
		#End Region
		
	End Class
