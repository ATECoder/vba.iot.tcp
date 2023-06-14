VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Kernel32Api"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''' <summary> Initializes this class. </summary>
Private Sub Classs_Initialize()
End Sub

''' <summary> Executed before all tests. </summary>
Public Sub BeforeAll()
End Sub

''' <summary> Executed after all tests. </summary>
Public Sub AfterAll()
End Sub

''' <summary> Executed before each test. </summary>
Public Sub BeforeEach()
End Sub

''' <summary> Executed after each test. </summary>
Public Sub AfterEach()
End Sub

''' <summary> </summary>
''' <returns> An Assert holding the test result. </returns> 
Public Sub AssertPremiseShouldPass() As Assert
	Dim expected as Variant
	Dim actual As Variant
	Set AssertPremiseShouldPass = Assert.AreEqual( expected, actual, _
		"Actual " & Cstr(actual) & " should equals " & Cstr( expected) )
End Sub
