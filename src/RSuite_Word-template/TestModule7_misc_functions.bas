Attribute VB_Name = "TestModule7_misc_functions"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    ' Load public vars:
    Application.ScreenUpdating = False
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    'reset loaded public vars
    Application.ScreenUpdating = True
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    
End Sub

'@TestMethod("MiscFunctions")
Private Sub Test_versionCompare() 'TODO Rename test
    Dim results_gt As String, results_gt_version As String, results_same As String, _
        results_same_version As String, results_lt As String, results_lt_version As String, _
        results_emptyA As String, results_emptyB As String, results_empty As String, _
        results_zeroA As String, results_zeroB As String, results_zero As String, _
        results_nondoubleA As String, results_nondoubleB As String, results_nondouble As String
    On Error GoTo TestFail
    'Arrange:
    'Act:
        results_gt = VersionCheck.versionCompare("6", "5.0")
        results_gt_version = VersionCheck.versionCompare("6.1.1", "6.0.1")
        results_same = VersionCheck.versionCompare("6.1", "6.1")
        results_same_version = VersionCheck.versionCompare("5.0.3", "5.0.2")
        results_lt = VersionCheck.versionCompare("3", "4.0")
        results_lt_version = VersionCheck.versionCompare("2.0.3", "5.0.2")
        results_emptyA = VersionCheck.versionCompare("", "5.0.2")
        results_emptyB = VersionCheck.versionCompare("2.0.3", "")
        results_empty = VersionCheck.versionCompare("", "")
        results_zeroA = VersionCheck.versionCompare("0", "42.5")
        results_zeroB = VersionCheck.versionCompare("3", "0.0")
        results_zero = VersionCheck.versionCompare("0", "0")
        results_nondoubleA = VersionCheck.versionCompare("vab6", "0")
        results_nondoubleB = VersionCheck.versionCompare("0", "v6.1.3")
        results_nondouble = VersionCheck.versionCompare("v6", "v7")
        

    'Assert:
        Assert.Succeed
        Assert.areequal ">", results_gt
        Assert.areequal ">", results_gt_version
        Assert.areequal "same", results_same
        Assert.areequal "same", results_same_version
        Assert.areequal "<", results_lt
        Assert.areequal "<", results_lt_version
        Assert.areequal "unable to compare", results_emptyA
        Assert.areequal "unable to compare", results_emptyB
        Assert.areequal "unable to compare", results_empty
        Assert.areequal "<", results_zeroA
        Assert.areequal ">", results_zeroB
        Assert.areequal "same", results_zero
        Assert.areequal "unable to compare", results_nondoubleA
        Assert.areequal "unable to compare", results_nondoubleB
        Assert.areequal "unable to compare", results_nondouble
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
