VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPActivityType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function createMultiple(pControllingArea As String, pTestRun As String, pData As Collection)
    On Error GoTo createMultiple_Error
    Dim lSAPFormat As New SAPFormat
    Dim boActivityType As Object
    Dim oReturn As Object
    Dim oActivityTypeList As Object
    Dim aControllingArea As String
    Dim aTestRun As String

    Set boActivityType = MySAPCon.SAPBapiControl.GetSAPObject("ActivityType")
    Set oReturn = MySAPCon.SAPBapiControl.DimAs(boActivityType, "CreateMultiple", "Return")
    Set oActivityTypeList = MySAPCon.SAPBapiControl.DimAs(boActivityType, "CreateMultiple", "ActivityTypeList")

    Dim aATLRow As Object
    Dim lRow As SAPActivityTypeList
    For Each lRow In pData
        Set aATLRow = oActivityTypeList.Rows.Add
        aATLRow.Value("ACTTYPE") = lRow.ACTTYPE
        aATLRow.Value("VALID_FROM") = lRow.VALID_FROM
        aATLRow.Value("VALID_TO") = lRow.VALID_TO
        aATLRow.Value("ACT_UNIT") = lRow.ACT_UNIT
        aATLRow.Value("COSTCENTER_CAT") = lRow.COSTCENTER_CAT
        aATLRow.Value("ACTIVITY_CAT") = lRow.ACTIVITY_CAT
        aATLRow.Value("ALLOC_COST_ELEM") = lRow.ALLOC_COST_ELEM
        aATLRow.Value("PRICE_INDICATOR") = lRow.PRICE_INDICATOR
        aATLRow.Value("AVERAGE_PRICE_IND") = lRow.AVERAGE_PRICE_IND
        aATLRow.Value("PREDIS_FXD_COST") = lRow.PREDIS_FXD_COST
        aATLRow.Value("MANUAL_ACT_QTY") = lRow.MANUAL_ACT_QTY
        aATLRow.Value("ACT_CAT_ACTUAL") = lRow.ACT_CAT_ACTUAL
        aATLRow.Value("ACT_PRICE_IND") = lRow.ACT_PRICE_IND
        aATLRow.Value("LOCK_INDICATOR") = lRow.LOCK_INDICATOR
        aATLRow.Value("OUTPUT_UNIT") = lRow.OUTPUT_UNIT
        aATLRow.Value("OUTPUT_FACTOR") = lRow.OUTPUT_FACTOR
        aATLRow.Value("NAME") = lRow.NAME
        aATLRow.Value("DESCRIPT") = lRow.DESCRIPT
    Next lRow
    Application.Cursor = xlWait
    If pTestRun = "X" Then
        boActivityType.createMultiple CoArea:=pControllingArea, _
        TestRun:=pTestRun, _
        ActivityTypeList:=oActivityTypeList, _
        Return:=oReturn
    Else
        boActivityType.createMultiple CoArea:=pControllingArea, _
        ActivityTypeList:=oActivityTypeList, _
        Return:=oReturn
    End If

    Application.Cursor = xlDefault

    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit

    If oReturn.RowCount = 0 Then
        createMultiple = "Success ???"
        aSAPBapiTranctionCommit.commit
    Else
        Dim retRow As Object
        Set retRow = oReturn.Rows(1)
        If retRow("TYPE") = "S" Then
            createMultiple = retRow("MESSAGE")
            aSAPBapiTranctionCommit.commit
        Else
            For Each retRow In oReturn.Rows
                createMultiple = createMultiple & ";" & retRow("MESSAGE")
            Next retRow
        End If
    End If
    Set boActivityType = Nothing
    Exit Function

createMultiple_Error:
    Application.Cursor = xlDefault
    createMultiple = "Error"
    MySAPErr.MSGProt "SAPActivityType", "createMultiple", "", Err.Number, Err.Description
    Exit Function
End Function

Function changeMultiple(pControllingArea As String, pLanguageKey As String, pTestRun As String, pData As Collection)
    On Error GoTo changeMultiple_Error
    Dim lSAPFormat As New SAPFormat
    Dim boActivityType As Object
    Dim oReturn As Object
    Dim oActivityTypeList As Object
    Dim aControllingArea As String
    Dim aTestRun As String
    Dim oLanguageKey As Object

    Set boActivityType = MySAPCon.SAPBapiControl.GetSAPObject("ActivityType")
    Set oReturn = MySAPCon.SAPBapiControl.DimAs(boActivityType, "ChangeMultiple", "Return")
    Set oActivityTypeList = MySAPCon.SAPBapiControl.DimAs(boActivityType, "ChangeMultiple", "ActivityTypeList")
    Set oLanguageKey = MySAPCon.SAPBapiControl.DimAs(boActivityType, "ChangeMultiple", "LanguageKey")
    oLanguageKey("LANGU") = aLanguageKey

    Dim aATLRow As Object
    Dim lRow As SAPActivityTypeList
    For Each lRow In pData
        Set aATLRow = oActivityTypeList.Rows.Add
        aATLRow.Value("ACTTYPE") = lRow.ACTTYPE
        aATLRow.Value("VALID_FROM") = lRow.VALID_FROM
        aATLRow.Value("VALID_TO") = lRow.VALID_TO
        aATLRow.Value("ACT_UNIT") = lRow.ACT_UNIT
        aATLRow.Value("COSTCENTER_CAT") = lRow.COSTCENTER_CAT
        aATLRow.Value("ACTIVITY_CAT") = lRow.ACTIVITY_CAT
        aATLRow.Value("ALLOC_COST_ELEM") = lRow.ALLOC_COST_ELEM
        aATLRow.Value("PRICE_INDICATOR") = lRow.PRICE_INDICATOR
        aATLRow.Value("AVERAGE_PRICE_IND") = lRow.AVERAGE_PRICE_IND
        aATLRow.Value("PREDIS_FXD_COST") = lRow.PREDIS_FXD_COST
        aATLRow.Value("MANUAL_ACT_QTY") = lRow.MANUAL_ACT_QTY
        aATLRow.Value("ACT_CAT_ACTUAL") = lRow.ACT_CAT_ACTUAL
        aATLRow.Value("ACT_PRICE_IND") = lRow.ACT_PRICE_IND
        aATLRow.Value("LOCK_INDICATOR") = lRow.LOCK_INDICATOR
        aATLRow.Value("OUTPUT_UNIT") = lRow.OUTPUT_UNIT
        aATLRow.Value("OUTPUT_FACTOR") = lRow.OUTPUT_FACTOR
        aATLRow.Value("NAME") = lRow.NAME
        aATLRow.Value("DESCRIPT") = lRow.DESCRIPT
    Next lRow
    Application.Cursor = xlWait
    If pTestRun = "X" Then
        boActivityType.changeMultiple CoArea:=pControllingArea, _
        TestRun:=pTestRun, _
        LanguageKey:=oLanguageKey, _
        ActivityTypeList:=oActivityTypeList, _
        Return:=oReturn
    Else
        boActivityType.changeMultiple CoArea:=pControllingArea, _
        ActivityTypeList:=oActivityTypeList, _
        LanguageKey:=oLanguageKey, _
        Return:=oReturn
    End If

    Application.Cursor = xlDefault

    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit

    If oReturn.RowCount = 0 Then
        changeMultiple = "Success ???"
        aSAPBapiTranctionCommit.commit
    Else
        Dim retRow As Object
        Set retRow = oReturn.Rows(1)
        If retRow("TYPE") = "S" Then
            changeMultiple = retRow("MESSAGE")
            aSAPBapiTranctionCommit.commit
        Else
            For Each retRow In oReturn.Rows
                changeMultiple = changeMultiple & ";" & retRow("MESSAGE")
            Next retRow
        End If
    End If
    Set boActivityType = Nothing
    Exit Function

changeMultiple_Error:
    Application.Cursor = xlDefault
    changeMultiple = "Error"
    MySAPErr.MSGProt "SAPActivityType", "changeMultiple", "", Err.Number, Err.Description
    Exit Function
End Function



