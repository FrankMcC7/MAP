Option Explicit

'==========================================================
' NewFundsIdentificationMacro (updated for TRANSPARENCY col)
' ---------------------------------------------------------
' • Inactive = Fund exists in SharePoint but NOT in HFTable.
' • “Tier” values come from HFTable column: TRANSPARENCY (1/2/3).
' • Filters use: IRR_last_update_date >= 01/01/2023 and TRANSPARENCY IN {1,2},
'   plus Strategy/Entity exclusion filters (unchanged).
'==========================================================

'=======================
' HELPERS
'=======================
Function GetOrClearSheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrClearSheet = wb.Sheets(sheetName)
    On Error GoTo 0
    If GetOrClearSheet Is Nothing Then
        Set GetOrClearSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        GetOrClearSheet.Name = sheetName
    Else
        GetOrClearSheet.Cells.Clear
    End If
End Function

Function EnsureTable(ws As Worksheet, tblName As String) As ListObject
    On Error Resume Next
    Set EnsureTable = ws.ListObjects(tblName)
    On Error GoTo 0
    If EnsureTable Is Nothing Then
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes)
        EnsureTable.Name = tblName
    End If
End Function

Function EnsureTableRange(ws As Worksheet, tblName As String, rng As Range) As ListObject
    On Error Resume Next
    Set EnsureTableRange = ws.ListObjects(tblName)
    On Error GoTo 0
    If EnsureTableRange Is Nothing Then
        Set EnsureTableRange = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        EnsureTableRange.Name = tblName
    Else
        EnsureTableRange.Resize rng
    End If
End Function

Function GetColumnIndex(lo As ListObject, headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If Trim(lo.HeaderRowRange.Cells(1, i).Value) = headerName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

Function ColumnExists(lo As ListObject, colName As String) As Boolean
    Dim cl As ListColumn
    For Each cl In lo.ListColumns
        If Trim(cl.Name) = colName Then ColumnExists = True: Exit Function
    Next cl
    ColumnExists = False
End Function

Function IsArrayNonEmpty(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayNonEmpty = (IsArray(arr) And (UBound(arr) >= LBound(arr)))
End Function

Function GetAllowedValues(lo As ListObject, fieldName As String, excludeArr As Variant) As Variant
    Dim colIdx As Long: colIdx = GetColumnIndex(lo, fieldName)
    If colIdx = 0 Then GetAllowedValues = Array(): Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cell As Range, valStr As String, i As Long, skipVal As Boolean

    For Each cell In lo.ListColumns(fieldName).DataBodyRange
        valStr = Trim(CStr(cell.Value))
        skipVal = False
        For i = LBound(excludeArr) To UBound(excludeArr)
            If StrComp(valStr, Trim(CStr(excludeArr(i))), vbTextCompare) = 0 Then
                skipVal = True: Exit For
            End If
        Next i
        If Not skipVal Then
            If Not dict.Exists(valStr) Then dict.Add valStr, valStr
        End If
    Next cell

    If dict.Count > 0 Then GetAllowedValues = dict.Keys Else GetAllowedValues = Array()
End Function

Sub ApplyStrategyFilter(ByVal loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Strategy")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Strategy", Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark"))
    If Not IsArrayNonEmpty(allowed) Then Exit Sub
    loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

Sub ApplyEntityFilter(ByVal loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Entity_type")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Entity_type", Array("Guaranteed subsidiary", "Investment Manager as Agent", _
                                "Managed Account", "Managed Account - No AF", "Loan Monitoring", "Loan FiF - No tracking", _
                                "Sleeve/share class/sub-account"))
    If Not IsArrayNonEmpty(allowed) Then Exit Sub
    loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

'=======================
' MAIN MACRO
'=======================
Sub NewFundsIdentificationMacro()
    '-----------------------
    ' 0. File paths
    '-----------------------
    Dim HFFilePath As String: HFFilePath = "C:\YourFolder\HFFile.xlsx"
    Dim SPFilePath As String: SPFilePath = "C:\YourFolder\SharePointFile.xlsx"

    '-----------------------
    ' 1. Workbooks/sheets
    '-----------------------
    Dim wbMain As Workbook: Set wbMain = ThisWorkbook
    Dim wbHF As Workbook, wbSP As Workbook
    Dim wsHFSource As Worksheet, wsSPSource As Worksheet
    Dim wsSourcePop As Worksheet, wsSPMain As Worksheet
    Dim wsUpload As Worksheet, wsInactive As Worksheet, wsCO As Worksheet

    '-----------------------
    ' 2. Tables
    '-----------------------
    Dim loHF As ListObject, loSP As ListObject
    Dim loMainHF As ListObject, loMainSP As ListObject
    Dim loUpload As ListObject, loInactive As ListObject, loCO As ListObject

    '-----------------------
    ' 3. Dictionaries & Collections
    '-----------------------
    Dim dictSP As Object: Set dictSP = CreateObject("Scripting.Dictionary"): dictSP.CompareMode = vbTextCompare
    Dim dictHF As Object: Set dictHF = CreateObject("Scripting.Dictionary"): dictHF.CompareMode = vbTextCompare
    Dim tierDict As Object: Set tierDict = CreateObject("Scripting.Dictionary"): tierDict.CompareMode = vbTextCompare

    Dim coDict As Object: Set coDict = CreateObject("Scripting.Dictionary"): coDict.CompareMode = vbTextCompare
    Dim imDict As Object: Set imDict = CreateObject("Scripting.Dictionary"): imDict.CompareMode = vbTextCompare
    Dim daysDict As Object: Set daysDict = CreateObject("Scripting.Dictionary"): daysDict.CompareMode = vbTextCompare

    Dim newFunds As Collection: Set newFunds = New Collection
    Dim inactiveFunds As Collection: Set inactiveFunds = New Collection

    '-----------------------
    ' 4. Other vars
    '-----------------------
    Dim colIndex As Long, i As Long, j As Long, rIdx As Long
    Dim rowCounter As Long
    Dim key As Variant, fundCoperID As String
    Dim rec As Variant
    Dim visData As Range, r As Range

    ' Column indexes placeholders
    Dim coCredCol As Long, coRegionCol As Long, coEmailCol As Long
    Dim sp_IMCol As Long, sp_NAVCol As Long, sp_FreqCol As Long, sp_AdHocCol As Long, sp_ParentFlagCol As Long
    Dim up_CredCol As Long, up_RegCol As Long, up_IMIDCol As Long, up_NAVCol As Long
    Dim up_FreqCol As Long, up_AdHocCol As Long, up_ParFlagCol As Long, up_DaysCol As Long, up_FundCol As Long
    Dim hfFundIDCol As Long, hfDaysCol As Long, idxTier As Long
    Dim share_CoperCol As Long, share_StatusCol As Long, share_CommentsCol As Long

    '=======================
    ' OPEN SOURCE WORKBOOKS
    '=======================
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Worksheets(1)
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, wsHFSource.UsedRange, , xlYes)
    End If
    loHF.Name = "HFTable"

    Set wbSP = Workbooks.Open(SPFilePath)
    Set wsSPSource = wbSP.Worksheets(1)
    If wsSPSource.ListObjects.Count > 0 Then
        Set loSP = wsSPSource.ListObjects(1)
    Else
        Set loSP = wsSPSource.ListObjects.Add(xlSrcRange, wsSPSource.UsedRange, , xlYes)
    End If
    loSP.Name = "SharePoint"

    '=======================
    ' COPY INTO MAIN WORKBOOK
    '=======================
    Set wsSourcePop = GetOrClearSheet(wbMain, "Source Population")
    Set wsSPMain    = GetOrClearSheet(wbMain, "SharePoint")

    loHF.Range.Copy wsSourcePop.Range("A1")
    loSP.Range.Copy wsSPMain.Range("A1")

    wbHF.Close False: wbSP.Close False

    Set loMainHF = EnsureTable(wsSourcePop, "HFTable")
    Set loMainSP = EnsureTable(wsSPMain,   "SharePoint")

    '=======================
    ' BUILD dictHF (ALL HF rows, independent of filters)
    '=======================
    hfFundIDCol = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    If hfFundIDCol > 0 Then
        For rIdx = 1 To loMainHF.DataBodyRange.Rows.Count
            key = Trim(CStr(loMainHF.DataBodyRange.Cells(rIdx, hfFundIDCol).Value))
            If Len(key) > 0 Then dictHF(key) = True
        Next rIdx
    End If

    '=======================
    ' FILTER HFTable (Date + Tier + Strategy/Entity)
    '=======================
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData

    colIndex = GetColumnIndex(loMainHF, "IRR_last_update_date")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=">=01/01/2023", Operator:=xlFilterValues
    End If

    colIndex = GetColumnIndex(loMainHF, "TRANSPARENCY")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=Array("1", "2"), Operator:=xlFilterValues
    End If

    ApplyStrategyFilter loMainHF
    ApplyEntityFilter  loMainHF

    '=======================
    ' BUILD dictSP (existing funds in SharePoint)
    '=======================
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    For i = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(i, colIndex).Value))
        If Len(key) > 0 Then dictSP(key) = True
    Next i

    '=======================
    ' BUILD tierDict (from FILTERED HF rows) using TRANSPARENCY
    '=======================
    idxTier = GetColumnIndex(loMainHF, "TRANSPARENCY")
    For rIdx = 1 To loMainHF.DataBodyRange.Rows.Count
        If Not loMainHF.DataBodyRange.Rows(rIdx).Hidden Then
            key = Trim(CStr(loMainHF.DataBodyRange.Cells(rIdx, hfFundIDCol).Value))
            If Len(key) > 0 Then
                If Not tierDict.Exists(key) Then tierDict(key) = loMainHF.DataBodyRange.Cells(rIdx, idxTier).Value
            End If
        End If
    Next rIdx

    '=======================
    ' COLLECT NEW FUNDS (filtered HF not present in SP)
    '=======================
    Dim idxName As Long, idxIMID As Long, idxIMName As Long, idxCred As Long
    idxName    = GetColumnIndex(loMainHF, "HFAD_Fund_Name")
    idxIMID    = GetColumnIndex(loMainHF, "HFAD_IM_CoperID")
    idxIMName  = GetColumnIndex(loMainHF, "HFAD_IM_Name")
    idxCred    = GetColumnIndex(loMainHF, "HFAD_Credit_Officer")

    On Error Resume Next
    Set visData = loMainHF.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visData Is Nothing Then
        For Each r In visData.Rows
            If Not r.EntireRow.Hidden Then
                fundCoperID = Trim(CStr(r.Cells(1, hfFundIDCol).Value))
                If Len(fundCoperID) > 0 Then
                    If Not dictSP.Exists(fundCoperID) Then
                        rec = Array(fundCoperID, r.Cells(1, idxName).Value, r.Cells(1, idxIMID).Value, _
                                    r.Cells(1, idxIMName).Value, r.Cells(1, idxCred).Value, _
                                    tierDict(fundCoperID), "Active")
                        newFunds.Add rec
                    End If
                End If
            End If
        Next r
    End If

    '=======================
    ' CREATE / POPULATE Upload to SP
    '=======================
    Set wsUpload = GetOrClearSheet(wbMain, "Upload to SP")
    Dim upHeaders As Variant: upHeaders = Array("HFAD_Fund_CoperID", "HFAD_Fund_Name", "HFAD_IM_CoperID", _
                                               "HFAD_IM_Name", "HFAD_Credit_Officer", "Tier", "Status")
    For j = LBound(upHeaders) To UBound(upHeaders)
        wsUpload.Cells(1, j + 1).Value = upHeaders(j)
    Next j

    rowCounter = 2
    For Each rec In newFunds
        For j = LBound(rec) To UBound(rec)
            wsUpload.Cells(rowCounter, j + 1).Value = rec(j)
        Next j
        rowCounter = rowCounter + 1
    Next rec

    Dim rngUpload As Range
    Set rngUpload = wsUpload.Range(wsUpload.Cells(1, 1), wsUpload.Cells(rowCounter - 1, UBound(upHeaders) + 1))
    Set loUpload = EnsureTableRange(wsUpload, "UploadHF", rngUpload)

    '=======================
    ' ADDITIONAL LOOKUPS & POPULATE EXTRA COLUMNS IN UploadHF
    '=======================
    '--- CO_Table (Credit Officer -> Region, Email) --------
    On Error Resume Next: Set wsCO = wbMain.Sheets("CO_Table"): On Error GoTo 0
    If Not wsCO Is Nothing Then
        Set loCO = wsCO.ListObjects("CO_Table")
        If Not loCO Is Nothing Then
            coCredCol   = GetColumnIndex(loCO, "Credit Officer")
            coRegionCol = GetColumnIndex(loCO, "Region")
            coEmailCol  = GetColumnIndex(loCO, "Email Address")
            For rIdx = 1 To loCO.DataBodyRange.Rows.Count
                key = Trim(CStr(loCO.DataBodyRange.Cells(rIdx, coCredCol).Value))
                If Len(key) > 0 Then
                    If Not coDict.Exists(key) Then coDict.Add key, Array(loCO.DataBodyRange.Cells(rIdx, coRegionCol).Value, loCO.DataBodyRange.Cells(rIdx, coEmailCol).Value)
                End If
            Next rIdx
        End If
    End If

    '--- SharePoint IM dictionary --------------------------
    sp_IMCol        = GetColumnIndex(loMainSP, "HFAD_IM_CoperID")
    sp_NAVCol       = GetColumnIndex(loMainSP, "NAV Source")
    sp_FreqCol      = GetColumnIndex(loMainSP, "Frequency")
    sp_AdHocCol     = GetColumnIndex(loMainSP, "Ad-Hoc Reporting")
    sp_ParentFlagCol = GetColumnIndex(loMainSP, "Parent/Flagship Reporting")

    For rIdx = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(rIdx, sp_IMCol).Value))
        If Len(key) > 0 Then
            If Not imDict.Exists(key) Then
                imDict.Add key, Array(loMainSP.DataBodyRange.Cells(rIdx, sp_NAVCol).Value, _
                                      loMainSP.DataBodyRange.Cells(rIdx, sp_FreqCol).Value, _
                                      loMainSP.DataBodyRange.Cells(rIdx, sp_AdHocCol).Value, _
                                      loMainSP.DataBodyRange.Cells(rIdx, sp_ParentFlagCol).Value)
            End If
        End If
    Next rIdx

    '--- Days to Report dict --------------------------------
    hfDaysCol = GetColumnIndex(loMainHF, "HFAD_Days_to_report")
    For rIdx = 1 To loMainHF.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainHF.DataBodyRange.Cells(rIdx, hfFundIDCol).Value))
        If Len(key) > 0 Then
            If Not daysDict.Exists(key) Then daysDict.Add key, loMainHF.DataBodyRange.Cells(rIdx, hfDaysCol).Value
        End If
    Next rIdx

    '--- Ensure & Fill extra columns ------------------------
    Dim extraCols As Variant
    extraCols = Array("Region", "NAV Source", "Frequency", "Ad-Hoc Reporting", "Parent/Flagship Reporting", "Days to Report")
    For Each key In extraCols
        If Not ColumnExists(loUpload, CStr(key)) Then loUpload.ListColumns.Add.Name = CStr(key)
    Next key

    up_CredCol    = GetColumnIndex(loUpload, "HFAD_Credit_Officer")
    up_RegCol     = GetColumnIndex(loUpload, "Region")
    up_IMIDCol    = GetColumnIndex(loUpload, "HFAD_IM_CoperID")
    up_NAVCol     = GetColumnIndex(loUpload, "NAV Source")
    up_FreqCol    = GetColumnIndex(loUpload, "Frequency")
    up_AdHocCol   = GetColumnIndex(loUpload, "Ad-Hoc Reporting")
    up_ParFlagCol = GetColumnIndex(loUpload, "Parent/Flagship Reporting")
    up_DaysCol    = GetColumnIndex(loUpload, "Days to Report")
    up_FundCol    = GetColumnIndex(loUpload, "HFAD_Fund_CoperID")

    Dim lrU As ListRow
    For Each lrU In loUpload.ListRows
        ' Credit officer / region / email
        key = Trim(CStr(lrU.Range.Cells(1, up_CredCol).Value))
        If coDict.Exists(key) Then
            lrU.Range.Cells(1, up_CredCol).Value = coDict(key)(1) 'Email
            lrU.Range.Cells(1, up_RegCol).Value = coDict(key)(0)
        End If

        ' IM-driven fields
        key = Trim(CStr(lrU.Range.Cells(1, up_IMIDCol).Value))
        If imDict.Exists(key) Then
            lrU.Range.Cells(1, up_NAVCol).Value     = imDict(key)(0)
            lrU.Range.Cells(1, up_FreqCol).Value    = imDict(key)(1)
            lrU.Range.Cells(1, up_AdHocCol).Value   = imDict(key)(2)
            lrU.Range.Cells(1, up_ParFlagCol).Value = imDict(key)(3)
        End If

        ' Days to report
        key = Trim(CStr(lrU.Range.Cells(1, up_FundCol).Value))
        If daysDict.Exists(key) Then lrU.Range.Cells(1, up_DaysCol).Value = daysDict(key)
    Next lrU

    '=======================
    ' INACTIVE FUNDS (SharePoint fund NOT in HFTable)
    '=======================
    share_CoperCol   = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    share_StatusCol  = GetColumnIndex(loMainSP, "Status")
    share_CommentsCol = GetColumnIndex(loMainSP, "Comments")

    For rIdx = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(rIdx, share_CoperCol).Value))
        If Len(key) > 0 Then
            If Not dictHF.Exists(key) Then   'fund not in HFTable (inactive)
                inactiveFunds.Add Array(key, loMainSP.DataBodyRange.Cells(rIdx, share_StatusCol).Value, _
                                        loMainSP.DataBodyRange.Cells(rIdx, share_CommentsCol).Value, "")
            End If
        End If
    Next rIdx

    '=======================
    ' BUILD Inactive Funds sheet
    '=======================
    Set wsInactive = GetOrClearSheet(wbMain, "Inactive Funds Tracking")
    Dim inactHeaders As Variant
    inactHeaders = Array("HFAD_Fund_CoperID", "Status", "Comments", "Tier")
    For j = LBound(inactHeaders) To UBound(inactHeaders)
        wsInactive.Cells(1, j + 1).Value = inactHeaders(j)
    Next j

    rowCounter = 2
    For Each rec In inactiveFunds
        For j = LBound(rec) To UBound(rec)
            wsInactive.Cells(rowCounter, j + 1).Value = rec(j)
        Next j
        rowCounter = rowCounter + 1
    Next rec

    Dim rngInactive As Range
    Set rngInactive = wsInactive.Range(wsInactive.Cells(1, 1), wsInactive.Cells(rowCounter - 1, UBound(inactHeaders) + 1))
    Set loInactive = EnsureTableRange(wsInactive, "InactiveHF", rngInactive)

    '=======================
    ' ADD TIER FROM tierDict AFTER TABLE BUILT
    '=======================
    Dim in_FundCol As Long, in_TierCol As Long
    in_FundCol = GetColumnIndex(loInactive, "HFAD_Fund_CoperID")
    in_TierCol = GetColumnIndex(loInactive, "Tier")

    Dim lr As ListRow, fundID As String
    For Each lr In loInactive.ListRows
        fundID = Trim(CStr(lr.Range.Cells(1, in_FundCol).Value))
        If tierDict.Exists(fundID) Then
            lr.Range.Cells(1, in_TierCol).Value = tierDict(fundID)
        End If
    Next lr

    '=======================
    MsgBox "Macro completed successfully.", vbInformation
End Sub
