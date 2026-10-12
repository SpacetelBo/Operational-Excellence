Option Explicit

'=============================================================
' WHATSAPP SUMMARY - FULL UPDATED VBA
'
' IMPORTANT:
' 1. Allocation logic is preserved: Allocation A/B/D is loaded
'    and matched by Location ID + exact SKU, as in the current code.
' 2. In Stock details ACTUAL columns:
'       B = Location Name
'       C = SKU
'       H = Status
' 3. Transfer uses AVAILABLE only and requires quantity > 1.
' 4. A source store is consumed by 1 unit after each recommendation.
'    Therefore a source with 2 units can be recommended once,
'    3 units twice, 5 units four times, etc., leaving 1 unit.
' 5. Exact SKU is searched first.
' 6. If exact SKU is unavailable, another sellable SKU from the
'    same Family is searched.
' 7. Haitham: same Area Manager first, then same Region.
' 8. Non-Haitham: same Sub Dealer only.
'=============================================================

Public Sub BuildWhatsAppSummary()

    Dim wb As Workbook
    Dim wsDC As Worksheet, wsStock As Worksheet, wsCall As Worksheet
    Dim wsFamily As Worksheet, wsAlloc As Worksheet, wsOut As Worksheet

    Dim dictCall As Object, dictStock As Object
    Dim dictAlloc As Object, dictFamily As Object, dictFamilySell As Object
    Dim dictUID As Object

    Dim lastRow As Long, r As Long, outRow As Long

    Dim locID As String, locationName As String, subDealer As String
    Dim uidText As String
    Dim areaManager As String, regionName As String, shortStore As String
    Dim familyName As String, baseFamily As String
    Dim demoSKU As String, sellableSKU As String, productName As String
    Dim demoCompliant As String, sellableCompliant As String
    Dim skuToOrder As String
    Dim allocationQty As Double
    Dim actionText As String, transferText As String

    Dim info As Variant
    Dim oldCalc As XlCalculation

    On Error GoTo ErrorHandler

    oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook

    '---------------- GET SHEETS ----------------
    Set wsDC = GetSheet(wb, "Device Compliance")
    Set wsStock = GetSheet(wb, "In Stock details")
    Set wsCall = GetSheet(wb, "Calling Tree")
    Set wsFamily = GetSheet(wb, "Family")
    Set wsAlloc = GetSheet(wb, "Allocation")
    Set wsOut = GetSheet(wb, "WhatsApp Summary")

    If wsDC Is Nothing Then Err.Raise vbObjectError + 100, , "Missing sheet: Device Compliance"
    If wsStock Is Nothing Then Err.Raise vbObjectError + 101, , "Missing sheet: In Stock details"
    If wsCall Is Nothing Then Err.Raise vbObjectError + 102, , "Missing sheet: Calling Tree"
    If wsFamily Is Nothing Then Err.Raise vbObjectError + 103, , "Missing sheet: Family"
    If wsAlloc Is Nothing Then Err.Raise vbObjectError + 104, , "Missing sheet: Allocation"
    If wsOut Is Nothing Then Err.Raise vbObjectError + 105, , "Missing sheet: WhatsApp Summary"

    '---------------- DICTIONARIES ----------------
    Set dictCall = CreateObject("Scripting.Dictionary")
    dictCall.CompareMode = vbTextCompare

    Set dictStock = CreateObject("Scripting.Dictionary")
    dictStock.CompareMode = vbTextCompare

    Set dictAlloc = CreateObject("Scripting.Dictionary")
    dictAlloc.CompareMode = vbTextCompare

    Set dictFamily = CreateObject("Scripting.Dictionary")
    dictFamily.CompareMode = vbTextCompare

    Set dictFamilySell = CreateObject("Scripting.Dictionary")
    dictFamilySell.CompareMode = vbTextCompare

    Set dictUID = CreateObject("Scripting.Dictionary")
    dictUID.CompareMode = vbTextCompare

    LoadCallingTree wsCall, dictCall, dictUID
    LoadStock wsStock, dictStock
    LoadAllocation wsAlloc, dictAlloc
    LoadFamily wsFamily, dictFamily, dictFamilySell

    '---------------- CLEAR OUTPUT + HEADERS ----------------
    wsOut.Cells.Clear

    wsOut.Cells(1, 1).Value = "Sub Dealers"
    wsOut.Cells(1, 2).Value = "Store"
    wsOut.Cells(1, 3).Value = "Family"
    wsOut.Cells(1, 4).Value = "Ops excellence missing"
    wsOut.Cells(1, 5).Value = "SKU to Order"
    wsOut.Cells(1, 6).Value = "Product Name"
    wsOut.Cells(1, 7).Value = "Allocation Qty to order"
    wsOut.Cells(1, 8).Value = "Action"
    wsOut.Cells(1, 9).Value = "Transfer Recommendation"

    outRow = 2

    '=========================================================
    ' DEVICE COMPLIANCE
    ' A=UID (labelled Location ID)  B=Sub Dealers
    ' C=Location Name  D=Demo SKU  E=Family Group
    ' F=Demo Compliant  I=Sellable Compliant
    '=========================================================
    lastRow = LastUsedRow(wsDC)

    For r = 2 To lastRow

        subDealer = CleanText(wsDC.Cells(r, 2).Value)
        locationName = CleanText(wsDC.Cells(r, 3).Value)

        locID = ExtractLocationID(locationName)

        ' Fallback: Device Compliance column A is UID.
        If Len(locID) = 0 Then
            uidText = CleanText(wsDC.Cells(r, 1).Value)
            If dictUID.Exists(uidText) Then locID = dictUID(uidText)
        ElseIf Not dictCall.Exists(locID) Then
            uidText = CleanText(wsDC.Cells(r, 1).Value)
            If dictUID.Exists(uidText) Then locID = dictUID(uidText)
        End If

        areaManager = ""
        regionName = ""
        shortStore = ""

        If Len(locID) > 0 Then
            If dictCall.Exists(locID) Then
                info = dictCall(locID)
                areaManager = CleanText(info(0))
                regionName = CleanText(info(1))
                shortStore = CleanText(info(2))
            End If
        End If

        If Len(subDealer) = 0 Then
            If dictCall.Exists(locID) Then
                subDealer = CleanText(dictCall(locID)(3))
            End If
        End If

        If UCase$(subDealer) = "HAITHAM" Then subDealer = "Haitham"

        '---------------- FAMILY ----------------
        familyName = CleanText(wsDC.Cells(r, 5).Value)
        If Len(familyName) = 0 Then GoTo NextDeviceRow

        baseFamily = RemoveDemoSuffix(familyName)

        demoSKU = ""
        sellableSKU = ""

        If dictFamily.Exists(NormalizeFamily(baseFamily)) Then
            info = dictFamily(NormalizeFamily(baseFamily))
            demoSKU = CleanSKU(info(0))
            sellableSKU = CleanSKU(info(1))
        End If

        If Len(demoSKU) = 0 Then demoSKU = CleanSKU(wsDC.Cells(r, 4).Value)

        '=========================================================
        ' DEMO
        '=========================================================
        demoCompliant = UCase$(CleanText(wsDC.Cells(r, 6).Value))

        If demoCompliant = "NO" Then

            skuToOrder = demoSKU

            If Len(skuToOrder) > 0 Then

                productName = GetFamilyProductNameFromSellables(dictFamilySell, baseFamily, skuToOrder)
                If Len(productName) = 0 Then productName = GetFamilyProductName(dictFamily, baseFamily, skuToOrder)
                If Len(productName) = 0 Then productName = baseFamily

                allocationQty = GetAllocationQty(dictAlloc, locID, locationName, skuToOrder)

                '=====================================================
                ' ONLY SHOW THE ROW IF ALLOCATION IS AVAILABLE
                '=====================================================
                If allocationQty > 0 Then

                    actionText = GetActionText(allocationQty)

                    transferText = BuildTransferRecommendation( _
                                        locID, locationName, subDealer, areaManager, _
                                        regionName, skuToOrder, dictCall, dictStock, _
                                        dictFamilySell, baseFamily)

                    AddOrMergeOutputRow wsOut, outRow, subDealer, _
                        BuildShortStoreName(locID, shortStore, locationName), _
                        baseFamily, "OpEx Missing - Demo", skuToOrder, productName, _
                        allocationQty, actionText, transferText

                End If
            End If
        End If

        '=========================================================
        ' SELLABLE
        '=========================================================
        sellableCompliant = UCase$(CleanText(wsDC.Cells(r, 9).Value))

        If sellableCompliant = "NO" Then

            '=====================================================
            ' IMPORTANT:
            ' Process ALL sellable SKUs in this family, but ONLY
            ' create an output row when Allocation Qty > 0.
            '=====================================================

            Dim sellOpts As Collection
            Dim sellItem As Variant
            Dim sellIndex As Long
            Dim familyKey As String

            familyKey = NormalizeFamily(baseFamily)

            If dictFamilySell.Exists(familyKey) Then

                Set sellOpts = dictFamilySell(familyKey)

                For sellIndex = 1 To sellOpts.Count

                    sellItem = sellOpts(sellIndex)

                    skuToOrder = CleanSKU(sellItem(0))
                    productName = CleanText(sellItem(1))

                    If Len(productName) = 0 Then
                        productName = GetFamilyProductName(dictFamily, baseFamily, skuToOrder)
                    End If

                    If Len(productName) = 0 Then
                        productName = baseFamily
                    End If

                    allocationQty = GetAllocationQty( _
                                        dictAlloc, _
                                        locID, _
                                        locationName, _
                                        skuToOrder)

                    '=================================================
                    ' DO NOT SHOW ZERO-ALLOCATION SKUs
                    '=================================================
                    If allocationQty > 0 Then

                        actionText = GetActionText(allocationQty)

                        transferText = BuildTransferRecommendation( _
                                            locID, locationName, subDealer, areaManager, _
                                            regionName, skuToOrder, dictCall, dictStock, _
                                            dictFamilySell, baseFamily)

                        AddOrMergeOutputRow wsOut, outRow, subDealer, _
                            BuildShortStoreName(locID, shortStore, locationName), _
                            baseFamily, "OpEx Missing - Sellable", skuToOrder, productName, _
                            allocationQty, actionText, transferText

                    End If

                Next sellIndex

            Else

                'Fallback only when the Family sheet has no sellable list.
                skuToOrder = sellableSKU

                If Len(skuToOrder) = 0 Then
                    skuToOrder = CleanSKU(wsDC.Cells(r, 4).Value)
                End If

                If Len(skuToOrder) > 0 Then

                    allocationQty = GetAllocationQty( _
                                        dictAlloc, _
                                        locID, _
                                        locationName, _
                                        skuToOrder)

                    If allocationQty > 0 Then

                        productName = GetFamilyProductName(dictFamily, baseFamily, skuToOrder)
                        If Len(productName) = 0 Then productName = baseFamily

                        actionText = GetActionText(allocationQty)

                        transferText = BuildTransferRecommendation( _
                                            locID, locationName, subDealer, areaManager, _
                                            regionName, skuToOrder, dictCall, dictStock, _
                                            dictFamilySell, baseFamily)

                        AddOrMergeOutputRow wsOut, outRow, subDealer, _
                            BuildShortStoreName(locID, shortStore, locationName), _
                            baseFamily, "OpEx Missing - Sellable", skuToOrder, productName, _
                            allocationQty, actionText, transferText

                    End If

                End If

            End If

        End If

NextDeviceRow:
    Next r

    FormatWhatsAppSummary wsOut, outRow - 1

    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "WhatsApp Summary updated successfully." & vbCrLf & _
           "Rows: " & (outRow - 2), vbInformation, "WhatsApp Summary"
    Exit Sub

ErrorHandler:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "WhatsApp Summary Error:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "WhatsApp Summary"

End Sub


'=============================================================
' TRANSFER RECOMMENDATION
'
' Exact SKU first.
' If exact SKU is unavailable, search another sellable SKU
' from the SAME FAMILY.
'
' HAITHAM:
'   1. Same Area Manager / Haitham
'   2. Same Region / Haitham
'
' NON-HAITHAM:
'   Same Sub Dealer only.
'
' IMPORTANT:
' Every successful recommendation consumes ONE unit from
' dictStock. This means:
'   2 units -> 1 recommendation
'   3 units -> 2 recommendations
'   5 units -> 4 recommendations
' Always leave 1 unit at source.
'=============================================================
Private Function BuildTransferRecommendation( _
        ByVal targetLocID As String, _
        ByVal targetLocationName As String, _
        ByVal targetSubDealer As String, _
        ByVal targetAM As String, _
        ByVal targetRegion As String, _
        ByVal sku As String, _
        ByVal dictCall As Object, _
        ByVal dictStock As Object, _
        ByVal dictFamilySell As Object, _
        ByVal familyName As String) As String

    Dim tKey As String
    Dim info As Variant
    Dim found As Boolean
    Dim bestQty As Double
    Dim bestLocID As String
    Dim bestShort As String
    Dim selectedSKU As String
    Dim selectedProduct As String
    Dim isHaitham As Boolean

    tKey = NormalizeLocationID(targetLocID)

    If Len(tKey) = 0 Then
        BuildTransferRecommendation = "No Qty available at other stores."
        Exit Function
    End If

    If Not dictCall.Exists(tKey) Then
        BuildTransferRecommendation = "No Qty available at other stores."
        Exit Function
    End If

    info = dictCall(tKey)
    targetAM = CleanText(info(0))
    targetRegion = CleanText(info(1))
    targetSubDealer = CleanText(info(3))
    isHaitham = (UCase$(targetSubDealer) = "HAITHAM")

    '=========================================================
    ' 1. EXACT SKU
    '=========================================================
    If isHaitham Then

        found = FindBestSource( _
                    1, tKey, sku, targetAM, targetRegion, targetSubDealer, _
                    dictCall, dictStock, bestQty, bestLocID, bestShort)

        If Not found Then
            found = FindBestSource( _
                        2, tKey, sku, targetAM, targetRegion, targetSubDealer, _
                        dictCall, dictStock, bestQty, bestLocID, bestShort)
        End If

    Else

        found = FindBestSource( _
                    3, tKey, sku, targetAM, targetRegion, targetSubDealer, _
                    dictCall, dictStock, bestQty, bestLocID, bestShort)

    End If

    If found Then

        selectedSKU = CleanSKU(sku)
        selectedProduct = GetFamilyProductNameFromSellables( _
                            dictFamilySell, familyName, selectedSKU)

        ConsumeTransferStock dictStock, bestLocID, selectedSKU

        ' Show remaining stock AFTER this transfer.
        bestQty = bestQty - 1

        BuildTransferRecommendation = _
            FormatQtyStore(bestQty, bestLocID, bestShort, selectedSKU, selectedProduct) & _
            ". You can transfer 1 Qty from this store."

        Exit Function

    End If

    '=========================================================
    ' 2. SAME FAMILY FALLBACK
    '=========================================================
    If FindBestFamilyTransfer( _
            tKey, sku, familyName, targetAM, targetRegion, targetSubDealer, _
            dictCall, dictStock, dictFamilySell, _
            bestQty, bestLocID, bestShort, selectedSKU, selectedProduct) Then

        ConsumeTransferStock dictStock, bestLocID, selectedSKU

        bestQty = bestQty - 1

        BuildTransferRecommendation = _
            FormatQtyStore(bestQty, bestLocID, bestShort, selectedSKU, selectedProduct) & _
            ". You can transfer 1 Qty from this store."

        Exit Function

    End If

    BuildTransferRecommendation = "No Qty available at other stores."

End Function


'=============================================================
' SAME FAMILY TRANSFER SEARCH
'
' Searches all other sellable SKUs in the same family.
' For each candidate SKU, the normal priority rules apply.
'=============================================================
Private Function FindBestFamilyTransfer( _
        ByVal tKey As String, _
        ByVal requestedSKU As String, _
        ByVal familyName As String, _
        ByVal tAM As String, _
        ByVal tRegion As String, _
        ByVal tSub As String, _
        ByVal dictCall As Object, _
        ByVal dictStock As Object, _
        ByVal dictFamilySell As Object, _
        ByRef bestQty As Double, _
        ByRef bestLocID As String, _
        ByRef bestShort As String, _
        ByRef selectedSKU As String, _
        ByRef selectedProduct As String) As Boolean

    Dim familyKey As String
    Dim opts As Collection
    Dim j As Long
    Dim item As Variant
    Dim altSKU As String
    Dim altName As String
    Dim found As Boolean
    Dim isHaitham As Boolean

    familyKey = NormalizeFamily(familyName)

    bestQty = 0
    bestLocID = ""
    bestShort = ""
    selectedSKU = ""
    selectedProduct = ""

    If Not dictFamilySell.Exists(familyKey) Then Exit Function

    Set opts = dictFamilySell(familyKey)
    isHaitham = (UCase$(tSub) = "HAITHAM")

    For j = 1 To opts.Count

        item = opts(j)
        altSKU = CleanSKU(item(0))
        altName = CleanText(item(1))

        If Len(altSKU) > 0 Then

            If altSKU <> CleanSKU(requestedSKU) Then

                found = False

                If isHaitham Then

                    found = FindBestSource( _
                                1, tKey, altSKU, tAM, tRegion, tSub, _
                                dictCall, dictStock, bestQty, bestLocID, bestShort)

                    If Not found Then
                        found = FindBestSource( _
                                    2, tKey, altSKU, tAM, tRegion, tSub, _
                                    dictCall, dictStock, bestQty, bestLocID, bestShort)
                    End If

                Else

                    found = FindBestSource( _
                                3, tKey, altSKU, tAM, tRegion, tSub, _
                                dictCall, dictStock, bestQty, bestLocID, bestShort)

                End If

                If found Then
                    selectedSKU = altSKU
                    selectedProduct = altName
                    FindBestFamilyTransfer = True
                    Exit Function
                End If

            End If

        End If

    Next j

End Function


'=============================================================
' FIND BEST SOURCE
'
' searchLevel:
'   1 = Same Area Manager + Haitham
'   2 = Same Region + Haitham
'   3 = Same Sub Dealer
'
' Exact SKU only.
' Qty must be > 1.
' Highest CURRENT quantity wins.
' Tie = lowest Location ID.
'=============================================================
Private Function FindBestSource( _
        ByVal searchLevel As Long, _
        ByVal tKey As String, _
        ByVal sku As String, _
        ByVal tAM As String, _
        ByVal tRegion As String, _
        ByVal tSub As String, _
        ByVal dictCall As Object, _
        ByVal dictStock As Object, _
        ByRef bestQty As Double, _
        ByRef bestLocID As String, _
        ByRef bestShort As String) As Boolean

    Dim k As Variant
    Dim sLoc As String
    Dim qty As Double
    Dim info As Variant
    Dim ok As Boolean
    Dim takeIt As Boolean
    Dim skuKey As String
    Dim sourceSub As String
    Dim sourceAM As String
    Dim sourceRegion As String

    skuKey = CleanSKU(sku)
    bestQty = 0
    bestLocID = ""
    bestShort = ""

    If Len(skuKey) = 0 Then Exit Function

    For Each k In dictStock.Keys

        If CleanSKU(GetKeySKU(CStr(k))) = skuKey Then

            qty = NumberValue(dictStock(k))

            If qty > 1 Then

                sLoc = NormalizeLocationID(GetKeyLocationID(CStr(k)))

                If sLoc <> NormalizeLocationID(tKey) Then

                    If dictCall.Exists(sLoc) Then

                        info = dictCall(sLoc)
                        sourceAM = CleanText(info(0))
                        sourceRegion = CleanText(info(1))
                        sourceSub = CleanText(info(3))

                        ok = False

                        Select Case searchLevel

                            Case 1
                                If Len(tAM) > 0 Then
                                    ok = (UCase$(sourceSub) = "HAITHAM") And _
                                         (UCase$(sourceAM) = UCase$(tAM))
                                End If

                            Case 2
                                If Len(tRegion) > 0 Then
                                    ok = (UCase$(sourceSub) = "HAITHAM") And _
                                         (UCase$(sourceRegion) = UCase$(tRegion))
                                End If

                            Case 3
                                If Len(tSub) > 0 Then
                                    ok = (UCase$(sourceSub) = UCase$(tSub))
                                End If

                        End Select

                        If ok Then

                            takeIt = False

                            If qty > bestQty Then
                                takeIt = True
                            ElseIf qty = bestQty Then
                                If LocIDLess(sLoc, bestLocID) Then takeIt = True
                            End If

                            If takeIt Then
                                bestQty = qty
                                bestLocID = sLoc
                                bestShort = CleanText(info(2))
                            End If

                        End If

                    End If

                End If

            End If

        End If

    Next k

    FindBestSource = (bestQty > 1)

End Function


'=============================================================
' CONSUME ONE UNIT FROM SOURCE STOCK
'=============================================================
Private Sub ConsumeTransferStock( _
        ByVal dictStock As Object, _
        ByVal sourceLocID As String, _
        ByVal sku As String)

    Dim key As String
    Dim currentQty As Double

    key = NormalizeLocationID(sourceLocID) & "|" & CleanSKU(sku)

    If dictStock.Exists(key) Then

        currentQty = NumberValue(dictStock(key))

        If currentQty > 0 Then
            dictStock(key) = currentQty - 1
        End If

    End If

End Sub


Private Function LocIDLess(ByVal a As String, ByVal b As String) As Boolean

    If Len(b) = 0 Then
        LocIDLess = True
    ElseIf IsNumeric(a) And IsNumeric(b) Then
        LocIDLess = (CDbl(a) < CDbl(b))
    Else
        LocIDLess = (UCase$(a) < UCase$(b))
    End If

End Function


'=============================================================
' LOAD CALLING TREE
'
' HAITHAM: B=Loc ID C=UID D=Area Manager E=Region
'          F=Short Store Name G=Location Name
'
' OTHER:   I=Loc ID J=UID K=Sub Dealer
'          L=Short Store Name M=Location Name
'
' dictCall(LocID):
'   (0)=Area Manager
'   (1)=Region
'   (2)=Short Store Name
'   (3)=Sub Dealer
'   (4)=Location Name
'=============================================================
Private Sub LoadCallingTree( _
        ByVal ws As Worksheet, _
        ByVal dictCall As Object, _
        ByVal dictUID As Object)

    Dim lastRow As Long
    Dim r As Long
    Dim locID As String
    Dim uid As String
    Dim arr As Variant

    lastRow = LastUsedRow(ws)

    '---------------- HAITHAM ----------------
    For r = 3 To lastRow

        locID = NormalizeLocationID(ws.Cells(r, 2).Value)

        If Len(locID) > 0 Then

            arr = Array( _
                CleanText(ws.Cells(r, 4).Value), _
                CleanText(ws.Cells(r, 5).Value), _
                CleanText(ws.Cells(r, 6).Value), _
                "Haitham", _
                CleanText(ws.Cells(r, 7).Value))

            dictCall(locID) = arr

            uid = CleanText(ws.Cells(r, 3).Value)
            If Len(uid) > 0 Then dictUID(uid) = locID

        End If

    Next r

    '---------------- OTHER SUB DEALERS ----------------
    For r = 3 To lastRow

        locID = NormalizeLocationID(ws.Cells(r, 9).Value)

        If Len(locID) > 0 Then

            arr = Array( _
                "", _
                "", _
                CleanText(ws.Cells(r, 12).Value), _
                CleanText(ws.Cells(r, 11).Value), _
                CleanText(ws.Cells(r, 13).Value))

            dictCall(locID) = arr

            uid = CleanText(ws.Cells(r, 10).Value)
            If Len(uid) > 0 Then dictUID(uid) = locID

        End If

    Next r

End Sub


'=============================================================
' LOAD STOCK
'
' ACTUAL In Stock details columns:
'   B = Location Name
'   C = SKU
'   H = Status
'
' AVAILABLE rows only.
' Each AVAILABLE row = one unit.
' Key = LocID|SKU
'=============================================================
Private Sub LoadStock( _
        ByVal ws As Worksheet, _
        ByVal dictStock As Object)

    Dim lastRow As Long
    Dim r As Long
    Dim locationName As String
    Dim locID As String
    Dim sku As String
    Dim statusText As String
    Dim key As String

    lastRow = LastUsedRow(ws)

    For r = 2 To lastRow

        locationName = CleanText(ws.Cells(r, "B").Value)
        sku = CleanSKU(ws.Cells(r, "C").Value)
        statusText = UCase$(CleanText(ws.Cells(r, "H").Value))

        If Len(locationName) > 0 And Len(sku) > 0 Then

            If statusText = "AVAILABLE" Then

                locID = ExtractLocationID(locationName)

                If Len(locID) > 0 Then

                    key = NormalizeLocationID(locID) & "|" & sku

                    If dictStock.Exists(key) Then
                        dictStock(key) = NumberValue(dictStock(key)) + 1
                    Else
                        dictStock.Add key, 1
                    End If

                End If

            End If

        End If

    Next r

End Sub


'=============================================================
' LOAD ALLOCATION
'
' A = Location Name
' B = Product SKU
' D = Allocated Qty
'
' PRESERVED FROM CURRENT VERSION:
' exact Location ID + exact SKU after extracting Location ID.
'=============================================================
Private Sub LoadAllocation( _
        ByVal ws As Worksheet, _
        ByVal dictAlloc As Object)

    Dim lastRow As Long
    Dim r As Long
    Dim locationName As String
    Dim sku As String
    Dim qty As Double
    Dim key As String

    lastRow = LastUsedRow(ws)

    For r = 2 To lastRow

        '=========================================================
        ' Allocation sheet:
        ' A = Location Name
        ' B = Product SKU
        ' D = Allocated Qty
        '
        ' IMPORTANT: keep the existing exact Location Name + SKU
        ' lookup. Do NOT convert the Allocation location to Loc ID.
        '=========================================================

        locationName = CleanText(ws.Cells(r, 1).Value)
        sku = CleanSKU(ws.Cells(r, 2).Value)
        qty = NumberValue(ws.Cells(r, 4).Value)

        If Len(locationName) > 0 And Len(sku) > 0 Then

            key = locationName & "||" & sku

            If dictAlloc.Exists(key) Then
                dictAlloc(key) = NumberValue(dictAlloc(key)) + qty
            Else
                dictAlloc.Add key, qty
            End If

        End If

    Next r

End Sub


'=============================================================
' LOAD FAMILY
'
' A = Family Group
' B = Demo SKU
' C = Sellable SKU
' D = Product Name
'
' Handles merged Family cells correctly.
'
' dictFamily:
'   family -> Array(Demo SKU, first Sellable SKU, first Product Name)
'
' dictFamilySell:
'   family -> Collection of Array(Sellable SKU, Product Name)
'=============================================================
Private Sub LoadFamily( _
        ByVal ws As Worksheet, _
        ByVal dictFamily As Object, _
        ByVal dictFamilySell As Object)

    Dim lastRow As Long
    Dim r As Long
    Dim familyName As String
    Dim baseFamily As String
    Dim key As String
    Dim demoSKU As String
    Dim sellableSKU As String
    Dim productName As String
    Dim arr As Variant
    Dim mergeTopRow As Long

    lastRow = LastUsedRow(ws)

    For r = 2 To lastRow

        familyName = ""

        ' IMPORTANT: merged family cells store their value in the
        ' top-left cell of the MergeArea.
        On Error Resume Next
        If ws.Cells(r, 1).MergeCells Then
            mergeTopRow = ws.Cells(r, 1).MergeArea.Row
            familyName = CleanText(ws.Cells(mergeTopRow, 1).Value)
        Else
            familyName = CleanText(ws.Cells(r, 1).Value)
        End If
        On Error GoTo 0

        demoSKU = CleanSKU(ws.Cells(r, 2).Value)
        sellableSKU = CleanSKU(ws.Cells(r, 3).Value)
        productName = CleanText(ws.Cells(r, 4).Value)

        If InStr(1, familyName, "To be compliant", vbTextCompare) > 0 Then
            GoTo NextFamilyRow
        End If

        If Len(familyName) = 0 Then GoTo NextFamilyRow

        baseFamily = RemoveDemoSuffix(familyName)
        key = NormalizeFamily(baseFamily)

        '---------------- Main family information ----------------
        If Not dictFamily.Exists(key) Then

            arr = Array(demoSKU, sellableSKU, productName)
            dictFamily.Add key, arr

        Else

            arr = dictFamily(key)

            If Len(CleanSKU(arr(0))) = 0 And Len(demoSKU) > 0 Then arr(0) = demoSKU
            If Len(CleanSKU(arr(1))) = 0 And Len(sellableSKU) > 0 Then arr(1) = sellableSKU
            If Len(CleanText(arr(2))) = 0 And Len(productName) > 0 Then arr(2) = productName

            dictFamily(key) = arr

        End If

        '---------------- ALL sellable SKUs ----------------
        If Len(sellableSKU) > 0 Then

            If Not dictFamilySell.Exists(key) Then
                Set dictFamilySell(key) = New Collection
            End If

            arr = Array(sellableSKU, productName)
            dictFamilySell(key).Add arr

        End If

NextFamilyRow:
    Next r

End Sub


'=============================================================
' ALLOCATION / PRODUCT / ACTION
'=============================================================
Private Function GetAllocationQty( _
        ByVal dictAlloc As Object, _
        ByVal locID As String, _
        ByVal locationName As String, _
        ByVal sku As String) As Double

    Dim key As String

    '=========================================================
    ' Exact Location Name + exact SKU.
    ' locID is retained only for compatibility with the rest
    ' of the module; Allocation itself is NOT converted to ID.
    '=========================================================
    key = CleanText(locationName) & "||" & CleanSKU(sku)

    If dictAlloc.Exists(key) Then
        GetAllocationQty = NumberValue(dictAlloc(key))
    Else
        GetAllocationQty = 0
    End If

End Function


Private Function GetFamilyProductName( _
        ByVal dictFamily As Object, _
        ByVal familyName As String, _
        ByVal sku As String) As String

    Dim key As String
    Dim info As Variant

    key = NormalizeFamily(familyName)

    If dictFamily.Exists(key) Then

        info = dictFamily(key)

        ' If the family has multiple sellable SKUs, try to find the
        ' exact SKU elsewhere in the Family sheet's stored collection
        ' through the helper below. Main fallback is the family name.
        If Len(CleanText(info(2))) > 0 Then
            GetFamilyProductName = CleanText(info(2))
        Else
            GetFamilyProductName = familyName
        End If

    Else
        GetFamilyProductName = familyName
    End If

End Function


Private Function GetFamilyProductNameFromSellables( _
        ByVal dictFamilySell As Object, _
        ByVal familyName As String, _
        ByVal sku As String) As String

    Dim key As String
    Dim opts As Collection
    Dim j As Long
    Dim item As Variant

    key = NormalizeFamily(familyName)

    If dictFamilySell.Exists(key) Then

        Set opts = dictFamilySell(key)

        For j = 1 To opts.Count

            item = opts(j)

            If CleanSKU(item(0)) = CleanSKU(sku) Then
                GetFamilyProductNameFromSellables = CleanText(item(1))
                Exit Function
            End If

        Next j

    End If

    GetFamilyProductNameFromSellables = ""

End Function


Private Function GetActionText(ByVal allocationQty As Double) As String
    GetActionText = "Order any of the available SKU (atleast 1 qty)"
End Function



'=============================================================
' ADD / MERGE OUTPUT ROW
'
' ONE ROW PER: Sub Dealer + Store + Family + Issue
'
' If more than one Sellable SKU in the same Family has
' allocation for the same store, all SKUs are combined into
' the same row.
'
' E = SKU list (one per line)
' F = Product list (one per line)
' G = Allocation list (one per line)
' H = ONE Action
' I = Transfer recommendations (one per line)
'=============================================================
Private Sub AddOrMergeOutputRow( _
        ByVal ws As Worksheet, _
        ByRef outRow As Long, _
        ByVal subDealer As String, _
        ByVal storeName As String, _
        ByVal familyName As String, _
        ByVal issueText As String, _
        ByVal sku As String, _
        ByVal productName As String, _
        ByVal allocationQty As Double, _
        ByVal actionText As String, _
        ByVal transferText As String)

    Dim r As Long
    Dim groupEnd As Long
    Dim existingSKU As String
    Dim keyMatch As Boolean

    '=============================================================
    ' ONE GROUP = Store + Family + Issue
    '
    ' IMPORTANT:
    ' Each SKU gets its own physical worksheet row so that:
    '   E = SKU
    '   F = Product Name
    '   G = Allocation Qty
    '
    ' Then FormatWhatsAppSummary merges:
    '   A:D vertically
    '   H vertically
    '   I vertically
    '
    ' This produces the requested layout:
    '
    ' A-D | SKU1 | Product1 | Qty1 | H | I
    '     | SKU2 | Product2 | Qty2 |   |  (merged)
    '=============================================================

    For r = 2 To outRow - 1

        keyMatch = _
            (UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = UCase$(Trim$(subDealer))) And _
            (UCase$(Trim$(CStr(ws.Cells(r, 2).Value))) = UCase$(Trim$(storeName))) And _
            (UCase$(Trim$(CStr(ws.Cells(r, 3).Value))) = UCase$(Trim$(familyName))) And _
            (UCase$(Trim$(CStr(ws.Cells(r, 4).Value))) = UCase$(Trim$(issueText)))

        If keyMatch Then

            '-----------------------------------------------------
            ' Find the end of this existing group.
            '-----------------------------------------------------
            groupEnd = r

            Do While groupEnd + 1 < outRow

                If _
                    UCase$(Trim$(CStr(ws.Cells(groupEnd + 1, 1).Value))) = UCase$(Trim$(subDealer)) And _
                    UCase$(Trim$(CStr(ws.Cells(groupEnd + 1, 2).Value))) = UCase$(Trim$(storeName)) And _
                    UCase$(Trim$(CStr(ws.Cells(groupEnd + 1, 3).Value))) = UCase$(Trim$(familyName)) And _
                    UCase$(Trim$(CStr(ws.Cells(groupEnd + 1, 4).Value))) = UCase$(Trim$(issueText)) Then

                    groupEnd = groupEnd + 1

                Else

                    Exit Do

                End If

            Loop

            '-----------------------------------------------------
            ' Do not add the same SKU twice.
            '-----------------------------------------------------
            Dim checkRow As Long

            For checkRow = r To groupEnd

                existingSKU = CStr(ws.Cells(checkRow, 5).Value)

                If UCase$(Trim$(existingSKU)) = UCase$(Trim$(sku)) Then
                    Exit Sub
                End If

            Next checkRow

            '-----------------------------------------------------
            ' Insert a new physical row immediately after the
            ' existing Store + Family group.
            '-----------------------------------------------------
            ws.Rows(groupEnd + 1).Insert Shift:=xlDown

            ' Copy formatting from the first row of the group.
            ws.Rows(r).Copy
            ws.Rows(groupEnd + 1).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False

            ws.Cells(groupEnd + 1, 1).Value = subDealer
            ws.Cells(groupEnd + 1, 2).Value = storeName
            ws.Cells(groupEnd + 1, 3).Value = familyName
            ws.Cells(groupEnd + 1, 4).Value = issueText
            ws.Cells(groupEnd + 1, 5).Value = sku
            ws.Cells(groupEnd + 1, 6).Value = productName
            ws.Cells(groupEnd + 1, 7).Value = allocationQty
            ws.Cells(groupEnd + 1, 8).Value = actionText
            ws.Cells(groupEnd + 1, 9).Value = transferText

            outRow = outRow + 1

            Exit Sub

        End If

    Next r

    '=============================================================
    ' No existing group -> create first row.
    '=============================================================
    ws.Cells(outRow, 1).Value = subDealer
    ws.Cells(outRow, 2).Value = storeName
    ws.Cells(outRow, 3).Value = familyName
    ws.Cells(outRow, 4).Value = issueText
    ws.Cells(outRow, 5).Value = sku
    ws.Cells(outRow, 6).Value = productName
    ws.Cells(outRow, 7).Value = allocationQty
    ws.Cells(outRow, 8).Value = actionText
    ws.Cells(outRow, 9).Value = transferText

    outRow = outRow + 1

End Sub


'=============================================================
' APPEND LINE WITHOUT DUPLICATE
'=============================================================
Private Function AppendOutputLine( _
        ByVal existingText As String, _
        ByVal newText As String) As String

    If Len(Trim$(newText)) = 0 Then
        AppendOutputLine = existingText
        Exit Function
    End If

    If Len(Trim$(existingText)) = 0 Then
        AppendOutputLine = newText
        Exit Function
    End If

    If OutputLineExists(existingText, newText) Then
        AppendOutputLine = existingText
    Else
        AppendOutputLine = existingText & vbLf & newText
    End If

End Function


'=============================================================
' CHECK WHETHER A VALUE ALREADY EXISTS AS A LINE
'=============================================================
Private Function OutputLineExists( _
        ByVal existingText As String, _
        ByVal searchText As String) As Boolean

    Dim parts As Variant
    Dim i As Long

    If Len(Trim$(existingText)) = 0 Then
        OutputLineExists = False
        Exit Function
    End If

    parts = Split(existingText, vbLf)

    For i = LBound(parts) To UBound(parts)

        If UCase$(Trim$(CStr(parts(i)))) = _
           UCase$(Trim$(searchText)) Then

            OutputLineExists = True
            Exit Function

        End If

    Next i

    OutputLineExists = False

End Function


'=============================================================
' WRITE OUTPUT ROW
'=============================================================
Private Sub WriteOutputRow( _
        ByVal ws As Worksheet, _
        ByVal r As Long, _
        ByVal subDealer As String, _
        ByVal storeName As String, _
        ByVal familyName As String, _
        ByVal issueText As String, _
        ByVal sku As String, _
        ByVal productName As String, _
        ByVal allocationQty As Double, _
        ByVal actionText As String, _
        ByVal transferText As String)

    ws.Cells(r, 1).Value = subDealer
    ws.Cells(r, 2).Value = storeName
    ws.Cells(r, 3).Value = familyName
    ws.Cells(r, 4).Value = issueText
    ws.Cells(r, 5).Value = sku
    ws.Cells(r, 6).Value = productName
    ws.Cells(r, 7).Value = allocationQty
    ws.Cells(r, 8).Value = actionText
    ws.Cells(r, 9).Value = transferText

End Sub


'=============================================================
' FORMAT WHATSAPP SUMMARY
'=============================================================
Private Sub FormatWhatsAppSummary( _
        ByVal ws As Worksheet, _
        ByVal lastRow As Long)

    Dim rng As Range
    Dim header As Range
    Dim r As Long
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim currentKey As String
    Dim nextKey As String

    If lastRow < 1 Then Exit Sub

    '=============================================================
    ' FIRST: UNMERGE ANY OLD CELLS
    '=============================================================
    ws.Range("A1:I" & lastRow).UnMerge

    Set rng = ws.Range("A1:I" & lastRow)
    Set header = ws.Range("A1:I1")

    rng.Font.Name = "Calibri"
    rng.Font.Size = 11
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.WrapText = True

    With header
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    If lastRow >= 2 Then

        ws.Range("A2:G" & lastRow).Interior.Color = RGB(221, 235, 247)

        With ws.Range("H2:H" & lastRow)
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
            .Font.Bold = True
        End With

        For r = 2 To lastRow

            If InStr(1, UCase$(CStr(ws.Cells(r, 9).Value)), "NO QTY", vbTextCompare) > 0 Or _
               InStr(1, UCase$(CStr(ws.Cells(r, 9).Value)), "NOT FOUND", vbTextCompare) > 0 Then

                ws.Cells(r, 9).Interior.Color = RGB(255, 235, 235)
                ws.Cells(r, 9).Font.Color = RGB(192, 0, 0)
                ws.Cells(r, 9).Font.Bold = True

            Else

                ws.Cells(r, 9).Interior.Color = RGB(221, 235, 247)
                ws.Cells(r, 9).Font.Color = RGB(0, 97, 0)
                ws.Cells(r, 9).Font.Bold = True

            End If

        Next r

    End If

    '=============================================================
    ' MERGE GROUP COLUMNS
    '
    ' Same Store + Family + Issue:
    '
    ' A:D = merged vertically
    ' H   = merged vertically
    ' I   = merged vertically
    '
    ' E:F:G remain separate for each SKU.
    '=============================================================

    groupStart = 2

    Do While groupStart <= lastRow

        currentKey = _
            UCase$(Trim$(CStr(ws.Cells(groupStart, 1).Value))) & "||" & _
            UCase$(Trim$(CStr(ws.Cells(groupStart, 2).Value))) & "||" & _
            UCase$(Trim$(CStr(ws.Cells(groupStart, 3).Value))) & "||" & _
            UCase$(Trim$(CStr(ws.Cells(groupStart, 4).Value)))

        groupEnd = groupStart

        Do While groupEnd < lastRow

            nextKey = _
                UCase$(Trim$(CStr(ws.Cells(groupEnd + 1, 1).Value))) & "||" & _
                UCase$(Trim$(CStr(ws.Cells(groupEnd + 1, 2).Value))) & "||" & _
                UCase$(Trim$(CStr(ws.Cells(groupEnd + 1, 3).Value))) & "||" & _
                UCase$(Trim$(CStr(ws.Cells(groupEnd + 1, 4).Value)))

            If nextKey = currentKey Then
                groupEnd = groupEnd + 1
            Else
                Exit Do
            End If

        Loop

        '---------------------------------------------------------
        ' Merge only when this family has more than one SKU row.
        '---------------------------------------------------------
        If groupEnd > groupStart Then

            ws.Range(ws.Cells(groupStart, 1), ws.Cells(groupEnd, 1)).Merge
            ws.Range(ws.Cells(groupStart, 2), ws.Cells(groupEnd, 2)).Merge
            ws.Range(ws.Cells(groupStart, 3), ws.Cells(groupEnd, 3)).Merge
            ws.Range(ws.Cells(groupStart, 4), ws.Cells(groupEnd, 4)).Merge
            ws.Range(ws.Cells(groupStart, 8), ws.Cells(groupEnd, 8)).Merge
            ws.Range(ws.Cells(groupStart, 9), ws.Cells(groupEnd, 9)).Merge

            With ws.Range(ws.Cells(groupStart, 1), ws.Cells(groupEnd, 4))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With

            With ws.Range(ws.Cells(groupStart, 8), ws.Cells(groupEnd, 9))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
            End With

        End If

        groupStart = groupEnd + 1

    Loop

    '=============================================================
    ' BORDERS
    '=============================================================

    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(150, 150, 150)
    End With

    '=============================================================
    ' FILTER
    '=============================================================

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:I" & lastRow).AutoFilter

    '=============================================================
    ' WIDTHS
    '=============================================================

    ws.Columns("A:I").AutoFit

    If ws.Columns("A").ColumnWidth > 25 Then ws.Columns("A").ColumnWidth = 25
    If ws.Columns("B").ColumnWidth > 32 Then ws.Columns("B").ColumnWidth = 32
    If ws.Columns("C").ColumnWidth > 28 Then ws.Columns("C").ColumnWidth = 28
    If ws.Columns("D").ColumnWidth > 28 Then ws.Columns("D").ColumnWidth = 28
    If ws.Columns("E").ColumnWidth > 18 Then ws.Columns("E").ColumnWidth = 18
    If ws.Columns("F").ColumnWidth > 35 Then ws.Columns("F").ColumnWidth = 35
    If ws.Columns("G").ColumnWidth > 18 Then ws.Columns("G").ColumnWidth = 18
    If ws.Columns("H").ColumnWidth > 45 Then ws.Columns("H").ColumnWidth = 45
    If ws.Columns("I").ColumnWidth > 75 Then ws.Columns("I").ColumnWidth = 75

    ws.Columns("A:I").WrapText = True

    ws.Rows("1:" & lastRow).AutoFit
    ws.Rows(1).RowHeight = 30

    '=============================================================
    ' FREEZE HEADER
    '=============================================================

    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

End Sub


'=============================================================
' DISPLAY HELPERS
'=============================================================
Private Function FormatQtyStore( _
        ByVal qty As Double, _
        ByVal locID As String, _
        ByVal shortName As String, _
        ByVal sku As String, _
        ByVal productName As String) As String

    Dim s As String

    s = FormatNumberSimple(qty) & " Qty " & CleanSKU(sku)

    If Len(productName) > 0 Then
        s = s & " - " & productName
    End If

    s = s & " available at " & NormalizeLocationID(locID)

    If Len(shortName) > 0 Then
        s = s & " - " & shortName
    End If

    FormatQtyStore = s

End Function


Private Function FormatNumberSimple(ByVal qty As Double) As String

    If qty = Int(qty) Then
        FormatNumberSimple = CStr(CLng(qty))
    Else
        FormatNumberSimple = Format$(qty, "0.##")
    End If

End Function


Private Function BuildShortStoreName( _
        ByVal locID As String, _
        ByVal shortName As String, _
        ByVal locationName As String) As String

    If Len(shortName) > 0 Then
        BuildShortStoreName = NormalizeLocationID(locID) & " - " & shortName
    ElseIf Len(locationName) > 0 Then
        BuildShortStoreName = BuildLocationDisplayName(locationName)
    Else
        BuildShortStoreName = NormalizeLocationID(locID)
    End If

End Function


Private Function BuildLocationDisplayName(ByVal locationName As String) As String

    Dim locID As String
    Dim storeName As String
    Dim p As Long

    locID = ExtractLocationID(locationName)
    p = InStr(1, locationName, ":", vbTextCompare)

    If p > 0 Then

        storeName = Trim$(Mid$(locationName, p + 1))

        If Len(locID) > 0 Then
            BuildLocationDisplayName = locID & " - " & ReplaceStorePrefix(storeName)
        Else
            BuildLocationDisplayName = storeName
        End If

    Else
        BuildLocationDisplayName = locationName
    End If

End Function


Private Function ReplaceStorePrefix(ByVal s As String) As String

    s = Trim$(s)

    If LCase$(Left$(s, 14)) = "spacetel llc -" Then
        ReplaceStorePrefix = Trim$(Mid$(s, 15))
    Else
        ReplaceStorePrefix = s
    End If

End Function


'=============================================================
' STOCK KEY HELPERS
'=============================================================
Private Function GetKeyLocationID(ByVal key As String) As String

    Dim p As Long

    p = InStr(1, key, "|", vbTextCompare)

    If p > 0 Then
        GetKeyLocationID = Left$(key, p - 1)
    Else
        GetKeyLocationID = key
    End If

End Function


Private Function GetKeySKU(ByVal key As String) As String

    Dim p As Long

    p = InStr(1, key, "|", vbTextCompare)

    If p > 0 Then
        GetKeySKU = Mid$(key, p + 1)
    Else
        GetKeySKU = ""
    End If

End Function


'=============================================================
' LOCATION HELPERS
'=============================================================
Private Function ExtractLocationID(ByVal value As Variant) As String

    Dim s As String
    Dim firstPart As String
    Dim p As Long

    s = CleanText(value)

    If Len(s) = 0 Then
        ExtractLocationID = ""
        Exit Function
    End If

    ' 138: Store Name
    p = InStr(1, s, ":", vbTextCompare)

    If p > 0 Then

        firstPart = Trim$(Left$(s, p - 1))

        If IsNumeric(firstPart) Then
            ExtractLocationID = NormalizeLocationID(firstPart)
            Exit Function
        End If

    End If

    ' 138 - Store Name
    p = InStr(1, s, " - ", vbTextCompare)

    If p > 0 Then

        firstPart = Trim$(Left$(s, p - 1))

        If IsNumeric(firstPart) Then
            ExtractLocationID = NormalizeLocationID(firstPart)
            Exit Function
        End If

    End If

    ' 138 Store Name
    p = InStr(1, s, " ", vbTextCompare)

    If p > 0 Then
        firstPart = Trim$(Left$(s, p - 1))
    Else
        firstPart = s
    End If

    If IsNumeric(firstPart) Then
        ExtractLocationID = NormalizeLocationID(firstPart)
    Else
        ExtractLocationID = ""
    End If

End Function


' Alias retained so old/new code cannot produce "Sub or Function not defined".
Private Function ExtractTransferLocationID(ByVal value As Variant) As String
    ExtractTransferLocationID = ExtractLocationID(value)
End Function


Private Function NormalizeLocationID(ByVal value As Variant) As String

    Dim s As String

    s = CleanText(value)

    If Len(s) = 0 Then
        NormalizeLocationID = ""
        Exit Function
    End If

    If IsNumeric(s) Then
        NormalizeLocationID = CStr(CLng(Val(s)))
    Else
        NormalizeLocationID = UCase$(s)
    End If

End Function


Private Function NormalizeTransferLocationID(ByVal value As Variant) As String
    NormalizeTransferLocationID = NormalizeLocationID(value)
End Function


'=============================================================
' FAMILY HELPERS
'=============================================================
Private Function NormalizeFamily(ByVal value As Variant) As String
    NormalizeFamily = LCase$(Trim$(RemoveDemoSuffix(CleanText(value))))
End Function


Private Function RemoveDemoSuffix(ByVal value As Variant) As String

    Dim s As String

    s = CleanText(value)

    If Len(s) >= 7 Then
        If LCase$(Right$(s, 7)) = " (demo)" Then
            s = Left$(s, Len(s) - 7)
        End If
    End If

    RemoveDemoSuffix = Trim$(s)

End Function


'=============================================================
' SKU / TEXT / NUMBER HELPERS
'=============================================================
Private Function CleanSKU(ByVal value As Variant) As String

    Dim s As String

    s = CleanText(value)
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")

    CleanSKU = UCase$(s)

End Function


Private Function CleanTransferText(ByVal value As Variant) As String
    CleanTransferText = CleanText(value)
End Function


Private Function CleanTransferSKU(ByVal value As Variant) As String
    CleanTransferSKU = CleanSKU(value)
End Function


Private Function NumberValue(ByVal value As Variant) As Double

    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
        NumberValue = 0
    ElseIf IsNumeric(value) Then
        NumberValue = CDbl(value)
    Else
        NumberValue = 0
    End If

End Function


Private Function TransferNumberValue(ByVal value As Variant) As Double
    TransferNumberValue = NumberValue(value)
End Function


Private Function FormatTransferQty(ByVal qty As Double) As String
    FormatTransferQty = FormatNumberSimple(qty)
End Function


Private Function CleanText(ByVal v As Variant) As String

    Dim s As String

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        CleanText = ""
        Exit Function
    End If

    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanText = s

End Function


'=============================================================
' GENERAL HELPERS
'=============================================================
Private Function GetSheet( _
        ByVal wb As Workbook, _
        ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

End Function


Private Function LastUsedRow(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    On Error Resume Next

    Set lastCell = ws.Cells.Find( _
            What:="*", _
            After:=ws.Cells(1, 1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)

    On Error GoTo 0

    If lastCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = lastCell.Row
    End If

End Function
