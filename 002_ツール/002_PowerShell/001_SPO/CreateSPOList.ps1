Import-Module PnP.PowerShell
Import-Module ImportExcel

# 自分で設定する。
$siteUrl = "$SPOサイトURL設定$"
$clientId = "Azureでアプリ登録したクライアントID"
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Interactive

$excelPath = "取り込みエクセルファイル名（拡張子付）"
$logFilePath = "ログファイル名（拡張子付）"

# ログ書き込み関数定義
function Write-Log {
    param (
        [string]$message,
        [string]$number,
        [string]$listName,
        [string]$columnName
    )
    $currentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $numberText = $number ? "No:「$($number)」" : ""
    $listNameText = $listName ? "List/Library:「$($listName)」" : ""
    $columnNameText = $columnName ? "列:「$($columnName)」" : ""
    $logMessage = "$($currentDateTime): $($message) $($numberText) $($listNameText) $($columnNameText)"

    Add-Content -Path $logFilePath -Value $logMessage
    Write-Host $logMessage
}

$successCount = 0
$errorCount = 0
$skipCount = 0
$number = 0
$listName = ""
$columnName = ""

try {
    Write-Log "List/Libraryの作成処理を開始"
    # シート名ListsからNo、リスト名、内部リスト名、リストタイプ、説明、バージョン、処理スキップされるかを取得
    $lists = Import-Excel -Path $excelPath -WorksheetName "Lists"

    # SPOリスト作成
    foreach ($list in $lists) {
        $number = $list.No
        $listName = $list.ListName
        $listInternal = $list.ListInternal   
        $listType = $list.ListType
        $description = $list.Description
        $majorVersions = $list.MajorVersions
        $skip = $list.Skip?.ToString() -eq "True" ? $true : $false

        # バリデーションチェック
        if ($skip) {
            Write-Log "Skip List/Libraryの作成をスキップしました" $number $listName
            $skipCount++
            continue
        }

        if (($null -eq $number) -or ($null -eq $listName) -or ($null -eq $listInternal) -or ($null -eq $listType)) {
            throw "「No」「ListName」「ListInternal」「ListType」は必須項目です"
        }

        if ($listType -ne "Library" -and $listType -ne "List") {
            throw "ListTypeは「Library」または「List」である必要があります"
        }

        if ($null -ne $majorVersions) {
            if (-not ($majorVersions -match "^\d+$")) {
                throw "MajorVersionsは整数である必要があります"
            }
            elseif ($majorVersions -lt 1 -or $majorVersions -gt 50000) {
                throw "MajorVersionsは「1」から「50000」の範囲内である必要があります"
            }
        }
      
        if ((Get-PnPList $listName -ErrorAction SilentlyContinue)) {
            Write-Log "Skip List/Libraryは既に存在します" $number $listName
            $skipCount++
            continue
        }

        if ($listType -eq "Library") {
            New-PnPList -Title $listName -Url $listInternal -Template DocumentLibrary -OnQuickLaunch:$false 
            Write-Log "Success DocumentLibraryを作成しました" $number $listName
            $successCount++
        }
        else {
            # Urlに「Lists/」が自動で付くよう、-Urlを指定せずlistInternalで登録してから-TitleをlistNameに更新する
            # listInternalに「Lists/」が含まれてしまい、Excel上のlistInternalと矛盾が生じることになるため、以降の処理ではlistNameでリストを特定するようにする
            New-PnPList -Title $listInternal -Template GenericList -OnQuickLaunch:$false 
            Set-PnPList $listInternal -Title $listName
            Write-Log "Success Listを作成しました" $number $listName
            $successCount++
        }

        if ($null -ne $description) {
            Set-PnPList -Identity $listName -Description $description
            Write-Log "Success Descriptionを設定しました" $number $listName
            $successCount++
        }

        if (($null -eq $majorVersions) -or ($majorVersions -eq "0")) {
            Set-PnPList -Identity $listName -EnableVersioning:$false            
        }
        else {
            Set-PnPList -Identity $listName -EnableVersioning:$true -MajorVersions $majorVersions
        }  

        Write-Log "Success MajorVersionsの設定をしました" $number $listName
        $successCount++

        Set-PnPField -List $listName -Identity "Title" -Values @{Required = $false }
        Write-Log "Success Title列を必須から外しました" $number $listName
        $successCount++
    }

    $number = 0
    $listName = ""
    Write-Log "列の作成処理を開始"

    $columns = Import-Excel -Path $excelPath -WorksheetName "Columns"

    # SPOリストの列定義作成
    foreach ($column in $columns) {
        # シート名ColumnsからNo、リスト名、列名、内部列名、列タイプ、説明、デフォルト値、選択値、複数選択可能か、ルックアップリスト名、ルックアップ内部列名、親追加親列内部名、必須項目か、ユニーク項目か、インデックス列か、処理スキップされるかを取得
        $number = $column.No
        $listName = $column.ListName
        $columnName = $column.ColumnName
        $columnInternal = $column.ColumnInternal
        $columnType = $column.ColumnType
        $description = $column.Description
        $defaultValue = $column.DefaultValue
        $choices = $column.Choices
        $multiple = $column.Multiple?.ToString() -eq "True" ? $true : $false
        $lookupListName = $column.LookupListName
        $lookupFieldInternal = $column.LookupFieldInternal
        $parentColumnInternal = $column.ParentColumnInternal
        $required = $column.Required?.ToString() -eq "True" ? $true : $false
        $unique = $column.Unique?.ToString() -eq "True" ? $true : $false
        $indexed = $column.Indexed?.ToString() -eq "True" ? $true : $false
        $skip = $column.Skip?.ToString() -eq "True" ? $true : $false

        # バリデーションチェック
        if ($skip) {
            Write-Log "Skip 列の作成をスキップしました" $number $listName $columnName
            $skipCount++
            continue
        }

        if (-not (Get-PnPList $listName -ErrorAction SilentlyContinue)) {
            throw "List/Libraryが存在しません"
        }

        if (Get-PnPField -List $listName -Identity $columnInternal -ErrorAction SilentlyContinue) {
            Write-Log "Skip 列が既に存在します" $number $listName $columnName
            $skipCount++
            continue
        }
        
        if (($null -eq $number) -or ($null -eq $listName) -or ($null -eq $columnName) -or ($null -eq $columnInternal) -or ($null -eq $columnType)) {
            throw "「No」「ListName」「ColumnName」「ColumnInternal」「ColumnType」は必須項目です"
        }

        If ((-not $multiple) -and ($column.Multiple -ne $false) -and ($null -ne $column.Multiple)) {
            throw "Multipleに「TRUE」または「FALSE」以外の値が指定されています"
        }

        If ((-not $required) -and ($column.Required -ne $false) -and ($null -ne $column.Required)) {
            throw "Requiredに「TRUE」または「FALSE」以外の値が指定されています"
        }

        If ((-not $unique) -and ($column.Unique -ne $false) -and ($null -ne $column.Unique)) {
            throw "Uniqueに「TRUE」または「FALSE」以外の値が指定されています"
        }

        If ((-not $indexed) -and ($column.Indexed -ne $false) -and ($null -ne $column.Indexed)) {
            throw "Indexedに「TRUE」または「FALSE」以外の値が指定されています"
        }

        if ((-not (@("Choice", "Lookup", "User", "UserGroup") -contains $columnType)) -and $multiple) {
            throw "ColumnType「$($columnType)」はMultipleオプションに対応していません"
        }

        if ((@("Note", "Boolean", "LookupAdd", "URL", "Thumbnail") -contains $columnType) -and $unique) {
            throw "ColumnType「$($columnType)」はUniqueオプションに対応していません"
        }

        if ((@("Note", "LookupAdd", "URL", "Thumbnail") -contains $columnType) -and $indexed) {
            throw "ColumnType「$($columnType)」はIndexedオプションに対応していません"
        }

        if ($unique -and $multiple) {
            throw "uniqueオプションはMultipleオプションと併用できません"
        }

        if ($indexed -and $multiple) {
            throw "indexedオプションはMultipleオプションと併用できません"
        }

        if ($unique -and (-not $indexed)) {
            throw "uniqueオプションを付ける場合はindexedオプションが必須です"
        }

        if (($null -ne $choices) -and ($columnType -ne "Choice")) {
            throw "Choicesが指定できるのはColumnType「Choice」のみです"
        }

        if (($null -ne $defaultValue) -and (@("Boolean", "Lookup", "LookupAdd", "User", "UserGroup", "URL", "Thumbnail") -contains $columnType)) {
            throw "ColumnType「Boolean」「Lookup」「Date」「LookupAdd」「User」「UserGroup」「URL」「Thumbnail」にはDefaultValueを指定できません"
        }

        if (($null -ne $lookupListName) -and (-not (@("Lookup", "LookupAdd") -contains $columnType))) {
            throw "LookupListNameが指定できるのはColumnType「Lookup」「LookupAdd」のみです"
        }

        if (($null -ne $lookupFieldInternal) -and (-not (@("Lookup", "LookupAdd") -contains $columnType))) {
            throw "lookupFieldInternalが指定できるのはColumnType「Lookup」「LookupAdd」のみです"
        }

        if (($null -ne $parentColumnInternal) -and ($columnType -ne "LookupAdd")) {
            throw "parentColumnInternalが指定できるのはColumnType「LookupAdd」のみです"
        }

        if (@("Lookup", "LookupAdd") -contains $columnType) {
            $lookupField = Get-PnPField -List $lookupListName -Identity $lookupFieldInternal -ErrorAction SilentlyContinue

            if (-not ($lookupField)) {
                throw "LookupFieldInternalに指定された列が存在しません"
            }

            if (-not (@("Text", "Number", "DateTime", "Counter") -contains $lookupField.TypeAsString)) {
                throw "LookupFieldInternalに指定するの列の種類は「Text」「Number」「Date」「DateTime」のいずれかである必要があります"
            }
        }

        if (($columnType -eq "Number" -or $columnType -eq "Currency") -and ($null -ne $defaultValue)) {
            try {
                [int]$defaultValue
            }
            catch {
                throw "ColumnType「Number」または「Currency」のDefaultValueは数値である必要があります"
            }
        }

        if (($columnType -eq "Date" -or $columnType -eq "DateTime") -and ($null -ne $defaultValue)) {
            try {
                [DateTime]::Parse($defaultValue)
            }
            catch {
                throw "ColumnType「Date」または「DateTime」のDefaultValueは日付形式である必要があります"
            }
        }
                    
        switch ($columnType) {
            "Choice" {

                if ($null -eq $choices) {                 
                    throw "Choice列の「Choices」は必須項目です"                    
                }

                $choicesArray = $choices -split ','

                if (($null -ne $defaultValue) -and (-not ($choicesArray -contains $defaultValue))) {                 
                    throw "DefaultValueに指定された値がChoicesに含まれていません"                    
                }

                $fieldType = if ($multiple) { "MultiChoice" } else { "Choice" }
                Add-PnPField -List $listName -DisplayName $columnName -InternalName $columnInternal -Type $fieldType -Choices $choicesArray -Required:$required
                Write-Log "Success Choice列を追加しました" $number $listName $columnName
                $successCount++
            }
            "Lookup" {

                $pnpField = Add-PnPField -List $listName -DisplayName $columnName -InternalName $columnInternal -type Lookup -Required:$required

                if ($multiple) {
                    [XML]$schemaXml = $pnpField.SchemaXml
                    $schemaXml.Field.SetAttribute("Mult", "TRUE")
                    $schemaXml.Field.SetAttribute("ShowField", $lookupFieldInternal.ToString())
                    $schemaXml.Field.SetAttribute("List", (Get-PnPList $lookupListName).Id.ToString())
                    $outerXML = $schemaXml.OuterXml.Replace('Field Type="Lookup"', 'Field Type="LookupMulti"')
                    Set-PnPField -List $listName -Identity $pnpField.Id -Values @{SchemaXml = $outerXML; LookupList = (Get-PnPList $lookupListName).Id.ToString(); LookupField = $lookupFieldInternal.ToString() } -UpdateExistingLists

                }
                else {
                    Set-PnPField -List $listName -Identity $pnpField.Id -Values @{LookupList = (Get-PnPList $lookupListName).Id.ToString(); LookupField = $lookupFieldInternal.ToString() } -UpdateExistingLists
                }

                Write-Log "Success Lookup列を追加しました" $number $listName $columnName
                $successCount++
            }
            "LookupAdd" {

                if ($required) {
                    throw "LookupAdd列はRequiredオプションを設定できません"
                }

                if ($null -eq $parentColumnInternal) {
                    throw "LookupAdd列の「ParentColumnInternal」は必須項目です"
                }

                $parentColumn = Get-PnPField -List $listName -Identity $parentColumnInternal -ErrorAction SilentlyContinue

                if (-not ($parentColumn)) {
                    throw "ParentColumnInternalに指定された列が存在しません"
                }

                $fieldTypeValue = $parentColumn.TypeAsString
                $multValue = $parentColumn.AllowMultipleValues ? "TRUE" : "FALSE"

                $xml = "<Field Type='$($fieldTypeValue)'                
                        ShowField='$($lookupFieldInternal)'
                        DisplayName='$($columnName)'
                        Name='$($columnInternal)'
                        StaticName='$($columnInternal)'
                        List='$((Get-PnPList $lookupListName).id)'
                        WebId='$((Get-PnPWeb).id)'
                        FieldRef='$(($parentColumn).id)'
                        ReadOnly='TRUE'
                        Mult='$($multValue)'
                        UnlimitedLengthInDocumentLibrary='FALSE'
                        SourceID='$((Get-PnPList $listName).id)' />"

                Add-PnPFieldFromXml -list $listName -FieldXml $xml
                Write-Log "Success LookupAdd列を追加しました" $number $listName $columnName
                $successCount++
            }                
            "Date" {
                $pnpField = Add-PnPField -List $listName -DisplayName $columnName -InternalName $columnInternal -Type DateTime -Required:$required
                [XML]$schemaXml = $pnpField.SchemaXml
                $schemaXml.Field.SetAttribute("Format", "DateOnly")
                Set-PnPField -List $listName -Identity $pnpField.Id -Values @{SchemaXml = $schemaXml.OuterXml } -UpdateExistingLists
                Write-Log "Success Date列を追加しました" $number $listName $columnName
                $successCount++
            }
            "DateTime" {
                Add-PnPField -List $listName -DisplayName $columnName -InternalName $columnInternal -Type DateTime -Required:$required
                Write-Log "Success DateTime列を追加しました" $number $listName $columnName
                $successCount++
            }
            "User" {
                $pnpField = Add-PnPField -List $listName -DisplayName $columnName -InternalName $columnInternal -Type User -Required:$required                    
                [XML]$schemaXml = $pnpField.SchemaXml                            

                if ($multiple) {
                    $schemaXml.Field.SetAttribute("UserSelectionMode", "0")  
                    $schemaXml.Field.SetAttribute("Mult", "TRUE")
                    $outerXML = $schemaXml.OuterXml.Replace('Field Type="User"', 'Field Type="UserMulti"')
                    Set-PnPField -List $listName -Identity $pnpField.Id -Values @{SchemaXml = $outerXML } -UpdateExistingLists
                }
                else {
                    $schemaXml.Field.SetAttribute("UserSelectionMode", "0")  
                    Set-PnPField -List $listName -Identity $pnpField.Id -Values @{SchemaXml = $schemaXml.OuterXml } -UpdateExistingLists  
                }
                               
                Write-Log "Success User列を追加しました" $number $listName $columnName 
                $successCount++                   
            }
            "UserGroup" {
                $pnpField = Add-PnPField -List $listName -DisplayName $columnName -InternalName $columnInternal -Type User -Required:$required

                if ($multiple) {
                    [XML]$schemaXml = $pnpField.SchemaXml
                    $schemaXml.Field.SetAttribute("Mult", "TRUE")
                    $outerXML = $schemaXml.OuterXml.Replace('Field Type="User"', 'Field Type="UserMulti"')
                    Set-PnPField -List $listName -Identity $pnpField.Id -Values @{SchemaXml = $outerXML } -UpdateExistingLists
                }
                    
                Write-Log "Success UserGroup列を追加しました" $number $listName $columnName   
                $successCount++                 
            }
            Default {
                Add-PnPField -List $listName -DisplayName $columnName -InternalName $columnInternal -Type $columnType -Required:$required
                Write-Log "Success $($columnType)列を追加しました" $number $listName $columnName
                $successCount++
            }
        }

        if ($null -ne $description) {
            Set-PnPField -List $listName -Identity $columnInternal -Values @{"Description" = $description }
            Write-Log "Success Descriptionを設定しました" $number $listName $columnName
            $successCount++
        }

        if ($null -ne $defaultValue) {
            Set-PnPField -List $listName -Identity $columnInternal -Values @{"DefaultValue" = $defaultValue.ToString() }
            Write-Log "Success DefaultValueを設定しました" $number $listName $columnName
            $successCount++
        }

        if ($indexed) {
            Set-PnpField -List $listName -Identity $columnInternal -Values @{"Indexed" = $true }
            Write-Log "Success Indexedを設定しました" $number $listName $columnName
            $successCount++
        }

        if ($unique) {
            Set-PnpField -List $listName -Identity $columnInternal -Values @{"EnforceUniqueValues" = $true }
            Write-Log "Success Uniqueを設定しました" $number $listName $columnName
            $successCount++
        }
    }

    $number = 0
    $listName = ""
    $columnName = ""
    Write-Log "DefaultViewの更新処理を開始"

    foreach ($list in $lists) {
        # シートListsからデフォルトビュー列を取得
        $number = $list.No
        $listName = $list.ListName
        $listInternal = $list.ListInternal
        $defaultViewColumns = $list.DefaultViewColumns
        $columnName = ""
        $skip = $list.Skip?.ToString() -eq "True" ? $true : $false

        # デフォルトビューを作成
        if ($skip) {
            Write-Log "Skip DefaultViewの更新をスキップしました" $number $listName
            $skipCount++
            continue
        }

        if (-not (Get-PnPList $listName -ErrorAction SilentlyContinue)) {
            throw "List/Libraryが存在しません"
        }

        if ($null -eq $defaultViewColumns) { continue }

        $columnArray = $defaultViewColumns -split ','

        foreach ($column in $columnArray ) {
            $columnName = $column

            if (-not ( Get-PnPField -List $listName -Identity $column -ErrorAction SilentlyContinue)) {
                throw "DefaultViewColumnsに指定された列が存在しません"
            }
        }

        $views = Get-PnPView -List $listName   

        foreach ($view in $views) {
            if (-not $view.DefaultView) { continue }                             
            Set-PnPView -List $listName -Identity $view.Id -Fields $columnArray
            Write-Log "Success DefaultViewを更新しました" $number $listName
            $successCount++
        }
    }
}
catch {
    Write-Log "Error $($_)" $number $listName $columnName
    $errorCount++
}
finally {
    Write-Log "処理が終了しました 成功件数:「$($successCount)」スキップ件数:「$($skipCount)」エラー件数:「$($errorCount)」"
}

