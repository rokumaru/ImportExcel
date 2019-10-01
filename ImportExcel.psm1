#region import everything we need
Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"
. $PSScriptRoot\AddConditionalFormatting.ps1
. $PSScriptRoot\AddDataValidation.ps1
. $PSScriptRoot\Charting.ps1
. $PSScriptRoot\ColorCompletion.ps1
. $PSScriptRoot\Compare-WorkSheet.ps1
. $PSScriptRoot\ConvertExcelToImageFile.ps1
. $PSScriptRoot\ConvertFromExcelData.ps1
. $PSScriptRoot\ConvertFromExcelToSQLInsert.ps1
. $PSScriptRoot\ConvertToExcelXlsx.ps1
. $PSScriptRoot\Copy-ExcelWorkSheet.ps1
. $PSScriptRoot\Export-Excel.ps1
. $PSScriptRoot\Export-ExcelSheet.ps1
. $PSScriptRoot\Export-StocksToExcel.ps1
. $PSScriptRoot\Get-ExcelColumnName.ps1
. $PSScriptRoot\Get-ExcelSheetInfo.ps1
. $PSScriptRoot\Get-ExcelWorkbookInfo.ps1
. $PSScriptRoot\Get-HtmlTable.ps1
. $PSScriptRoot\Get-Range.ps1
. $PSScriptRoot\Get-XYRange.ps1
. $PSScriptRoot\Import-Html.ps1
. $PSScriptRoot\InferData.ps1
. $PSScriptRoot\Invoke-Sum.ps1
. $PSScriptRoot\Join-Worksheet.ps1
. $PSScriptRoot\Merge-Worksheet.ps1
. $PSScriptRoot\New-ConditionalFormattingIconSet.ps1
. $PSScriptRoot\New-ConditionalText.ps1
. $PSScriptRoot\New-ExcelChart.ps1
. $PSScriptRoot\New-PSItem.ps1
. $PSScriptRoot\Open-ExcelPackage.ps1
. $PSScriptRoot\Pivot.ps1
. $PSScriptRoot\PivotTable.ps1
#. $PSScriptRoot\Plot.ps1
. $PSScriptRoot\RemoveWorksheet.ps1
. $PSScriptRoot\Send-SQLDataToExcel.ps1
. $PSScriptRoot\Set-CellStyle.ps1
. $PSScriptRoot\Set-Column.ps1
. $PSScriptRoot\Set-Row.ps1
. $PSScriptRoot\Set-WorkSheetProtection.ps1
. $PSScriptRoot\SetFormat.ps1
. $PSScriptRoot\TrackingUtils.ps1
. $PSScriptRoot\Update-FirstObjectProperties.ps1


New-Alias -Name Use-ExcelData -Value "ConvertFrom-ExcelData" -Force

if ($PSVersionTable.PSVersion.Major -ge 5) {
    . $PSScriptRoot\Plot.ps1

    Function New-Plot {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'New-Plot does not change system state')]
        Param()

        [PSPlot]::new()
    }

}
else {
    Write-Warning 'PowerShell 5 is required for plot.ps1'
    Write-Warning 'PowerShell Excel is ready, except for that functionality'
}
if ($IsLinux -or $IsMacOS) {
    $ExcelPackage = [OfficeOpenXml.ExcelPackage]::new()
    $Cells = ($ExcelPackage | Add-WorkSheet).Cells['A1']
    $Cells.Value = 'Test'
    try {
        $Cells.AutoFitColumns()
    }
    catch {
        if ($IsLinux) {
            Write-Warning -Message 'ImportExcel Module Cannot Autosize. Please run the following command to install dependencies: "sudo apt-get install -y --no-install-recommends libgdiplus libc6-dev"'
        }
        if ($IsMacOS) {
            Write-Warning -Message 'ImportExcel Module Cannot Autosize. Please run the following command to install dependencies: "brew install mono-libgdiplus"'
        }
    }
    finally {
        $ExcelPackage | Close-ExcelPackage -NoSave
    }
}
#endregion
function Import-Excel {
    <#
   .SYNOPSIS
       Create custom objects from the rows in an Excel worksheet.

   .DESCRIPTION
       The Import-Excel cmdlet creates custom objects from the rows in an Excel worksheet. Each row represents one object. All of this is possible without installing Microsoft Excel and by using the .NET library ‘EPPLus.dll’.

       By default, the property names of the objects are retrieved from the column headers. Because an object cannot have a blank property name, only columns with column headers will be imported.

       If the default behavior is not desired and you want to import the complete worksheet ‘as is’, the parameter ‘-NoHeader’ can be used. In case you want to provide your own property names, you can use the parameter ‘-HeaderName’.

   .PARAMETER Path
       Specifies the path to the Excel file.
   .PARAMETER ExcelPackage
       Instead of specifying a path provides an Excel Package object (from Open-ExcelPackage)
       Using this avoids re-reading the whole file when importing multiple parts of it.
       To allow multiple read operations Import-Excel does NOT close the package, and you should use
       Close-ExcelPackage -noSave to close it.
   .PARAMETER WorksheetName
       Specifies the name of the worksheet in the Excel workbook to import. By default, if no name is provided, the first worksheet will be imported.

   .PARAMETER DataOnly
       Import only rows and columns that contain data, empty rows and empty columns are not imported.

   .PARAMETER HeaderOnly
       Import only columns that contain header data, empty headers are not imported.

   .PARAMETER HeaderName
       Specifies custom property names to use, instead of the values defined in the column headers of the TopRow.
       If you provide fewer header names than there are columns of data in the worksheet, then data will only be imported from that number of columns - the others will be ignored.
       If you provide more header names than there are columns of data in the worksheet, it will result in blank properties being added to the objects returned.

   .PARAMETER NoHeader
       Automatically generate property names (P1, P2, P3, ..) instead of the ones defined in the column headers of the TopRow.
       This switch is best used when you want to import the complete worksheet ‘as is’ and are not concerned with the property names.

   .PARAMETER StartRow
       The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row.
       When the parameters ‘-NoHeader’ and ‘-HeaderName’ are not provided, this row will contain the column headers that will be used as property names. When one of both parameters are provided, the property names are automatically created and this row will be treated as a regular row containing data.

   .PARAMETER EndRow
       By default all rows up to the last cell in the sheet will be imported. If specified, import stops at this row.

   .PARAMETER StartColumn
        The number of the first column to read data from (1 by default).

   .PARAMETER EndColumn
        By default the import reads up to the last populated column, -EndColumn tells the import to stop at an earlier number.

   .PARAMETER IncludeRow
       

   .PARAMETER IncludeColumn
       

   .PARAMETER Password
       Accepts a string that will be used to open a password protected Excel file.

   .EXAMPLE
       Import data from an Excel worksheet. One object is created for each row. If the headers are duplicated or do not exist, the sequence number is automatically added.

       ----------------------------------------------
       | File: Movies.xlsx     -      Sheet: Actors |
       ----------------------------------------------
       |           A           B            C       |
       |1     First Name                 Address    |
       |2     Chuck         Norris       California |
       |3     Jean-Claude   Vandamme     Brussels   |
       ----------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors

       First Name: Chuck
       #1        : Norris
       Address   : California

       First Name: Jean-Claude
       #1        : Vandamme
       Address   : Brussels

   .EXAMPLE
       Ignore columns without headers.

       ----------------------------------------------
       | File: Movies.xlsx     -      Sheet: Actors |
       ----------------------------------------------
       |           A           B            C       |
       |1     First Name                 Address    |
       |2     Chuck         Norris       California |
       |3     Jean-Claude   Vandamme     Brussels   |
       ----------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors -HeaderOnly

       First Name: Chuck
       Address   : California

       First Name: Jean-Claude
       Address   : Brussels

   .EXAMPLE
       Import the complete Excel worksheet ‘as is’ by using the ‘-NoHeader’ switch. One object is created for each row. The property names of the objects will be automatically generated (P1, P2, P3, ..).

       ----------------------------------------------
       | File: Movies.xlsx     -      Sheet: Actors |
       ----------------------------------------------
       |           A           B            C       |
       |1     First Name                 Address    |
       |2     Chuck         Norris       California |
       |3     Jean-Claude   Vandamme     Brussels   |
       ----------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors -NoHeader

       P1: First Name
       P2:
       P3: Address

       P1: Chuck
       P2: Norris
       P3: California

       P1: Jean-Claude
       P2: Vandamme
       P3: Brussels

    .EXAMPLE
       Import data from an Excel worksheet. One object is created for each row. The property names of the objects consist of the names defined in the parameter ‘-HeaderName’. The properties are named starting from the most left column (A) to the right. In case no value is present in one of the columns, that property will have an empty value.

       ----------------------------------------------------------
       | File: Movies.xlsx            -           Sheet: Movies |
       ----------------------------------------------------------
       |           A            B            C          D       |
       |1     The Bodyguard   1992           9                  |
       |2     The Matrix      1999           8                  |
       |3                                                       |
       |4     Skyfall         2012           9                  |
       ----------------------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Movies -HeaderName 'Movie name', 'Year', 'Rating', 'Genre'

       Movie name: The Bodyguard
       Year      : 1992
       Rating    : 9
       Genre     :

       Movie name: The Matrix
       Year      : 1999
       Rating    : 8
       Genre     :

       Movie name:
       Year      :
       Rating    :
       Genre     :

       Movie name: Skyfall
       Year      : 2012
       Rating    : 9
       Genre     :

       Notice that empty rows are imported and that data for the property 'Genre' is not present in the worksheet. As such, the 'Genre' property will be blanc for all objects.

    .EXAMPLE
       Import data from an Excel worksheet. One object is created for each row. The property names of the objects are automatically generated by using the switch ‘-NoHeader’ (P1, P@, P#, ..). The switch ‘-DataOnly’ will speed up the import because empty rows and empty columns are not imported.

       ----------------------------------------------------------
       | File: Movies.xlsx            -           Sheet: Movies |
       ----------------------------------------------------------
       |           A            B            C          D       |
       |1     The Bodyguard   1992           9                  |
       |2     The Matrix      1999           8                  |
       |3                                                       |
       |4     Skyfall         2012           9                  |
       ----------------------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Movies –NoHeader -DataOnly -HeaderOnly

       P1: The Bodyguard
       P2: 1992
       P3: 9

       P1: The Matrix
       P2: 1999
       P3: 8

       P1: Skyfall
       P2: 2012
       P3: 9

       Notice that empty rows and empty columns are not imported.

    .EXAMPLE
       Import data from an Excel worksheet. One object is created for each row. The property names are provided with the ‘-HeaderName’ parameter. The import will start from row 2 and empty columns and rows are not imported.

       ----------------------------------------------------------
       | File: Movies.xlsx            -           Sheet: Actors |
       ----------------------------------------------------------
       |           A           B           C            D       |
       |1     Chuck                     Norris       California |
       |2                                                       |
       |3     Jean-Claude               Vandamme     Brussels   |
       ----------------------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors -DataOnly -HeaderOnly -HeaderName 'FirstName','SecondName','City' –StartRow 2

       FirstName : Jean-Claude
       SecondName: Vandamme
       City      : Brussels

       Notice that only 1 object is imported with only 3 properties. Column B and row 2 are empty and have been disregarded by using the switch '-DataOnly'. The property names have been named with the values provided with the parameter '-HeaderName'. Row number 1 with ‘Chuck Norris’ has not been imported, because we started the import from row 2 with the parameter ‘-StartRow 2’.

    .EXAMPLE

       ----------------------------------------------------------
       | File: Movies.xlsx            -           Sheet: Actors |
       ----------------------------------------------------------
       |           A           B           C            D       |
       |1     Chuck                     Norris       California |
       |2                                                       |
       |3     Jean-Claude               Vandamme     Brussels   |
       ----------------------------------------------------------

       PS C:\> Import-Excel -Path 'C:\Movies.xlsx' -WorkSheetname Actors -Column 1,3,4 -Row 3 -NoHeader -HeaderName 'FirstName','SecondName','City'

       FirstName : Jean-Claude
       SecondName: Vandamme
       City      : Brussels

    .EXAMPLE
        >
        PS> ,(Import-Excel -Path .\SysTables_AdventureWorks2014.xlsx) |
            Write-SqlTableData -ServerInstance localhost\DEFAULT -Database BlankDB -SchemaName dbo -TableName MyNewTable_fromExcel -Force

            Imports data from an Excel file and pipe the data to the Write-SqlTableData to be INSERTed into a table in a SQL Server database.
            The ",( ... )" around the Import-Excel command allows all rows to be imported from the Excel file, prior to pipelining to the Write-SqlTableData cmdlet.  This helps prevent a RBAR scenario and is important when trying to import thousands of rows.
            The -Force parameter will be ignored if the table already exists.  However, if a table is not found that matches the values provided by -SchemaName and -TableName parameters, it will create a new table in SQL Server database.  The Write-SqlTableData cmdlet will inherit the column names & datatypes for the new table from the object being piped in.
            NOTE: You need to install the SqlServer module from the PowerShell Gallery in oder to get the Write-SqlTableData cmdlet.

   .LINK
       https://github.com/dfinke/ImportExcel

   .NOTES
  #>

    [CmdLetBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
   
    Param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Path')]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Path-NoHeader')]
        [ValidateNotNullOrEmpty()]
        [Alias('PSPath')]
        [String]
        $Path,

        [Parameter(Mandatory, ParameterSetName = 'Package')]
        [Parameter(Mandatory, ParameterSetName = 'Package-NoHeader')]
        [ValidateNotNull()]
        [OfficeOpenXml.ExcelPackage]
        $ExcelPackage,
        
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [Alias('Sheet')]
        [String]
        $WorksheetName,

        [String[]]
        $HeaderName,

        [Parameter(Mandatory, ParameterSetName = 'Path-NoHeader')]
        [Parameter(Mandatory, ParameterSetName = 'Package-NoHeader')]
        [Switch]
        $NoHeader,

        [ValidateRange(1, 1048576)]
        [Alias('TopRow')]
        [Int]
        $StartRow,
        
        [ValidateRange(1, 1048576)]
        [Alias('BottomRow')]
        [Int]
        $EndRow,

        [ValidateRange(1, 16384)]
        [Alias('LeftColumn')]
        [Int]
        $StartColumn,
        
        [ValidateRange(1, 16384)]
        [Alias('RightColumn')]
        [Int]
        $EndColumn,

        [ValidateCount(1, 1048576)]
        [ValidateRange(1, 1048576)]
        [Alias('Row')]
        [Int[]]
        $IncludeRow,

        [ValidateCount(1, 16384)]
        [ValidateRange(1, 16384)]
        [Alias('Column')]
        [Int[]]
        $IncludeColumn,

        [Switch]
        $DataOnly,

        [Parameter(ParameterSetName = 'Path')]
        [Parameter(ParameterSetName = 'Package')]
        [Switch]
        $HeaderOnly,

        [ValidateNotNullOrEmpty()]
        [String]
        $Password
    )

    begin {

        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($StartRow -or $EndRow -and $IncludeRow) {
            $PSCmdlet.WriteWarning("'-StartRow' and '-EndRow' will be ignored because the '-IncludeRow' parameter is specified")
        }
        if ($StartColumn -or $EndColumn -and $IncludeColumn) {
            $PSCmdlet.WriteWarning("'-StartColumn' and '-EndColumn' will be ignored because the '-IncludeColumn' parameter is specified.")
        }
    }

    process {

        if ($Path) {
            try {
                $file = $PSCmdlet.InvokeProvider.Item.Get($path, $false, $true)
            }
            catch {
                # path not found
                $e = [exception]::new("'${Path}' file not found.", $_.Exception)
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($e, 'PathNotFound', 'ObjectNotFound', $Path))
                return
            }

            # check item type
            if ($file.PSIsContainer) {
                $e = [exception]::new("'${Path}' A directory cannot be specified for the '-Path' parameter.")
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($e, 'InvalidItemType', 'ObjectNotFound', $Path))
                return
            }

            # check file type
            if ($file.Extension -notin '.xlsx','.xlsm') {
                $e = [exception]::new("Import-Excel does not support reading this extension type $($file.Extension)")
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($e, 'InvalidFileType', 'NotSpecified', $Path))
                return
            }

            try {
                $stream = $file.Open('Open', 'Read', 'ReadWrite')
            }
            catch {
                $e = [exception]::new("The process cannot access the file '$($file.FullName)' because it is being used by another process.", $_.Exception)
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($e, 'IOException', 'NotSpecified', $Path))
                return
            }
            
            try {
                $ExcelPackage = if ($Password) { [OfficeOpenXml.ExcelPackage]::new($stream, $Password) } else { [OfficeOpenXml.ExcelPackage]::new($stream) }
            }
            catch {
                $e = [exception]::new("Failed loading Excel package '$Path'. The file is corrupted or the password is incorrect.", $_.Exception)
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($e, 'Exception', 'NotSpecified', $Path))
                return
            }
        }

        try {

            # select worksheet
            $workSheet = $ExcelPackage.Workbook.Worksheets[$(if ($WorkSheetName) { $WorkSheetName } else { 1 })]

            # check for existence of worksheet
            if (!$workSheet) {
                $e = [exception]::new("Worksheet '$WorksheetName' not found, the workbook only contains the worksheets '$($ExcelPackage.Workbook.Worksheets)'. If you only wish to select the first worksheet, please remove the '-WorksheetName' parameter.")
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($e, "InvalidWorksheetName", "NotSpecified", $WorksheetName))
                return
            }
 
            $PSCmdlet.WriteDebug($sw.Elapsed.TotalMilliseconds)

            if (!$IncludeRow) { $IncludeRow = $(if ($StartRow) { $StartRow } else { 1 })..$(if ($EndRow) { $EndRow } else { $Worksheet.Dimension.End.Row }) }
            if (!$IncludeColumn) { $IncludeColumn = $(if ($StartColumn) { $StartColumn } else { 1 })..$(if ($EndColumn) { $EndColumn } else { $Worksheet.Dimension.End.Column }) }

            #region Get rows and columns
            
            # If we are doing dataonly it is quicker to work out which rows to ignore before processing the cells.
            $dataRow = $IncludeRow[(!$NoHeader)..($IncludeRow.Length - 1)]
            if ($DataOnly) {

                # We're going to look at every cell and build 2 hash tables holding rows & columns which contain data.
                # Want to Avoid 'select unique' operations & large Sorts, becuse time time taken increases with square
                # of number of items (PS uses heapsort at large size). Instead keep a list of what we have seen,
                # using Hash tables: "we've seen it" is all we need, no need to worry about "seen it before" / "Seen it many times".
                
                $colHash = @{}
                $rowHash = @{}
                foreach ($row in $dataRow) {
                    foreach ($col in $IncludeColumn) {
                        if ($null -ne $workSheet.GetValue($row, $col)) { $colHash[$col] = $rowHash[$row] = $true }
                    }
                }
                $rows = foreach ($row in $dataRow)  { if ($rowHash[$row]) { $row } }
                $columns = foreach ($col in $IncludeColumn)  { if ($colHash[$col]) { $col } }
            }
            else {
                $rows = $dataRow
                $columns = $IncludeColumn
            }
            #endregion


            #region Create property names

            [array]$propertyNames =
                if ($NoHeader -and !$HeaderName) {
                    foreach ($i in 1..$columns.Length) { "P$i" }
                }
                else {
                    if ($HeaderOnly -and $HeaderName) {
                        $columns = foreach ($col in $columns) { if($workSheet.GetValue($IncludeRow[0], $col)) { $col } }
                    }
                    elseif ($HeaderOnly) {
                        $headerList = [System.Collections.Generic.List[string]]::new()
                        $columns = foreach ($col in $columns) {
                            if($header = $workSheet.GetValue($IncludeRow[0], $col)) {
                                $headerList.Add($header)
                                $col
                            }
                        }
                        $HeaderName = $headerList
                    }
                    elseif (!$HeaderName) {
                        $HeaderName = foreach ($col in $columns) { $workSheet.GetValue($IncludeRow[0], $col) }
                    }
                    
                    if ($HeaderName.Length -lt $columns.Length) { $HeaderName += @("") * ($columns.Length - $HeaderName.Length) }

                    # check duplicates
                    $h = @{ "" = 1 }
                    foreach($header in $HeaderName) { $h["$header"]++ }
                    foreach($header in $HeaderName) {
                        if ($h["$header"] -ge 2) {
                            $i = 1
                            while ($h.Contains(($newHeader = $header + "#$i"))) { $i++ }
                            $h[$newHeader]++
                            $newHeader
                        }
                        else { $header }
                    }
                }
            #endregion
            
            if (!$columns) {
                $PSCmdlet.WriteWarning("No valid columns found.")
                return
            }
            if (!$rows) {
                $PSCmdlet.WriteWarning("Worksheet '$WorksheetName' in workbook '$Path' contains no data in the rows after top row '$StartRow'")
                return
            }

            $PSCmdlet.WriteDebug($sw.Elapsed.TotalMilliseconds)

            #region Create one object per row

            $newRow = Select-Object -InputObject 0 -Property $propertyNames
            
            foreach ($row in $rows) {
                # Disabled write-verbose for speed
                #$PSCmdlet.WriteVerbose("Import row '$row'")

                $i = 0
                foreach ($col in $columns) {
                    $newRow.($propertyNames[$i]) = $workSheet.GetValue($row, $col)
                    #$PSCmdlet.WriteVerbose("Import cell '$($Worksheet.Cells[$row, $col].Address)' with property name '$($propertyNames[$i])' and value '$($Worksheet.Cells[$row, $col].Value)'.")
                    $i++
                }
                $newRow.psobject.Copy()
            }
            #endregion

            $PSCmdlet.WriteDebug($sw.Elapsed.TotalMilliseconds)
        }
        catch {
            $e = [exception]::new("Failed importing the Excel workbook '$Path' with worksheet '$($workSheet.Name)':`r`n ($($_.Exception.Message)[$(($_.ScriptStackTrace -split "`r`n")[0])])", $_.Exception)
            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($e, 'Unknown', 'NotSpecified', $WorksheetName))
            return
        }
        finally {
            if ($Path) {
                $stream.Close()
                $ExcelPackage.Dispose()
            }
        }
    }
}

function ConvertFrom-ExcelSheet {
    <#
      .Synopsis
        Reads an Excel file an converts the data to a delimited text file.

      .Example
        ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data
        Reads each sheet in TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt.

      .Example
        ConvertFrom-ExcelSheet .\TestSheets.xlsx .\data sheet?0
        Reads and outputs sheets like Sheet10 and Sheet20 form TestSheets.xlsx and outputs it to the data directory as the sheet name with the extension .txt.
    #>

    [CmdletBinding()]
    param
    (
        [Alias("FullName")]
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
        [String]
        $OutputPath = '.\',
        [String]
        $SheetName = "*",
        [ValidateSet('ASCII', 'BigEndianUniCode', 'Default', 'OEM', 'UniCode', 'UTF32', 'UTF7', 'UTF8')]
        [string]
        $Encoding = 'UTF8',
        [ValidateSet('.txt', '.log', '.csv')]
        [string]
        $Extension = '.csv',
        [ValidateSet(';', ',')]
        [string]
        $Delimiter = ';'
    )

    $Path = (Resolve-Path $Path).Path
    $Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path, "Open", "Read", "ReadWrite"
    $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Stream
    $workbook = $xl.Workbook

    $targetSheets = $workbook.Worksheets | Where-Object { $_.Name -like $SheetName }

    $params = @{ } + $PSBoundParameters
    $params.Remove("OutputPath")
    $params.Remove("SheetName")
    $params.Remove('Extension')
    $params.NoTypeInformation = $true

    Foreach ($sheet in $targetSheets) {
        Write-Verbose "Exporting sheet: $($sheet.Name)"

        $params.Path = "$OutputPath\$($Sheet.Name)$Extension"

        Import-Excel $Path -Sheet $($sheet.Name) | Export-Csv @params
    }

    $Stream.Close()
    $Stream.Dispose()
    $xl.Dispose()
}

function Export-MultipleExcelSheets {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
    param(
        [Parameter(Mandatory = $true)]
        $Path,
        [Parameter(Mandatory = $true)]
        [hashtable]$InfoMap,
        [string]$Password,
        [Switch]$Show,
        [Switch]$AutoSize
    )

    $parameters = @{ } + $PSBoundParameters
    $parameters.Remove("InfoMap")
    $parameters.Remove("Show")

    $parameters.Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

    foreach ($entry in $InfoMap.GetEnumerator()) {
        Write-Progress -Activity "Exporting" -Status "$($entry.Key)"
        $parameters.WorkSheetname = $entry.Key

        & $entry.Value | Export-Excel @parameters
    }

    if ($Show) { Invoke-Item $Path }
}

Function WorksheetArgumentCompleter {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $xlPath = $fakeBoundParameter['Path']
    if (Test-Path -Path $xlPath) {
        $xlpkg = Open-ExcelPackage -ReadOnly -Path $xlPath
        $WorksheetNames = $xlPkg.Workbook.Worksheets.Name
        Close-ExcelPackage -nosave -ExcelPackage $xlpkg
        $WorksheetNames.where( { $_ -like "*$wordToComplete*" }) | foreach-object {
            New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList "'$_'",
            $_ , ([System.Management.Automation.CompletionResultType]::ParameterValue) , $_
        }
    }
}
If (Get-Command -ErrorAction SilentlyContinue -name Register-ArgumentCompleter) {
    Register-ArgumentCompleter -CommandName 'Import-Excel' -ParameterName 'WorksheetName' -ScriptBlock $Function:WorksheetArgumentCompleter
}
