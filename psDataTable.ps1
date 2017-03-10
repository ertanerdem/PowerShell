# ===================================================
#    Program : DataTable Fonksiyonları
# Hazırlayan : Ertan Erdem
#      Tarih : 10.03.2017 Cuma 16:00
#      Dosya : psDataTable.ps1
# ===================================================

Function fnDataTableOlustur
{    
    $global:tblTemp = New-Object System.Data.DataSet 
}

# ===================================================

Function fnDataTableEkle ( $pTblName )
{    
    $global:tblTemp.Tables.Add($pTblName)    
}

# ===================================================

Function fnDataTableClear ( $pTblName )
{    
    $global:tblTemp.Tables[$pTblName].Clear()      
}

# ===================================================

Function fnDataTableAlanEkle ( $pTblName , $pColName , $pColType )
{
    if ( $pColType -eq $null ) { $pColType = "string"} 

    $tblCol = New-Object System.Data.DataColumn
    $tblCol.ColumnName = $pColName
    $tblCol.DataType   = $pColType    
    $global:tblTemp.Tables[$pTblName].Columns.Add($tblCol)   
}

# ===================================================

Function fnDataTableSatirEkle ( $pTblName , $pColNames , $pColValues , $pAyr )
{        
    if ( $pAyr -eq $null ) { $pAyr = ";" }
    $pUzn = $pColNames.Split($pAyr).Count

    $tblRow = $global:tblTemp.Tables[$pTblName].NewRow()
    
    For ( $pSay = 0 ; $pSay -lt $pUzn ; $pSay++ )
    {
        $pColName  = $pColNames.Split($pAyr)[$pSay].Trim()
        $pColValue = $pColValues.Split($pAyr)[$pSay].Trim()
      
        $tblRow[$pColName] = $pColValue
    }

    $global:tblTemp.Tables[$pTblName].Rows.Add($tblRow)
}

# ===================================================

Function fnDataTableSatirSil ( $pTblName , $pColName , $pColValue )
{
    $rows = $global:tblTemp.Tables[$pTblName].Select($pColName + "='" + $pColValue + "'")
    
    ForEach($row in $rows) 
    { 
        $row.Delete()
    }     
}

# ===================================================

Function fnDataTableSatirGuncelle ( $pTblName , $pColName , $pColValue , $pSetCols , $pSetValues , $pAyr )
{      
    if ( $pAyr -eq $null ) { $pAyr = ";" }
    $pUzn = $pSetCols.Split($pAyr).Count
      
    $rows = $global:tblTemp.Tables[$pTblName].Select($pColName + "='" + $pColValue + "'")
    
    ForEach($tblRow in $rows) 
    {        
        For ( $pSay = 0 ; $pSay -lt $pUzn ; $pSay++ )
        {
            $pSetCol   = $pSetCols.Split($pAyr)[$pSay].Trim()
            $pSetValue = $pSetValues.Split($pAyr)[$pSay].Trim()
      
            $tblRow[$pSetCol] = $pSetValue
        }
    }                    
}

# ===================================================
<#

$pTmpName = "Category"

fnDataTableOlustur
fnDataTableEkle -pTblName $pTmpName

fnDataTableAlanEkle -pTblName $pTmpName -pColName "CatID" -pColType "int"
fnDataTableAlanEkle -pTblName $pTmpName -pColName "CatName" -pColType "string"

fnDataTableSatirEkle -pTblName $pTmpName -pColNames "CatID,CatName" -pColValues "3,Spor" -pAyr ","
fnDataTableSatirEkle -pTblName $pTmpName -pColNames "CatID,CatName" -pColValues "1,Sağlık" -pAyr ","
fnDataTableSatirEkle -pTblName $pTmpName -pColNames "CatID,CatName" -pColValues "2,Müzik" -pAyr ","

ForEach ( $DataRow in $global:tblTemp.Tables[$pTmpName])
{
    Write-Host $DataRow.CatName $DataRow.CatID   
}

fnDataTableSatirSil -pTblName $pTmpName -pColName "CatName" -pColValue "Spor"

ForEach ( $DataRow in $global:tblTemp.Tables[$pTmpName])
{
    Write-Host $DataRow.CatName $DataRow.CatID   
}

fnDataTableSatirGuncelle -pTblName $pTmpName -pColName "CatName" -pColValue "Sağlık" -pSetCols "CatID" -pSetValues "33" -pAyr ","

ForEach ( $DataRow in $global:tblTemp.Tables[$pTmpName])
{
    Write-Host $DataRow.CatName $DataRow.CatID   
}

fnDataTableClear -pTblName $pTmpName

ForEach ( $DataRow in $global:tblTemp.Tables[$pTmpName])
{
    Write-Host $DataRow.CatName $DataRow.CatID   
}
#>
# ===================================================
