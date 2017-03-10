# ===================================================
#    Program : Excel Application
# Hazırlayan : Ertan Erdem
#      Tarih : 04.02.2015 Çarşamba 18:30
#      Dosya : 
# ===================================================
# http://www.petri.com/export-to-excel-with-powershell.htm
# http://powershell.org/wp/forums/topic/dataview-with-datas-from-several-tables/
# ===================================================

Function fnCiktiExcelTanim
{
    $global:dsExcel = new-object System.Data.DataSet

    $global:tblExcelBicim = $global:dsExcel.Tables.Add("tblExcelBicim")
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("SayfaNo",[int])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("SiraNo",[int])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("ExcIslem",[string])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("BasSatir",[int])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("BitSatir",[int])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("BasSutun",[int])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("BitSutun",[int])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("ExcSutun",[string])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("ExcDeger",[string])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("RenkYazi",[int])
    [void]$global:dsExcel.Tables["tblExcelBicim"].Columns.Add("RenkArka",[int])

    $global:tblExcelDosya = $global:dsExcel.Tables.Add("tblExcelDosya")
    [void]$global:dsExcel.Tables["tblExcelDosya"].Columns.Add("SayfaNo",[int])
    [void]$global:dsExcel.Tables["tblExcelDosya"].Columns.Add("SayfaAd",[string])
    [void]$global:dsExcel.Tables["tblExcelDosya"].Columns.Add("DosyaAd",[string])
    [void]$global:dsExcel.Tables["tblExcelDosya"].Columns.Add("Rapor",[string])
}

# ===================================================

Function fnCiktiExcelDosya ( $SayfaNo , $SayfaAd , $DosyaAd , $Rapor )
{
    $rowExcel = $global:tblExcelDosya.NewRow()  
    $rowExcel["SayfaNo"] = $SayfaNo
    $rowExcel["SayfaAd"] = $SayfaAd
    $rowExcel["DosyaAd"] = $DosyaAd
    $rowExcel["Rapor"] = $Rapor
    $global:tblExcelDosya.Rows.Add($rowExcel)
}

# ===================================================

Function fnCiktiExcelBicimEkle ( $SayfaNo,$SiraNo,$ExcIslem,$BasSatir,$BitSatir,$BasSutun,$BitSutun,$ExcSutun,$ExcDeger,$RenkYazi,$RenkArka )
{
    if ( $BasSatir -eq $null ) { $BasSatir = 0 }
    if ( $BitSatir -eq $null ) { $BitSatir = 0 }
    if ( $BasSutun -eq $null ) { $BasSutun = 0 }
    if ( $BitSutun -eq $null ) { $BitSutun = 0 }
    if ( $RenkYazi -eq $null ) { $RenkYazi = 0 }
    if ( $RenkArka -eq $null ) { $RenkArka = 0 }
    if ( $ExcSutun -eq $null ) { }
    if ( $ExcDeger -eq $null ) { }

    $rowExcel = $global:tblExcelBicim.NewRow()  
    $rowExcel["SayfaNo"] = $SayfaNo
    $rowExcel["SiraNo"] = $SiraNo
    $rowExcel["ExcIslem"] = $ExcIslem
    $rowExcel["BasSatir"] = $BasSatir
    $rowExcel["BitSatir"] = $BitSatir
    $rowExcel["BasSutun"] = $BasSutun
    $rowExcel["BitSutun"] = $BitSutun
    $rowExcel["ExcSutun"] = $ExcSutun
    $rowExcel["ExcDeger"] = $ExcDeger
    $rowExcel["RenkYazi"] = $RenkYazi
    $rowExcel["RenkArka"] = $RenkArka
    $global:tblExcelBicim.Rows.Add($rowExcel)
}

# ===================================================

Function fnCiktiExcelCsv ( $pDosyaKaydet , $pKapat , $pDelimiter )
{  
    $pExcelApp = New-Object -ComObject "Excel.Application"
    $pExcelApp.Visible = $False
    $pExcelApp.ScreenUpdating = $True
    $pWorkBook = $pExcelApp.Workbooks.Add()
    $pExcelApp.DisplayAlerts = $False

    ForEach ( $DataRow in $global:dsExcel.Tables["tblExcelDosya"])
    {    
        $pWorkSheet = $pExcelApp.Sheets.Add()
    }

    ForEach ( $DataRow in $global:dsExcel.Tables["tblExcelDosya"])
    {    
        if ( ( $DataRow.DosyaAd -le 3 ) -or ( $DataRow.DosyaAd -eq "" ) -or ( $DataRow.DosyaAd -eq $null ) ) { continue }
                
        $pWorkSheet = $pExcelApp.Sheets.ITEM($DataRow.SayfaNo) 
        $pWorkSheet.Name = $DataRow.SayfaAd

        $pQueryTable = $pWorkSheet.QueryTables.ADD("TEXT;" + $DataRow.DosyaAd , $pWorkSheet.Cells( 1 , 1) )
        $pQueryTable.TextFileParseType = 1
        $pQueryTable.FieldNames = $True            
        $pQueryTable.RowNumbers = $False           
        $pQueryTable.FillAdjacentFormulas = $False
        $pQueryTable.PreserveFormatting = $True    
        $pQueryTable.TextFileConsecutiveDelimiter = $False
        $pQueryTable.TextFileTabDelimiter = $False        
        $pQueryTable.TextFileCommaDelimiter = $False
        $pQueryTable.TextFileSpaceDelimiter = $False
        $pQueryTable.TextFileTrailingMinusNumbers = $True

        if ( $pDelimiter -eq $null ) { $pQueryTable.TextFileSemicolonDelimiter = $True }
        if ( $pDelimiter -eq ";"   ) { $pQueryTable.TextFileSemicolonDelimiter = $True }
        if ( $pDelimiter -eq "|"   ) { $pQueryTable.TextFileOtherDelimiter = "|" }
        
        $pQueryTable.Refresh()
    }

    fnCiktiExcelBicimle

    $pExcelApp.ActiveWindow.ScrollRow = 1

    $pWorkSheet = $pExcelApp.Sheets.ITEM(1)
    $pWorkSheet:Activate    

    if ( $pDosyaXls.Length -gt 5 ) 
    { 
        $pWorkBook.SaveAs( $pDosyaKaydet , 1 ) 
    }

    $pExcelApp.Visible = $True

    if ( $pKapat -eq "E" )
    {
        $pExcelApp.Workbooks.Close()
        $pExcelApp.Quit()
    }
}

# ===================================================

Function fnCiktiExcelBicimle
{

    ForEach ( $DataRowDosya in $global:dsExcel.Tables["tblExcelDosya"])
    {
        if ( ( $DataRowDosya.DosyaAd -le 3 ) -or ( $DataRowDosya.DosyaAd -eq "" ) -or ( $DataRowDosya.DosyaAd -eq $null ) ) { continue }

        $pWorkSheet = $pExcelApp.Sheets.ITEM($DataRowDosya.SayfaNo) 
        $pWorkSheet.Name = $DataRowDosya.SayfaAd
        $pWorkSheet.Select()

        ForEach ( $DataRowBicim in $global:dsExcel.Tables["tblExcelBicim"])
        {
            if ( $DataRowBicim.SayfaNo -ne $DataRowDosya.SayfaNo )      # Bu sayfaya Uygulanacak bir biçimlendirme ise
            {
                if ( $DataRowBicim.SayfaNo -ne 0 ) { continue }         # Genel Uygulanacak biçimlendirme değil ise                                
            }

            if ( $DataRowBicim.ExcIslem -EQ "baslikdondur" )            # Başlık Dondurur
            {
                $pExcelApp.ActiveWindow.SplitRow = ( $DataRowBicim.ExcDeger )
                $pExcelApp.ActiveWindow.FreezePanes = $TRUE 
            }

            if ( $DataRowBicim.ExcIslem -EQ "sayfakaydir" )             # Başlık Dondurur Kaydır
            {
                $pExcelApp.ActiveWindow.ScrollRow = ($DataRowBicim.ExcDeger)
            }

            if ( $DataRowBicim.ExcIslem -EQ "cerceve" )                 # Çerçeve Çizer
            {
                if ( ( $DataRowBicim.BasSatir -ne 0 ) -and ( $DataRowBicim.BitSatir -ne 0 ) -and ( $DataRowBicim.BasSutun -ne 0 ) -and ( $DataRowBicim.BitSutun -ne 0 ) )
                {       
                    $pWorkSheet.Range($pWorkSheet.Cells( $DataRowBicim.BasSatir,$DataRowBicim.BasSutun ),$pWorkSheet.Cells( $DataRowBicim.BitSatir,$DataRowBicim.BitSutun )).SELECT()
                    $pExcelApp.SELECTION.Borders(1).LineStyle = 1
                    $pExcelApp.SELECTION.Borders(2).LineStyle = 1
                    $pExcelApp.SELECTION.Borders(3).LineStyle = 1
                    $pExcelApp.SELECTION.Borders(4).LineStyle = 1                 
                }
            }

            if ( $DataRowBicim.ExcIslem -eq "sayfarenk" )
            {
                $pExcelApp.Sheets($DataRowDosya.SayfaAd).Tab.Color = ($DataRowBicim.ExcDeger)
            }

            if ( $DataRowBicim.ExcIslem -eq "grupla" )
            { 
                $pWorkSheet.Rows($DataRowBicim.ExcDeger).GROUP()
                $pWorkSheet.Outline.SummaryRow = 0
                $pWorkSheet.Outline.ShowLevels( 1 )
            }

            if ( $DataRowBicim.ExcIslem -eq "gruplasutun" )
            {
                $pWorkSheet.Columns( $DataRowBicim.ExcDeger ).GROUP()
                $pWorkSheet.Outline.SummaryRow = 0
                $pWorkSheet.Outline.ShowLevels( 1 )
            }


            if ( $DataRowBicim.ExcIslem -EQ "cercevetek" )
            {
                $pWorkSheet.range($pWorkSheet.cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).BorderAround( 1, 3, 1, 1)
            }

            if ( $DataRowBicim.ExcIslem -EQ "birlestir" )
            {
                $pWorkSheet.Range($pWorkSheet.cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).MergeCells = $TRUE
            }

            if ( $DataRowBicim.ExcIslem -EQ "yerlesdikey" )
            {
                $pWorkSheet.Range($pWorkSheet.Cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).HorizontalAlignment = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "yerlesyatay" )
            {
                $pWorkSheet.Range($pWorkSheet.Cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).VerticalAlignment = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "yerleskaydir" )
            {
                $pWorkSheet.Range($pWorkSheet.Cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).WrapText = $TRUE 
            }

            if ( $DataRowBicim.ExcIslem -EQ "fontkalin" )
            {
                $pWorkSheet.range($pWorkSheet.cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).FONT.bold = $TRUE
            }

            if ( $DataRowBicim.ExcIslem -EQ "renkarka" )
            {
                $pWorkSheet.Range($pWorkSheet.Cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).Interior.ColorIndex = $DataRowBicim.RenkArka
            }

            if ( $DataRowBicim.ExcIslem -EQ "renkyazi" )
            {
                $pWorkSheet.Range($pWorkSheet.Cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).FONT.ColorIndex = ( $DataRowBicim.ExcDeger )
            }

            if ( $DataRowBicim.ExcIslem -EQ "sayi" )
            {
                $pWorkSheet.Range($pWorkSheet.cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).NumberFormat = "#.##0,00"
            }

            if ( $DataRowBicim.ExcIslem -EQ "yaziformat" )
            {
                $pWorkSheet.Range($pWorkSheet.cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).NumberFormat = $DataRowBicim.ExcDeger 
            }

            if ( $DataRowBicim.ExcIslem -EQ "genislik" )
            {
                $pWorkSheet.COLUMNS( $DataRowBicim.ExcSutun ).ColumnWidth = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "yukseklik" )
            {
                $pWorkSheet.Rows( $DataRowBicim.ExcSutun ).RowHeight = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "otomatikgenislik" )
            {
                $pWorkSheet.Cells.EntireColumn.AutoFit
            }

            if ( $DataRowBicim.ExcIslem -EQ "sayfafontad" )                      # Sayfanýn Fontunu Belirler
            {
                $pWorkSheet.Cells.FONT.NAME = $DataRowBicim.ExcDeger             # Arial
            }

            if ( $DataRowBicim.ExcIslem -EQ "sayfafontboyut" )                 # Sayfanýn Fon Foyutunu Belirler 
            {
                $pWorkSheet.Cells.FONT.SIZE = $DataRowBicim.ExcDeger             #  10 
            }

            if ( $DataRowBicim.ExcIslem -EQ "fontboyut" )                   # Sayfanın Fon Foyutunu Belirler 
            {
                $pWorkSheet.Range($pWorkSheet.cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).FONT.SIZE = $DataRowBicim.ExcDeger            # 10 
            }

            if ( $DataRowBicim.ExcIslem -EQ "baslikustsatir" )
            {
                $pWorkSheet.PageSetup.PrintTitleRows = $DataRowBicim.ExcDeger   # "$1:$1"
            }

            if ( $DataRowBicim.ExcIslem -EQ "baslikustorta" )
            {
                $pWorkSheet.PageSetup.CenterHeader = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -eq "baslikustsol" )
            {
                $pWorkSheet.PageSetup.LeftHeader = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "baslikustsag" )
            {
                $pWorkSheet.PageSetup.RightHeader = $DataRowBicim.ExcDeger 
            }

            if ( $DataRowBicim.ExcIslem -EQ "baslikaltorta" )
            {
                $pWorkSheet.PageSetup.CenterFooter = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "baslikaltsol" )
            {
                $pWorkSheet.PageSetup.LeftFooter = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "baslikaltsag" )
            {
                $pWorkSheet.PageSetup.RightFooter = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "sayfayerlesim" )
            {
                $pWorkSheet.PageSetup.ORIENTATION = $DataRowBicim.ExcDeger
            }

            if ( $DataRowBicim.ExcIslem -EQ "filtrele" )
            {
                $pWorkSheet.Range($pWorkSheet.Cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ),$pWorkSheet.Cells( $DataRowBicim.BitSatir , $DataRowBicim.BitSutun )).SELECT() 
                $pWorkSheet.SELECT()
                $pExcelApp.SELECTION.AutoFilter()
            }

            if ( $DataRowBicim.ExcIslem -EQ "linkekle" )
            {
                $pHyperlink = $pWorkSheet.Hyperlinks.ADD( $pWorkSheet.cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ) ,  "" , $DataRowBicim.ExcSutun , "" , $DataRowBicim.ExcDeger )
            }

            if ( $DataRowBicim.ExcIslem -EQ "resimekle" )
            {
                $pExcelApp.ActiveSheet.Pictures.INSERT( $DataRowBicim.ExcDeger ).SELECT()
                $pExcelApp.SELECTION.ShapeRange.HEIGHT = 50
                $pExcelApp.SELECTION.ShapeRange.WIDTH  = 50                 
                $pExcelApp.SELECTION.TOP = $pWorkSheet.Cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ).TOP()
                $pExcelApp.SELECTION.LEFT = $pWorkSheet.Cells( $DataRowBicim.BasSatir , $DataRowBicim.BasSutun ).LEFT()
                $pExcelApp.SELECTION.ShapeRange.IncrementLeft("3")
                $pExcelApp.SELECTION.ShapeRange.IncrementTop("3")
                $pWorkSheet.Rows( $DataRowBicim.BasSatir.ToString() + ":" + $DataRowBicim.BasSatir.ToString() ).RowHeight = 55
            }
            
            if ( $DataRowBicim.ExcIslem -EQ "sayfakenarbosluk" )
            {
                $pWorkSheet.PageSetup.LeftMargin   = $pExcelApp.InchesToPoints(0.8)
                $pWorkSheet.PageSetup.RightMargin  = $pExcelApp.InchesToPoints(0.2)
                $pWorkSheet.PageSetup.TopMargin    = $pExcelApp.InchesToPoints(0.7)
                $pWorkSheet.PageSetup.BottomMargin = $pExcelApp.InchesToPoints(0.8)
                $pWorkSheet.PageSetup.HeaderMargin = $pExcelApp.InchesToPoints(0.2)
                $pWorkSheet.PageSetup.FooterMargin = $pExcelApp.InchesToPoints(0.2)
            }

            if ( $DataRowBicim.ExcIslem -EQ "grafikayri" )
            {
                $pChart = $pExcelApp.Charts.Add()
                $pChart:ChartType = 11 
                # $pChart:SetSourceData($pExcelApp.WorkSheets("Sheet1").Range("A1:A5;C1:C5") , 1 )
            }


            if ( $DataRowBicim.ExcIslem -EQ "grafikyersec" )
            {              
                
                $pWorksheetRange = $pWorkSheet.Range($DataRowBicim.ExcDeger)
                $pWorkSheet.ChartObjects.Add(10,150,425,300).Activate # Aynı Sayfada Konum Belirtiliyor 
                $pExcelApp.ActiveChart.ChartWizard( $pWorksheetRange , 3 , 1 , 2 , 1 , 1 , $TRUE , ( ENTRY ( 1 ,  $DataRowBicim.ExcSutun , "|" ) ) , ( ENTRY ( 2 ,  $DataRowBicim.ExcSutun , "|" ) ) , ( ENTRY ( 3 ,  $DataRowBicim.ExcSutun , "|" ) ) )
               
                #chExcelApplication:ActiveChart:ChartWizard( chWorksheetRange , 3 , 1 , 2 , 1 , 1 , TRUE , "Üst Grafik Baþlýk", "Alt Taraf Personel", "Sol Taraf Tutarlar" ).                 
            }
        }                              
    }

    

    ForEach ( $DataRowDosya in $global:dsExcel.Tables["tblExcelDosya"])         # Özet Tablo Yapma
    {
        if ( ( $DataRowDosya.DosyaAd -le 3 ) -or ( $DataRowDosya.DosyaAd -eq "" ) -or ( $DataRowDosya.DosyaAd -eq $null ) ) { continue }

        ForEach ( $DataRowBicim in ( $global:dsExcel.Tables["tblExcelBicim"] | Where-Object { $_.SayfaNo -eq $DataRowDosya.SayfaNo } ) )
        {            
            IF ( $DataRowBicim.ExcIslem -EQ "ozettabloyap" )
            {
                $pWorkSheet = $pExcelApp.Sheets.Item($DataRowDosya.SayfaAd)
                $pWorkSheet.Select()
                $pWorkSheet.Activate()                  
                $pPivotTable = $pExcelApp.ActiveWorkbook.PivotCaches().Create( 1 , $DataRowBicim.ExcSutun )
                $pPivotTable.CreatePivotTable("", $DataRowBicim.ExcDeger) | Out-Null                 
            }
                    
            IF ( $DataRowBicim.ExcIslem -EQ "ozettabloyerbas" )
            {
                $pExcelApp.ActiveSheet.PivotTables($DataRowBicim.ExcSutun).PivotFields( $DataRowBicim.ExcDeger ).Orientation = 3 # 1:Satýr 2:Sütun 3:Üst 4:Deðer
            }

            IF ( $DataRowBicim.ExcIslem -EQ "ozettabloyersatir" )
            {
                $pExcelApp.ActiveSheet.PivotTables($DataRowBicim.ExcSutun).PivotFields( $DataRowBicim.ExcDeger ).Orientation = 1 # 1:Satýr 2:Sütun 3:Üst 4:Deðer
            }

            IF ( $DataRowBicim.ExcIslem -EQ "ozettabloyersutun" )
            {
                $pExcelApp.ActiveSheet.PivotTables($DataRowBicim.ExcSutun).PivotFields( $DataRowBicim.ExcDeger ).Orientation = 2 # 1:Satýr 2:Sütun 3:Üst 4:Deðer
            }

            IF ( $DataRowBicim.ExcIslem -EQ "ozettabloyerdeger" )
            {
                $pExcelApp.ActiveSheet.PivotTables($DataRowBicim.ExcSutun).PivotFields( $DataRowBicim.ExcDeger ).Orientation = 4 # 1:Satýr 2:Sütun 3:Üst 4:Deðer           
            }

            IF ( $DataRowBicim.ExcIslem -EQ "ozettablotoplamyok" )   # Çalışmadı 
            {
                $pExcelApp.ActiveSheet.PivotTables($DataRowBicim.ExcSutun).ColumnGrand = $false
                $pExcelApp.ActiveSheet.PivotTables($DataRowBicim.ExcSutun).RowGrand = $false    
            }
        }
    }
}

# ===================================================
<#
fnCiktiExcelTanim
fnCiktiExcelDosya -SayfaNo 1 -SayfaAd "Rapor1" -DosyaAd "C:\psLog\Rok\Duzeltmeler\RokListKon20170208112110.csv" -Rapor ""

fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 1 -ExcIslem "cerceve" -BasSatir 1 -BitSatir 1198 -BasSutun 1 -BitSutun 12 
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 1 -ExcIslem "renkarka" -BasSatir 1 -BitSatir 1 -BasSutun 1 -BitSutun 12 -RenkArka 15
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 1 -ExcIslem "baslikdondur" -ExcDeger "1"
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 1 -ExcIslem "otomatikgenislik"
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 1 -ExcIslem "filtrele" -BasSatir 1 -BitSatir 1198 -BasSutun 1 -BitSutun 12

fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 2 -ExcIslem "baslikaltsol" -ExcDeger "Kontrol"
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 2 -ExcIslem "baslikaltsag" -ExcDeger "Sayfa : &P / &N"

fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 10 -ExcIslem "ozettabloyap" -ExcSutun "A1:L1198" -ExcDeger "PivotTable1"
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 11 -ExcIslem "ozettabloyerbas" -ExcSutun "PivotTable1" -ExcDeger "Lst"    
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 13 -ExcIslem "ozettabloyersatir" -ExcSutun "PivotTable1" -ExcDeger "CategoryName"
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 14 -ExcIslem "ozettabloyersatir" -ExcSutun "PivotTable1" -ExcDeger "Domain"    
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 16 -ExcIslem "ozettabloyersutun" -ExcSutun "PivotTable1" -ExcDeger "Tablo"
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 17 -ExcIslem "ozettabloyerdeger" -ExcSutun "PivotTable1" -ExcDeger "SourceName"
fnCiktiExcelBicimEkle -SayfaNo 1 -SiraNo 18 -ExcIslem "ozettablotoplamyok" -ExcSutun "PivotTable1"

fnCiktiExcelCsv -pDelimiter "|"

#>