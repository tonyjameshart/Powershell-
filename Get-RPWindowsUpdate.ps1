cls
# Writer : Ritesh Parab
# Used .Net [Activator] class to create instance and to send comobj to remote server 
# Required 2.0 Framework on Remote 

Function Get-RPWindowsUpdates ($comp) {
	Try {
		$UpdateSession = $null
		$TotalUpdateCount = $null
		$UpdateListObj = $null	
		[array]$TotalUpdates = $null
			$UpdateSession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session","$comp")) 
			$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
			$SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software'")
			$PendingPatches = $SearchResult.updates
			$Critical = $SearchResult.updates | where { $_.MsrcSeverity -eq "Critical" }
			$important = $SearchResult.updates | where { $_.MsrcSeverity -eq "Important" }
			$other = $SearchResult.updates | where { $_.MsrcSeverity -eq $null }
			$Moderate = $SearchResult.updates | where { $_.MsrcSeverity -eq "Moderate" }
			
				for($i=0 ; $i -lt $PendingPatches.Count; $i++ ) {
				$UpdateListObj = New-Object PSObject –Prop @{
											'PatchTitle'= $($PendingPatches.Item($i).title);
											'MsrcSeverity' = $($PendingPatches.Item($i).MsrcSeverity);
											'RebootRequired' = $($PendingPatches.Item($i).rebootrequired);
											}
				$TotalUpdates += $UpdateListObj 
			}
			
			$TotalUpdateCount 	= New-Object PSObject –Prop @{
										'Total'= $($SearchResult.updates.count);
					                	'Critical'=$($Critical.count);
										'Important'=$($Important.count) ;
										'Moderate'=$($Moderate.count) ;
					                	'Other'=$($other.count);
									 } 
			
			$WidnowsUpdateResult	= New-Object PSObject –Prop @{
										'TotalUpdates'= $TotalUpdates ;
					                	'TotalUpdateCount'=$TotalUpdateCount;
										}
			
			return $WidnowsUpdateResult
	}
	Catch {
			
			if ($_.Exception.Message -like "*CreateInstance*" ){
				$WUError = "Failed to Create Instance on Remote"			
			}
			Elseif ($_.Exception.Message -like "*Search*"){
				$WUError = "Seems No Internet on Remote"			
			}
			Else {
				$WUError = "Unexpected Kind of Error"			
			}
			$WidnowsUpdateResult 	= New-Object PSObject –Prop @{
										'TotalUpdates'= "$WUError" ;
					                	'TotalUpdateCount'="$WUError";
										}
			#$_.Exception.message   
			Return $WidnowsUpdateResult	
	}
}	
#region Windows Update Info 
$comp = "Ritz-DreamPC"
$GetWindowsUpdates = Get-RPWindowsUpdates $comp

$GetWindowsUpdates.TotalUpdateCount.Total
#Endregion 

$excel = new-object -comobject Excel.Application
$excel.Visible = $True
$excel.DisplayAlerts = $False
$workbooks = $excel.Workbooks.Add()
$WSheetWU = $workbooks.worksheets
$WSheetWU= $WSheetWU.Item(1)

#region Windows Update
	$WSheetWU.Activate() 
#region Windows Update Heading 
		$MergeCellsWU= $WSheetWU.Range("A1:F2")
		$MergeCellsWU.Select() 
		$MergeCellsWU.MergeCells = $true
		$MergeCellsWU.Cells = "Pending Windows Updates"
		$MergeCellsWU.WrapText = $True   
		$MergeCellsWU.Interior.ColorIndex = 23 
		$MergeCellsWU.Font.ColorIndex = 27
		$MergeCellsWU.Font.Bold = $true 
#endregion 

	$WUStartrow = 4
		if ($GetWindowsUpdates.TotalUpdateCount -contains "Failed to Create Instance on Remote"){
			$WSheetWU.Cells.Item($WUStartrow ,2) = "Failed to Create Instance on Remote"
			$WSheetWU.Cells.Item(4,2).font.ColorIndex = 3
			# continue; 
		}	
		Elseif ($GetWindowsUpdates.TotalUpdateCount -contains "Seems No Internet on Remote") {	
			$WSheetWU.Cells.Item($WUStartrow ,2) = "Seems No Internet on Remote"
			$WSheetWU.Cells.Item(4,2).font.ColorIndex = 3
		}
		
		$WSheetWU.Cells.Item($WUStartrow ,1).font.bold = $true
		$WSheetWU.Cells.Item($WUStartrow ,1) = "Summary of Pending Windows Updates"
		$WSheetWU.Cells.Item($WUStartrow + 1 ,1) = "Critical" ; $WSheetWU.Cells.Item($WUStartrow + 1,2) = $GetWindowsUpdates.TotalUpdateCount.Critical
		$WSheetWU.Cells.Item($WUStartrow + 2 ,1) = "Important" ;  $WSheetWU.Cells.Item($WUStartrow +2 ,2) = $GetWindowsUpdates.TotalUpdateCount.Important
		$WSheetWU.Cells.Item($WUStartrow + 3 ,1) = "Other" ;  $WSheetWU.Cells.Item($WUStartrow + 3,2) = $GetWindowsUpdates.TotalUpdateCount.Other
		$WSheetWU.Cells.Item($WUStartrow + 4 ,1) = "Moderate" ;  $WSheetWU.Cells.Item($WUStartrow + 4,2) = $GetWindowsUpdates.TotalUpdateCount.Moderate
		$WSheetWU.Cells.Item(9,2).font.bold = $true
		$WSheetWU.Cells.Item($WUStartrow + 5 ,1) = "Total" ;  $WSheetWU.Cells.Item($WUStartrow + 5,2) = $GetWindowsUpdates.TotalUpdateCount.Total
		
	$WUforloop = 13	
		$WSheetWU.Cells.Item(12 ,1).font.bold = $true
		$WSheetWU.Cells.Item(12,1) = "Patch Details"
		
		$WSheetWU.Cells.Item(12 ,2).font.bold = $true
		$WSheetWU.Cells.Item(12,2) = "Severity"
		
#		$WSheetWU.Cells.Item(12 ,3).font.bold = $true
#		$WSheetWU.Cells.Item(12,3) = "RebootRequired"
		
		
		
		for ($u = 0; $u -lt $GetWindowsUpdates.TotalUpdates.count; $u++  ){
			$WSheetWU.Cells.Item($WUforloop+$U,1) = $GetWindowsUpdates.TotalUpdates[$U].'PatchTitle'  
			$WSheetWU.Cells.Item($WUforloop+$U,2) = $GetWindowsUpdates.TotalUpdates[$U].'MsrcSeverity'
			# $WSheetWU.Cells.Item($WUforloop+$U,3) = $GetWindowsUpdates.TotalUpdates[$U].'RebootRequired'
			$WUforloopend = $WUforloop+$U 
		}
	
		$WUStartRange = "A"+$WUforloop
		$WUEndRange = "B"+$WUforloopend 
		$WURange = $WUStartRange+':'+$WUEndRange
		$WSheetWU.Range($WURange).Wraptext = $true
		$WSheetWU.Range($WURange).EntireRow.Group()
		
		$rangeautofilter = "A12"
		$WSheetWU.Range($rangeautofilter).AutoFilter()
		$WSheetWU.Range("A12").ColumnWidth = 80
		$WSheetWU.Range("B12").ColumnWidth = 20
	#	$WSheetWU.Range("C12").ColumnWidth = 20
		
		$WUMakeupStartRange = "A"+ $WUStartrow
		$WUMakeupEndRange = "B"+ $WUforloopend 
		$WUMakeup = $WUMakeupStartRange+':'+$WUMakeupEndRange
		$WSheetWUMakeup  = $WSheetWU.Range($WUMakeup)
		$WSheetWUMakeup.BorderAround(2,2,23)
		$WSheetWUMakeup.Interior.ColorIndex = 19 
		$WSheetWUMakeup.Font.ColorIndex = 23
		
		$WSheetWU.Cells.Item(5,2).font.ColorIndex = 3			
		$WSheetWU.Cells.Item(6,2).font.ColorIndex = 3
#endregion 