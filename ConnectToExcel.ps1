
Function ET-PBIDesktopODCConnection
{  
# modified the https://github.com/DevScope/powerbi-powershell-modules/blob/master/Modules/PowerBIPS.Tools/PowerBIPS.Tools.psm1
# the Function Export-PBIDesktopODCConnection

	[CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]        
		[string]
        $port,
        [Parameter(Mandatory = $false)]        
		[string]
        $path	
    )
    
        $port = $port

        $odcXml = "<html xmlns:o=""urn:schemas-microsoft-com:office:office""xmlns=""http://www.w3.org/TR/REC-html40""><head><meta http-equiv=Content-Type content=""text/x-ms-odc; charset=utf-8""><meta name=ProgId content=ODC.Cube><meta name=SourceType content=OLEDB><meta name=Catalog content=164af183-2454-4f45-964a-c200f51bcd59><meta name=Table content=Model><title>PBIDesktop Model</title><xml id=docprops><o:DocumentProperties  xmlns:o=""urn:schemas-microsoft-com:office:office""  xmlns=""http://www.w3.org/TR/REC-html40"">  <o:Name>PBIDesktop Model</o:Name> </o:DocumentProperties></xml><xml id=msodc><odc:OfficeDataConnection  xmlns:odc=""urn:schemas-microsoft-com:office:odc""  xmlns=""http://www.w3.org/TR/REC-html40"">  <odc:Connection odc:Type=""OLEDB"">   
        <odc:ConnectionString>Provider=MSOLAP;Integrated Security=ClaimsToken;Data Source=$port;MDX Compatibility= 1; MDX Missing Member Mode= Error; Safety Options= 2; Update Isolation Level= 2; Locale Identifier= 1033</odc:ConnectionString>   
        <odc:CommandType>Cube</odc:CommandType>   <odc:CommandText>Model</odc:CommandText>  </odc:Connection> </odc:OfficeDataConnection></xml></head></html>"   
                
        #the location of the odc file to be opened
        $odcFile = "$path\excelconnector.odc"

        $odcXml | Out-File $odcFile -Force	

        # Create an Object Excel.Application using Com interface
        $objExcel = New-Object -ComObject Excel.Application

        # Disable the 'visible' property so the document won't open in excel
        $objExcel.Visible = $true

        # Open the Excel file and save it in $WorkBook
        $WorkBook = $objExcel.Workbooks.Open($odcFile)

}


write $args[0]

ET-PBIDesktopODCConnection -port $args[0] -path "C:\Temp"
