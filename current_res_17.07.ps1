AllFolderShortcuts -f $true #����� �������� �������
 
function AllFolderShortcuts ($f) {
        $Domain = $env:USERDNSDOMAIN #��� ������
        $CurrentSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name #��� �����
        $UserName = $env:username #��� ������������       
        $strFamilia_IO = GetFIO -UserName $UserName        
        $UserGroups = ([ADSISEARCHER]"samaccountname=$UserName").FindOne().Properties.memberof #������ ����� �������� ������������
        $GroupList =''
        foreach ( $Group in $UserGroups ) {
                $GroupList+=$Group #������ ����� ������� ������������ � ������� string
        } 
        $flag = $false
        $WSShell =New-Object -com WScript.Shell  
        $NetHood = $WSShell.SpecialFolders.Item("Nethood")
        Get-ChildItem *.lnk -Path $NetHood| Remove-Item -Force -Recurse
        $AliasPrintServer = $CurrentSite.ToLower()+'_printserver'
        $PrintServer ="\\" +( GetPrintSerName -AliasPrintSer $AliasPrintServer)
       $setprinters =  Get-CimInstance -Class Win32_Printer | where-object{$_.name -like "\\*\*"}
        foreach ($printer in $setprinters) { 
                $PrinterN= $printer.name
               (New-Object -ComObject WScript.Network).RemovePrinterConnection("$PrinterN")
        }
        foreach ($Group in $UserGroups) #�������������� ��� ������ ������������
        {
            $GroupName=($Group -split ',')[0]
            $GroupName=$GroupName.Substring(3);
            $GroupInfo = GetGroupInfo -Group $Group
            if (($GroupInfo -split '\n')[0] -like "lnk*") 
            { #���� ������, � ������� ������ ������� "lnk*"                 
              
                    $WSShell1 =New-Object -com WScript.Shell  
                    $NetHood1 = $WSShell1.SpecialFolders.Item("Nethood")
                    $DirTitle = ($GroupInfo -split '\n')[2];
                    #��� ����� �����
                    $DirTitle = $DirTitle.Remove($DirTitle.Length-1)
                    $DirPath = ($GroupInfo -split '\n')[1];
                    $DirPath = $DirPath.Remove($DirPath.Length-1)
                    $ShortcutPath1 = Join-Path -Path $NetHood1 -ChildPath " $DirTitle.lnk"
                    $NewShortcut1 = $WSShell1.CreateShortcut($ShortcutPath1)
                    $NewShortcut1.TargetPath = "$DirPath"
                    $NewShortcut1.Save() #������ �� ����� �����
                
                    #������ ����� �������������
                    if ((($GroupList -like "*krk_Users*") -and ($CurrentSite -eq "KRK")) -or (($GroupList -like "*prp_Users*") -and ($CurrentSite -eq "PRP"))) { #������ � ������� ���� ���������
                        MakeShortcut -PathFile "\\$Domain\dfs\KRK\docs" -strFamilia_IO $strFamilia_IO -Location "" 
                    }
                    elseif (($GroupList -like "*krk_Users*") -and  ($GroupList -like "*prp_Users*") ) { #������������ ������� � ������� prp_ � krk_ 
                        MakeShortcut -PathFile "\\$Domain\dfs\KRK\docs" -strFamilia_IO $strFamilia_IO -Location "����������" 
                        MakeShortcut -PathFile "\\$Domain\dfs\PRP\docs" -strFamilia_IO $strFamilia_IO -Location "������������" 
                    }
                    elseif (($GroupList -like "*krk_Users*") -and ($CurrentSite -eq 'PRP')) { #������������ �� ����������� �� ��������
                        MakeShortcut -PathFile "\\$Domain\dfs\KRK\docs" -strFamilia_IO $strFamilia_IO -Location "����������"   
                    }
                    elseif(($GroupList -like "*prp_Users*") -and ($CurrentSite -eq 'KRK')) { #������������ � �������� � �����������
                        MakeShortcut -PathFile "\\$Domain\dfs\PRP\docs" -strFamilia_IO $strFamilia_IO -Location "������������"                   
                    }      
                }  
               if (($GroupInfo -split '\n')[0] -like "prn*") 
                { #���� ������ � ���������� "prn*" 
                     $AliasPrintServer = ($GroupInfo -split '\n')[1] #��� �����-������� �� ����������
                     $PrintServer ="\\" + ( GetPrintSerName -AliasPrintSer $AliasPrintServer)
                     $tempPrinterName = ($GroupInfo -split '\n')[2] #��� �������� �� ����������                    
                     $tempPrinterNameres = $tempPrinterName.Remove($tempPrinterName.Length-1)
                     $AddPrinterName = $PrintServer + "\" + $tempPrinterNameres                
                        if(($GroupInfo -split '\n')[0] -like "prnd*") {  #���� prnd - ������������� ������� ���������
                                AddPrinter -AddPrinterName $AddPrinterName -isDef $true #��������� ��� ���������
                        } 
                        else {
                                AddPrinter -AddPrinterName $AddPrinterName -isDef $false   #���� Prn - ������������� ��� �������                              
                        }   
                }    
                    
        }    
        $setPrinters = Get-CimInstance -Class Win32_Printer
        $setPrintersList=""
        foreach ( $print in $setPrinters ) {
                $setPrintersList+=$print #������ ������������ ��������� [string]
        } 

        $RoomName = (GetRoomName) #�������� ��� ��������  
        $RoomName = $RoomName.toLower()     
        $AliasPrintServer = $CurrentSite.ToLower()+'_printserver'
        $PrintServer ="\\" +( GetPrintSerName -AliasPrintSer $AliasPrintServer) #�������� ��� �����-�������
        $Printers = Get-Printer -ComputerName $PrintServer | where-object{$_.devicetype -eq 0} #�������� ������ ��������� � �������
        foreach ($Printer in $Printers)
        {
                $PrinterNameCur = $Printer.name #�������� ��� �������� ��������
                if ($PrinterNameCur -like "????$RoomName*") #������� ��������� �� �� � ������� ��������
                {     if (!($setPrintersList -like "*$PrinterNameCur*"))    #������� �� ��������� �� �� ���
                        {    
                                $AddPrinterName = $PrintServer+"\"+$PrinterNameCur #\\��� �����-�������\��� ��������
                                if ($flag -eq $true ) { #���� ��� ���� ��������� �������     
                                 AddPrinter -AddPrinterName $AddPrinterName -isDef $false #�������� ��� �������                               
                                }
                                else { #���� ���������� �������� ���   
                                        if ($PrinterNameCur -like "*pr01") {
                                                AddPrinter -AddPrinterName $AddPrinterName -isDef $true #��������� ��� ���������
                                        }
                                AddPrinter -AddPrinterName $AddPrinterName -isDef $false #�������� ��� �������
                                }                       
                        }
                }
        }
}


function  MakeShortcut { #�������� ������� 
        param (
           $PathFile, $strFamilia_IO, $Location     #�������� Location - ��� ����� �������� ����� �������_��
        )
        $WSShell =New-Object -com WScript.Shell  
        $NetHood = $WSShell.SpecialFolders.Item("Nethood")
        $ShortcutPath = Join-Path -Path $NetHood -ChildPath "$strFamilia_IO $Location.lnk"
        $NewShortcut = $WSShell.CreateShortcut($ShortcutPath)
        $NewShortcut.TargetPath = "$PathFile\$strFamilia_IO" 
        $NewShortcut.Save()        
}

function GetFIO { #�������� �������_��
        param ($UserName)
        $UserFilter = "(&(objectCategory=User)(samAccountName=$UserName))"
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.Filter	= $UserFilter
        $ADUserPath = $Searcher.FindOne()
        $ADUser = $ADUserPath.GetDirectoryEntry()
        $ADDisplayName = $ADUser.DisplayName.ToString()
        $strFIO = $ADDisplayName.Split("")
        $strLastName = $strFIO[0]
        $strFirstName = $strFIO[1]
        $strSecondName = $strFIO[2]
        $strFamilia_IO = $strLastName + "_" + $strFirstName.Substring(0, 1) + $strSecondName.Substring(0, 1) 
        return $strFamilia_IO
}

function GetGroupInfo { #�������� ������� � ������
        param ($Group)
        $GroupName=$Group.Substring(3)     
        $GroupName=($GroupName -split ',')[0]  
        $Search = New-Object DirectoryServices.DirectorySearcher "LDAP://DC=VR,DC=NET"
        $GroupFilter= "(&(objectCategory=group)(cn=$GroupName))" 
        $Search.Filter = $GroupFilter
        $Search.SearchScope = 2
        $result = $Search.FindOne()
        $obj = $result.GetDirectoryEntry()
        $GroupInfo = $obj.psbase.properties.info 
        return $GroupInfo
}

function GetPrintSerName { #�� ���������� ����_printserver �������� ��� �����-�������
        param($AliasPrintSer)
        $NslookupRes = (nslookup $AliasPrintSer)
        $DNSPrintSerName = $NslookupRes[3]
        $temp=$DNSPrintSerName.substring(9)
        $PrintSerName = $temp.Remove($temp.Length-7)
        return $PrintSerName
}

function GetRoomName { #�������� ��� �������� �� ����� ����������
        param()
        $ComputerName = $env:computername
        $FindPrinterName = $ComputerName.Substring(4)
        $FindPrinterName = $FindPrinterName.Remove($FindPrinterName.Length-4) #�������� ���
        return $FindPrinterName
}

function AddPrinter { #��������� �������. isDef- ����� �� ������� ���������
       param($AddPrinterName, $isDef)
        $net = new-object -com wscript.network
        $net.AddWindowsPrinterConnection("$AddPrinterName")
        if ($isDef -eq $true) {
        $net.SetDefaultPrinter("$AddPrinterName") 
        }
        
}

