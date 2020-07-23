AllFolderShortcuts -f $true #Вызов основной функции
 
function AllFolderShortcuts ($f) {
        $Domain = $env:USERDNSDOMAIN #Имя домена
        $CurrentSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name #Имя сайта
        #$CurrentSite = "PRP"
        $UserName = $env:username #Имя пользователя       
        $strFamilia_IO = GetFIO -UserName $UserName        
        $UserGroups = ([ADSISEARCHER]"samaccountname=$UserName").FindOne().Properties.memberof #Список групп текущего пользователя
        $GroupList =''
        foreach ( $Group in $UserGroups ) {
                $GroupList+=$Group #Список групп текщего пользователя в формате string
        } 
        $flag = $false
        $WSShell =New-Object -com WScript.Shell  
        $NetHood = $WSShell.SpecialFolders.Item("Nethood")
        Get-ChildItem *.lnk -Path $NetHood| Remove-Item -Force -Recurse
        foreach ($Group in $UserGroups) #Простматриваем все группы пользователя
        {
            $GroupName=($Group -split ',')[0]
            $GroupName=$GroupName.Substring(3);
            $GroupInfo = GetGroupInfo -Group $Group
            if (($GroupInfo -split '\n')[0] -like "lnk*") 
            { #Ищем группу, у которой первый парметр "lnk*"                 
              
                    $WSShell1 =New-Object -com WScript.Shell  
                    $NetHood1 = $WSShell1.SpecialFolders.Item("Nethood")
                    $DirTitle = ($GroupInfo -split '\n')[2];
                    #Имя общей папки
                    $DirTitle = $DirTitle.Remove($DirTitle.Length-1)
                    $DirPath = ($GroupInfo -split '\n')[1];
                    $DirPath = $DirPath.Remove($DirPath.Length-1)
                    $ShortcutPath1 = Join-Path -Path $NetHood1 -ChildPath " $DirTitle.lnk"
                    $NewShortcut1 = $WSShell1.CreateShortcut($ShortcutPath1)
                    $NewShortcut1.TargetPath = "$DirPath"
                    $NewShortcut1.Save() #Ссылка на общую папку
                
                    #Личные папки пользователей
                    if ((($GroupList -like "*krk_Users*") -and ($CurrentSite -eq "KRK")) -or (($GroupList -like "*prp_Users*") -and ($CurrentSite -eq "PRP"))) { #Группа и текущий сайт совпадают
                        MakeShortcut -PathFile "\\$Domain\dfs\KRK\docs" -strFamilia_IO $strFamilia_IO -Location "" 
                    }
                    elseif (($GroupList -like "*krk_Users*") -and  ($GroupList -like "*prp_Users*") ) { #Пользователь состоит в группах prp_ и krk_ 
                        MakeShortcut -PathFile "\\$Domain\dfs\KRK\docs" -strFamilia_IO $strFamilia_IO -Location "Красноярск" 
                        MakeShortcut -PathFile "\\$Domain\dfs\PRP\docs" -strFamilia_IO $strFamilia_IO -Location "Промплощадка" 
                    }
                    elseif (($GroupList -like "*krk_Users*") -and ($CurrentSite -eq 'PRP')) { #Пользователь из красноярска на площадке
                        MakeShortcut -PathFile "\\$Domain\dfs\KRK\docs" -strFamilia_IO $strFamilia_IO -Location "Красноярск"   
                    }
                    elseif(($GroupList -like "*prp_Users*") -and ($CurrentSite -eq 'KRK')) { #Пользователь с площадки в Красноярске
                        MakeShortcut -PathFile "\\$Domain\dfs\PRP\docs" -strFamilia_IO $strFamilia_IO -Location "Промплощадка"                   
                    }      
                } <#
                if (($GroupInfo -split '\n')[0] -like "prn*") 
                { #Ищем группу с параметром "prn*" 
                     $AliasPrintServer = ($GroupInfo -split '\n')[1] #Имя принт-сервера из параметров
                     $PrintServer ="\\" + ( GetPrintSerName -AliasPrintSer $AliasPrintServer)
                     $tempPrinterName = ($GroupInfo -split '\n')[2] #Имя принтера из параметров
                     
                     $tempPrinterNameres = $tempPrinterName.Remove($tempPrinterName.Length-1)

                     $AddPrinterName = $PrintServer + "\" + $tempPrinterNameres                
                     Write-Host $PrinterName
                        if(($GroupInfo -split '\n')[0] -like "prnd*")   #если prnd - устанавливаем принтер дефолтным
                        {
                                AddPrinter -AddPrinterName $AddPrinterName -isDef $true #Добавляем как дефолтный
                        } 
                        else 
                        {
                                AddPrinter -AddPrinterName $AddPrinterName -isDef $false   #если Prn - устанавливаем как обычный                              
                        }   #>
                #}    
                      
                      
                    
        }   
        $setPrinters = Get-CimInstance -Class Win32_Printer
        $setPrintersList=""
        foreach ( $print in $setPrinters ) {
                $setPrintersList+=$print #Список существующих принтеров [string]
        } 

        $RoomName = (GetRoomName) #Получаем имя кабинета  
        $RoomName = $RoomName.toLower()     
        $AliasPrintServer = $CurrentSite.ToLower()+'_printserver'
        $PrintServer ="\\" +( GetPrintSerName -AliasPrintSer $AliasPrintServer) #Получаем имя принт-сервера
        $Printers = Get-Printer -ComputerName $PrintServer | where-object{$_.devicetype -eq 0} #Получаем список принтеров с сервера
        foreach ($Printer in $Printers)
        {
                $PrinterNameCur = ($Printer -split '"')[1] #Получаем имя текущего принтера
                if ($PrinterNameCur -like "????$RoomName*") #Смотрим совпадает ли он с номером кабинета
                {     if (!($setPrintersList -like "*$PrinterNameCur*"))    #Смотрим не подключен ли он уже
                        {    
                                $AddPrinterName = $PrintServer+"\"+$PrinterNameCur #\\имя принт-сервера\имя принтера
                                if ($flag -eq $true) { #если уже есть дефолтный принтер     
                                 AddPrinter -AddPrinterName $AddPrinterName -isDef $false #доваляем как обычный                               
                                }
                                else { #если дефолтного принтера нет   
                                 AddPrinter -AddPrinterName $AddPrinterName -isDef $true #Добавляем как дефолтный
                                }
                        
                        }
                }
        }
        

}


function  MakeShortcut { #Создание ярлыков 
        param (
           $PathFile, $strFamilia_IO, $Location     #параметр Location - что будет написано после Фамилия_ИО
        )
        $WSShell =New-Object -com WScript.Shell  
        $NetHood = $WSShell.SpecialFolders.Item("Nethood")
        $ShortcutPath = Join-Path -Path $NetHood -ChildPath "$strFamilia_IO $Location.lnk"
        $NewShortcut = $WSShell.CreateShortcut($ShortcutPath)
        $NewShortcut.TargetPath = "$PathFile\$strFamilia_IO" 
        $NewShortcut.Save()        
}

function GetFIO { #Получаем Фамилия_ИО
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

function GetGroupInfo { #Получаем заметки о группе
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

function GetPrintSerName { #по псевдониму сайт_printserver получаем имя принт-сервера
        param($AliasPrintSer)
        $NslookupRes = (nslookup $AliasPrintSer)
        $DNSPrintSerName = $NslookupRes[3]
        $temp=$DNSPrintSerName.substring(9)
        $PrintSerName = $temp.Remove($temp.Length-7)
        return $PrintSerName
}

function GetRoomName { #получаем имя кабинета из имени компьютера
        param()
        $ComputerName = $env:computername
        $FindPrinterName = $ComputerName.Substring(4)
        $FindPrinterName = $FindPrinterName.Remove($FindPrinterName.Length-4) #Получаем имя
        return $FindPrinterName
}

function AddPrinter { #Добвление принтра. isDef- будет ли принтер дефолтным
        param($AddPrinterName, $isDef)
        $net = new-object -com wscript.network
        $net.AddWindowsPrinterConnection("$AddPrinterName")
        if ($isDef -eq $true) {
        $net.SetDefaultPrinter("$AddPrinterName") 
        }
}

