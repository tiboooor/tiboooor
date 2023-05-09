# .SYNOPSIS
#  Generiert FATCA XML und SEI XML anhand einer CSV Datengrundlage und einer vorgegebenen XML Schema Definition.
# 
# .NOTES 
# ===================================================================================================================
#  Version: 1.2
#  Erstellt am:             21.03.2023 15:57 Uhr
#  Erstellt von:            Tibor Blasko (UEX17050)
#  Betrieb:                 Raiffeisen Schweiz Genossenschaft
#  Ort:                     St. Gallen
#  Departement:             IT
#  Abteilung:               Lernende Plattformentwicklung
#  Betreuende Abteilung:    Systemtechnik
#  Gruppe:                  Windows Server
#  Betreuer:                Robin Calis, Stefan Aschenbrenner
#  Dateiname:               1.2anotherJsonTest.ps1
# ===================================================================================================================
# 
# .DESCRIPTION
#  Erstellt anhand einer CSV Datengrundlage und einer vorgegebenen XML Schema Definition die für die US-Steuerbehoerde notwendigen FATCA und SEI XML Dateien.
# 
# .INPUTS
#  CSV mit den Grunddaten.
#  XML Schema Definition für FATCA und SEI XML.
# 
# .OUTPUTS
#  FATCA und SEI XML Dateien.
# 
# .LINK
#  https://www.swissbanking.ch/de/themen/steuern/fatca
#  https://www.sif.admin.ch/sif/de/home/bilateral/lander/vereinigen-staaten-von-amerika-usa/fatca-abkommen.html

Param (
    [Parameter(Mandatory = $false,
        Position = 0,
        HelpMessage = "Path to CSV source file",
        ValueFromPipeline = $true)]
    [String]$CsvFile = "H:\RCH_tblasko\Skripts\FATCA\Datengrundlage\Datengrundlage FATCA Gruppenersuchen.csv",

    [Parameter(Mandatory = $false,
        Position = 1,
        HelpMessage = "FATCA XML destination path",
        ValueFromPipeline = $true)]
    [String]$XmlDestinationPath = "H:/RCH_tblasko/Skripts/FATCA/testXML/"
)

Function Convert-JsonToHashtable {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0)]
        [String]$JsonObject
    )

    BEGIN {
        Add-Type -AssemblyName "System.Web.Extensions"
        Write-Host "Converting JSON Object to hashtable..." -ForegroundColor Cyan
    }
    PROCESS {
        $jsonSerializer = [System.Web.Script.Serialization.JavaScriptSerializer]::new()
        $jsonSerializer.Deserialize($JsonObject, [System.Collections.Hashtable])
    }
    END {
        $jsonSerializer = $null
        Write-Host "Hashtable created" -ForegroundColor Cyan
    }
}

Function Remove-EmptyNodes {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
        	Position = 0)]
        [System.Xml.XmlDocument]$XmlDoc
    )

    BEGIN {
        [int]$nodesRemovedCounter = 0
        Write-Host "Removing empty nodes..." -ForegroundColor Cyan
    }
    PROCESS {
        Do {
            $nodesRemoved = $false

            ForEach ( $node in $XmlDoc.SelectNodes("//*[not(node())]") ) {
                $node.ParentNode.RemoveChild($node) | Out-Null
                $nodesRemovedCounter += 1
                $nodesRemoved = $true
            }

            ForEach ( $node in $XmlDoc.SelectNodes("//*[not(*) and not(normalize-space())]") ) {
                $node.ParentNode.RemoveChild($node) | Out-Null
                $nodesRemovedCounter += 1
                $nodesRemoved = $true
            }
        } While ( ($nodesRemoved) -eq ($true) )
    }
    END {
        Write-Host "$($nodesRemovedCounter) empty nodes removed." -ForegroundColor Cyan
    }
}

Function Ignore-XmlComments {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0)]
        [String]$XmlFileName
    )
    
    BEGIN {
        Write-Host "Ignoring comments..." -ForegroundColor Cyan
        $xmlPath = [System.IO.Path]::Combine($XmlDestinationPath, $XmlFileName)
        $xmlSettings = [System.Xml.XmlReaderSettings]::new()
        $xmlSettings.IgnoreComments = $true
    }
    PROCESS {
        $xmlReader = [System.Xml.XmlReader]::Create($xmlPath, $xmlSettings)
        $newXml = [System.Xml.XmlDocument]::new()
        $newXml.Load($xmlReader)
    }
    END {
        $xmlReader.Dispose() | Out-Null
        $newXml.Save($xmlPath)
        Write-Host "Done" -ForegroundColor Cyan
    }
}

Function Create-XmlElement {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
        	Position = 0)]
        [System.Xml.XmlDocument]$XmlDoc,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
        	Position = 1)]
        [System.Xml.XmlLinkedNode]$XmlNode,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
        	Position = 2)]
        [String]$XmlElementName,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
        	Position = 3)]
        [String]$XmlNameSpace,

        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
        	Position = 4)]
        [String]$XmlNameSpaceLink,
        
        [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        Position = 5)]
        [String]$XmlInnerText
    )

    BEGIN {
        Write-Host "Creating new XML element $($XmlElementName)" -ForegroundColor Cyan
    }
    PROCESS {
        [System.Xml.XmlElement]$newXmlElement = $XmlDoc.CreateElement($XmlNameSpace, $XmlElementName, $XmlNameSpaceLink)
        $XmlNode.AppendChild($newXmlElement) | Out-Null
        $newTextNode = $XmlDoc.CreateTextNode($XmlInnerText)
        $commentToRemove = $XmlDoc.CreateComment('comment to remove')
    }
    END {
        $newXmlElement.AppendChild($newTextNode) | Out-Null
        $newXmlElement.AppendChild($commentToRemove) | Out-Null
    }
}

Function Import-Xsd {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true,
            Position = 2,
            HelpMessage = "Path to FATCA XML schemas",
            ValueFromPipeline = $true)]
        [String]$XsdPath
    )
    
    BEGIN {
        [String[]]$xmlNodes = @()
        $xsdFiles = Get-ChildItem -Path $XsdPath
    }
    PROCESS {
        ForEach ( $file in $xsdFiles ) {
            Write-Host "Collecting schema information from $($file.Name)..." -ForegroundColor Cyan
            $xsdContent = Get-Content -Path "$($file.FullName)"
            $xmlSchema = [System.Xml.XmlDocument]::new()
            $xmlSchema.LoadXml($xsdContent)
            $xmlSchemaElementNames = $xmlSchema.SelectNodes("//*[@name]")
            $xmlNodes += $xmlSchemaElementNames.Name
        }
    }
    END {
        Write-Host "Done" -ForegroundColor Cyan
        return $xmlNodes
    }
}

Function Create-FatcaXml {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0)]
        [System.Collections.Hashtable]$Hashtable,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0)]
        [String[]]$XmlNodes
    )
    
    BEGIN {
        [System.Collections.Hashtable]$fatcaXmlNameSpaces = @{
            stf = "urn:oecd:ties:stf:v4"
            iso = "urn:oecd:ties:isofatcatypes:v1"
            sfa = "urn:oecd:ties:stffatcatypes:v2"
            ftc = "urn:oecd:ties:fatca:v2"
            xsi = "http://www.w3.org/2001/XMLSchema-instance"
        }
        [String[]]$fatcaMessageSpec = @(
            "SendingCompanyIN",
            "TransmittingCountry",
            "ReceivingCountry",
            "MessageType",
            "Warning",
            "Contact",
            "MessageRefId",
            "CorrMessageRefId",
            "ReportingPeriod",
            "Timestamp"
        )
        [String[]]$fatcaReportingFI = @(
            "ResCountryCode",
            "TIN",
            "Name",
            "Address",
            "FilerCategory",
            "DocSpec"
        )
        [String[]]$fatcaReportingFIAddress = @(
            "CountryCode",
            "AddressFix"
        )
        [String[]]$fatcaReportingFIAddressFix = @(
            "Street",
            "BuildingIdentifier",
            "SuiteIdentifier",
            "FloorIdentifier",
            "DistrictName",
            "POB",
            "PostCode",
            "City",
            "CountrySubentity"
        )
        [String[]]$fatcaReportingGroup = @(
            "Sponsor",
            "Intermediary",
            "NilReport",
            "AccountReport",
            "PoolReport"
        )
    }
    PROCESS {
        Write-Host "Creating XML files..." -ForegroundColor Cyan

        $fatcaXml = [System.Xml.XmlDocument]::new()
        $fatcaXml.LoadXml("<?xml version=`"1.0`" encoding=`"utf-8`"?>
            <ftc:FATCA_OECD xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:ftc=`"urn:oecd:ties:fatca:v2`" xmlns:sfa=`"urn:oecd:ties:stffatcatypes:v2`" xmlns:iso=`"urn:oecd:ties:isofatcatypes:v1`" xmlns:stf=`"urn:oecd:ties:stf:v4`" xsi:schemaLocation=`"urn:oecd:ties:fatca:v2 FatcaXML_v2.0.xsd`" version=`"2.0`">    
                <ftc:MessageSpec>
                    <!--comment to remove-->
                </ftc:MessageSpec>
                <ftc:FATCA>
                    <ftc:ReportingFI>
                        <!--comment to remove-->
                    </ftc:ReportingFI>
                    <ftc:ReportingGroup>
                        <!--comment to remove-->
                    </ftc:ReportingGroup>
                </ftc:FATCA>
            </ftc:FATCA_OECD>
        ")

        For ( $i = 0; $null -ne ($xmlNodes[$i]); $i ++ ) {
            [String]$fatcaXmlElementName = $xmlNodes[$i].Split("_")[0]

            If ( $fatcaXmlElementName -in $fatcaMessageSpec ) {
                If ( $null -eq $fatcaXml.FATCA_OECD.MessageSpec.$fatcaXmlElementName ) {
                    Create-XmlElement -XmlDoc $fatcaXml -XmlNode $fatcaXml.FATCA_OECD.MessageSpec -XmlElementName $fatcaXmlElementName -XmlNameSpace "sfa" -XmlNameSpaceLink $fatcaXmlNameSpaces.sfa -XmlInnerText "test"
                }
            }
            ElseIf ( $fatcaXmlElementName -in $fatcaReportingFI ) {
                If ( $null -eq $fatcaXml.FATCA_OECD.FATCA.ReportingFI.$fatcaXmlElementName ) {
                    Create-XmlElement -XmlDoc $fatcaXml -XmlNode $fatcaXml.FATCA_OECD.FATCA.ReportingFI -XmlElementName $fatcaXmlElementName -XmlNameSpace "ftc" -XmlNameSpaceLink $fatcaXmlNameSpaces.ftc -XmlInnerText "test"
                }
            }
            ElseIf ( $fatcaXmlElementName -in $fatcaReportingGroup ) {
                If ( $null -eq $fatcaXml.FATCA_OECD.FATCA.ReportingGroup.$fatcaXmlElementName ) {
                    Create-XmlElement -XmlDoc $fatcaXml -XmlNode $fatcaXml.FATCA_OECD.FATCA.ReportingGroup -XmlElementName $fatcaXmlElementName -XmlNameSpace "ftc" -XmlNameSpaceLink $fatcaXmlNameSpaces.ftc -XmlInnerText "test"
                }
            }

            [int]$percent = ($i / $xmlNodes.Length) * 100
            If ( $percent -eq 99 ) {
                $percent = 100
            }

            Write-Progress -Activity "Creating FATCA XML..." -Status "$($percent)% completed..." -PercentComplete $percent
        }
    }
    END {
        [String]$fatcaXmlFileName = "$($newHashTable.SEIDossierNumber + "-FATCA.xml")"
        Remove-EmptyNodes -XmlDoc $fatcaXml
        $fatcaXml.Save("$($XmlDestinationPath)/$($fatcaXmlFileName)")
        Ignore-XmlComments -XmlFileName $fatcaXmlFileName
        Write-Host "FATCA XML created" -ForegroundColor Cyan
    }
}

Function Create-SeiXml {
    Param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0)]
        [System.Collections.Hashtable]$Hashtable,
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0)]
        [String[]]$XmlNodes
    )

    BEGIN {
        Write-Host "Function for SEI XML"
        
        $seiXml = [System.Xml.XmlDocument]::new()
        $seiXml.LoadXml("$($newHashTable.'Header SEI XML')
                <!--comment to remove-->
            </FatcaEditedInformationMetaData>
        ")
    }
    PROCESS {
        [int]$i = 0
        Do {
            [String]$seiXmlElementName = $xmlNodes[$i].Split("_")[0]
            [System.Xml.XmlElement]$newSeiXmlElement = $seiXml.CreateElement($seiXmlElementName)
            $seiXml.FirstChild.AppendChild($newSeiXmlElement) | Out-Null

            ForEach ( $seiKey in ($newHashTable.Keys) ) {
                If ( ($seiKey) -like "*$($seiXmlElementName)*" ) {
                    [String]$seiValue = "$($newHashTable.$seiKey)"
                }
            }

            [String]$seiValue = "$($newHashTable.$seiXmlElementName)"
            $newSeiTextNode = $seiXml.CreateTextNode("$($seiValue)")
            $newSeiXmlElement.AppendChild($newSeiTextNode) | Out-Null

            [int]$percent = ($i / $xmlNodes.Length) * 100
            If ( $percent -eq 99 ) {
                $percent = 100    
            }

            Write-Progress -Activity "Creating SEI XML..." -Status "$($percent)% completed..." -PercentComplete $percent
            $i ++
        } While ( $null -ne ($xmlNodes[$i]) )
    }
    END {
        [String]$seiXmlFileName = "$($newHashTable.SEIDossierNumber + ".xml")"
        Remove-EmptyNodes -XmlDoc $seiXml
        $seiXml.Save("$($XmlDestinationPath)/$($seiXmlFileName)")
        Ignore-XmlComments -XmlFileName $seiXmlFileName
        Write-Host "SEI XML created" -ForegroundColor Cyan
    }
}

$csv = Import-CSV -Path $CsvFile -Delimiter ";"

ForEach ($element in $csv) {    
    $newJsonObject = ConvertTo-Json -InputObject $element
    [System.Collections.Hashtable]$newHashTable = Convert-JsonToHashtable -JsonObject $newJsonObject
    [String[]]$fatcaXmlNodes = Import-Xsd -XsdPath "H:/RCH_tblasko/Skripts/FATCA/importSchemas/FATCA"
    [String[]]$seiXmlNodes = Import-Xsd -XsdPath "H:/RCH_tblasko/Skripts/FATCA/importSchemas/SEI"
    Create-FatcaXml -Hashtable $newHashTable -XmlNodes $fatcaXmlNodes
    Create-SeiXml -Hashtable $newHashTable -XmlNodes $seiXmlNodes
}

# ===================================================================================================================
# 
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣀⣀⣀⣀⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⢰⠖⢶⠖⢦⡀⠀⢀⣠⠤⠒⠚⠉⠉⠁⠀⠀⠀⠀⠉⠉⠙⠒⠦⣄⡀⠀⣠⢤⣄⢤⡄⠀⠀⠀⠀⠀⠀⠀
# ⣀⣠⡴⢻⠟⢹⡄⠈⣇⠈⣇⠀⣧⠖⠋⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠉⣻⠁⢰⠇⢰⠃⢠⡗⢲⡦⢄⡀⠀
# ⣹⠀⡇⢸⡀⢸⡅⠀⢹⠀⢹⠄⢹⠀⠀⠀⠋⠀⠀⣀⠔⠲⣄⠀⢀⡤⠲⣄⠀⠀⠀⠀⢠⡏⠀⡏⠀⣾⠀⢸⡇⢸⠀⢸⠉⡇
# ⢿⠀⢳⠀⣇⠈⣇⠀⣸⠀⠏⠀⣸⠀⠰⣄⠀⠀⠀⠧⣀⠀⠈⠙⠁⠀⣠⠼⠃⠂⠀⠀⠸⡇⠀⢧⠀⣇⠀⢸⠁⢸⠀⣸⠀⡇
# ⠘⣇⠈⢧⠘⠆⠈⢻⠏⠀⢠⣴⠇⠀⠀⢠⣤⠀⠀⣀⡼⠧⠄⠘⠣⠽⠧⡄⠀⠐⠇⣀⠀⢿⣶⣄⣆⢹⠖⠃⢰⠃⢠⠇⣸⠃
# ⠀⠘⢦⣌⣁⣀⣀⣠⣾⣷⡿⢻⠀⠀⠠⠀⠀⠀⣀⣻⣕⣩⣞⣉⣳⣄⣵⠋⠀⠀⠀⠀⠀⠀⢳⣝⡿⣷⣤⣄⣀⣰⣫⡴⠃⠀
# ⠀⠀⠀⠉⠛⠛⠛⠿⣏⡉⠓⣊⣹⣆⠤⠖⠚⠉⠁⠀⠀⢠⡀⠀⠀⠀⠈⠉⠙⠒⠤⣀⠀⢴⣿⣿⣍⠽⠛⠛⠛⠋⠁⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⢠⡿⣯⠞⠋⢀⡀⠀⠀⢀⣀⠤⠤⠖⠷⢤⡤⣀⣀⣀⠀⠀⠀⠈⠙⢾⣋⡩⣽⡄⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⣠⠟⠡⠀⠀⣀⠤⠒⠉⢁⣀⠠⠤⠀⠐⠛⠳⠤⢤⣈⠉⠑⠺⢵⡆⠀⠙⢦⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⣼⠃⠀⣀⠴⠋⠀⣤⠴⠋⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠙⠲⣄⠀⠉⠳⣄⠀⠱⡄⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠘⣿⠀⡼⠃⠠⢀⡞⠁⠀⠀⢑⢦⡄⠀⠀⠀⠀⠀⠀⣤⢤⡆⠀⠈⠳⡀⠀⠘⣆⠀⢹⡄⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⢿⣼⠣⢨⡀⠘⡆⠀⢠⣾⣿⣭⣷⠀⠀⠀⠀⠀⠀⣿⣿⣯⣕⡆⠀⣹⡦⢀⣘⡆⢸⠇⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠻⣍⠠⠁⣄⣹⡀⠘⣿⣿⣿⡽⠀⢴⣒⡶⠄⠀⢿⣿⣿⡯⠃⢠⠇⠁⠐⢢⣿⡎⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠙⠲⢤⣄⣳⡀⠀⠀⠀⠐⠦⣤⠿⣤⡼⠂⠀⠀⠀⢀⣴⠃⣈⣣⡤⠖⠉⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠓⠦⢤⣀⣀⣙⣶⡋⢀⣤⣠⣤⣶⡯⠼⠿⣏⡇⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣴⡿⣿⣿⣏⣿⣿⡟⢭⡧⢭⡷⣞⣶⠟⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣴⣿⠟⠁⢻⡿⣯⣿⣏⡿⠀⠹⡄⠹⡄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣴⣟⣉⣻⡄⢀⡀⠑⠊⠀⠁⣀⣀⠀⣳⠀⢱⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢰⢯⡧⠼⢻⡋⠉⠉⠉⠉⠉⠉⠉⣉⣹⡇⠀⠀⢳⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢠⣞⡀⠀⠀⠀⠻⣽⣷⣶⣶⣶⣿⣿⡿⠋⠀⠀⠀⠈⣇⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠉⠙⢦⠀⠀⣀⣈⣽⣿⡤⢨⣿⣭⣤⠤⢤⠤⠖⠋⠉⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠉⠉⠻⠥⠷⠹⠀⠛⠒⠣⠿⠂⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# 
# ===================================================================================================================
