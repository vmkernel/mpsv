# Path to input XML file
$strXmlDataFile = 'D:\tmp\vm20180515_xml\vm20180515.xml';

# Loading XML from file
[xml]$xmlData = Get-Content -Path $strXmlDataFile -Encoding UTF8;

# Getting XML with vacancies list
$xmlVacancies = ($xmlData.ChildNodes | where Name -eq 'VOLNAMISTA').VOLNEMISTO;

# Converting XML to CSV
$arrResultingVacancies = @();
foreach ($xmlVacancy in $xmlVacancies) {
    
    $objVacancy = New-Object PSObject -Property ([ordered]@{

        # PRACPRAVNI_VZTAH tag skipped due to unknown purpose
        # VHODNE_PRO tag skipped due to unknown purpose
        # URAD_PRACE (updated by) tag skipped due to no need
        # AGENTURA_PRACE_PRIDELENI tag skipped due to unknown purpose
        # PRACOVISTE tag skipped due to unknown purpose

        VacancyId = $null
        VacancyBlueCardPoints = $null
        VacancyLastChanged = $null
        VacancyContactMethod = $null
        ProfessionId = $null
        ProfessionName = $null
        JobName = $null
        CompanyName = $null
        CompanyId = $null
        ShiftId = $null
        ShiftName = $null
        EducationMinLevel = $null
        EducationMinLevelName = $null
        ContactFirstName = $null
        ContactLastName = $null
        ContactPhone = $null
        ContactEmail = $null
        AddressRegionId = $null
        AddressRegionName = $null 
        AddressCity = $null
        AddressDistrict = $null
        AddressStreet = $null
        AddressBuilding = $null
        AddressOffice = $null
        AddressPostCode1 = $null
        AddressPostCode2 = $null
        AddressAttendee = $null
        SalaryPeriod = $null
        SalaryMin = $null
        SalaryMax = $null
        TimeOfEmployment = $null
        JobDescription = $null
        BlueCardPocetVmProMk = $null
        BlueCardCelkemVmProMk = $null
        BlueCardVmRezervProMk = $null
        BlueCardVmRezervProPodanMk = $null
        BlueCardVmRezervProVyhovMk = $null
        BlueCardVmRezervProVydanMk = $null
        WorkingCardPocetVmProZm = $null
        WorkingCardCelkemVmProZm = $null
        WorkingCardVmRezervProZm = $null
        WorkingCardVmRezervProPodanZm = $null
        WorkingCardVmRezervProVyhovZm = $null
        WorkingCardVmRezervProVydanZm = $null
        WorkingCardJenDalsiZamZm = $null
        IndustryId = $null 
        IndustryName = $null 
        RequiredSkills = $null
        RequiredLanguage = $null 
        RawXmlData = $null
    });


    $objVacancy.RawXmlData = $xmlVacancy.InnerXml.Replace( "`n", '' )

    $objVacancy.VacancyId = $xmlVacancy.uid # Vacancy ID
    $objVacancy.VacancyBlueCardPoints = $xmlVacancy.celkemVm # Blue Card points (?)
    $objVacancy.VacancyLastChanged = $xmlVacancy.zmena # Last change date and time
    $objVacancy.VacancyContactMethod = $xmlVacancy.jakKontaktovat # Preferred contact method

    $objVacancy.ProfessionId = $xmlVacancy.PROFESE.kod # Job classification code
    $objVacancy.ProfessionName = $xmlVacancy.PROFESE.nazev # Job classification name
    $objVacancy.JobName = $xmlVacancy.PROFESE.doplnek # Job name

    $objVacancy.CompanyName = $xmlVacancy.FIRMA.nazev # Company name
    $objVacancy.CompanyId = $xmlVacancy.FIRMA.ic # Company id

    $objVacancy.ShiftId = $xmlVacancy.SMENNOST.kod # Interchangeability code
    $objVacancy.ShiftName = $xmlVacancy.SMENNOST.nazev # Interchangeability name

    $objVacancy.EducationMinLevel = $xmlVacancy.MIN_VZDELANI.kod # Minimum level of education code
    $objVacancy.EducationMinLevelName = $xmlVacancy.MIN_VZDELANI.nazev  # Minimum level of education name

    $objVacancy.ContactFirstName = $xmlVacancy.KONOS.jmeno # Contact first name
    $objVacancy.ContactLastName = $xmlVacancy.KONOS.prijmeni # Contact last name
    $objVacancy.ContactPhone = $xmlVacancy.KONOS.telefon # Contact phone number
    $objVacancy.ContactEmail = $xmlVacancy.KONOS.email # Contact email

    $objVacancy.AddressRegionId = $xmlVacancy.PRACOVISTE.okresKod # Address region code
    $objVacancy.AddressRegionName = $xmlVacancy.PRACOVISTE.okres # Address region name
    $objVacancy.AddressCity = $xmlVacancy.PRACOVISTE.obec # Address city
    $objVacancy.AddressDistrict = $xmlVacancy.PRACOVISTE.cobce # Address district
    $objVacancy.AddressStreet = $xmlVacancy.PRACOVISTE.ulice # Address street
    $objVacancy.AddressBuilding = $xmlVacancy.PRACOVISTE.cp # Address building
    $objVacancy.AddressOffice = $xmlVacancy.PRACOVISTE.co # Address office
    $objVacancy.AddressPostCode1 = $xmlVacancy.PRACOVISTE.psc # Address post idx
    $objVacancy.AddressPostCode2 = $xmlVacancy.PRACOVISTE.posta # Address post idx2
    $objVacancy.AddressAttendee = $xmlVacancy.PRACOVISTE.nazev # Address Attendee

    $objVacancy.SalaryPeriod = $xmlVacancy.MZDA.typMzdy # Salary interval
    $objVacancy.SalaryMin = $xmlVacancy.MZDA.min # Salary min
    $objVacancy.SalaryMax = $xmlVacancy.MZDA.max # Salary max

    $objVacancy.TimeOfEmployment = $xmlVacancy.PRAC_POMER.od # Employment start date

    if ( -not [System.String]::IsNullOrEmpty( $xmlVacancy.POZNAMKA ) ) {
        $objVacancy.JobDescription = $xmlVacancy.POZNAMKA.Replace("`n", '') # Description (TAG)
    }

    $objVacancy.BlueCardPocetVmProMk = $xmlVacancy.MODRE_KARTY.pocetVmProMk # Blue card
    $objVacancy.BlueCardCelkemVmProMk = $xmlVacancy.MODRE_KARTY.celkemVmProMk # Blue card
    $objVacancy.BlueCardVmRezervProMk = $xmlVacancy.MODRE_KARTY.vmRezervProMk # Blue card
    $objVacancy.BlueCardVmRezervProPodanMk = $xmlVacancy.MODRE_KARTY.vmRezervProPodanMk # Blue card
    $objVacancy.BlueCardVmRezervProVyhovMk = $xmlVacancy.MODRE_KARTY.vmRezervProVyhovMk # Blue card
    $objVacancy.BlueCardVmRezervProVydanMk = $xmlVacancy.MODRE_KARTY.vmRezervProVydanMk # Blue card

    $objVacancy.WorkingCardPocetVmProZm = $xmlVacancy.ZAMEST_KARTY.pocetVmProZm # Worker card
    $objVacancy.WorkingCardCelkemVmProZm = $xmlVacancy.ZAMEST_KARTY.celkemVmProZm # Worker card
    $objVacancy.WorkingCardVmRezervProZm = $xmlVacancy.ZAMEST_KARTY.vmRezervProZm # Worker card
    $objVacancy.WorkingCardVmRezervProPodanZm = $xmlVacancy.ZAMEST_KARTY.vmRezervProPodanZm # Worker card
    $objVacancy.WorkingCardVmRezervProVyhovZm = $xmlVacancy.ZAMEST_KARTY.vmRezervProVyhovZm # Worker card
    $objVacancy.WorkingCardVmRezervProVydanZm = $xmlVacancy.ZAMEST_KARTY.vmRezervProVydanZm # Worker card
    $objVacancy.WorkingCardJenDalsiZamZm = $xmlVacancy.ZAMEST_KARTY.jenDalsiZamZm # Worker card

    $objVacancy.IndustryId = $xmlVacancy.OBOR.kod # Industry code
    $objVacancy.IndustryName = $xmlVacancy.OBOR.nazev # Industry name

    # Filling required skills field
    if ( $xmlVacancy.DOVEDNOST -ne $null ) {

        if ( $xmlVacancy.DOVEDNOST.GetType().BaseType -eq [System.Array] ) {

            $strRequiredSkills = "";
            for ($idx = 0; $idx -lt $xmlVacancy.DOVEDNOST.Count; $idx++ ) {
                if ( $idx -gt 0 ) {
                    $strRequiredSkills += ", ";
                }
                $strRequiredSkills += "$($xmlVacancy.DOVEDNOST[$idx].nazev) ($($xmlVacancy.DOVEDNOST[$idx].popis))";
            }
            $objVacancy.RequiredSkills = $strRequiredSkills;

        } else {
            $objVacancy.RequiredSkills = "$($xmlVacancy.DOVEDNOST.nazev) ($($xmlVacancy.DOVEDNOST.popis))";
        }

    }
    # Filling required language fiels
    if ( $xmlVacancy.JAZYK -ne $null ) {

        if ( $xmlVacancy.JAZYK.GetType().BaseType -eq [System.Array] ) {

            $strRequiredLanguages = "";
            for ($idx = 0; $idx -lt $xmlVacancy.JAZYK.Count; $idx++ ) {
                if ( $idx -gt 0 ) {
                    $strRequiredLanguages += ", ";
                }
                $strRequiredLanguages += "$($xmlVacancy.JAZYK[$idx].nazev) ($($xmlVacancy.JAZYK[$idx].uroven))";
            }
            $objVacancy.RequiredLanguage = $strRequiredLanguages;

        } else {
            $objVacancy.RequiredLanguage = "$($xmlVacancy.JAZYK[$idx].nazev) ($($xmlVacancy.JAZYK[$idx].uroven))";
        }

    }
    $arrResultingVacancies += $objVacancy;
}

$objXmlDataFile = Get-Item -Path $strXmlDataFile;
$strCsvFilePath = "$($objXmlDataFile.DirectoryName)\$([system.io.path]::ChangeExtension($objXmlDataFile.Name, '.csv'))";

$arrResultingVacancies | select `
    BlueCardVmRezervProMk, `    
    VacancyId, `
    VacancyLastChanged, `
    IndustryName, `
    IndustryId, `
    ProfessionName, `
    ProfessionId, `
    JobName, `
    SalaryMin, `
    SalaryMax, `
    SalaryPeriod, `
    CompanyName, `
    CompanyId, `
    ShiftName, `
    EducationMinLevelName, `
    RequiredSkills, `
    RequiredLanguage, `
    JobDescription, `
    TimeOfEmployment, `
    ContactFirstName, `
    ContactLastName, `
    ContactPhone, `
    ContactEmail, `
    AddressCity, `
    AddressDistrict, `
    AddressStreet, `
    AddressBuilding, `
    AddressOffice, `
    AddressAttendee, `
    BlueCardPocetVmProMk, `
    BlueCardCelkemVmProMk, `
    BlueCardVmRezervProPodanMk, `
    BlueCardVmRezervProVyhovMk, `
    BlueCardVmRezervProVydanMk, `
    WorkingCardPocetVmProZm, `
    WorkingCardCelkemVmProZm, `
    WorkingCardVmRezervProZm, `
    WorkingCardVmRezervProPodanZm, `
    WorkingCardVmRezervProVyhovZm, `
    WorkingCardVmRezervProVydanZm, `
    WorkingCardJenDalsiZamZm, `
    RawXmlData | `
Export-Csv -Force -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Path $strCsvFilePath;