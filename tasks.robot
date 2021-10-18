*** Settings ***
Documentation   A Robot to get data from website and save it in excel
Library         RPA.HTTP
Library         RPA.Browser.Selenium
Library         RPA.Excel.Files
Library         String



*** Keywords ***
Open Excel-sheet's links and get data  
    Open Workbook     Testsheet.xlsx
    Set Active Worksheet    testsheet
    ${contents}=      Read Worksheet as Table   header=True
    Create Worksheet  test_output2
    
    FOR   ${row}  IN  @{contents}
        Open links and get data    ${row}  
    END
    Save Workbook

# +
*** Keywords ***
Open links and get data
    [Arguments]     ${row}
    Open Available Browser   ${row}[Linkki]
    Wait Until Element Is Visible      css:div.field__items
    ${page_contents}=   Get Text  css:div.field__items
    ${links}=   Get Text  css:div.page__additional-information
    ${linksLower}=  Convert To Lower Case   ${links}
    ${contacts}=    Get Lines Containing String   ${page_contents}   0
    ${emails}=      Get Lines Containing String      ${page_contents}    @porinperusturva.fi    case-insensitive
    Close Browser
    &{dictionary}=    Create Dictionary    Name=${row}[Nimi]    Data=${page_contents}   Links=${linksLower}  Numbers=${contacts}  emails=${emails}
    Append Rows To Worksheet  ${dictionary}  header=True
    

# -

*** Tasks ***
Get content from a website
    Open Excel-sheet's links and get data
