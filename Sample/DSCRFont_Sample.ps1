$output = 'C:\MOF'

Configuration DSCR_Font_Sample
{
    Import-DscResource -ModuleName DSCR_Font
    Node localhost
    {
        cFont Add_Font_Sample
        {
            Ensure   = 'Present'
            FontName = 'Noto Serif (TrueType)'
            FontFile = 'C:\temp\NotoSerif-Regular.ttf'
        }

        cFont Remove_Font_Sample
        {
            Ensure   = 'Absent'
            FontName = 'Noto Sans Regular (TrueType)'
            FontFile = 'NotoSans-Regular.ttf'
        }
    }
}

DSCR_Font_Sample -OutputPath $output
#Test-DscConfiguration -Path  $output -Verbose
Start-DscConfiguration -Path  $output -Verbose -Wait -Force

