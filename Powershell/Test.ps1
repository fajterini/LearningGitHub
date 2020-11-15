# First Script
# NYTimes_COVID-19_data source https://github.com/nytimes/covid-19-data
# 11/15/2020

$url = 'https://raw.githubusercontent.com/nytimes/covid-19-data/master/us.csv'
#Invoke-WebRequest -Uri $url -OutFile c:\temp\us.csv
$DataUS = Import-Csv -Path C:\temp\us.csv
$TimeGenerated = Get-Date

# Generating Graph https://bytecookie.wordpress.com/2012/04/13/tutorial-powershell-and-microsoft-chart-controls-or-how-to-spice-up-your-reports/

[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
$scriptpath = 'c:\temp'

    #Chart object
    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 1920
    $chart1.Height = 1080
    $chart1.AntiAliasing  = 3 # Not working
    $chart1.BackColor = [System.Drawing.Color]::White
    
    # Title 
    [void]$chart1.Titles.Add("US Covid 19 Stats (NY Times Data) - Generated: $TimeGenerated")
    $chart1.Titles[0].Font = "Arial,14pt"
    $chart1.Titles[0].Alignment = "topLeft"

    # Chart area 
    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = "ChartArea1"
    $chartarea.AxisY.Title = "Cases/Deaths"
    $chartarea.AxisY.TitleFont = "Arial,13pt"
    $chartarea.AxisX.Title = "Date"
    $chartarea.AxisX.TitleFont = "Arial,13pt"
    #$chartarea.AxisY.IsLogarithmic = $true
    #$chartarea.AxisX.Interval = 10
    $chart1.ChartAreas.Add($chartarea)

    # Legend 
    $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $legend.name = "Legend1"
    $chart1.Legends.Add($legend)
    

  
    # Data source
    $datasource = $DataUS #| Select-Object -First 70

    # Data series (Active Cases)

    [void]$chart1.Series.Add("Covid-19 cases")
    $chart1.Series["Covid-19 cases"].ChartType = "SplineArea"
    $chart1.Series["Covid-19 cases"].BorderWidth  = 1
    $chart1.Series["Covid-19 cases"].IsVisibleInLegend = $true
    $chart1.Series["Covid-19 cases"].chartarea = "ChartArea1"
    $chart1.Series["Covid-19 cases"].Legend = "Legend1"
    $chart1.Series["Covid-19 cases"].color = "#E3B64C"
    $datasource | ForEach-Object {$chart1.Series["Covid-19 cases"].Points.addxy( $_.date ,$_.cases ) }

    # Data series (Deaths)

    [void]$chart1.Series.Add("Deaths")
    $chart1.Series["Deaths"].ChartType = "SplineArea"
    $chart1.Series["Deaths"].BorderWidth  = 1
    $chart1.Series["Deaths"].IsVisibleInLegend = $true
    $chart1.Series["Deaths"].chartarea = "ChartArea1"
    $chart1.Series["Deaths"].Legend = "Legend1"
    $chart1.Series["Deaths"].color = "#62B5CC"
    $datasource | ForEach-Object {$chart1.Series["Deaths"].Points.addxy( $_.date ,$_.deaths ) }

  

 # Data series (Deaths)
 
 [void]$chart1.Series.Add("Deaths ratio")
 $chart1.Series["Deaths ratio"].ChartType = "SplineArea" #https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.datavisualization.charting.seriescharttype?view=netframework-4.8
 $chart1.Series["Deaths ratio"].BorderWidth  = 1
 $chart1.Series["Deaths ratio"].IsVisibleInLegend = $true
 $chart1.Series["Deaths ratio"].chartarea = "ChartArea2"
 $chart1.Series["Deaths ratio"].Legend = "Legend2"
 #$chart1.Series["Deaths"].color = "#62B5CC"
 $datasource | ForEach-Object {$chart1.Series["Deaths ratio"].Points.addxy( $_.date ,($($_.deaths/$_.cases)*100)) }


# save chart
$chart1.SaveImage("$scriptpath\Covid-19_US.png","png")

