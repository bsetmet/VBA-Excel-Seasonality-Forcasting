# VBA-Excel-Seasonality-Forcasting 

This annylasis uses the Multiplicative Decomposition Model to Deseasonalize the data to smooth.
Yt=Trend(levels)*Season(t)*Irregular(t)

## Benefits

Measuring seasonality of a data set can help organizations plan for the future. Seasonality can be caused by various factors such as weather vacation or holidays. 

## Public Domain

This project is in the public domain within the United States, and
copyright and related rights in the work worldwide are waived through
the [CC0 1.0 Universal public domain dedication](https://creativecommons.org/publicdomain/zero/1.0/).

All contributions to this project will be released under the CC0 dedication. By submitting a pull request, you are agreeing to comply with this waiver of copyright interest

For more information, see [license](https://github.com/seakintruth/VBA-Excel-Seasonality-Forcasting/blob/master/LICENSE.md).

## Privacy

All comments, messages, pull requests, and other submissions received including this GitHub page may be subject to archiving requirements. See the [Privacy Statement](http://www.archives.gov/global-pages/privacy.html) for more information.

## Contributions

Contributions are welcome. If you would like to contribute to the project you can do so by forking the repository and submitting your changes in a pull request. You can submit issues using [GitHub Issues](https://github.com/seakintruth/VBA-Excel-Seasonality-Forcasting/issues).

## Deployment
...

## Limitations

Note: the excel limitations that consern this tool are as follows:
"Cover" sheet Cells C2, F2, F5,  and F7 are static (Can't be moved) 
"Data" Sheet must start with three uniqe values, "Category, Date, Values" (the fourth Column is created at runtime) 
"Historical Periods" is the number of periods to display on the chart, all data will be used to calculate the forecast line 
"Dataset Has Values of '0' " is set to False by default if your data set has 0's then is must also have a value for every category on each date entry (or you will need to make your 0's  = 1x10^-20)  
All week Calculations and groupings conform to ISO 8601 i.e. a year has 53 weeks every 5.6338 years (starting in 2004) 
The number of Categories can not exceed 126 The number of Rows can not exceed 1048575

