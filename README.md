# Pre-processing Excel Sheets#

While cleaning data for a data analysis project in R, I came across a very big (> 100 MB) Excel file (xlsx) containing most of the data. There was no codebook, so I naturally wanted to open the file in Excel first, and have a look at it before loading it into R. It took a very long time to open and when it did, Excel often produced a "not enough resources" error whenever I tried to modify the file.

Trying to open the file in R using `read.table()`, `read.xlsx()`, `read.xlsx2()`, or using most of the solutions mentioned [here](http://www.r-bloggers.com/read-excel-files-from-r/), took even longer than opening it in Excel and wasn't always successful. 

In this short guide, I am sharing how I quickly worked around this issue. In particular, how I pre-processed the Excel sheet before reading it into R.

Although my goal was to read the file into R to clean it and analyse it, I learnt some valuable lessons by detouring from the standard flow. I hope you find this guide useful.

## How I pre-processed the Excel file ##

### Reducing the file size

My first objective was to reduce the size of the file so I can manipulate it. I was able to achieve this by doing the following:

1. The file had 6 sheets. **I split each sheet into a separate file**. I made copies of the file, iteratively opened each copy in Excel and deleted all but a single sheet. Trying to copy or move a sheet into another file while it is open in Excel consumed more memory so it was not possible in this case. 

2. I checked the file for any color-coded information in preparation for the next step. In my case, colors did not carry any information, but in case it does, I suggest you **encode the color-coded information into a variable**.  

3. I **cleared formats** of all the cells in the file.

4. I then **cleared all the data validation** rules. These were there to ensure that data entered in certain columns were selected from look-up tables. These look-up tables were available as sheets in the same file. 

5. **All formulas were substituted by values**. Columns were copied and pasted in-place as values.

6. **All conditional formatting formulas were deleted**.

The above steps reduced the file size significantly. It is important to mention that the first 4 steps paved the way to steps 5 and 6. Copying and pasting columns (step 5) was impossible without reducing the file size considerably first. 

### Data cleansing

I could have loaded the new Excel file resulting from step 5, into R, but I decided to do more cleaning in Excel first:
   
7. I **removed empty columns between populated ones**, so that all the variables are stacked to the left.

8. The first row represented the variable names. They were not in English, so to avoid character encoding issues while loading the data into R and to make it easier to manipulate the variables, I **translated the variable names to English.**

9. **Some records (observations) were invalid, missing required variables. These were filtered and deleted.** This could have been easily done in R after the data was loaded, but it was worth doing in Excel since it contributed a lot to reducing the file size, and in some cases, could be an essential step in order to load the data into R. The number of records went down from more than 100,000 to about 9,000. 

10. Since the version of Excel I had had an issue saving the file in Unicode, I saved the file as csv and changed the encoding in Notepad. To do that, I **replaced all the `","` characters in the file into `" - "`.** This enabled me to preserve the number of columns in the csv file while using `","` as a delimiter. You might want to choose a different substitution or delimiter based on the type of data you are dealing with.
 
11. Since the paragraph direction in Excel cells depend on the language of the text in the cell, the sheet had mixed paragraph directions. This could cause columns in the csv file to move before or after another adjacent column, so I **set the paragraph direction for all columns to be the same**.

The file was finally saved in csv format, re-encoded in Notepad, and loaded into R. But the tidy Excel sheet was also shared with the data providers which they liked and lead to improving the data entry process.

## License 

<a rel="license" href="http://creativecommons.org/licenses/by/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by/4.0/88x31.png" /></a><br />This guide is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by/4.0/">Creative Commons Attribution 4.0 International License</a>.
