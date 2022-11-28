# CHART HELPER
#### A Python web-scraping script for music charts with Excel sheet organization
#### Version 1.0

  Chart Helper was created by myself in 2020 because I am a music nerd. I love listening to tons of albums, and with that I also love internet music charts. Some websites have very nice chart interfaces, like RateYourMusic. Others do not... in particular, I was super frustrated with using ProgArchives. ProgArchives is a very outdated website with some notable flaws in its structure. For example, album descriptions are not consistently formatted, making metadata finding a hassle. Another problem is that ProgArchives is very cumbersome and slow loading. So, instead of using the website itself, I began to create Excel sheets corresponding to certain charts that I liked. However, this was a lengthy process that was prone to error. That is when I decided to learn about web scraping and to take a crack at automating this process.
  
  Originally, I created Chart Helper for my friends and I to track our listening habits. I wanted to be able to see who's listened to what off from some given chart, and also to check whether certain albums exist on streaming services (or, if I have them on my disc). Right now, Chart Helper is limited to ProgArchives charts. Right now the content of the Excel sheets it generates is fixed to the format that I use, but that will hopefully change in the future. This program makes us of the python packages OpenPyxl, SpotiPy, and LXML.
  
  As of right now, I have developed a rough CLI for the program. To "install" it, clone my repo and use conda to make an environment with the right packages (or, reference the list and manually install). This was developed on a M1 MacBook, but there really shouldn't be any reason this can't work on another Mac with the right adjustments. There might need to be some adjustments for Windows or Linux, so I cannot say that this program supports it yet.
  
# WARNING
  USE THIS SCRIPT AT YOUR OWN RISK. This script was made with the utmost respect for ProgArchives' servers. This includes manual sleep commands between http requests. If you are producing a large batch of files, consider editing the time for every sleep command to be bigger. 

# SETUP

  To use my Chart Helper right now, you will need a couple things. You will need an app that can open .xlsx files (I use LibreOffice and Google Sheets). You will need a Spotify API client ID and secret (these are not hard to get). If you want this script to reference your digital music library on file, you need to provide a link to a directory formatted like so: [directory]/[first letter of artist]/[artist]/[album]. I cannot guarantee that this script will search your library in a dignified and sophisticated manner, but this works for most cases. The first time you run this script, you will need to run the setup command.
 
# THE FORMAT

  The way this script formats Excel might seem a bit quirky, but I thought it conveys the important information I'd want to know before checking out an album I haven't heard: the ranking, its popularity, whether I have listened to it before (or not), and whether it's easily accessible. I've used with my close friends and they've grown accustomed to it, so I like to think it's worth promoting. I don't want to expand the types of info about each album too much because what fits on a screen is more important to me. The way this script keeps track of listening is fairly portable: it stores what it can when you give it that info, but it's easy to add new people and update things. This doesn't replace something like last.fm for tracking your listening, but I want to expand this format in the future to work with sites like RYM (or others?). The format of the Excel sheet is below
  
| Album | Artist | Year | Genre | Rating | # of Ratings | Duration | Country | Listener(s) status(es) | On Spotify? | Bought? | In Library?
--- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
  
# COMMANDS

## setup
  Sets up the necessary configuration file for you.
#### parameters

*id*: a Spotify API client ID number.

*secret*: a Spotify API client secret.

*directory*: a link to the parent file of your music directory. If you don't have one, just make sure this is a valid directory, that shouldn't crash this script.

(optional) *--albumlib | -al* [FILENAME]: manually enter the name of the file which python will store the album library on disc.

(optional) *--chartlib | -cl* [FILENAME]: manually enter the name of the file which python will store the album library on disc.

## scanchart
  Scanchart takes a link to a ProgArchives chart and scans it into an internal chart which can later be printed or updated.
#### parameters:

*chartname*: a string representing the name of the chart being scanned (descriptive, up to you. For example "Top 100 Prog Electronic Albums")

*link*: a string representing the link of the chart being scanned. I usually get this from progarchives itself, but I might create a chart link generator in the future.

*listeners*: list however many people you want on this chart. This really should be optional unless you choose the write option, but that's for a future update. This doesn't mean you can't generate the chart with other people's names in the future, it's just a starting place. 

(optional)*--update-spotify | -us*: If an album on the chart is already in the system, this tells the script to overwrite its Spotify status. If unspecified, the script will not modify this field.

(optional)*--update-duration | -ud*: Similarly, if an album on the chart is already in the system, this tells the script to overwrite its duration.

(optional)*--write | -w* [filename]: Specifies an Excel .xlsx file to write this chart into. If unspecified, the script will only store the chart in the system.

(optional)*--new | -n* [filename]: Specifies a new filename to write the workbook to. The new file will have all of the charts from the file specified by --write

## readchart
  Opens a .xlsx file and reads in a specific chart. Note that this does not create an internal chart, but it syncs the information from each album to the album library (eg, listening info, spotify status, duration...)
#### parameters:

*filename*: a string representing the .xlsx file from which a chart will be read

*chartname*: a string representing the name of the chart to be read

*listeners*: IMPORTANT: this must correspond to the listener sections of the chart, otherwise this will not work. Provide the names of the people (IN ORDER) that are on the chart to be read

(optional) *--overwrite | -o*: if specified, the script will overwrite any interally stored albums' info with the info on the chart. Otherwise, this script does not touch the duration or spotify status (because they are unpredictable, so a user should be able to specify them manually if need be)

## updatechart
  Internally updates a chart with its stored link (or a new link that you provide)
#### parameters:

*chartname*: a string representing the name of a chart from the chart library.

*listeners*: a list of names to appear on the charts. This is really only necessary if you specify --write.

(optional) *--newlink* [LINK]: if specified, will update the chart against the provided URL.

(optional) *--update-spotify*: if specified, updates the Spotify status of the albums on the chart.

(optional) *--update-duration*: if specified, updates the duration of the albums on the chart.

(optional) *--write | -w* [FILENAME]: if specified, the script will write the updated chart to the provided filename (.xlsx)

## writechart
  Writes an internally stored chart into an Excel file
#### parameters:

*chartname*: a string representing the name of a chart from the chart library.

*listeners*: a list of names to appear on the written chart.

*filename*: a string representing the file to write or save the chart to.

(optional) *--new | -n*: THIS IS DIFFERENT FROM OTHER --new FLAGS. This flag specifies that when the script writes the chart into an Excel file, it will write it in as explicitly a new chart (it will not overwrite an existing chart).

## getrejectchart
  Given an internal chart, a reject chart is created out of albums that once were on the chart but are not anymore. Note that a reject chart is a special chart which must always be regenerated when its master chart is updated. It cannot be updated in and of itself. This chart will be split into sections by dates when albums were rejected.
#### parameters:

*chartname*: a string representing the name of a chart from which the reject chart will be derived.

*listeners*: a list of names to appear on the written chart

(optional) *--write | -w *[FILENAME]: if specified, the script will write the reject chart to the file specified (.xlsx)

(optional) *--new | -n*: if specified, the script will create a new sheet on the file when writing (instead of overwriting an existing sheet)

## updateworkbook
  Given a workbook (.xlsx) file, this script will run through every sheet in the book and update them. Every sheet MUST have the same listeners. If they do not, please use updatechart instead.
#### parameters:

*filename*: a string representing the name of the file (.xlsx) to be processed.

*listeners*: a list of the names which appear on these charts (ORDER MATTERS).

(optional) *--update-spotify*: if specified, updates the Spotify status of the albums on each chart.

(optional) *--update-duration*: if specified, updates the duration of the albums on each chart.

(optional) *--write | -w* [FILENAME]: if specified, will write the chart to the specified filename (.xlsx).

(optional) *--new | -n* [FILENAME]: if specified, will write the chart to a new filename (.xlsx) [META COMMENT: between you and me, I think I may have made the same function twice. My bad. Will be fixed in the next version. Maybe don't use this?].

## readworkbook
  Given a workbook (.xlsx) file, this script will read every sheet in the book and sync that information to the internal library of albums.
#### parameters:

*filename*: a string representing the name of the file (.xlsx) to be processed.

*listeners*: a list of the names which appear on these charts (ORDER MATTERS).

(optional) *--overwrite | -o*: if specified, will overwrite both the Spotify status and duration of every album specified.

## scancharts
  Given a file formatted with each row like so: [Chart title here]::[Chart link here], this script will batch scan these charts.
#### parameters:

*chartfile*: a file with pairs representing the chart title and chart link.

*listeners*: a list of the names which appear on these charts.

(optional) *--update-spotify*: if specified, updates the Spotify status of the albums on each chart.

(optional) *--update-duration*: if specified, updates the duration of the albums on each chart.

(optional) *--write | -w* [FILENAME]: if specified, will write the chart to the specified filename (.xlsx).

(optional) *--new | -n* [FILENAME]: if specified, will write the chart to a new filename (.xlsx) [META COMMENT: yep, I did it twice].

CHANGELOG:
1.0:

* Uploading everything for the first time

* Created basic documentation

