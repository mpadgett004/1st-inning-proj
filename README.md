# MLB1stInning-Scapper

Created a small web scraper to collect MLB 2021 season 1st inning scores from baseball-reference.com website.

## Setup
> pip install beautifulsoup4

## Process
Using Python and BeautifulSoup, I am able to webscrape baseball-reference.com to collect the first inning data. The script is able to take in new URLs for new game results as the season continues.

## Extract
1. Web scraper utlizes a baseball-reference URL that specifies a date to extract all the links to each game that occured on that date.
2. The scraper then loops through the game URLs to pull the team data and 1st inning scores.

## Transform
Within the Jupyter Notebook, utlizing the .append function in order to concatenate the new dataframe containing the most recent 1st inning scores data to the existing dataframe for historic 1st inning score data.

## Load
Loading all of the date 1st inning score data to a CSV document using 'to_csv'.


VBA ANALYSIS
Once the first inning data is collected and wrote into a Excel file, I am able to write a VBA script to parse through the data to give me a summary of each teams results. In order to not have repetative code, I used a For loop to loop through an array of team names. 
