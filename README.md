# MLB1stInning-Scapper

Project focused on tracking the 1st inning results of the 2021 MLB season. 

Sporting betting is growing in popularity as more and more states begin to legalize it. Some of the most popular sportsbooks in the industry include Fanduel, Draft Kings, and Barstool. Bettors are able to choose from thousands of different game lines, player props, and game props, but one of the most exhiliarting sports bets in a NRFI!

So what is a NRFI? A No Run First Inning occurs when the first inning of a baseball game results in neither team scoring a run. This type of bet cause high levels of exhilration as well as a immediately sobriety once a run is scored. However, due to the quickness of results and the lucreative odds, this fun bet can be one that if you can find an edge, it could result into some postive gains for a bettor's bankroll.

In order to find this edge, I wanted to first do some analysis on the current data of the 2021 MLB season based on the 1st inning results. With different pitching line ups and stadiums that vary in size, there an be a lot of factors that play into the results and are important to consider when looking at the trends. 

WEBSCRAPING
Using Python and BeautifulSoup, I am able to webscrape baseball-reference.com to collect the first inning data. The script is able to take in new URLs for new game results as the season continues.

VBA ANALYSIS
Once the first inning data is collected and wrote into a Excel file, I am able to write a VBA script to parse through the data to give me a summary of each teams results. In order to not have repetative code, I used a For loop to loop through an array of team names. 
