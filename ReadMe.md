# Motivation
This repository will store my project for scraping data off [RightMove](https://www.rightmove.co.uk/)
# What can be found here?
A program that was written to go through particular areas, identified through zone codes, and find houses that come up after some basic filtration for price, number of beds and 
type of property. When it would find these houses it would also do a rent search in their area to see what the rental demand is. After applying some maths it would print in a
word document the houses that were over a certain Return On Investment threshold.
# Current status
This program was written a while back and an update to the rightmove site has caused incompatibility.
# Current ideas for important improvements
  1. Selenium is not something meant for efficient web scraping but rather for something like automating website testing routines as it mimics a human's interration with the
website. Due to this it wouldn't scale up well as the demand for data is increased and also will get you banned by the websites automatic anti-scraping monitor. These exist to
protect the website against not so thoughful scrapers. Yes, I did get a temporari warning ban. A more efficient web scraper needs to be used or potentially the website API,
which RighMove offers.
  2. As the demand for data increases it becomes impractical for a person to scroll through word files to find the information required. An SQL database can be used to store the
data which would enable modern data science tools to analyze the data and provide summaries, recommendation and trends.
# What do you need to run this?
  1. Python 3
  2. selenium
  3. Chromedriver
  4. python-docx
# Aknowledgements
I would like to thank RightMove for being tolerant of people just learning data scraping and not immediately handing out permanent bans.
