# Stock Options Analysis!

Hi! This project was inspired by the **Gamestop** fiasco in 2019. Like many others, I began trading during the pandemic where I realized I'd need more data to make better decisions.

## What does this project do?

It scrapes options data from **Barchart.com** which is then transformed, analyzed, and summarized into an end-of-day email report.

## How did you do it?

I used *Python*, *Selenium*, *BeautifulSoup*, and a handful of other libraries to complete this project. The script opens a Chrome browser, loads to the appropriate URL, waits for the data load, then scrapes the data into a *Pandas* DataFrame. From there, the data is analyzed into a report which would be sent out in an email using a SMTP library.
