# Stock Analysis!

Hi! This project was inspired by the **Gamestop** Fiasco in 2019. Like many, I got interested in trading when the pandemic, resulting in the biggest bull market we've had in a while. Upon losing some money, I realized I needed more data to make smart and calculated decisions.

## What does this project do?

It's a very simple project. It scrapes options data from **Barchart.com** which then is transformed, analyzed, and summarized into a end-of-day email report.

## How did you do it?

I used *Python*, *Selenium*, *BeautifulSoup*, and a handful of other libraries to complete this project. The script opens a Chrome browser, loads to the appropriate URL, waits for the data load, then scrapes the data into a *Pandas* DataFrame. From there, the data is analyzed into a report which would be sent out in an email using a SMTP library.
