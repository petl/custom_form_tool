# Customs form tool for https://www.post.at/en/n/f/customs-form
Script to automatically create customs forms for the Austrian Post (Österreichische Post - Zollformular). It takes a csv file and feeds the data into the customs form via the chromium chromedriver.

<img src='https://github.com/petl/customs_form_tool/blob/main/screenshot/postat_customs.png' width='100%'>

#### How to get it running

I've tested this on Ubuntu 22.04, but other operating systems like Windows or MacOS should work as well. You need to install chromium:
> apt install chromium

Then download the fitting version of chromedriver:
https://chromedriver.chromium.org/downloads

> wget https://chromedriver.storage.googleapis.com/index.html?path=110.0.5481.30

After that start the tool with 
> python3 main_v3.py

It opens up a chromium, inputs all the data from address.csv and downloads the created form to the ./pdf folder. 

#### Disclaimer

I'm in no way associated with the Austrian Post (Österreichische Post) and this is no official tool. I take no responsibility and am not liable for anything that you do with this script. I have created this, because manually inputting data into this form takes ages and there is no official way to input multiple recipients automatically that I'm aware of.  
