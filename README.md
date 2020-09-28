# DiskScraper
## Introduction

A Python script that reads the metadata of every file in the given directory and
all subdirectories. The data is saved as a csv file (using semicolons as
separators). Using the config file, the user can specify which metadata 
categories should be read. Extracting only a few categories greatly increases 
speed.

Project progress can be viewed at: https://trello.com/b/jGlZOZjt/diskcrawler 

## Technologies
* Python3.8.5
* pyinstaller (to create executable)

## Launch
1. Change the values in "config.txt" based on your preferences
2. Launch "disk_scraper_windows.py" or "disk_scraper_windows.exe"

