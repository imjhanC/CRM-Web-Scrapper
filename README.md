# CRM-Web-Scrapper
[![Build Workflow](https://github.com/imjhanC/CRM-Web-Scrapper/actions/workflows/build.yml/badge.svg)](https://github.com/imjhanC/CRM-Web-Scrapper/actions/workflows/build.yml)

## Introduction
* This application is able to edit and read the VTiger CRM data in real-time. Moreover, This CRM Web Scraping Application is designed to scrape the product details from the VTiger CRM website. This tool automatically extracts all the product details based on quote ID , helping you keep your CRM database up-to-date with minimal effort.

## How does this application works ?
Regarding the flow of this application 
* There are two flows while interacting with this application which are __editing products__  and __adding products__

__Signing in__
What you need to input in the sign in windows
* Credentials (Username and Password)
* Quote ID 

__Editing products__
> 1. Sign in > Extraction process > Edit / Add windows pops up > Press save button >  Product's details are being altered > Log out

__Adding products__
> 2. Sign in > Extraction process > Edit / Add windows pops up > Press save button > Product's details are being added > Log out 

## Key advantages of this program
> 1. User is able to add multiple products and insert it back to the VTiger CRM web application automatically.
> 2. User is able to add multiple products' subproducts or comments by using the desired products list , however user must know what to put in the desired products list. 
> 3. This application aims to  reduce human error and improve the efficiency of adding product details into the VTiger CRM Web Application 

## Constraint 
> 1. User might need to delete certain row they want themselves.
> 2. User should not interact with the testing browser during the automation process is running because this website is a dynamic website as it needs to scroll to certain position to extract value out from the HTML elements. 
> 3. User should not close all the CRM Web Scrapping website when it is launched.

## Security and Privacy 
* This program will not save any of the products details nor the credentials that inputted into the application. It will destroy and wipe out all the details after the __save__ button is pressed. 

## How to deploy this application ( for maintainance and development )
> 1. Set up a github repository ( recommended )
> 2. Double check all the required files before deploying the py file into executables file by using PyInstaller
> 3. To install the dependencies like ( pyInstaller, tkinter, .... ) , you will need to include the exact or earlier version of all the dependencies namely requirement.txt 
> 4. Next in the local repository's command prompt type in `pyinstaller --onefile --windowed --icon=icon.ico name_of_the_application.py`.
> 5. You might need to configure the `spec` file as it can disable some executables' requirements such as disabled console or fullscreen mode when it is launched.
> 6. The dist and build file should appeared in both local repository and GitHub repository 
> 7. ( To deploy the CI pipeline of transforming py file to exe release by using Git Actions ) Go to Actions then name any __yaml__ file ( E.g. build.yml ) and refer to https://github.com/imjhanC/CRM-Web-Scrapper/blob/master/.github/workflows/build.yml 
> 8. Go back to the repository's dashboard and go to __release__ to download the exe & zip folder 
Summary of what you need for deployment (spec)

## Can't find the exe or zip file ?
* It might be due to windows defender is quarantined the folder or exe file. Follow these steps to allow your computer to download the executable file 
> 1. Go to search and type in Windows Security 
> 2. Once you go in the Windows Security, go to `Virus & threat protection` 
> 3. Then, go to `Protection History` 
> 4. Look for the executables's name and allow it 
> 5. Download the exe file again from the GitHub Repository