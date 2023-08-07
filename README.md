# Currency converter

This is a desktop app developed with Visual Basic for Applications. The purpose of the project was to create a currency converter that works with real-time data fetched from the following website (currency API): https://www.xe.com/currencytables/?from=USD&date=2021-07-14 

<br/>

<img src="https://github.com/DanielHzp/WebScrappingApp-CurrencyConverter-VB/assets/124480168/be0deca5-1b12-43b0-b3ae-03547c4cd7d0" width="700" height="400">

<br/>

<br/>
<br/>

The app lets the user track the currency exchange behaviour and compare the converted values with a log of estimations and a graphical analysis.
The user form is executed in runtime in Excel when the user clicks an action button in the main worksheet.
<br/>

<br/>

## Layout
When the user form is loaded the conversion currencies are updated automatically and the current date is set as default:

<br/>

<img src="https://github.com/DanielHzp/WebScrapingCurrencyConverter-Vb/assets/124480168/a80ebf31-6f2c-4011-bac1-f29a9fd62ae4" width="700" height="450">



<img src="https://github.com/DanielHzp/WebScrapingCurrencyConverter-Vb/assets/124480168/5ba3c13c-4273-43dc-991d-0250ad4fe270" width="700" height="450">




When all input values are added the user must click on 'Update Currencies' to get the conversion rates according to the selected date and the currencies chosen in the combo boxes. 
<br/>

<br/>

## Update Currencies

Based on the input date, a visual basic method is executed to fetch the currency conversion rates directly from the web service. The following syntax creates the query connection using the necessary parameters to pull the dataset:

  
<br/>
<img src="https://github.com/DanielHzp/WebScrapingCurrencyConverter-Vb/assets/124480168/89708e75-4456-4375-9ae9-7c8dee3fe32e" width="800" height="400">

<br/>

<br/>

In order to handle any connection runtime error, a try-catch block triggers display messages if needed with an 'On Eror GoTo' form command. However, some dates may not be available if the website provider has internal constraints or fails to update the conversion metadata. In this case, the conversion button will not work and a error pop-up window will alert the user.

i.e

<br/>

<img src="https://github.com/DanielHzp/WebScrapingCurrencyConverter-Vb/assets/124480168/9a6a7bf2-b45e-4bc0-af75-aa263a4417ef" width="700" height="450">

<br/>

<br/>
<br/>

## Convert Currency

The conversion output will be displayed in the 'Converted Amount' field when the user clicks the button. If the conversion rates haven't been updated ('Refresh Currencies' is not clicked) all currencies will be automatically updated forcing the previous method to be executed. The following syntax illustrates how the conversion is estimated using the fetched data:

<br/>

![image](https://github.com/DanielHzp/WebScrapingCurrencyConverter-Vb/assets/124480168/51d36532-f7bb-44fa-afb1-7c56d564e73f)

<br/>


This result will be rendered in the user form as follows:

<br/>


<img src="https://github.com/DanielHzp/WebScrapingCurrencyConverter-Vb/assets/124480168/821ac86e-b099-4aaf-8472-d56943fd6279" width="700" height="450">

<br/>

<br/>

## View Conversions Log
 
Every conversion request is saved in an internal log of changes and It is possible to view the selected currency behaviour of the last 30 days (previous to the input date) in a data plot. This will be automatically displayed in a spreadsheet when the user clicks 'Plot Last 30 Days' button. 

<br/>

 i.e Using sample data for USD - GBP conversion behaviour over time:
 
<br/>

<br/>


<img src="https://github.com/DanielHzp/WebScrapingCurrencyConverter-Vb/assets/124480168/a347e90f-1614-4649-ad09-d191f70e945c" width="700" height="450">

<br/>


<br/>

<br/>

In order to dynamically populate this plot with different conversions, a query connection is created recursively per daily rate. The following syntax partially illustrates a visual basic method that extracts the daily rates and iterates over the last 30 days estimating the conversions:

<br/>

<br/>



![image](https://github.com/DanielHzp/WebScrapingCurrencyConverter-Vb/assets/124480168/28287c3d-19fb-4833-bf29-0544f38cf9b0)

<br/>

<br/>


In order to handle any connection runtime error, a try-catch block launches display messages if needed with an 'On Eror GoTo' form command. However, some dates may not be available if the website provider has internal constraints or fails to update the conversion metadata. In this case, the conversion button will not work and a error pop-up window will alert the user.

<br/>

<br/>

## .

### Usage

Import the .bas files to a VB/VBA Excel developer editor (module1 and RibbonX customization are optional) and add a macro button to execute the user form.
Add three worksheets without display name which will be automatically updated when the data is pulled from the XE website.

<br/>

<br/>























