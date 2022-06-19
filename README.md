# AllStocks Analysis with VBA Code and Measure Peformance

Analysis of AllStocls Campaign

## Overview of Project amd prupose 

Steve and his Parents to expand the dataset to include the stire stok market ovet the past few year not only few stock but they wants entire stock. and they want to know which stock was better performance in 2017 and 2018. 

So we started first with fewer dataset and collect total volume of those stock and see what return comes. we create botton to clear sheet and run data analysis to obtain each year information but its not enough so we refector code and edit some code and make that data set for whole stocks and run faster.
 


## Results

- Compare stocks performace between 2017 and 2018 as well as the execution times.


In Comparision of stock data 2017 and 2018 we found out that in 2017 mostly all stock was good and give profit except one stock which was "TERP" was giveing loss. highest gain was in "DQ" and in 2018 mostly all was negative but only "ENPH" and "RUN" stocks was gain profit or you can call successful out of those two "Run" was more percentage 84%. i am atttachig two image so can glance through

![refactored 2017 data](https://user-images.githubusercontent.com/103727169/174497193-90f7f3be-51dd-4909-9943-e84d0b0c0d5b.png)
![refactored 2018 data](https://user-images.githubusercontent.com/103727169/174497216-4bfdf94c-e7f9-4bd8-87aa-c06832b14148.png)


in execution times of original script was taking longer time to run, after we refactored the script it execution time much faster then original script which image are follow to compare how faster refactored script run

this out put you can see first image is original runtime and second image of refactored run time in 2017. we can see original run time 1.171875 second and refectored time is 0.164625 second we got more faster then original script. and we can see blelow for 2018 too.

![original 2017 runtime](https://user-images.githubusercontent.com/103727169/174497264-7a326a8b-d137-4831-964a-921946f627f7.png)
![refactored 2017 runtime](https://user-images.githubusercontent.com/103727169/174497276-ef132fab-24c8-4e36-be1e-ea0e9d3c4fc3.png)


This out put you can see first image is original runtime and second image of refactored run time in 2018

![original2018runtime](https://user-images.githubusercontent.com/103727169/174497347-3a48ce5b-7d3e-46b4-864a-1bdab985e3b3.png)
![refactored 2018 runtime](https://user-images.githubusercontent.com/103727169/174497357-57ba7efa-3af7-4597-9711-1fcdee0aa84b.png)


also we use some formatting row of coloum return red and green so client can see which one is bettr and which one is not using following code. add some color change font style using merge cells top header.

![color](https://user-images.githubusercontent.com/103727169/174497293-6671d29f-dac6-481e-9b57-3d23e5b8265f.png)
![formating](https://user-images.githubusercontent.com/103727169/174497302-4ead99b7-84c2-4694-a116-900946243ddb.png)


also we use some formatting row of coloum return red and green so client can see which one is bettr and which one is not using following code. add some color change font style using merge cells top header.

![color](https://user-images.githubusercontent.com/103727169/174497293-6671d29f-dac6-481e-9b57-3d23e5b8265f.png)
![formating](https://user-images.githubusercontent.com/103727169/174497302-4ead99b7-84c2-4694-a116-900946243ddb.png)





### Summary

   - what are the advantage or disavrantages of refactoring code ?

    The advantage of refactoring code to get run our analysis much faster then original script and to get data set for all stock and to view some images to compare on 


   - How do these Pros and cons apply to refactoring the original VBA script?

    In original script we give steve and his parents for analysis for 2017 and 2018, after that we refactored script with editing fewar code to make the code more      efficient using less memory and improving logic code to make it easier. in this refactoring we edit create tickerIndex as variable and create out put arrays "tickerVolumes","tickrStartingPrices" and "TikerEndingPrices". i run for loop over all the rows in spreadsheet. and inside for loop i wrote code to increase the current stock volume with ticer?index using following code 


    ![refactred_code1](https://user-images.githubusercontent.com/103727169/174497498-7eecaa96-b1f0-484d-b32d-05ead9f47510.png)

    
   then using if then statement i check the current row is the last row with tickerIndex.


    ![currentrowislastrow](https://user-images.githubusercontent.com/103727169/174497510-303b393d-6427-4225-bf61-502dad4a53bf.png)

    


   we got result as follow by creating botton to run for our analysis and output of both year 2017 and 2018 esult as follow images

   ![refactored 2017 data](https://user-images.githubusercontent.com/103727169/174497660-9fe66328-3fe8-4ed2-9a94-0b0927a02cc1.png)
   ![refactored 2018 data](https://user-images.githubusercontent.com/103727169/174497682-8d7bf128-9234-4554-a598-0be84c39f835.png)
   
