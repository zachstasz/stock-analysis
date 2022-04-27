# Stock Analysis

## Project Overview
Steve is researching stock trends for his parents. I've already created a macro that will analyze the stock data he gave me, but we want to improve it so he can analyze larger sets of data. 

1. Refactor the VBA code
2. Measure the performance of the code.
3. Take note of pros and cons of refactoring the code

## Results
In order to change the code to loop over all the data one time, I

#### Changed the output variables to be arrays

![refactor_array](https://user-images.githubusercontent.com/103209236/165620839-fb0ff445-3d7c-46f3-a009-4c45902300d8.PNG)

#### Started the tickerVolume at 0 for every loop

![refactor_volume_loop](https://user-images.githubusercontent.com/103209236/165620951-cfe1facd-e621-479b-9ddc-2e63eb879737.PNG)

#### Looped through all the rows

![refactor_rows](https://user-images.githubusercontent.com/103209236/165621027-d4e4ecdb-4d64-4f61-b9f6-615f04589cfc.PNG)

#### Increased the volume of the tickerVolumes by using the tickerIndex variable as the index

![refractor_index](https://user-images.githubusercontent.com/103209236/165621127-2698416f-39df-4d28-b695-4f14398f14a9.PNG)

By looping through all the data at once, this dramatically decreased the amount of time the code took to run.

Here are the original times the code took to ran,

![og_2017](https://user-images.githubusercontent.com/103209236/165624639-fc39ccb1-c1b3-4ca9-a8d2-65a0e6f2c04d.PNG)
![og_2018](https://user-images.githubusercontent.com/103209236/165624644-0f682cdb-7e3b-445b-9865-10f570a450d3.PNG)


And here are the run times with the refactored code,

![VBA_Challenge_2017](https://user-images.githubusercontent.com/103209236/165624649-0dd0f7ba-6062-40cf-ae05-83f5ac36ae98.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/103209236/165624655-268c5eec-3f2f-49b9-9a71-00126802aa75.PNG)


For the stocks themselves, many of them had positive returns in 2017, but only two stocks had a net positive return in 2018; ENPH & RUN.

![Stocks_2017](https://user-images.githubusercontent.com/103209236/165623116-05f4d907-30c7-4f8a-bd53-bcbb756f2911.PNG)
![Stocks_2018](https://user-images.githubusercontent.com/103209236/165622918-7092b481-1668-467f-9996-d2c7aec55b14.PNG)

RUN went from 5.5% return in 2017 to 84% return in 2018, so I'd recommend RUN stock for Steve's parents.

## Summary

In general, refactoring can provide several advantages and disadvantages.
  - Advantages
    - Clean and easy to read
    - Improves performance
    - Contains the code for easy changes
  - Disadvantages
    - Can be confusing for someone who didn't work on the code
    - Can be time consuming if done at the end
    - Can introduce new errors and bugs

For Steve's VBA code, refactoring increased the speed in which the code ran. It also cleaned up the code and made it easier to read. Since this code was small and I wasn't pressed for time, there wasn't any disadvantages in refactoring this VBA code.
  
  
