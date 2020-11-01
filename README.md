# Stock Analysis
## Overview
The purpose of this analysis is to refactor a VBA code that was used to analyze a series of stock markets from a Excel spreadsheet.

## Results

First, we need to analyse the original code to understand what can be done to make it more eficient - to refactor it.
The main issue with the first code is that it has a nested loop that would run through **all the data** each time in order to find the specific ticker and gather its information.

<img width="527" alt="Old_code_1" src="https://user-images.githubusercontent.com/72593264/97816917-99563780-1c5e-11eb-824e-3d990913799f.png">


Because of this, the time that this code would take to run is displayed below, for the years of 2017 and 2018:

![Old_code_time](https://user-images.githubusercontent.com/72593264/97816969-010c8280-1c5f-11eb-80b6-7096f2e84fd7.png)
![old_code_timer_2018](https://user-images.githubusercontent.com/72593264/97817291-fce16480-1c60-11eb-8b4a-46987a6cbcb7.png)


Now, because the data is already sorted by Tickers, it is possible to run this code in a much faster way, using one loop (instead of 2 nested) and using a variable (tickerIndex) to link the results.

Below is a new way to code it:

1) First we create a ticker Index, 
2) then, we create three autoput arrays with the information we would like to gather (Volumes, Starting prices and Ending Prices)
3) we loop only one time throug the whole data, reseting the ticker Index when it changes.

Here is an example of that new code:
<img width="681" alt="New_Code" src="https://user-images.githubusercontent.com/72593264/97817175-41b8cb80-1c60-11eb-9aee-94e411b56e35.png">

And with this simple change we can get a much better time when running the same data set, with the same results:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/72593264/97817216-92c8bf80-1c60-11eb-82d2-4799faba83e6.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/72593264/97817304-12ef2500-1c61-11eb-90f4-9b0a23cf1038.png)

## Summary

There is a clear advantage in succefully recaftoring a code. It can make your script work in much more efficent way.

As can be seen in this example, the exact same results could be achieved almost **5 times faster** with the refactored script. That is a huge difference.
One can argue that the diference is actually not humanly noticiable in this case, but when we transport this refactoring to much bigger data sets, that kind of difference can mean a lot of time saved, which usually translates in saving money.

However, what needs to be taken in consideration is that refactoring code is laborous and also time consuming - and you never know what can be reafctored until you start diggin into the code. In the example above, the new script intorduced a big time saving, but that does not mean the it will happen every time.

In this specific case, it was totally worthy.

