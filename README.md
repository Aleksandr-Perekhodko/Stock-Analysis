# Stock-Analysis

## Overview of Project
The purpose of this project is to show how VBA is used to create macro's for an Excel document with stock data that has highs, lows, close, and volume. We used the data from these sheets in Excel and used the provided VBA client to create our own sheet to calculate the total daily volume and the return on these stock options. In this project we used a refractored method which means to create a single **Sub()** to and have all the code neatly run togeather instead of making various different **Sub()** to have the code run this macro.
## Results
The results of the differences between the stocks return in 2017 and 2018 are staggering, in 2017 the retrun off all the stock options were up with the exception of TERP which was down -7.2%, and with multiple stock option up over 100%.
### 2017 Stock Options Returns over 100%:
> - DQ : +199.4%
> - ENPH : +129.5%
> - FSLR : +101.3%
> - SEDG : +184.5%
<img width="193" alt="2017_Stock_Analysis" src="https://user-images.githubusercontent.com/97326526/158070437-182fd823-8d9e-4b4d-a87e-7adc82a2ad00.PNG">


Now the opposite can be said about the same stock options in 2018 many of the same stock options in the following year were all down in the returns compared to 2017. Many of the stocks that were up in return the year prior were down in 2018. And even in year 2018 TERP that was down in -5.0% in 2017 was still going down by -7.2% in 2018. RUN also had a larger increase from the year prior with a return increase of 5.5% in 2017 to a 84.0% return increase in 2018 making it one of the only other stock that did well in 2018. Seems that ENPH and RUN are the stock to be keeping an eye for.
### 2018 Stock Options Returns that were over 100% in 2017:
> -  DQ : -62.6%
> - ENPH : +81.9%
> - FSLR : -39.7%
> - SEDG : -7.8%
<img width="193" alt="2018_Stock_Analysis" src="https://user-images.githubusercontent.com/97326526/158070448-24ceb90f-0bcc-4779-ae91-de428c8ec05c.PNG">


## Summary
With refractoring code there are some advantages and disadvanages, the follwing are some of the advantages and disadvantages.
### Advantages of Refractoring Code
> - Easy to look at and define the code
> - Debugging is simply done when in one **Sub()**
> - Better to plan out all your steps when using comments
### Disadvantages of using Refractored code
> - Repeating patterns in code
>  - More work to create a completely new code
>  - Could be more complex than first derived code
### Pros and Cons of refractoring to this Project
The disadvatages and advatages to this project apply, since in the original code it was already done, such as having the formating of the color, and the fonts of the text to use to bring up in the excel sheet. But on the other hand the code is also easier to apply the macro to a button to have everything done with one code instead of having a seperate button to do so or even to apply two macros on to one button. Another advantage of having refractored code that was mentioned was debugging is very easlily done since VBA can identify what seems out of place and u can fix it when trying to run. In this project an advantage of using refractored code is that everything was planned out to make the code work. Planning with comments and showing what it is that you're doing is a very important  step when making a code that you could forget what you did, or even when someone else is looking at it.





