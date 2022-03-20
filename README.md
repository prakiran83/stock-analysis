# stock-analysis
## Overview of Project

This project, "All Stocks Analysis Refactored," is designed for the students to learn the technique of 
restructuring an existing codes. Although we are using or re-using the existing body of codes, we do have to alter 
some of its component, without impacting the results.

We are using the data from the exercise in our module to retrieve the return value of 12 stocks, from two different years.
We are refactoring the dataset, through VBA, to determine if there are any significant advantages of refactoring codes.

## Results
Initially, I copied the code for my "All Stocks Analysis" exercise to a new Module, in order to use it as a reference to format the "AllStocksAnalysisRefactored" exercise. 


![codes](https://user-images.githubusercontent.com/100887673/159148985-46a3865f-67a1-4bed-af68-12b85d669299.png)


I opened the README.md file, with the assignment, in notebook and then copied the data in the file. I then pasted the data in the VBA editor, which I was using during the challenge. Then I just followed the instruction that was supplied in the file and started refactoring the codes.

According to my analysis, stocks had higher return in 2017 than 2018. As you can see in the images below, 11 out of 12 stocks had some kind of return in 2017, compared to 2 out of 12 in 2018. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/100887673/159149202-bed3b9bb-362d-4310-8f27-59ca29fa3bc4.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/100887673/159149208-97766392-b92d-4b41-a93a-c2708a9862a5.png)
Also, i noticed that once the refactoring was completed the tables were loading much faster than before. For example, as it can be seen in the images below, the 2017 data ran in 0.75 seconds after refactoring, compare to 0.8359375 seconds previously.![VBA_Challenge_2017](https://user-images.githubusercontent.com/100887673/159149259-f34f7087-bfc5-4ab1-8734-a3216d4cf878.png)

![VBA_Challenge_2017_refac](https://user-images.githubusercontent.com/100887673/159149266-198b6cc5-5956-41d7-bcee-edb5d4a2b54e.png)

same was the result for 2018.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/100887673/159149281-3d18b797-a23e-46a4-ac7c-72d76a0617dc.png)

![VBA_Challenge_2018_refac](https://user-images.githubusercontent.com/100887673/159149282-a4618528-47b7-4c0a-b876-1f91f6415c56.png)

## Summary

##  1.what are the advantages or disadvantages of refactoring code?
I believe the main advatage of refactoring code is that it declutters the script. This, in turn, makes the script readable for anyone who is trying the work on it. Also, as this exercise shows, refactoring can also make your program run faster.
The only disadvatge that I think refactoring might have is that if not done properly, it might cause errors in the functionality of the programme.

##  2.How do these pros and cons apply to refactoring the original VBA script?
Refactoring a script is to restructure a code without compromising the output. The pros to refactoring a script is that it gives us an opprtunity to declutter the existing code and make it understandable to anyone that will be using it. It can be considered as paving a graveled road for a smoother ride. But, on the other hand, if the person refactoring is not capable, he or she can create more errors in the exisitng script and make the program fail.
