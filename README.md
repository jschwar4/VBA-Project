# VBA-Project
1)	Overview of Project: Explain the purpose of this analysis.
The purpose of this analysis was to determine the volume and yearly return for stocks based on trading data on daily trade volume, starting price, and ending price. While useful on a descriptive basis, this analysis did not attempt to draw conclusions about a stock’s expected future value, correlate input or output variables, or determine causation.
2)	Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Stock Analysis
Stock performance, on net, between 2017 and 2018 fell. Total volume of shares traded fell by 11,616,592 between 2017 and 2018, while the average return on investment fell by 75.8% over the same period. Importantly, this average calculating the average change in ROI does not account for any variations in outstanding shares among the stocks. At a first glance, however, 2018 was not a banner year for these shares in comparison to 2017.
 While the market overall between 2017 and 2018 fell, some stocks did continue to appreciate. ENPH continued to gain value, albeit at a slower rate, while RUN’s growth actually accelerated by 78.4%. Trading volume also increased for AY, CSIQ, FSLR, and SPWR while falling across all other stocks. Notably, RUN and ENPH both had a significantly decreased trading volume while being the only stocks to appreciate in both years.
 
<img width="392" alt="Screen Shot 2021-09-27 at 1 14 08 AM" src="https://user-images.githubusercontent.com/90073490/134957517-cc0bf4c2-415c-45b8-b321-052b302701f7.png">

This correlation may lead to the hypothesis that appreciation is inversely correlated to trading volume. However, the below scatterplots show that there is little correlation between the change in trading volume and year-over-year change in return on investment (R2 = 0.0516). Furthermore, there is an almost astonishing lack of relationship between yearly trading volume and yearly appreciation (R2 = .0001). The data shows that using trading volume as a predictive variable for return or ROI change may not be descriptive.

<img width="471" alt="Screen Shot 2021-09-27 at 1 13 32 AM" src="https://user-images.githubusercontent.com/90073490/134957388-af0ad85e-6dc9-4d15-a66a-23d8b58359a1.png">

<img width="465" alt="Screen Shot 2021-09-27 at 1 13 37 AM" src="https://user-images.githubusercontent.com/90073490/134957469-2ea8dd82-7a01-491a-a14f-ab09ce62c1b3.png">

Refactoring analysis

The refactored script had a significantly reduced run time when compared to the original script. Screenshots are posted below:
2018 Original:

<img width="422" alt="Screen Shot 2021-09-26 at 12 18 31 PM" src="https://user-images.githubusercontent.com/90073490/134957597-32195d7f-c24d-4eb9-94d4-926e256deba9.png">

2018 Refactored:

<img width="422" alt="Screen Shot 2021-09-26 at 12 19 09 PM" src="https://user-images.githubusercontent.com/90073490/134957653-90b6e180-3023-45f4-91e8-15b68c8affdf.png">

2017 Original:

<img width="423" alt="Screen Shot 2021-09-26 at 12 14 45 PM" src="https://user-images.githubusercontent.com/90073490/134957749-0f8bb128-f38c-46cb-93b5-7cdc354c346e.png">

2017 Refactored:

 <img width="420" alt="Screen Shot 2021-09-26 at 12 15 38 PM" src="https://user-images.githubusercontent.com/90073490/134957766-f7132d56-3dd8-47de-b064-fade3524e49b.png">

3)	Summary: In a summary statement, address the following questions.
a)	What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code are significant. Refactoring leads to fewer operations and less working memory used, leading to faster run times and a greater capacity to handle large datasets. For code that involves iterative loops, increasing the efficiency of the loops or eliminating them altogether reduces runtime at an exponential rate.
Of course, the downside is that it involves rewriting code. This introduces the potential for errors or misunderstandings which could lead to code that doesn’t work at all. Refactoring also could cause bugs, which may take significant additional time to chase down.
b)	How do these pros and cons apply to refactoring the original VBA script?
In application to the VBA script, the code reduced the runtime of the code for each year significantly. This is likely due to the elimination of the nested for loop in the refactored code, which would have exponentially decreased the number of operations required for the script to run. Instead of using a for loop to store the ticker increment, the refactored code uses an index. Use of an index in place of a for loop leads to far fewer times that the program must iterate through a section of code, improving code performance.

Original loops:

<img width="626" alt="Screen Shot 2021-09-27 at 1 53 10 AM" src="https://user-images.githubusercontent.com/90073490/134957869-c158151f-7093-4235-b899-f878a25f6903.png">

Refactored loop:

<img width="767" alt="Screen Shot 2021-09-27 at 1 53 34 AM" src="https://user-images.githubusercontent.com/90073490/134957918-e838c3dd-5bb4-4fb9-a662-d6c22228385d.png">

Of course, this could only have been possible with a time investment, not only to determine the best way to refactor the code but also to debug what resulted. There were several instances where the index wasn’t initialized or incremented correctly, leading to compilation errors or other erroneous results. These were fixed, with a bit of effort.
