# VBA-Project
1)	Overview of Project: Explain the purpose of this analysis.
The purpose of this analysis was to determine the volume and yearly return for stocks based on trading data on daily trade volume, starting price, and ending price. While useful on a descriptive basis, this analysis did not attempt to draw conclusions about a stock’s expected future value, correlate input or output variables, or determine causation.
2)	Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Stock Analysis
Stock performance, on net, between 2017 and 2018 fell. Total volume of shares traded fell by 11,616,592 between 2017 and 2018, while the average return on investment fell by 75.8% over the same period. Importantly, this average calculating the average change in ROI does not account for any variations in outstanding shares among the stocks. At a first glance, however, 2018 was not a banner year for these shares in comparison to 2017.
 While the market overall between 2017 and 2018 fell, some stocks did continue to appreciate. ENPH continued to gain value, albeit at a slower rate, while RUN’s growth actually accelerated by 78.4%. Trading volume also increased for AY, CSIQ, FSLR, and SPWR while falling across all other stocks. Notably, RUN and ENPH both had a significantly decreased trading volume while being the only stocks to appreciate in both years.
![image](https://user-images.githubusercontent.com/90073490/134867575-c56e65c0-0e6c-490e-800b-b993fa95b894.png)
This correlation may lead to the hypothesis that appreciation is inversely correlated to trading volume. However, the below scatterplots show that there is little correlation between the change in trading volume and year-over-year change in return on investment (R2 = 0.0516). Furthermore, there is an almost astonishing lack of relationship between yearly trading volume and yearly appreciation (R2 = .0001). The data shows that using trading volume as a predictive variable for return or ROI change may not be descriptive.![image](https://user-images.githubusercontent.com/90073490/134867602-300f7fea-b7f1-4e83-8b0f-2bf20a6096ab.png)
![image](https://user-images.githubusercontent.com/90073490/134867629-c73a16ca-e9fc-4862-83d4-81cf53d2bc58.png)
Refactoring analysis
The refactored script had a significantly reduced run time when compared to the original script. Screenshots are posted below:
2018 Original:
 ![image](https://user-images.githubusercontent.com/90073490/134867683-0dcfd83a-e9b5-476d-94a5-420a1fb2e28d.png)

2018 Refactored:
 ![image](https://user-images.githubusercontent.com/90073490/134867696-bbaecd76-4cd6-46fc-8680-b7ecff9e9891.png)

2017 Original:
 ![image](https://user-images.githubusercontent.com/90073490/134867716-30ebde02-b44a-4b4a-b146-f9b66c62e639.png)

2017 Refactored:
 
![image](https://user-images.githubusercontent.com/90073490/134867659-fd92b176-d933-49ac-b387-87c929a762e9.png)
3)	Summary: In a summary statement, address the following questions.
a)	What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code are significant. Refactoring leads to fewer operations and less working memory used, leading to faster run times and a greater capacity to handle large datasets. For code that involves iterative loops, increasing the efficiency of the loops or eliminating them altogether reduces runtime at an exponential rate.
Of course, the downside is that it involves rewriting code. This introduces the potential for errors or misunderstandings which could lead to code that doesn’t work at all. Refactoring also could cause bugs, which may take significant additional time to chase down.
b)	How do these pros and cons apply to refactoring the original VBA script?
In application to the VBA script, the code reduced the runtime of the code for each year significantly. This is likely due to the elimination of the nested for loop in the refactored code, which would have exponentially decreased the number of operations required for the script to run. Instead of using a for loop to store the ticker increment, the refactored code uses an index. Use of an index in place of a for loop leads to far fewer times that the program must iterate through a section of code, improving code performance.

Original loops:
![image](https://user-images.githubusercontent.com/90073490/134867759-1d27321d-ddde-40c4-8c03-8a13d6a0b08b.png)

Refactored loop:
![image](https://user-images.githubusercontent.com/90073490/134867779-c2de4fc0-1c98-4591-bc21-811fc3cdee9c.png)
Of course, this could only have been possible with a time investment, not only to determine the best way to refactor the code but also to debug what resulted. There were several instances where the index wasn’t initialized or incremented correctly, leading to compilation errors or other erroneous results. These were fixed, with a bit of effort![image](https://user-images.githubusercontent.com/90073490/134867806-5e08e9c0-023b-4f0b-bb3b-fb1c1e8cc475.png)
