# stock-analysis
##  Overview of Project:
This project showcases the scripting features of Visual Basic for Application. \With the tool we will script instructions to analyze stock  \
data for the year 2017, 2018. \
In addition, the result of our analysis will be confined to a result cheet showing the ticker symbols, as tracked, respective their returns. \
Since we are dealing with code, we will refactor it and make it as efficient as can be done.

## Results:
Running the macro before refactoring yielded the expected results: tables and contents identical to reqs.
Metrics captured during run show the following results:
1) when processing the 2017 and 2018 dataset and presenting the summary, the macro ran for ~5 seconds; \
    [Results and Metrics 2017 Stock data](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2017.png)
    [Results and Metrics 2018 Stock data](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2018.png)
3) after refactoring, the content of the reports remain unchnaged. However, the macro crunched the data and produced the report in a shorter \
time. Here are the results:
    [Results and Metrics 2017 Stock data(refactored)](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2017_after-refactoring.png)
    [Results and Metrics 2018 Stock data(refactored)](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2018_after-refactoring.png)
As an observation the refactored code improved runtime by 3 seconds.
I went back to the code and decided to make some additional changes to a version of the code I've kept. I was able to bring the execution of the code to less than a second. See attachments below"
    [Results and Metrics 2017 Stock data(refactored)](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2017_after-refactoring_mine.png)
    [Results and Metrics 2018 Stock data(refactored)](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2018_after-refactoring_mine.png)
    
##Summary:
What are the advantages or disadvantages of refactoring code?
    - Refactoring helps pinpoint flaws of contructs in one's code. It offers great opportunity to improve mechanisms with metrics captured through execution. In this challenge, I have included the 'DoEvents' keyword in areas where significant number crunching was taking place. The idea was to insure that calculations run to completion. I turns out that the addition of 'DoEvents' added a whoping 3 seconds to the overall execution of the code.
    - Refactoring may foster greater collaboration 
How do these pros and cons apply to refactoring the original VBA script?

