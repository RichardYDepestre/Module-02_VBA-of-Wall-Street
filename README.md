# stock-analysis
##  Overview of Project:
This project showcases the scripting features of Visual Basic for Application. \With the tool we will script instructions to analyze stock  \
data for the year 2017, 2018. \
In addition, the result of our analysis will be confined to a result cheet showing the ticker symbols, as tracked, respective their returns. \
Since we are dealing with code, we will refactor it and make it as efficient as can be done.

## Results:
Running the macro before refactoring yielded the expected results: tables and contents identical to reqs.
Metrics captured during run show the following results:
1) when processing the 2017 and 2018 dataset and presenting the summary, the macro ran for ~7 seconds; \
    [Results and Metrics 2017 Stock data](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2017.png) \
    [Results and Metrics 2018 Stock data](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2018.png)
3) after refactoring, the content of the reports remain unchnaged. However, the macro crunched the data and produced the report in a shorter \
time. Here are the results: \
    [Results and Metrics 2017 Stock data(refactored)](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2017_after-refactoring.png) \
    [Results and Metrics 2018 Stock data(refactored)](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2018_after-refactoring.png)
As an observation the refactored code improved runtime by 3 seconds.
I went back to the code and decided to make some additional changes to a version of the code I've kept. I was able to bring the execution of the code to less than a second. See attachments below: \
    [Results and Metrics 2017 Stock data(refactored encore)](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2017_after-refactoring_mine.png) \
    [Results and Metrics 2018 Stock data(refactored encore)](https://github.com/RichardYDepestre/stock-analysis/blob/main/vb_challenge_2018_after-refactoring_mine.png)
    
Since we were dealing with number crunching, I thought it would be prudent to add 'DoEvents' and allow for completion of operations. This resulted in increasing the processing time by a couple of seconds. I also assessed using 'Range' versus 'Cells' when referencing cells on the worksheet. This yielded a reducing the processing time by a couple of seconds. The improvement from these small changes will materialize when dealing a large-size stocks data.

## Summary:
Refactoring helps pinpoint flaws in complex code contructs that handle massive operations: looping, external data requests, application memory management, etc. It also offers an opportunity for technical leads to define coding patterns, design recommendations, compilation of application metrics, etc. for use throughout an IT division therefore making applications hightly maintainable. \
Such an endeavor may not be welcomed by management as it comes with risks. Time consuming, introduction of defects, etc. are some of the risks associated with refactoring.

How do these pros and cons apply to refactoring the original VBA script?
The first sets of metrics compiled indicated that the VBA_Challenge macro initially ran for 7 seconds. The first attempts at refactoring looked at the variable types. Using a variable of type long when single is more appropriate. Substituting Cells to Range when referencing areas of a worksheet. Activating a worksheet before using it. Revisiting 'For Loops'. These changes resulted in significantly reducing the run-time from 7 seconds to less than 1 second.
This reduction in run-time required a lot of testing time, trying out other constructs. One can understand that refactoring may be frown upon as it requires time that may put in jeopardy delivery deadlines, application health etc.

### Resources from the Web:
[Tutorialspoint](https://www.tutorialspoint.com/vba/vba_for_loop.html)
[GitHub Docs](https://docs.github.com/en/github/writing-on-github/basic-writing-and-formatting-syntax#quoting-text)
