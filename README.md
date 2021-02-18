# stock-analysis
##  Overview of Project:
This project showcases the scripting features of Visual Basic for Application. \With the tool we will script instructions to analyze stock  \
data for the year 2017, 2018. \
In addition, the result of our analysis will be confined to a result cheet showing the ticker symbols, as tracked, respective their returns. \
Since we are dealing with code we will refactor it and make it as efficient as can be done.

## Results:
Running the macro before refactoring yielded the expected results: tables and contents identical to reqs.
Metrics captured during run show the following results:
1) when processing the 2017 and 2018 dataset and presenting the summary, the macro ran for ~5 seconds;
    [2017 Stocks Data](https://github.com/RichardYDepestre/kickstarter-analysis/blob/main/vb_challenge_2017.png)
3) after refactoring
4) Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
    - Refactoring helps pinpoint flaws of contructs in one's code. It offers great opportunity to improve mechanisms with metrics captured through execution. \In this challenge, I have included the 'DoEvents' keyword in areas where significant number crunching was taking place. The idea was to insure that calculations run to completion. I turns out that the addition of 'DoEvents' added a whoping 3 seconds to the overall execution of the code.
How do these pros and cons apply to refactoring the original VBA script?
# Kickstarting with Excel
  This project serves as an introduction to Data Analysis using Excel.\
  We must analyze raw data about crowd-funded projects initiated all over the globe, and uncover any *hidden trends*.

## Overview of Project
  This project is to present Louise, whom requested the analysis, with our findings. We will use data analysis
  and presentation features in MS Excel's to mine the raw data, annotate the trends identified and visually present \
  the outcomes of *all the categories*.

### Purpose
  This project is also meant to extrapolate, from the raw data, reasonable causes for their individual
  outcomes:\
  successful, failed or canceled.\
  Did the level or lack thereof played a role in the outcome of the project?\
  Which country enjoy the most successful projects in light of their funding levels?

## Analysis and Challenges
### Analysis of Outcomes Based on Launch Date
  [Chart Outcomes Based on Launch Date](https://github.com/RichardYDepestre/kickstarter-analysis/blob/main/chart_outcomes-based-on-launch-date.png)
### Analysis of Outcomes Based on Goals
  [Chart Analysis of Outcomes Based on Goals](https://github.com/RichardYDepestre/kickstarter-analysis/blob/main/chart_outcomes-based-on-launch-date.png)
### Challenges and Difficulties Encountered
  My challenge consisted within the boundaries of the assignment that is:
  1.  keep formulas and charts simple;
  2.  follow the instructions;
  3.  focus on what I could realistically extrapolate from the data given my current knowledge.
### Results
- What are two conclusions you can draw about the Outcomes based on Launch Date?
   - in the *theater* category, the months of May and June showcase the most successful months. The period of April through June has traditionally been the most productive. The fate of the projects that have failed appear to be proportional the funding they garner. For instance, if we looked at kickstarter projects (theater) that failed in the US, we'd see that these projects didn't raise enough interest from the backers. 
- What can you conclude about the Outcomes based on Goals?
  - On one side, this chart confirms that projects that have been canceled are those that have not secured enough funding. On the other side, the success or failure outcomes of other projects is not related at all to the level of funding. We see that  projects funded between 40000 and 49000 dollars have all failded and those funded at above 50000 dollars have also failed.
- What are some limitations of this dataset?
  Some columns could have been fully discribed and could have played a role in better understanding the data. What does staff_pick mean? what does spotlight mean?
 - What are some other possible tables and/or graphs that we could create?
  - 
