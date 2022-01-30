# VBA Projects: Decision Support Systems
Excel Macros Projects

## [1] One Million Savings
### Problem
You are going to be wealthy after graduating from business school; thus, you want a financial tool that will determine exactly how long it will take for you to get rich. You are to build an Excel model that determines exactly how long it will take for your savings to reach $1 million.

### Model Assumptions
1. At the beginning of the first year, your savings balance should always be $0.
2. Any contributions in a given year are assumed to be made at the end of the year. Thus, there are no investment earnings for the year in which monies are contributed.
3. The following tax rates apply to investment income you earn depending on salary.

| Salary | Federal Tax Rate | State Tax Rate
| ----------- | ----------- | ----------- |
| =< $30,000 | 15% | 6% |
| $30,000.01 - $70,000 | 25% | 6% |
| $70,000.01 - $150,000 | 28% | 6% |
| $150,000.01 - $250,000 | 33% | 6% |
| >= $250,000.01 | 35% | 6% |

4. Salary is just what you earn from your employer and should not include investment earnings.
5. You cannot claim tax credits for investment losses (i,e. IF taxes < 0, THEN taxes = 0).
6. You can assume salary is net taxes (i.e., you do not need to worry about taxes that you would pay on your salary for these calculations).
Basic requirements of the system:

### Requirements
1. Your system should allow the user to vary the input variable values:
  a) Current age (Default: 35)
  b) Current salary (Default: $135,000)
  c) Percent of salary (before taxes) that will be contributed to taxable accounts per year (Default: 15%)
  d) Expected annual salary growth rate (Minimum: 0%, Maximum: 10%, Default: 4% per year)
  e) Expected annual investment return rate (Minimum: -10%, Maximum: 40%, Default: 4% per year)

2. The salary brackets (=< $30,000, $30,000.01 - $70,000, etc.) do not have to change, but the Federal and State Tax rates applied to those brackets should be set up to allow change.
3. When the file is opened, there are no gridlines.
4. The user can only enter valid data for the five main input variables (e.g., current age must be positive, salary must be nonnegative, etc.). You only need to check for invalid numbers as you assume that your user will not enter random text. You do not have to perform data validity checks on the tax rates (you can if you know how, but it is not a requirement).
5. Data validity checks are performed by Visual Basic code using IF statements. You should not use the built in Excel spreadsheet data validity tool.
6. All numbers are formatted properly (e.g., currency with dollar signs, rates as percentages) on the interface and in the output message box.
7. A Command Button is used to launch the code.
8. A Message Box is used to display the results (see sample output).
9. Under some conditions, you will not be able to ever save $1 million. Your system needs to check for this and provide an appropriate user message (instead of the standard message box that displays the time until you save $1 million).

## [2] One Million Savings
### Problem
Assume that the user is not convinced that saving is all that important and does not believe the estimates returned by the model.

### Requirements
1. The user would like to have the option of seeing the investment data details.
2. The user should be able to select or deselect the optional “View Investment Detail” with a CheckBox,  but the calculations should still launch with a command button. NOTE: Clicking the CheckBox should not launch anything.)

If the user selects the CheckBox to see the savings details and clicks the command button, the program  may or may not still show the MessageBox (your choice), and then the program must end with the  detailed output showing on the screen (i.e., the InvestmentData worksheet).

If the user does not select the CheckBox, when s/he clicks the command button, the program should  still show the original MessageBox from Assignment 1.

3. The user should have the option to allow the investments to grow tax-free. Two option buttons should  be used to allow the user to choose taxable vs. tax-free growth.

4. Since it is unrealistic that you would realize the same investment return rate each year, the model  should be able to randomly simulate an investment return rate for each year (think Management  Science). The simulated value should be sampled from a triangular distribution with the most likely,  minimum, and maximum values specified by scrollbars in the user interface. The code for generating  random numbers according to the triangular distribution is provided in the appendix and will be  discussed in class.
5. A set of option buttons should allow the user to choose whether to use the most likely investment  return rate for the duration of the investment period or to use the triangular distribution to generate a  random investment return rate for each year according to the specified distribution.
6. The user should be able to use scroll bars to adjust all the parameters. Use the following ScrollBar  specifications:
**Salary** Min: $0; Max: $500,000; Large Change: $1,000; Small Change: $100
**Current Age** Min: 12; Max: 100; Large Change: 5; Small Change: 1
**Percent of Salary to Contribute** Min: 0%; Max: 100%; Large Change: 5%; Small Change: 1%
**Salary Growth Rate** Min: 0%; Max: 20%; Large Change: 1%; Small Change: 0.5%
**Investment Return Rate** Absolute Min: -10%; Absolute Max: 40%; Large Change: 1%; Small  Change: 0.5%

7. As with Assignment 1, all numbers should be formatted properly (in the message boxes and on the  worksheet(s)), and the user should only be able to enter valid data.
8. Advanced Option: You must still provide the user the option of writing the output to a second  worksheet, but in the advanced option you also show the investment detail to the user in a ListBox on  the Input worksheet. Simply writing the investment detail to the InvestmentData sheet will not earn  you these points. Also, when the file opens, there should not be old data in the ListBox, and if the user  does not want to see the investment detail, the screen should appear as in the basic implementation (i.e., ListBox is not visible). Additionally, if you complete this advanced option, the user should end up  on the Input worksheet after the code has run and not the InvestmentData sheet as stated above.  Figure 3 shows an acceptable interface for the advanced option, and hints are provided for suggested  ways to implement this.

## [3] Small Town Zoo
### Problem
The Small Town Zoo (STZ) wants to emulate the Cincinnati Zoo in its use analytics. They want to better understand their operations and make smarter decisions to increase revenues and to reduce costs. The STZ is a nonprofit 100-acre facility serving an average of 250,000 people annually. The zoo houses over 100 animals. In order to track customers, STZ issues virtually all family visitors a card (one per family) on their first visit that accumulates points based on purchases. Visitors can choose from multiple cards with images of their favorite animals (e.g., a lion, an elephant, a polar bear, or a monkey), and points can eventually be redeemed for free zoo giveaways.

Most importantly to the marketing staff at STZ, the card captures how much families spend during their zoo visits.

Like most successful analytics companies, STZ experiments on a small scale with alternative marketing programs, conducts post-mortems on the tests to determine whether or not proposed programs are profitable, and then continues and expands successful programs while modifying or discontinuing the unsuccessful ones.  In order to analyze the programs, SMZ relies on decision support systems.

You are to develop for STZ a decision support system (DSS) to analyze the success of promotional e-mailings. An e-mailing is determined to be successful based on return visit behavior.  That is, (a) whether the family returned to STZ or not and (b) how profitable the visitors were if they did return after considering the cost of any coupon sent (e.g. free ice cream cone, etc.). Because the program is an e-mail, you can assume that program operating and administration costs are one fixed cost of $200 per e-mail effort regardless of how many people you e-mail. You can ignore any time-value of money. You are to consider the initial data set of visitors as first time visitors, and the system you develop should evaluate the promotion based on success (or lack thereof) of getting those visitors to return for a second visit.

### Part I: Expected Value Analysis
The DSS should provide the user the ability to enter the input data listed below to determine if the STZ e-mailing under consideration is profitable. The key output should be expected net profit for both the email promotion and no email promotion options (i.e., relying on visitors to return on their own). The model output should be a recommendation based on how the expected net profit from the two alternatives.

NOTE: Analysis of the email promotion evaluates first time zoo visitors (1 visitor = 1 family), and the impact of different offers to convince the visitors (i.e., the family) to return for a second visit.  For calculation purposes, assume 1,000 first-time visitors, and then consider percentages of these 1,000 visitors as described above to perform your calculations.  A numerical example for each input based on the default values is given to help you develop the computational model.   Your recommendations should be based on the expected value of the e-mail campaign (i.e., a simulation is NOT the correct tool to analyze this problem).

The percentage of first-time visitors to whom you send the e-mail.  For example, the e-mail may be a message from the STZ alerting families to any baby animals recently born at the zoo with (or without) a coupon good for something, (e.g., a free ice cream cone) on their next visit.
Default value: e-mail to 40% of first-time visitors
[total visitors receiving email = 1000*.4]

The percentage of visitors to whom an e-mail is sent that includes a coupon for something of value (i.e., a free ice cream cone).
Default value: 50% of the families who receive the e-mail receive the coupon. Thus, the remaining 50% would only receive the announcement of the new baby animals without the coupon.
[total receiving email with coupon = (1000*.4)*.5]

The expected percentage of visitors who return for a 2nd visit on their own, i.e., do not receive any e-mail but come back anyway.
Default value: 20% return on their own
When there is a promotion: [total returning without email = (1000-(1000*.4))*.2]
When there is NOT a promotion: [total returning without email = 1000*.2]

The expected percentage of visitors who receive the e-mail without the coupon and return for a 2nd visit
Default value: 30% of those who receive the e-mail (but no coupon) return
[total returning with email but no coupon = (1000*.4) * .5 *.3]

The expected percentage of visitors who receive the e-mail with the coupon and return for a 2nd visit
Default value: 40% of those who received the e-mail with the coupon return
[total returning with email and coupon = (1000*.4) * .5 *.4]
NOTE: based on bullets 4 and 5, 70% of visitors who received any email returned.

The cost of the coupon awarded in this e-mail (this value will need to be subtracted from revenue per visitor to determine the profit per visitor)
Default value: $2 ice cream coupon

The cost of the e-mail campaign
Default value: $200

The average revenue of visitors on their 2nd visit who return on their own without some type of e-mail to attract them.
Default value: $15

The average revenue of visitors on their 2nd visit who return after receiving the e-mail without the coupon
Default value: $13

The average revenue of visitors on their 2nd visit who return after receiving the e-mail with the coupon (assume 100% redemption of coupons if they receive one and return)
Default value: $9

### Part II: Sensitivity Analysis
In addition to the expected value computations, the user would like to see 5 sensitivity graphs that examine the sensitivity of the recommendation to the first five parameters described above (i.e., (1) percentage of first-time visitors to whom you send the e-mail, (2) the percentage of visitors to whom an e-mail with a coupon is sent, (3) the expected percentage of visitors who return for a second visit on their own, (4) the expected percentage of visitors who receive just the e-mail with no coupon and return for a second visit, and (5) the expected percentage of visitors who receive the e-mail with the coupon that return for a second visit. Percentages should vary from 0-100% in 5 or 10 percent increments.

Each sensitivity graph should show 2 lines – the expected profit from sending the e-mail and the expected profit from the “no e-mail” option (i.e., relying on all 1000 visitors to comeback on their own). Additionally, you need to have some kind of “help” to explain to the users how to interpret a sensitivity analysis graphs.

NOTE (and this is repeated in the requirements below): Calculations for sensitivity analysis must be generated with the use of VBA code.  This means they must “use” the model that you constructed and loop through possible alternative outcomes given different inputs.  Points will be deducted if you rely on data tables on a worksheet to calculate the sensitivity analysis.

ALSO NOTE: Some interpretation/recommendation of the sensitivity analysis for the base case is required, i.e., what can the marketing department learn from this analysis about sending e-mail promotions. Your detailed answer does not need to be dynamic and can specifically address just the base case.

### Part III: Real-data Validity
The marketing group at STZ regularly conducts trial experiments and collects data on first time visitors, promotional offers made, and the results of those trials.  A sample of the data collected is shown in Figure 1 below.  The first column is the Member ID numbered 1 to 1,000 (you can assume that the marketing group always conducts trials on 1,000 visitors).  The second column is the revenue of the family visitors on their first visit.  The third column is whether the marketing group e-mailed that customer or not (1 = yes, 0 = no).  The fourth column is whether the e-mail included a coupon or not. The fifth column is whether or not the customer came back (1 = yes, 0 = no). The sixth column is the revenue of the family visitor on the second visit.

The user would like the system to be able to analyze trial data, compare the expected profit of sending the e-mail versus the expected profit of not sending the e-mail, and make a recommendation if the net profit is positive and thus justifies the e-mail program. This part of the assignment should use the model you constructed for Parts I and II. You should open the selected file, and use the percentages and average revenues from that data set in your model constructed for Parts I and II. You need not redo the e-mail vs. no e-mail evaluation (see the ALSO NOTE above).

Three test data sets are provided.  The user should be able to select any dataset, and the system should update the recommendations based on the data set used.  Each dataset provided represents different e-mailings to a fraction of the 1,000 first time customers identified in the data set.  The value of the coupon offered (if one was offered) was $2.

Your system will be tested against additional datasets that have the same data columns and formatting (1,000 customers) so the user must be able to choose the data file to be analyzed by the DSS.

### Requirements
1. Your system should allow the user to change any of the 10 variables identified in Part I, and make a recommendation regarding the profitability of the e-mail program versus the no email option.
2. Your system should allow the user to examine the sensitivity of 5 parameters identified in Part II to consider how changing these parameters over a range of values will affect the profitability of the e-mail program.
3. Your system should allow the user to open/read/use a dataset of 1,000 customers provided by the Marketing department to determine if the trial was profitable. (Part III)
4. The system must include a UserForm that displays HELP instructions to aid the user in data entry. You can have additional user forms for additional help such as for how to interpret the sensitivity graphs, but you must have one that provides help for the data entry portion.
5. You must have a command button that returns all input parameters in Part I to the default values stated in this assignment.
6. You can use the spreadsheet to perform relevant calculations.  Calculations do not need to be performed in VBA code except for the sensitivity analysis component.
7. Calculations for sensitivity analysis should use the model that you constructed and loop through possible alternative outcomes given different inputs.  Points will be deducted if you rely on data tables to calculate the sensitivity analysis. The graphs of the sensitivity analysis should show the two alternatives (e-mail vs. no e-mail) and have some kind of help to explain how to interpret the graphs.
8. You must make sure the user can only enter valid data.  How you do that is up to you.
9. I do not specify which types of controls you should use in this assignment (i.e., scrollbars, option buttons, spin buttons, list boxes, combo boxes, etc.), but 25% of the grade is for your implementation of Controls.  You should be creative but make sure they are useful.
10. The system should look good, and its use should be obvious to any user. One alternative is to create UserForms for the interface. Alternatively, if you are showing results in the spreadsheet, you can create a dashboard for your output (i.e., all the data results with charts summarized in a single screen).

## [4] Job Choice
### Problem
Since you are an outstanding student, many of your peers come to ask you for advice on which job offer to  accept. You are finding these requests for advice significantly diminish the time you have available to watch  Scandal. Thus, in an attempt to minimize the amount of time you spend providing advice to your peers, you  decide to create a Decision Support System that will support them in ranking several job options.

### Part 1: Basic Analysis
At a minimum (see discussion of minimum below), the system must allow the user to enter the salary, the  location, and the enjoyment of the work. The system must also be able to rank 3 different job prospects. The  decision support system should provide a clear recommendation that describes both the (1) ranking of job  options and the (2) rationale for why the best alternative is ranked as such.  

### Part 2: Sensitivity Analysis
The user would like to see 1 graph that shows how sensitive the ranking is to some important variable that you  identify.

### Requirements
1. The system must include a UserForm that specifically explains to the user how the recommendation (final decision) is being derived. That is, the decision model(s) used must be clearly stated.
2. Model must allow the user to input at least the 3 job attributes of salary, location, and enjoyment of  the work.
3. Model must allow the user to rank at least 3 different job prospects.
4. No specific Controls (i.e., scrollbars, option buttons, spin buttons, list boxes, combo boxes, etc.) are  required, but you should try to be creative about what you do choose to implement. Make sure the  controls you develop are useful.
5. The system should look good and be obvious to use (think of your “non-DSS” friends opening the file  and trying to use it).  
6. All data should be formatted correctly, all formulas should be protected, and recommendations should  be clearly articulated.
7. Completing the basic system requirements will earn you at most 90 points out of a possible 100. The  final 10 points are earned by adding to your system over and above the basic system requirements.  Examples include, but are not limited to, adding extra job attributes, allowing the user to make  changes to the attributes objectives matrix, allowing users to run different decision models, or adding  features that improve the system recommendations or insight, etc.
