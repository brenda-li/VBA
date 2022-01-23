# VBA
Excel Macros Projects

## One Million Savings
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

##
