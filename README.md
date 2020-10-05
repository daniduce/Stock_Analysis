# Stock_Analysis
Using VBA for Stock Analysis

##Overview of Project

  We originally created the workbook in order to analyze the DQ stocks that Steve’s parents were interested investing in. We were able to see the DQ stock along with about a dozen others that may have been a good investment. Since none of those were the best options for them, we wanted to come up with a better alternative to expand the list of stocks to the entire stock market. With the refactored code, Steve will now be able to analyze a dataset of the stock market over the last few years to possibly find a better fit for them while taking less time and effort. 

##Results 

After refactoring the code and analyzing the results, I have established that the refactored code took longer to run than the original code. If you take a look at the window examples below, you can see that it took longer to run the script for 2017 and 2018 on the refactored code than the original. 

###Original Code Script

<img width="449" alt="Screen Shot 2020-10-04 at 12 03 21 PM" src="https://user-images.githubusercontent.com/71396367/95030072-e7394900-067a-11eb-93ca-884f400569ed.png">

<img width="444" alt="Screen Shot 2020-10-04 at 12 03 36 PM" src="https://user-images.githubusercontent.com/71396367/95030076-ebfdfd00-067a-11eb-863c-7c51ad31a9cb.png">

### Refactored Code Script

<img width="468" alt="Screen Shot 2020-10-04 at 11 44 53 AM" src="https://user-images.githubusercontent.com/71396367/95030065-db4d8700-067a-11eb-9764-98c35715d0a7.png">

<img width="462" alt="Screen Shot 2020-10-04 at 11 45 18 AM" src="https://user-images.githubusercontent.com/71396367/95030068-e1dbfe80-067a-11eb-9512-51055f95b8aa.png">

When analyzing over the stock sheets, it appears the original and refactored code never pulled from the 2017 worksheet. I was only able to activate the 2018 worksheet to look over the stocks for that year. Taking a look at the images above, the original code never properly formatted the “All Stocks” worksheet. 

<img width="487" alt="Screen Shot 2020-10-04 at 12 05 57 PM" src="https://user-images.githubusercontent.com/71396367/95030144-8b22f480-067b-11eb-89c3-1563703dde20.png">

<img width="485" alt="Screen Shot 2020-10-04 at 12 04 31 PM" src="https://user-images.githubusercontent.com/71396367/95030147-95dd8980-067b-11eb-953d-c39b3a2aa59f.png">

##Summary

	In conclusion, the advantages to refactoring the code is it will make things run smoother, but not necessarily faster. It also is a great way to re-summarize your code to be more simplified. The original VBA script we ran was very long with several separate Macros. Everything was broken up with each Macro separately running. By Refactoring the code, one pro is that it everything was able to run at the same time with a great amount of stocks than the original. A con of refactoring the code, is it is not separate so if something messes up it may be harder to debug and it may also take longer since there is more that its taking in to consideration. 

