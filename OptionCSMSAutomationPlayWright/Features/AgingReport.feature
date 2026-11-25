Feature: AgingReport

Captures the Acutis dynamic report (Id=106) with every record shown, sorted by Aging Days (largest first),
Aging Days above 90 highlighted, and inserts the full screenshot in the "Aging Report" sheet of today's audit workbook.

@Sprint_01 @Functional
Scenario: 001_Capture the Aging Days report with Aging Days above 90 highlighted and add it to the audit workbook
	Given Acutis User has successfully launched
	| URL                         | Username             | Password  |
	| https://acutis.optionc.com/ | jclement@optionc.com | viper@123 |
	And Open the dynamic report "https://acutis.optionc.com/dynamic-report?Id=106"
	When Display all the records using the paging option
	And Sort the report by Aging Days from the largest to the smallest
	And Highlight the Aging Days greater than 90
	Then Take a full screenshot of the report and insert it in the next sheet of the audit workbook for Aging Days greater than 90
	# eCheck Service Fee Report (yesterday; Friday to Sunday of the previous week when run on a Monday)
	When Go to the Reports menu, eCheck Report and open the Service Fee Report tab
	And Select the Processing Start Date and Processing End Date as yesterday (Friday to Sunday of the previous week on a Monday) and click Run Report
	And Resize the page to take the screenshot
	And Highlight the Total data in yellow with a thick border
	And Display the text "Service Fee to OptionC on the selected date range" in big bold red font
	Then Attach the Service Fee Report screenshot in the Aging Report sheet
	# Auto Withdrawals scheduled for today
	When Go to the Reports menu and open Auto Withdrawals
	And Select the current date as Start Date and End Date and click Filter
	And Display the count of auto withdrawal transactions scheduled today, or that none are scheduled
	Then Attach the Auto Withdrawals screenshot in the Aging Report sheet
