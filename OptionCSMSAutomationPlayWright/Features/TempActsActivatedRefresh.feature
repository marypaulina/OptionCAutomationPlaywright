Feature: TempActsActivatedRefresh

TEMPORARY - re-counts only "Acts Activated (Last 24 hrs)" for the listed schools and updates just that column
in today's audit workbook. Delete this feature once it has been used.

#@Temp
#Scenario: 001_Refresh only the Acts Activated count in the audit workbook
#	Given Acutis User has successfully launched
#	| URL                         | Username             | Password  |
#	| https://acutis.optionc.com/ | jclement@optionc.com | viper@123 |
#	And Refresh only the Acts Activated count for the listed schools
#		| SchoolCode | StartDate  |
#		| 7291       | 07/01/2026 |
#		| 7304       | 07/01/2026 |
#		| 8763       | 07/08/2026 |
#		| 8308       | 06/18/2026 |
#		| 142        | 05/22/2026 |
#		| 7296       | 07/20/2026 |
#		| 8417       | 07/20/2026 |
#		| 8369       | 07/20/2026 |
#		| 8407       | 07/20/2026 |
#		| 8188       | 07/20/2026 |
#		| 8232       | 07/20/2026 |
#		| 8545       | 06/14/2026 |
#		| 8235       | 06/26/2026 |
#		| 6904       | 07/20/2026 |
#		| 8351       | 06/22/2026 |
