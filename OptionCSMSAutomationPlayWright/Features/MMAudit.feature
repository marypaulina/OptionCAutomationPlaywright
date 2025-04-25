Feature: MMAudit


This is to verify whether MM Transactions are happened without any fail by cross checking all the MM Reports

@Sprint_01 @Functional
Scenario: 001_To generate the Daily MM audit report for all the MM schools and confirm the transactions are correct
	Given Acutis User has successfully launched
	| URL                         | Username             | Password  |
	| https://acutis.optionc.com/ | jclement@optionc.com | viper@123 |
	And Open all the MM schools and audit the fee details everyday
	| SchoolCode                            | StartDate  |
	| 7304                                  | 06/10/2024 |
	| 5158                                  | 06/18/2024 |
	| 8518                                  | 06/20/2024 |
	| 221                                   | 08/15/2024 |
	| 7301 St. Bridget School - River Falls | 07/23/2024 |
	| 8232                                  | 08/21/2024 |
	| 7292                                  | 08/19/2024 |
	| 7285                                  | 08/28/2024 |
	| 6929                                  | 07/26/2024 |
	| 8407                                  | 08/20/2024 |
	| 8958                                  | 08/27/2024 |
	| 6904                                  | 08/16/2024 |
	| 8298                                  | 02/19/2025 |
	| 7291                                  | 07/01/2024 |
	| 8417                                  | 05/15/2024 |

	#| 3340       | 08/27/2024 |
	#| 8351       | 07/11/2024 |
	#| 6142       | 07/05/2024 |
	#| 4291       | 02/19/2025 |
	#| 3932       | 01/02/2025 |
	#| 8763       | 06/25/2024 |
	#| 7128       | 07/23/2024 |
	#| 8400       | 06/01/2024 |
	#| 8545       | 08/19/2024 |
	#| 8990       | 08/18/2024 |
	#| 142        | 09/10/2024 |
	#| 7296       | 09/17/2024 |
	#| 8507       | 06/21/2024 |
	#| 7303       | 06/30/2024 |
	#| 8465       | 07/17/2024 |



	