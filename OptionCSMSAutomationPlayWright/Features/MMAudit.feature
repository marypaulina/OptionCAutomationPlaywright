Feature: MMAudit


This is to verify whether MM Transactions are happened without any fail by cross checking all the MM Reports

@Sprint_01 @Functional
Scenario: 001_To generate the Daily MM audit report for all the MM schools and confirm the transactions are correct
	Given Acutis User has successfully launched
	| URL                         | Username             | Password  |
	| https://acutis.optionc.com/ | jclement@optionc.com | viper@123 |
	And Open all the MM schools and audit the fee details everyday
	| SchoolCode                       | StartDate  |
	| 7291                             | 07/02/2025 |
	| 8417                             | 05/24/2025 |
	| 7292                             | 08/05/2025 |
	| 7285                             | 06/26/2025 |
	| 5940                             | 08/13/2025 |
	| 8308                             | 09/30/2025 |
	| 8958                             | 06/30/2025 |
	| 8298                             | 06/09/2025 |
	| 7304                             | 05/28/2025 |
	| 5158                             | 06/28/2025 |
	| 8518                             | 08/05/2025 |
	| 221                              | 08/15/2025 |
	| St. Bridget School - River Falls | 06/26/2025 |
	| 8232                             | 08/05/2025 |
	| 16000                            | 05/06/2025 |
	| 6929                             | 08/05/2025 |
	| 6904                             | 08/05/2025 |

	#| 3340       | 08/05/2025 |
	#| 8351       | 08/05/2025 |
	#| 6142       | 06/25/2025 |
	#| 1126       | 08/05/2025 |
	#| 3932       | 07/15/2025 |
	#| 8763       | 06/26/2025 |
	#| 7128       | 06/17/2025 |
	#| 8400       | 06/01/2025 |
	#| 8545       | 07/28/2025 |
	#| 8990       | 07/22/2025 |
	#| 8507       | 06/10/2025 |
	#| 7303       | 06/30/2025 |
	#| 142        | 08/05/2025 |
	#| 7296       | 08/05/2025 |
	#| 8188       | 11/21/2025 |
	#| 8235       | 08/14/2025 |
	#| 8407       | 08/05/2025 |



	
	#6929,6904 - New yr  created (Verified on 08/28/2025)	
	#8465, 4291 - New yr not created

	#8188, 7301 -St. Bridget School - River Falls
