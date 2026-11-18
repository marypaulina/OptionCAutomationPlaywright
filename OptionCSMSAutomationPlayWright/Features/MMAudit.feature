Feature: MMAudit


This is to verify whether MM Transactions are happened without any fail by cross checking all the MM Reports

@Sprint_01 @Functional
Scenario: 001_To generate the Daily MM audit report for all the MM schools and confirm the transactions are correct
	Given Acutis User has successfully launched
	| URL                         | Username             | Password  |
	| https://acutis.optionc.com/ | jclement@optionc.com | viper@123 |
	And Open all the MM schools and audit the fee details everyday	
		| SchoolCode | StartDate  |
		| 7296       | 07/20/2026 |
		| 8417       | 07/20/2026 |
		| 8369       | 07/20/2026 |
		| 8407       | 07/20/2026 |
		| 8188       | 07/20/2026 |
		| 8232       | 07/20/2026 |
		| 8545       | 06/14/2026 |
		| 8235       | 06/26/2026 |
		| 6904       | 07/20/2026 |
		| 8351       | 06/22/2026 |
		| 7291       | 07/01/2026 |
		| 7304       | 07/01/2026 |
		| 8763       | 07/08/2026 |
		| 8308       | 06/18/2026 |
		| 142        | 05/22/2026 |
		
		
		
		

		#| 16000      | 07/20/2026 |
		#| 3340       | 07/20/2026 |		
		

		#	| 6142       | 06/25/2026 |
		#| 8507       | 06/19/2026 |
		#| 7285       | 07/20/2026 |
		
		

#----------------------------------------------------------------
#	| 5158                             | 06/25/2026 |
#| 3932                             | 07/20/2026 |
#| 8400                             | 06/10/2026 |
#| 7292                             | 06/12/2026 |
#| 7127                             | 07/23/2026 |
#| 7303                             | 07/06/2026 |
#| 6929                             | 05/28/2026 |
#| St. Bridget School - River Falls | 07/08/2026 |
#| 221                              | 07/09/2026 |
#| 8298                             | 07/20/2026 |
#| 8958                             | 06/25/2026 |
#| 7128                             | 06/17/2026 |
#| 6142                             | 06/25/2026 |
#| 8507                             | 06/19/2026 |
#| 7285                             | 07/20/2026 |




	
	#6929,6904 - New yr  created (Verified on 08/28/2025)	
	#8465, 4291 - New yr not created

	#8188, 7301 -St. Bridget School - River Falls
