Feature: ParishBasicFlow

This feature is to verify the basic flow of the parish activities
@Sprint_01 @Functional
Scenario: 001_To sign in as parish admin and view the home page and open the daily readings and saint of the day
	Given I sign in using Parish Acutis user credentials
	| URL                                | Username             | Password |
	| https://parish-acutis.optionc.com/ | jclement@optionc.com | password |
	And I Verify whether all the parish dasboard elements are displayed for the parish admin user

@Sprint_01 @Functional
Scenario: 002_Verify total tickets count and latest submitted ticket details
	Given I sign in using Parish Acutis user credentials
	| URL                                | Username             | Password |
	| https://parish-acutis.optionc.com/ | jclement@optionc.com | password |
	When I navigate to the Recent Tickets page
    Then I should see the list of tickets displayed
    And I should display the total count of tickets
    And I should get the latest submitted ticket details including Ticket ID, School ID, School Name, Submitted By, and Submitted Date