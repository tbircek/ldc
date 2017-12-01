Feature: UCL
	In order to avoid calculation mistakes
	I want to verify setup file entries

@setupFileVerification
Scenario Outline: Verify Setup File setting values between max and min
	Given I have access to setupFile
	When I enter <caseName> and <caseNumber> and 'value'
	Then the result should be <result>
	Examples: 
	| caseName                     | caseNumber | result   |
	| 'version'                    | 0          | 1        |
	| 'totaltestnumbers'           | 0          | 79       |
	| 'totalregulartest'           | 0          | 72       |
	| 'maxnetworkdevice'           | 0          | 3        |
	| 'primarybasecurrent'         | 0          | 25       |
	| 'sensitivity'                | 0          | 0.0      |
	| 'sensitivity'                | 1          | 4.0      |
	| 'bandcenter'                 | 0          | 120.0    |
	| 'bandcentertolerancepercent' | 0          | 0.5      |
	| 'deltatime'                  | 0          | 1.0      |
	| 'deltavoltage'               | 0          | 0.1      |
	| 'bandwidth'                  | 0          | 2.0      |
	| 'bandwidthtolerancepercent'  | 0          | 0.666667 |
	| 'rampingoffset'              | 0          | 1.0      |
	| 'voltagephase'               | 0          | 0        |
	| 'voltagephase'               | 0          | 0        |
	| 'voltagephase'               | 0          | 0        |
	| 'ldcrsetting'                | 0          | 0        |
	| 'ldcxsetting'                | 0          | 0        |
	| 'ldccurrent'                 | 0          | 0        |
	| 'ldcrsetting'                | 1          | 0        |
	| 'ldcxsetting'                | 1          | 0        |
	| 'ldccurrent'                 | 1          | 0        |
	| 'ldcrsetting'                | 2          | 0        |
	| 'ldcxsetting'                | 2          | 0        |
	| 'ldccurrent'                 | 2          | 0        |
	| 'mva'                        | 0          | 100      |
	| 'mva'                        | 1          | 50       |
	| 'ct'                         | 0          | 5000     |
	| 'ct'                         | 1          | 4500     |
	| 'ct'                         | 2          | 2500     |
	| 'impedance'                  | 0          | 10       |
	| 'impedance'                  | 1          | 8        |
	| 'breakers'                   | 0          | 111      |
	| 'breakers'                   | 1          | 110      |
	| 'breakers'                   | 2          | 101      |
	| 'breakers'                   | 3          | 100      |
	| 'breakers'                   | 4          | 011      |
