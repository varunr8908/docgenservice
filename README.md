# docgenservice
REST API which accepts JSON input and merges and generates word and pdf templates


Project on Python 3.7.1

<B>Dependencies</B>
```
pip install flask
pip install lxml
pip install docx-mailmerge
pip install flask
pip install flask_compress

```
<p>
jre > 7 to run pdf conversion on non windows platforms
</p>
<br/>
To Start server - python run.py

<b>Sample JSON input for template</b>
```
{
	"DocFormat" : "docx",
	"TemplateName" : "SG_Proposal_Full",
	"GroupName" : "Test Employer Group",
	"EffectiveDate" :"11/01/2018",
	"DateCreated" :"11/07/2018",
	"PreparedBy" :"Johan Bluth",
	"AgencyName" :"Mark Henry Agency",
	"PBEmail" :"johan.bluth@nomail.com",
	"QuoteId" : "100531990001221",
	"Zipcode" : "33156",
	"RatingRegion" : "Miami-Dade",
	"QuotedPlans" :[
		{
			"BusinessPackageId" : "Agility LS300-SG17",
			"MonthlyPremium" : "$4,300",
			"SBC" : [
				{
					"Name" : "Coinsurance",
					"INN" : "0%",
					"ONN" : "50%"
				},
				{
					"Name" : "Deductible",
					"INN" : "$2,500 individual/$5,000 family Doesn't apply to preventive care",
					"ONN" : "$7,500 individual/$15,000 family Doesn't apply to preventive care"
				},
				{
					"Name" : "Other Deductible",
					"INN" : "$65 per child for Pediatric Dental. Doesn’t apply to overall deductible. There are no other specific deductibles.",
					"ONN" : "Not Applicable"
				},
				{
					"Name" : "Out of Pocket Max (Includes Deductible)",
					"INN" : "$6,800 individual/$13,600 family. Pediatric Dental is limited to $350 per child, or $700 for 2 or more children.",
					"ONN" : "$20,400 individual/$40,800 family. Pediatric Dental is limited to $350 per child, or $700 for 2 or more children."
				},
				{
					"Name" : "PCP Cost Share",
					"INN" : "No charge for first non-preventive visit; $35 copay/ visit thereafter",
					"ONN" : "50% Coinsurance after deductible"
				},
				{
					"Name" : "Specialist Cost Share (No Referral Needed)",
					"INN" : "$75 copay/ visit",
					"ONN" : "50% Coinsurance after deductible"
				},
				{
					"Name" : "Inpatient Hospital Cost Share",
					"INN" : "$750 copay/day for the first 3 days per admission, after deductible",
					"ONN" : "50% Coinsurance after deductible"
				},
				{
					"Name" : "ER Cost Share",
					"INN" : "$600 copay/ visit",
					"ONN" : "Same as AvMed Network"
				},
				{
					"Name" : "Urgent Care Cost Share",
					"INN" : "$125 copay/ visit at urgent care facilities; $35 copay/ visit at retail clinics",
					"ONN" : "50% coinsurance after deductible at urgent care facilities or retail clinics"
				},
				{
					"Name" : "Outpatient Surgery Cost Share",
					"INN" : "$500 copay/ visit at independent facilities; $1000 copay/ visit after deductible at all other facilities",
					"ONN" : "50% Coinsurance after deductible"
				},
				{
					"Name" : "Imaging Tests (CT / PET scans / MRI's) Cost Share",
					"INN" : "$350 copay/ visit at independent facilities; $1,000 copay/ visit after deductible at all other facilities",
					"ONN" : "50% Coinsurance after deductible"
				},
				{
					"Name" : "Drug Cost Share",
					"INN" : "Generic - $25 copay (retail)/ $62.50 copay (mail order) Preferred Brand - $55 copay (retail)/ $137.50 copay (mail order) Non-Preferred Brand - $95 copay (retail)/ $237.50 copay (mail order) Specialty - 50% coinsurance after deductible (retail only)",
					"ONN" : "Not Covered"
				}
				],
			"QuoteCensus" : [
				{
					"EmployeeName" : "Mark Levingston",
					"EmployeeNumber" : "A001",
					"BirthDate" : "11/12/1983",
					"NumDependents": "2",
					"FamilyRate" : "$1,150"
				},
				{
					"EmployeeName" : "Fred Therou",
					"EmployeeNumber" : "A002",
					"BirthDate" : "03/04/1989",
					"NumDependents": "1",
					"FamilyRate" : "$780"
				},
				{
					"EmployeeName" : "Henry Davidson",
					"EmployeeNumber" : "A003",
					"BirthDate" : "11/12/1983",
					"NumDependents": "2",
					"FamilyRate" : "$830"
				},
				{
					"EmployeeName" : "Bruce Campbell",
					"EmployeeNumber" : "A004",
					"BirthDate" : "05/04/1983",
					"NumDependents": "0",
					"FamilyRate" : "$390"
				}
				]
		}
		]
}
```
