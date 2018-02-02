<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'##################################################################################################
'# Declaration of constants
'##################################################################################################

'##################################################################################################
'# Declaration of variables
'##################################################################################################
Dim objFS
Dim objEDI
Dim objFieldName

Dim record_line
Dim record_id
Dim field_id

Dim objApplication
Dim objTitleHolder
Dim objDownPayment
Dim objOtherCredit
Dim objApplicant
Dim objItem
Dim strKey

'Applicant(s)
Dim ssn
'Other Credit Type 'Down Payment
Dim typeCode
'Title Holder
Dim titleName
'##################################################################################################
'# Initializing Page
'##################################################################################################

'##################################################################################################
'# Loading Page
'##################################################################################################
Set objFieldName = SetFieldIdName()

Set objFS	= Server.CreateObject("Scripting.FileSystemObject")
'================================================
'= Application
'================================================
Set objApplication = Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Applicant(s)"		, Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Other Credit Type"	, Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Title Holder"		, Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Down Payment"		, Server.CreateObject("Scripting.Dictionary")
'================================================
'= Reading EDI File
'================================================
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/C0101904_1.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test2.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test3.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test4.txt"),1,true)
Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test5.txt"),1,true)
'================================================
'= Parsing EDI File
'================================================
Do Until objEDI.AtEndOfStream
	'================================================
	'= Reading Record Line
	'================================================
	record_line = objEDI.ReadLine
	'================================================
	'= Getting Record ID
	'================================================
	record_id = LeftCut(record_line,3)
	'================================================
	'= Parsing Record Line
	'================================================
	Select Case UCase(record_id)
		'--------------------------------------------------------------------------------------------------
		'- [Application]
		'--------------------------------------------------------------------------------------------------
		'------------------------------------------------
		'- [00A] Top of Form
		'------------------------------------------------
		Case "00A"
			objApplication.Add "00A-020", Mid(record_line,4,1) '[LoanQualificationUsed]
			objApplication.Add "00A-030", Mid(record_line,5,1) '[LoanQualificationNotUsed]
		'------------------------------------------------
		'- [01A] I	 Mortgage Type and Terms
		'------------------------------------------------
		Case "01A"
			objApplication.Add "01A-020", Mid(record_line,4,2)     'Mortgage Applied For [MortgageAppliedForOther]
			objApplication.Add "01A-030", Mid(record_line,6,80)    'Mortgage Applied For (Other) [MortgageAppliedForOther]
			objApplication.Add "01A-040", Mid(record_line,86,30)   'Agency Case Number [AgencyCaseNumber]
			objApplication.Add "01A-050", Mid(record_line,116,15)  'Case Number [CaseNumber]
			objApplication.Add "01A-060", Mid(record_line,131,15)  'Loan Amount [LoanAmt]
			objApplication.Add "01A-070", Mid(record_line,146,7)   'Interest Rate [InterestRate]
			objApplication.Add "01A-080", Mid(record_line,153,3)   'No. of Months [NoOfMonth]
			objApplication.Add "01A-090", Mid(record_line,156,2)   'Amortization Type [AmortizationType]
			objApplication.Add "01A-100", Mid(record_line,158,80)  'Amortization Type Other Explanation [AmortizationTypeOtherExplain]
			objApplication.Add "01A-110", Mid(record_line,238,80)  'ARM Textual Description [ARMTrxtualDesc]
		'------------------------------------------------
		'- [02A] II	Property Information
		'------------------------------------------------
		Case "02A" 
			objApplication.Add "02A-020", Mid(record_line,4,50) 	'Property Street Address [PropertyStreetAddress]
			objApplication.Add "02A-030", Mid(record_line,54,35)	'Property City [PropertyCity]
			objApplication.Add "02A-040", Mid(record_line,89,2) 	'Property State [PropertyState]
			objApplication.Add "02A-050", Mid(record_line,91,5) 	'Property Zip Code [PropertyZip]
			objApplication.Add "02A-060", Mid(record_line,96,4)		'Property Zip Code Plus Four [PropertyZipPlusFour]
			objApplication.Add "02A-070", Mid(record_line,100,3)	'No. of Units [NoOfUnits]
			objApplication.Add "02A-080", Mid(record_line,103,2)	'Legal Description of Subject Property Code [LegalDescSubjPropCode]
			objApplication.Add "02A-090", Mid(record_line,105,80)	'Legal Description of Subject Property [LegalDescSubjProp]
			objApplication.Add "02A-100", Mid(record_line,185,4)	'Year Built [YearBuilt]
		'------------------------------------------------
		'- [02B]	II	Purpose of Loan
		'------------------------------------------------
		Case "02B" 
			objApplication.Add "02B-030", Mid(record_line,6,2)	'Purpose of Loan [PurposeOfLoan]
			objApplication.Add "02B-040", Mid(record_line,8,80)	'Purpose of Loan (Other) [PurposeOfLoanOther]
			objApplication.Add "02B-050", Mid(record_line,88,1)	'Property will be [PropertyWillBe]
			objApplication.Add "02B-060", Mid(record_line,89,60)'Manner in which Title will be held [MannerTitleWillBeHeld] 
			objApplication.Add "02B-070", Mid(record_line,149,1)'Estate will be held in [EstateWillBeHeldIn]
			objApplication.Add "02B-080", Mid(record_line,150,8)'(Estate will be held in) Leasehold expiration date [LeaseholdExpirationDate]
		'------------------------------------------------
		'- [02C]	II	 Title Holder
		'------------------------------------------------
		Case "02C" 
			titleName = Mid(record_line,4,60)
			Set objTitleHolder = GetDuplicateData(objApplication("Title Holder"),titleName)
			objTitleHolder.Add "02C-020", Mid(record_line,4,60) 'Titleholder Name [TitleholderName]
		'------------------------------------------------
		'- [02D]	II	 Construction or Refinance Data
		'------------------------------------------------
		Case "02D" 
			objApplication.Add "02D-020", Mid(record_line,4,4)	'Year Lot Acquired (Construction) or Year Acquired (Refinance) [YearAcquired]
			objApplication.Add "02D-030", Mid(record_line,8,15)	'Original Cost (Construction or Refinance) [OriginalCost]
			objApplication.Add "02D-040", Mid(record_line,23,15)'Amount Existing Liens (Construction or Refinance) [AmtExistingLiens]
			objApplication.Add "02D-050", Mid(record_line,38,15)'(a) Present Value of Lot [PresentValueOfLot]
			objApplication.Add "02D-060", Mid(record_line,53,15)'(b) Cost of Improvements [CostOfImprovements]
			objApplication.Add "02D-070", Mid(record_line,68,2) 'Purpose of Refinance [PurposeOfRefinance]
			objApplication.Add "02D-080", Mid(record_line,70,80)'Describe Improvements [DescribeImprovements]
			objApplication.Add "02D-090", Mid(record_line,150,1)'(Describe Improvements) made/to be made [DescImporvMadeToBeMade]
			objApplication.Add "02D-100", Mid(record_line,151,15)'(Describe Improvements) Cost [DescImporvCost]
		'------------------------------------------------
		'- [02E]	II	 Down Payment
		'------------------------------------------------
		Case "02E" 
			typeCode = Mid(record_line,4,2)
			Set objDownPayment = GetDuplicateData(objApplication("Down Payment"),typeCode)
			objDownPayment.Add "02E-020", Mid(record_line,4,2) 	'Down Payment Type Code [DownPaymentTypeCode]
			objDownPayment.Add "02E-030", Mid(record_line,6,15) 'Down Payment Amount [DownPaymentamt]
			objDownPayment.Add "02E-040", Mid(record_line,21,80)'Down Payment Explanation [DownPaymentExplanation]
		'------------------------------------------------
		'- [07A]	VII	 Details of Transaction
		'------------------------------------------------
		Case "07A"
			objApplication.Add "07A-020", Mid(record_line,4,15)   'a. Purchase price [PurchasePrice]
			objApplication.Add "07A-030", Mid(record_line,19,15)  'b. Alterations, improvements, repairs [AlterationsImprovRepair]
			objApplication.Add "07A-040", Mid(record_line,34,15)  'c. Land [Land]
			objApplication.Add "07A-050", Mid(record_line,49,15)  'd. Refinance (Inc. debts to be paid off) [Refinance]
			objApplication.Add "07A-060", Mid(record_line,64,15)  'e. Estimated prepaid items [EstimatedPrepaidItems]
			objApplication.Add "07A-070", Mid(record_line,79,15)  'f. Estimated closing costs [EstimatedClosingCosts]
			objApplication.Add "07A-080", Mid(record_line,94,15)  'g. PMI MIP, Funding Fee [PMIMIPFundingFee]
			objApplication.Add "07A-090", Mid(record_line,109,15) 'h. Discount [Discount]
			objApplication.Add "07A-100", Mid(record_line,124,15) 'j. Subordinate financing [SubordinateFinancing]
			objApplication.Add "07A-110", Mid(record_line,139,15) 'k. Applicant's closing costs paid by Seller [ClosingCostPaidBySeller]
			objApplication.Add "07A-120", Mid(record_line,154,15) 'n. PMI, MIP, Funding Fee financed [PMIMIPFundingFeeFinan]
		'------------------------------------------------
		'- [07B]	VII	 Other Credits
		'------------------------------------------------
		Case "07B" 
			typeCode = Mid(record_line,4,2)
			Set objOtherCredit = GetDuplicateData(objApplication("Other Credit Type"),typeCode)
			objOtherCredit.Add "07B-020", Mid(record_line,4,2) 		'Other Credit Type Code [OtherCreditTypeCode]
			objOtherCredit.Add "07B-030", Mid(record_line,6,15) 	'Amount of Other Credit [AmtOfOtherCredit]
		'------------------------------------------------
		'- [10B]	X	 Loan Originator Information
		'------------------------------------------------
		Case "10B" 
			objApplication.Add "10B-020", Mid(record_line,4,1)		'This application was taken by [ThisAppWasTakenBy]
			objApplication.Add "10B-030", Mid(record_line,5,60)		'Loan Originator's Name [LoanOriginatorName]
			objApplication.Add "10B-040", Mid(record_line,65,8)		'Interview Date [InterviewDate]
			objApplication.Add "10B-050", Mid(record_line,73,10)	'Loan Originator's Phone Number [LOPhoneNo]
			objApplication.Add "10B-060", Mid(record_line,83,35)	'Loan Origination Company's Name [LOCompanyName]
			objApplication.Add "10B-070", Mid(record_line,118,35)	'Loan Origination Company's Street Address [LOCompanyStAddr]
			objApplication.Add "10B-080", Mid(record_line,153,35)	'Loan Origination Company's Street Address 2 [LOCompanyStAddr2]
			objApplication.Add "10B-090", Mid(record_line,188,35)	'Loan Origination Company's City [LOCompanyCity]
			objApplication.Add "10B-100", Mid(record_line,223,2)	'Loan Origination Company's State Code [LOCompanyStateCode]
			objApplication.Add "10B-110", Mid(record_line,225,5)	'Loan Origination Company's Zip Code [LOCompanyZip]
			objApplication.Add "10B-120", Mid(record_line,230,4)	'Loan Origination Company's Zip Code Plus Four [LOCompanyZipFour]
		'--------------------------------------------------------------------------------------------------
		'- [Applicant]
		'--------------------------------------------------------------------------------------------------
		'------------------------------------------------
		'- 03A	III	 Applicant(s) Data
		'------------------------------------------------
		Case "03A"
			ssn = Mid(record_line,6,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "03A-020", Mid(record_line,4,2) 		'Applicant / Co-Applicant Indicator [ApplicantIndicator]
			objApplicant.Add "03A-040", Mid(record_line,15,35) 		'Applicant First Name [[ApplicantFirstName]
			objApplicant.Add "03A-050", Mid(record_line,50,35) 		'Applicant Middle Name [ApplicantMidName]
			objApplicant.Add "03A-060", Mid(record_line,85,35) 		'Applicant Last Name [ApplicantLastName]
			objApplicant.Add "03A-070", Mid(record_line,120,4) 		'Applicant Generation [ApplicantGeneration]
			objApplicant.Add "03A-080", Mid(record_line,124,10)		'Home Phone [HomePhone]
			objApplicant.Add "03A-090", Mid(record_line,134,3)		'Age [Age]
			objApplicant.Add "03A-100", Mid(record_line,137,2)		'Yrs. School [YrsSchool]
			objApplicant.Add "03A-110", Mid(record_line,139,1)		'Marital Status [MaritalStatus]
			objApplicant.Add "03A-120", Mid(record_line,140,2)		'Dependents (no.) [DependantsNo]
			objApplicant.Add "03A-130", Mid(record_line,142,1)		'Completed Jointly/Not Jointly [CompletedJoinNotJoin]
			objApplicant.Add "03A-140", Mid(record_line,143,9)		'Cross-Reference Number [CrossRefNumber]
			objApplicant.Add "03A-150", Mid(record_line,152,8)		'Date of Birth [DateOfBirth]
			objApplicant.Add "03A-160", Mid(record_line,160,80)		'Email Address [EmailAddr]
		'------------------------------------------------
		'- 03B	III	 Dependent's Age.
		'------------------------------------------------
		Case "03B"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "03B-030", Mid(record_line,13,3) 		'Dependant's age [DependantsAge]
		'------------------------------------------------
		'- 03C	III	 Applicant(s) Address
		'------------------------------------------------
		Case "03C"
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Address")
			strKey = Mid(record_line,13,2)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "03C-030", Mid(record_line,13,2)	'(Present/Former) [PresentFormer]
			objItem(strKey).Add "03C-040", Mid(record_line,15,50)	'Residence Street Address [ResidenceStAddr] 
			objItem(strKey).Add "03C-050", Mid(record_line,65,35)	'Residence City [ResidenceCity]
			objItem(strKey).Add "03C-060", Mid(record_line,100,2)	'Residence State [ResidenceState]
			objItem(strKey).Add "03C-070", Mid(record_line,102,5)	'Residence Zip Code [ResidenceZip]
			objItem(strKey).Add "03C-080", Mid(record_line,107,4)	'Residence Zip Code Plus Four [ResidenceZipFour]
			objItem(strKey).Add "03C-090", Mid(record_line,111,1)	'Own/Rent/Living Rent Free [OwnRentLivingRentFree]
			objItem(strKey).Add "03C-100", Mid(record_line,112,2)	'No. Yrs. [AddrNoYrs]
			objItem(strKey).Add "03C-110", Mid(record_line,114,2)	'No. Months [AddrNoMonth]
			objItem(strKey).Add "03C-120", Mid(record_line,116,50)	'Country [ApplicantAddrCounrtry]
		'------------------------------------------------
		'- 04A	IV	 Primary Current Employer(s)
		'------------------------------------------------
		Case "04A"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "04A-030", Mid(record_line,13,35)		'Employer Name [EmpName]
			objApplicant.Add "04A-040", Mid(record_line,48,35)		'Employer Street Address [EmpStAddr]
			objApplicant.Add "04A-050", Mid(record_line,83,35)		'Employer City [EmpCity]
			objApplicant.Add "04A-060", Mid(record_line,118,2)		'Employer State [EmpState]
			objApplicant.Add "04A-070", Mid(record_line,120,5)		'Employer Zip Code [EmpZip]
			objApplicant.Add "04A-080", Mid(record_line,125,4)		'Employer Zip Code Plus Four [EmpZipFour]
			objApplicant.Add "04A-090", Mid(record_line,129,1)		'Self Employed [SelfEmployed]
			objApplicant.Add "04A-100", Mid(record_line,130,2)		'Yrs. on this job [YrsOnThisJob]
			objApplicant.Add "04A-110", Mid(record_line,132,2)		'Months on this job [MonthsOnThisJob]
			objApplicant.Add "04A-120", Mid(record_line,134,2)		'Yrs. employed in this line of work/profession [YrsEmpInThisLineWork]
			objApplicant.Add "04A-130", Mid(record_line,136,25)		'Position / Title / Type of Business [PositionTitleTypeBiz]
			objApplicant.Add "04A-140", Mid(record_line,161,10)		'Business Phone [BizPhone]
		'------------------------------------------------
		'- 04B	IV	 Secondary/Previous Employer(s)
		'------------------------------------------------
		Case "04B"
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Secondary/Previous Employer(s)")
			strKey = Mid(record_line,13,35)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "04B-030", Mid(record_line,13,35)	'Employer Name [SPrevEmpName]
			objItem(strKey).Add "04B-040", Mid(record_line,48,35)	'Employer Street Address [SPrevEmpStAddr]
			objItem(strKey).Add "04B-050", Mid(record_line,83,35)	'Employer City [SPrevEmpCity]
			objItem(strKey).Add "04B-060", Mid(record_line,118,2)	'Employer State [SPrevEmpState]
			objItem(strKey).Add "04B-070", Mid(record_line,120,5)	'Employer Zip Code [SPrevEmpZip]
			objItem(strKey).Add "04B-080", Mid(record_line,125,4)	'Employer Zip Code Plus Four [SPrevEmpZipFour]
			objItem(strKey).Add "04B-090", Mid(record_line,129,1)	'Self Employed [SPrevSelfEmployed]
			objItem(strKey).Add "04B-100", Mid(record_line,130,1)	'Current Employment Flag [SPrevYrsOnThisJob]
			objItem(strKey).Add "04B-110", Mid(record_line,131,8)	'From Date [SPrevFromDate]
			objItem(strKey).Add "04B-120", Mid(record_line,139,8)	'To Date [SPrevToDate]
			objItem(strKey).Add "04B-130", Mid(record_line,147,15)	'Monthly Income [SPrevMonthlyIncome]
			objItem(strKey).Add "04B-140", Mid(record_line,162,25)	'Position / Title / Type of Business [SPrevPositionTitleTypeBiz]
			objItem(strKey).Add "04B-150", Mid(record_line,187,10)	'Business Phone [SPrevBizPhone]
		'------------------------------------------------
		'- 05H	V	 Present/Proposed Housing Expense 
		'------------------------------------------------
		Case "05H"
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Present/Proposed Housing Expense")
			strKey = Mid(record_line,13,1) & "-" & Mid(record_line,14,2)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "05H-030", Mid(record_line,13,1)	'Present/Proposed Indicator [HousingExpensePresentIndicator]
			objItem(strKey).Add "05H-040", Mid(record_line,14,2)	'Housing Payment Type Code [HousingPaymentTypeCode]
			objItem(strKey).Add "05H-050", Mid(record_line,16,15)	'Housing Payment Amount (Monthly Housing Exp.) [HousingPaymentAmt]
		'------------------------------------------------
		'- 05I	V	 Income
		'------------------------------------------------
		Case "05I"		
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Income")
			strKey =  Mid(record_line,13,2)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "05I-030", Mid(record_line,13,2)	'Type of Income Code [TypeOfIncomeCode]
			objItem(strKey).Add "05I-040", Mid(record_line,15,15)	'Income Amount (Monthly Income) [IncomeAmt]
		'------------------------------------------------
		'- 06A	VI	 For all asset types, enter data in the 06C assets segment.
		'------------------------------------------------	
		Case "06A"		
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "06A-030", Mid(record_line,13,35)	'Cash deposit toward purchase held by [CashDepositPurcHeldBy]
			objApplicant.Add "06A-040", Mid(record_line,48,15)	'Cash or Market Value [CashOrMarketValue]
		'------------------------------------------------
		'- 06B	VI	 Life Insurance
		'------------------------------------------------
		Case "06B"		
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "06B-030", Mid(record_line,13,30)	'Acct. no. [LifeInsurAcctNo]
			objApplicant.Add "06B-040", Mid(record_line,43,15)	'Life Insurance Cash or Market Value [LifeInsurCashMarketVal]
			objApplicant.Add "06B-050", Mid(record_line,58,15)	'Life insurance Face Amount	[LifeInsurFaceAmt]
		'------------------------------------------------
		'- 06C	VI	 Assets
		'------------------------------------------------
		Case "06C"		
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Assets")
			strKey = Mid(record_line,13,3) & "-" & Mid(record_line,16,35) & "-" & Mid(record_line,132,30)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "06C-030", Mid(record_line,13,3)	'Account/Asset Type [AccountAssetType]
			objItem(strKey).Add "06C-040", Mid(record_line,16,35)	'Depository/Stock/Bond Institution Name [DepositoryStockBondName]
			objItem(strKey).Add "06C-050", Mid(record_line,51,35)	'Depository Street Address [DepositoryStAddr]
			objItem(strKey).Add "06C-060", Mid(record_line,86,35)	'Depository City [DepositoryCity]
			objItem(strKey).Add "06C-070", Mid(record_line,121,2)	'Depository State [DepositoryState]
			objItem(strKey).Add "06C-080", Mid(record_line,123,5)	'Depository Zip Code [DepositoryZip]
			objItem(strKey).Add "06C-090", Mid(record_line,128,4)	'Depository Zip Code Plus Four [DepositoryZipFour]
			objItem(strKey).Add "06C-100", Mid(record_line,132,30)	'Acct. no. [AssetAcctNo]
			objItem(strKey).Add "06C-110", Mid(record_line,162,15)	'Cash or Market Value [AssetCashMarketVal]
			objItem(strKey).Add "06C-120", Mid(record_line,177,7)	'Number of Stock/Bond Shares [NumberOfStockBondShares]
			objItem(strKey).Add "06C-130", Mid(record_line,184,80)	'Asset Description [AssetDesc]
		'------------------------------------------------
		'- 06D	VI	 Automobile(s)
		'------------------------------------------------
		Case "06D"		
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Automobile(s)")
			strKey = Mid(record_line,13,30)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "06D-030", Mid(record_line,13,30)	'Automobile Make/ Model [AutomobileMakeModel]
			objItem(strKey).Add "06D-040", Mid(record_line,43,4)	'Automobile Year [AutomobileYear]
			objItem(strKey).Add "06D-050", Mid(record_line,47,15)	'Cash or Market Value [AutomobileCashMarketVal]
		'------------------------------------------------
		'- 06F	VI	 Alimony, Child Support/ Separate Maintenance and/or Job Related Expense(s)
		'------------------------------------------------
		Case "06F"		
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Alimony, Child Support/ Separate Maintenance and/or Job Related Expense(s)")
			strKey = Mid(record_line,13,3)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "06F-030", Mid(record_line,13,3)	'Expense Type Code [ExpenseTypeCode]
			objItem(strKey).Add "06F-040", Mid(record_line,16,15) 	'Monthly Payment Amount [MonthlyPaymentAmt]
			objItem(strKey).Add "06F-050", Mid(record_line,31,3)	'Months Left to Pay [MonthsLeftToPay]
			objItem(strKey).Add "06F-060", Mid(record_line,34,60)	'Alimony/ Child Support/ Separate Maintenance Owed To [AlimonyCSSperateOwedTo]
		'------------------------------------------------
		'- 06G	VI	 Real Estate Owned
		'------------------------------------------------
		Case "06G"		
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Real Estate Owned")
			strKey = Mid(record_line,13,35)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "06G-030", Mid(record_line,13,35)	'Property Street Address [REOPropStAddr]
			objItem(strKey).Add "06G-040", Mid(record_line,48,35)	'Property City [REOPropCity]
			objItem(strKey).Add "06G-050", Mid(record_line,83,2)	'Property State [REOPropState]
			objItem(strKey).Add "06G-060", Mid(record_line,85,5)	'Property Zip Code [REOPropZip]
			objItem(strKey).Add "06G-070", Mid(record_line,90,4)	'Property Zip Code Plus Four [REOPropZipFour]
			objItem(strKey).Add "06G-080", Mid(record_line,94,1)	'Property Disposition [REOPropDisposition]
			objItem(strKey).Add "06G-090", Mid(record_line,95,2)	'Type of Property [REOTypeOfProperty]
			objItem(strKey).Add "06G-100", Mid(record_line,97,15)	'Present Market Value [REOPresentMarketValue]
			objItem(strKey).Add "06G-110", Mid(record_line,112,15)	'Amount of Mortgages & Liens [REOAmtMortgageLiens]
			objItem(strKey).Add "06G-120", Mid(record_line,127,15)	'Gross Rental Income [REOGrossRentalIncome]
			objItem(strKey).Add "06G-130", Mid(record_line,142,15)	'Mortgage Payments [REOMortgagePayment]
			objItem(strKey).Add "06G-140", Mid(record_line,157,15)	'Insurance, Maintenance Taxes & Misc. [InsurMaintenanceTaxMisc]
			objItem(strKey).Add "06G-150", Mid(record_line,172,25)	'Net Rental Income [REONetRentalIncome]
			objItem(strKey).Add "06G-160", Mid(record_line,187,1)	'Current Residence Indicator [REOCurResidenceIndicator]
			objItem(strKey).Add "06G-170", Mid(record_line,188,1)	'Subject Property Indicator [REOSubjectPropIndicator]
			objItem(strKey).Add "06G-180", Mid(record_line,189,2)	'REO Asset ID [REOAssetID]
			objItem(strKey).Add "06G-190", Mid(record_line,191,15)	'Reserved for Future Use [REOReservedForFutureUse]
		'------------------------------------------------
		'- 06H	VI	 Alias
		'------------------------------------------------
		Case "06H"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "06H-030", Mid(record_line,13,35) 'Alternate First Name [AliasFirstName]
			objApplicant.Add "06H-040", Mid(record_line,48,35) 'Alternate Middle Name [AliasMidName]
			objApplicant.Add "06H-050", Mid(record_line,83,35) 'Alternate Last Name [AliasLastName]
			objApplicant.Add "06H-060", Mid(record_line,118,15)'Reserved for Future Use 
			objApplicant.Add "06H-070", Mid(record_line,153,15)'Reserved for Future Use
		'------------------------------------------------
		'- 06L	VI	 Liabilities
		'------------------------------------------------
		Case "06L"		
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Liabilities")
			strKey = Mid(record_line,13,2) & "-" & Mid(record_line,15,35) & "-" &  Mid(record_line,131,30)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "06L-030", Mid(record_line,13,2)  'Liability Type [LiabilityType]
			objItem(strKey).Add "06L-040", Mid(record_line,15,35) 'Creditor Name [CreditorName]
			objItem(strKey).Add "06L-050", Mid(record_line,50,35) 'Creditor Street Address [CreditorStAddr]
			objItem(strKey).Add "06L-060", Mid(record_line,85,35) 'Creditor City [CreditorCity]
			objItem(strKey).Add "06L-070", Mid(record_line,120,2) 'Creditor State [CreditorState]
			objItem(strKey).Add "06L-080", Mid(record_line,122,5) 'Creditor Zip Code [CreditorZip]
			objItem(strKey).Add "06L-090", Mid(record_line,127,4) 'Creditor Zip Code Plus Four [CreditorZipFour]
			objItem(strKey).Add "06L-100", Mid(record_line,131,30)'Acct. no. [LiabilityAcctNo]
			objItem(strKey).Add "06L-110", Mid(record_line,161,15)'Monthly Payment Amount [LiabilityMonPaymentAmt]
			objItem(strKey).Add "06L-120", Mid(record_line,176,3) 'Months Left to Pay [LiabilityMonLeftToPay]
			objItem(strKey).Add "06L-130", Mid(record_line,179,15)'Unpaid Balance [UnpaidBalance]
			objItem(strKey).Add "06L-140", Mid(record_line,194,1) 'Liability will be paid prior to closing [LiabilityPaidClosing]
			objItem(strKey).Add "06L-150", Mid(record_line,195,2) 'REO Asset ID [REOAssetID]
			objItem(strKey).Add "06L-160", Mid(record_line,197,1) 'Resubordinated Indicator [ResubordinatedIndicator]
			objItem(strKey).Add "06L-170", Mid(record_line,198,1) 'Omitted Indicator [OmittedIndicator]
			objItem(strKey).Add "06L-180", Mid(record_line,199,1) 'Subject Property Indicator [SubjectPropIndicator]
			objItem(strKey).Add "06L-190", Mid(record_line,200,1) 'Rental Property Indicator [RentalPropIndicator]
		'------------------------------------------------
		'- 06S	VI	 Undrawn HELOC and IPCs
		'------------------------------------------------
		Case "06S"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "06S-030", Mid(record_line,13,3) 'Summary Amount Type Code [HELOCSummaryAmtTypeCode]
			objApplicant.Add "06S-040", Mid(record_line,16,15)'Amount [HELEOCAmt]
		'------------------------------------------------
		'- 08A	VIII Declarations
		'------------------------------------------------
		Case "08A"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "08A-030", Mid(record_line,13,1)   'a. Are there any outstanding judgments against you? [DeclarationsA]
			objApplicant.Add "08A-040", Mid(record_line,14,1)   'b. Have you been declared bankrupt within the past 7 years?[DeclarationsB]
			objApplicant.Add "08A-050", Mid(record_line,15,1)   'c. Have you had property foreclosed upon or given title or deed in lieu thereof in the last 7 years? [DeclarationsC]
			objApplicant.Add "08A-060", Mid(record_line,16,1)   'd. Are you a party to a lawsuit? [DeclarationsD]
			objApplicant.Add "08A-070", Mid(record_line,17,1)   'e. Have you directly or indirectly been obligated on any loan [DeclarationsE]
			objApplicant.Add "08A-080", Mid(record_line,18,1)   'f. Are you presently delinquent or in default on any Federal debt [DeclarationsF]
			objApplicant.Add "08A-090", Mid(record_line,19,1)   'g. Are you obligated to pay alimony child support or separate maintenance? [DeclarationsG]
			objApplicant.Add "08A-100", Mid(record_line,20,1)   'h. Is any part of the down payment borrowed? [DeclarationsH]
			objApplicant.Add "08A-110", Mid(record_line,21,1)   'i. Are you a co-maker or [DeclarationsI]
			objApplicant.Add "08A-120", Mid(record_line,22,2)   'j. Are you a U.S. citizen?'k. Are you a permanent resident alien? [DeclarationsJ]
			objApplicant.Add "08A-130", Mid(record_line,24,1)   'l. Do you intend to occupy [DeclarationsL]
			objApplicant.Add "08A-140", Mid(record_line,25,1)   'm. Have you had an ownership interest [DeclarationsM]
			objApplicant.Add "08A-150", Mid(record_line,26,1)   'm. (1) What type of property [DeclarationsM1]
			objApplicant.Add "08A-160", Mid(record_line,27,2)   'm. (2) How did you hold title [DeclarationsM2]
		'------------------------------------------------
		'- 08B	VIII Declaration Explanations
		'------------------------------------------------
		Case "08B"
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Declaration Explanations")
			strKey = Mid(record_line,13,2)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			objItem(strKey).Add "08B-030", Mid(record_line,13,2)   'Declaration Type Code [DeclarationTypeCode]
			objItem(strKey).Add "08B-040", Mid(record_line,15,255) 'Declaration Explanation [DeclarationExplanation]
		'------------------------------------------------
		'- 09A	IX	 Acknowledgment and Agreement
		'------------------------------------------------
		Case "09A"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "09A-030", Mid(record_line,13,8) 'Signature Date [SignatureDate]
		'------------------------------------------------
		'- 10A	X	 Information for Government Monitoring Purposes.
		'------------------------------------------------
		Case "10A"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "10A-030", Mid(record_line,13,1) 	'I do not wish to furnish this information [IDoNotFurnishMyInfo]
			objApplicant.Add "10A-040", Mid(record_line,14,1)	'Ethnicity [Ethnicity]
			objApplicant.Add "10A-050", Mid(record_line,15,30)	'Filler [Filler]
			objApplicant.Add "10A-060", Mid(record_line,45,1)	'Sex [Sex]
		'------------------------------------------------
		'- 10R	X	 Information for Government Monitoring Purposes
		'------------------------------------------------
		Case "10R"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "10R-030", Mid(record_line,13,2)	'Race type [RaceType]
	End Select
	
Loop
'------------------------------------------------
'- Printing Application
'------------------------------------------------
Call PrintApplication(objApplication)
 
'##################################################################################################
'# Unloading Page
'##################################################################################################
Set objFieldName = Nothing
Set objFS= Nothing

'##################################################################################################
'# Functions
'##################################################################################################
'==================================================================================================
'= PrintApplication
'==================================================================================================
Sub PrintApplication(ByRef objApplication)
	Dim fld_application
	Dim fld_applicant
	Dim ssn
	Dim objApplicant
	'------------------------------------------------
	'- Application
	'------------------------------------------------
	For Each fld_application In objApplication.Keys
		Select Case fld_application
			Case "Applicant(s)" '03A
				For Each ssn In objApplication("Applicant(s)")
					Response.Write "<hr>"
					Response.Write "<h1>Applicant : " & ssn & "</h1>"
					Set objApplicant = objApplication("Applicant(s)")(ssn)
					'------------------------------------------------
					'- Applicant
					'------------------------------------------------
					For Each fld_applicant In objApplicant.Keys
						Select Case fld_applicant
							Case "Address" '03C
								Call print_code("Address", ssn)
							Case "Secondary/Previous Employer(s)" '04B
								Call print_code("Secondary/Previous Employer(s)", ssn)
							Case "Present/Proposed Housing Expense" '05H
								Call print_code("Present/Proposed Housing Expense", ssn)
							Case "Income" '05I
								Call print_code("Income", ssn)	
							Case "Assets" '06C
								Call print_code("Assets", ssn)	
							Case "Automobile(s)" '06D
								Call print_code("Automobile(s)", ssn)	
							Case "Alimony, Child Support/ Separate Maintenance and/or Job Related Expense(s)" '06F
								Call print_code("Alimony, Child Support/ Separate Maintenance and/or Job Related Expense(s)", ssn)	
							Case "Real Estate Owned" '06G
								Call print_code("Real Estate Owned", ssn)	
							Case "Liabilities" '06L
								Call print_code("Liabilities", ssn)	
							Case "Declaration Explanations" '08B
								Call print_code("Declaration Explanations", ssn)	
							Case Else
								Response.Write fld_applicant & "(" & objFieldName(fld_applicant) & "): <strong>" & objApplicant(fld_applicant) & "</strong><br>"
						End Select
					Next
				Next
			Case "Other Credit Type" '07B
				print_dupicate_code("Other Credit Type")
			Case "Title Holder" '02C
				print_dupicate_code("Title Holder")
			Case "Down Payment" '02E
				print_dupicate_code("Down Payment")
			Case Else
				Response.Write fld_application & "(" & objFieldName(fld_application) & "): <strong>" & objApplication(fld_application) & "</strong><br>"
		End Select
	Next
End Sub
'==================================================================================================
'= print_code
'==================================================================================================
Function print_code(ByVal fld_applicant, ByRef ssn)
	Dim str_key
	Dim objFields
	Dim fld_item
	Dim objItem
	
	Set objApplicant = objApplication("Applicant(s)")(ssn)
	Response.Write "<strong>" & fld_applicant & "</storng><br>"
	Set objItem = objApplicant(fld_applicant)
	For Each str_key In objItem.Keys
		Response.Write "- <strong>" & str_key & "</storng><br>"
		Set objFields = objItem(str_key)
		For Each fld_item In objFields.Keys
			Response.Write fld_item & "(" & objFieldName(fld_item) & "): <strong>" & objFields(fld_item) & "</strong><br>"
		Next
	Next
	Response.Write "<p>"
End Function
'==================================================================================================
'= print_dupicate_code
'==================================================================================================
Function print_dupicate_code(ByVal fld_data)
	Dim typeCode
	Dim objDuplicateData
	Dim fld_code
	
	For Each typeCode In objApplication(fld_data)
		Response.Write "<hr>"
		Response.Write "<h1>" & fld_data & " : " & typeCode & "</h1>"
		Set objDuplicateData = objApplication(fld_data)(typeCode)
		For Each fld_code In objDuplicateData.Keys
			Select Case fld_code
			Case ""
				
			Case Else
				Response.Write fld_code & "(" & objFieldName(fld_code) & "): <strong>" & objDuplicateData(fld_code) & "</strong><br>"
			End Select
		Next
	Next
End Function
'==================================================================================================
'= GetApplicant
'==================================================================================================
Function GetApplicant(ByRef obj_applicants, ByVal ssn)
	If obj_applicants.Exists(ssn) = FALSE Then
		obj_applicants.Add ssn, Server.CreateObject("Scripting.Dictionary")
		'------------------------------------------------
		'-
		'------------------------------------------------
		obj_applicants(ssn).Add "Address",Server.CreateObject("Scripting.Dictionary") '03C
		obj_applicants(ssn).Add "Secondary/Previous Employer(s)",Server.CreateObject("Scripting.Dictionary") '04B
		obj_applicants(ssn).Add "Present/Proposed Housing Expense",Server.CreateObject("Scripting.Dictionary") '05H
		obj_applicants(ssn).Add "Income",Server.CreateObject("Scripting.Dictionary") '05I
		obj_applicants(ssn).Add "Assets",Server.CreateObject("Scripting.Dictionary") '06C
		obj_applicants(ssn).Add "Automobile(s)",Server.CreateObject("Scripting.Dictionary") '06D
		obj_applicants(ssn).Add "Alimony, Child Support/ Separate Maintenance and/or Job Related Expense(s)",Server.CreateObject("Scripting.Dictionary") '06F
		obj_applicants(ssn).Add "Real Estate Owned",Server.CreateObject("Scripting.Dictionary") '06G
		obj_applicants(ssn).Add "Liabilities",Server.CreateObject("Scripting.Dictionary") '06L
		obj_applicants(ssn).Add "Declaration Explanations",Server.CreateObject("Scripting.Dictionary") '08B
	End If
	Set GetApplicant = obj_applicants(ssn)
End Function
'==================================================================================================
'= GetDuplicateData
'==================================================================================================
Function GetDuplicateData(ByRef obj_type, ByVal code)
	If obj_type.Exists(code) = FALSE Then
		obj_type.Add code, Server.CreateObject("Scripting.Dictionary")
	End If
	Set GetDuplicateData = obj_type(code)
End Function
'==================================================================================================
'= LeftCut
'==================================================================================================
Function LeftCut(strString, intCut)
    dim intPos, chrTemp, strCut, intLength
    'Initial String length 
    intLength = 0
    intPos = 1
    'Loop until string length
    do while ( intPos <= Len( strString ))
       'compare with one word
        chrTemp = ASC(Mid( strString, intPos, 1))

        if chrTemp < 0 then 'if (-) then Korean
          strCut = strCut & Mid( strString, intPos, 1 ) 
          intLength = intLength + 2  'If Korean then string length + 2
        else
          strCut = strCut & Mid( strString, intPos, 1 )            
          intLength = intLength + 1  'If it is not Korean then string length + 1
        end If

        if intLength >= intCut  then
           exit do
        end if

        intPos = intPos + 1
    Loop
    'Return value
    LeftCut = strCut
End function	
%>
<!--#include file="SetFieldIdName_function.asp"-->