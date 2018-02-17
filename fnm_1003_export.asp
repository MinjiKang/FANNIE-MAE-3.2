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
'Other Credit 'Down Payment
Dim typeCode
'Title Holder
Dim titleName

Dim i
Dim result_str 'Printing EDI

Dim LoanQualificationUsed, LoanQualificationNotUsed '00A
Dim MortgageAppliedFor, MortgageAppliedForOther, AgencyCaseNumber, CaseNumber,LoanAmt, InterestRate, NoOfMonth, AmortizationType, AmortizationTypeOtherExplain, ARMTextualDesc'01A
Dim PropStAddress, PropCity, PropState, PropZip, PropZipPlusFour, NoOfUnits, LegalDescSubjPropCode, LegalDescSubjProp, YearBuilt '02A
Dim ReservedFutureUse, PurposeOfLoan, PurposeOfLoanOther, PropWillBe, MannerTitleWillBeHeld, EstateWillBeHeldIn, LeaseholdExpirationDate '02B
Dim TitleholderName '02C
Dim YearAcquired, OriginalCost, AmtExistingLiens, PresentValueOfLot, CostOfImprovements, PurposeOfRefinance, DescribeImprovements, DescImporvMadeToBeMade, DescImporvCost '02D
Dim DownPaymentTypeCode, DownPaymentAmt, DownPaymentExplanation '02E
Dim PurchasePrice, AlterationsImprovRepair, Land, Refinance, EstimatedPrepaidItems, EstimatedClosingCosts '07A
Dim PMIMIPFundingFee, Discount, SubordinateFinancing, ClosingCostPaidBySeller, PMIMIPFundingFeeFinan '07A
Dim OtherCreditTypeCode, AmtOfOtherCredit '07B
Dim LoanOriginatorName, InterviewDate, LOPhoneNo, LOCompanyName, LOCompanyStAddr, LOCompanyStAddr2, LOCompanyCity, LOCompanyStateCode, LOCompanyZip, LOCompanyZipFour '10B
Dim ApplicantIndicator, ApplicantFirstName, ApplicantMidName, ApplicantLastName,ApplicantGeneration,HomePhone,Age
Dim YrsSchool,MaritalStatus,DependantsNo,CompletedJoinNotJoin,CrossRefNumber,DateOfBirth,EmailAddr '03A
Dim DependantsAge '03B
Dim PresentFormer,ResidenceStAddr,ResidenceCity,ResidenceState,ResidenceZip,ResidenceZipFour,OwnRentLivingRentFree,AddrNoYrs,AddrNoMonth,ApplicantAddrCountry '03C
Dim EmpName,EmpStAddr,EmpCity,EmpState,EmpZip,EmpZipFour,SelfEmployed,YrsOnThisJob,MonthsOnThisJob, YrsEmpInThisLineWork,PositionTitleTypeBiz,BizPhone '04A
Dim SPrevEmpName,SPrevEmpStAddr,SPrevEmpCity,SPrevEmpState,SPrevEmpZip,SPrevEmpZipFour,SPrevSelfEmployed,SPrevCurrentEmpFlag
Dim SPrevFromDate,SPrevToDate,SPrevMonthlyIncome,SPrevPositionTitleTypeBiz,SPrevBizPhone '04B
Dim HousingExpensePresentIndicator,HousingPaymentTypeCode,HousingPaymentAmt '05H
Dim TypeOfIncomeCode,IncomeAmt '05I
Dim CashDepositPurcHeldBy,CashOrMarketValue '06A
Dim LifeInsurAcctNo,LifeInsurCashMarketVal,LifeInsurFaceAmt '06B
Dim AccountAssetType,DepositoryStockBondName,DepositoryStAddr,DepositoryCity,DepositoryState,DepositoryZip,DepositoryZipFour,AssetAcctNo
Dim AssetCashMarketVal,NumberOfStockBondShares,AssetDesc '06C
Dim AutomobileMakeModel, AutomobileYear,AutomobileCashMarketVal '06D
Dim MonthlyPaymentAmt,MonthsLeftToPay,AlimonyCSSperateOwedTo '06F
Dim REOPropStAddr,REOPropCity,REOPropState,REOPropZip,REOPropZipFour,REOPropDisposition,REOTypeOfProp,REOPresentMarketValue
Dim REOAmtMortgageLiens,REOGrossRentalIncome,REOMortgagePayment,InsurMaintenanceTaxMisc
Dim REONetRentalIncome,REOCurResidenceIndicator,REOSubjectPropIndicator,REOAssetID,REOReservedFutureUse '06G
Dim AliasMidNam,AliasLastName,ReservedFutureUse6_1,ReservedFutureUse6_2 '06H
Dim LiabilityType, CreditorName, CreditorStAddr,CreditorCity,CreditorState,CreditorZip,CreditorZipFour
Dim LiabilityAcctNo,LiabilityMonPaymentAmt,LiabilityMonLeftToPay,UnpaidBalance,LiabilityPaidClosing,ResubordinatedIndicator
Dim OmittedIndicator,SubjectPropIndicator,RentalPropIndicator '06L
Dim HELOCSummaryAmtTypeCode,HELEOCAmt '06S
Dim DeclarationsA '08A
Dim DeclarationsB
Dim DeclarationsC
Dim DeclarationsD
Dim DeclarationsE
Dim DeclarationsF
Dim DeclarationsG
Dim DeclarationsH
Dim DeclarationsI 
Dim DeclarationsJ
Dim DeclarationsK
Dim DeclarationsL
Dim DeclarationsM
Dim DeclarationsM1
Dim DeclarationsM2
Dim DeclarationTypeCode,DeclarationExplanation,SignatureDate '08B
Dim IDoNotFurnishMyInfo,Ethnicity,Filler,Sex,ThisAppWasTakenBy '10A
Dim RaceType '10R
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
objApplication.Add "Other Credit"		, Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Title Holder"		, Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Down Payment"		, Server.CreateObject("Scripting.Dictionary")
'================================================
'= Reading EDI File
'================================================
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/C0101904_1.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test2.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test3.txt"),1,true)
Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test4.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test5.txt"),1,true)

'================================================
'= Printing EDI Header
'================================================
	'------------------------------------------------
	'- EH line
	'------------------------------------------------
	result_str = ""
	'Header
	result_str = result_str & WriteEDI("EH",3,"F")
	result_str = result_str & WriteEDI("",6,"F")
	result_str = result_str & WriteEDI("",25,"F")
	result_str = result_str & WriteEDI("",11,"F")
	result_str = result_str & WriteEDI("",9,"F")
	Response.write result_str & "<br>"
	'------------------------------------------------
	'- TH line
	'------------------------------------------------
	result_str = ""
	'Header
	result_str = result_str & WriteEDI("TH",3,"F")
	result_str = result_str & WriteEDI("T100099-002",11,"F")
	result_str = result_str & WriteEDI("",9,"F")
	Response.write result_str & "<br>"
	'------------------------------------------------
	'- TPI line
	'------------------------------------------------
	result_str = ""
	'Header
	result_str = result_str & WriteEDI("TPI",3,"F")
	result_str = result_str & WriteEDI("1.00",5,"E")
	result_str = result_str & WriteEDI("01",2,"F")
	result_str = result_str & WriteEDI("",2,"F")
	result_str = result_str & WriteEDI("",30,"F")
	Response.write result_str & "<br>"
	'------------------------------------------------
	'- 000 line
	'------------------------------------------------
	result_str = ""
	'Header
	result_str = result_str & WriteEDI("000",3,"F")
	result_str = result_str & WriteEDI("1",3,"F")
	result_str = result_str & WriteEDI("3.20",5,"F")
	result_str = result_str & WriteEDI("W",1,"F")
	Response.write result_str & "<br>"
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
			
			LoanQualificationUsed = Trim(Mid(record_line,4,1))
			LoanQualificationNotUsed = Trim(Mid(record_line,5,1))
			'------------------------------------------------
			'- [00A] printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("00A",3,"F")
			'The income or assets of a person other than the borrower
			result_str = result_str & WriteEDI(LoanQualificationUsed,1,"F")
			'The income or assets of the borrower’s spouse will not be used
			result_str = result_str & WriteEDI(LoanQualificationNotUsed,1,"F")
			
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- [01A] I	 Mortgage Type and Terms
		'------------------------------------------------
		Case "01A"
			objApplication.Add "01A-020", Mid(record_line,4,2)     'Mortgage Applied For [MortgageAppliedFor]
			objApplication.Add "01A-030", Mid(record_line,6,80)    'Mortgage Applied For (Other) [MortgageAppliedForOther]
			objApplication.Add "01A-040", Mid(record_line,86,30)   'Agency Case Number [AgencyCaseNumber]
			objApplication.Add "01A-050", Mid(record_line,116,15)  'Case Number [CaseNumber]
			objApplication.Add "01A-060", Mid(record_line,131,15)  'Loan Amount [LoanAmt]
			objApplication.Add "01A-070", Mid(record_line,146,7)   'Interest Rate [InterestRate]
			objApplication.Add "01A-080", Mid(record_line,153,3)   'No. of Months [NoOfMonth]
			objApplication.Add "01A-090", Mid(record_line,156,2)   'Amortization Type [AmortizationType]
			objApplication.Add "01A-100", Mid(record_line,158,80)  'Amortization Type Other Explanation [AmortizationTypeOtherExplain]
			objApplication.Add "01A-110", Mid(record_line,238,80)  'ARM Textual Description [ARMTextualDesc]
			
			MortgageAppliedFor = Trim(Mid(record_line,4,2))
			MortgageAppliedForOther = Trim(Mid(record_line,6,80))
			AgencyCaseNumber =Trim( Mid(record_line,86,30))
			CaseNumber = Trim(Mid(record_line,116,15))
			LoanAmt = Trim(Mid(record_line,131,15)) 
			InterestRate = Trim(Mid(record_line,146,7))
			NoOfMonth = Trim(Mid(record_line,153,3))
			AmortizationType = Trim(Mid(record_line,156,2))
			AmortizationTypeOtherExplain = Trim(Mid(record_line,158,80))
			ARMTextualDesc = Trim(Mid(record_line,238,80))
			'------------------------------------------------
			'- [01A] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("01A",3,"F")
			'Mortgage Applied For
			result_str = result_str & WriteEDI(MortgageAppliedFor,2,"F")
			'Mortgage Applied For (Other)
			result_str = result_str & WriteEDI(MortgageAppliedForOther,80,"F")
			'Agency Case Number
			result_str = result_str & WriteEDI(AgencyCaseNumber,30,"F")
			'Case Number
			result_str = result_str & WriteEDI(CaseNumber,15,"F")
			'Loan Amount
			result_str = result_str & WriteEDI(LoanAmt,15,"E")
			'Interest Rate
			result_str = result_str & WriteEDI(InterestRate,7,"E")
			'No. of Months
			result_str = result_str & WriteEDI(NoOfMonth,3,"F")
			'Amortization Type
			result_str = result_str & WriteEDI(AmortizationType,2,"F")
			'Amortization Type Other Explanation
			result_str = result_str & WriteEDI(AmortizationTypeOtherExplain,80,"F")
			'ARM Textual Description
			result_str = result_str & WriteEDI(ARMTextualDesc,80,"F")
			
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- [02A] II	Property InFormation
		'------------------------------------------------
		Case "02A" 
			objApplication.Add "02A-020", Mid(record_line,4,50) 	'Property Street Address [PropStAddress]
			objApplication.Add "02A-030", Mid(record_line,54,35)	'Property City [PropCity]
			objApplication.Add "02A-040", Mid(record_line,89,2) 	'Property State [PropState]
			objApplication.Add "02A-050", Mid(record_line,91,5) 	'Property Zip Code [PropZip]
			objApplication.Add "02A-060", Mid(record_line,96,4)		'Property Zip Code Plus Four [PropZipPlusFour]
			objApplication.Add "02A-070", Mid(record_line,100,3)	'No. of Units [NoOfUnits]
			objApplication.Add "02A-080", Mid(record_line,103,2)	'Legal Description of Subject Property Code [LegalDescSubjPropCode]
			objApplication.Add "02A-090", Mid(record_line,105,80)	'Legal Description of Subject Property [LegalDescSubjProp]
			objApplication.Add "02A-100", Mid(record_line,185,4)	'Year Built [YearBuilt]
			
			PropStAddress = Trim(Mid(record_line,4,50))
			PropCity = Trim(Mid(record_line,54,35))
			PropState = Trim(Mid(record_line,89,2))
			PropZip = Trim(Mid(record_line,91,5))
			PropZipPlusFour = Trim(Mid(record_line,96,4))
			NoOfUnits = Trim(Mid(record_line,100,3))
			LegalDescSubjPropCode = Trim(Mid(record_line,103,2))
			LegalDescSubjProp = Trim(Mid(record_line,105,80))
			YearBuilt = Trim(Mid(record_line,185,4))
			'------------------------------------------------
			'- [02A] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("02A",3,"F")
			'Property Street Address
			result_str = result_str & WriteEDI(PropStAddress,50,"F")
			'Property City
			result_str = result_str & WriteEDI(PropCity,35,"F")
			'Property State 
			result_str = result_str & WriteEDI(PropState,2,"F")
			'Property Zip Code
			result_str = result_str & WriteEDI(PropZip,5,"F")
			'Property Zip Code Plus Four
			result_str = result_str & WriteEDI(PropZipPlusFour,4,"F")
			'No. of Units
			result_str = result_str & WriteEDI(NoOfUnits,3,"F")
			'Legal Description of Subject Property Code
			result_str = result_str & WriteEDI(LegalDescSubjPropCode,2,"F")
			'Legal Description of Subject Property
			result_str = result_str & WriteEDI(LegalDescSubjProp,80,"F")
			'Year Built
			result_str = result_str & WriteEDI(YearBuilt,4,"F")
			
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- [02B]	II	Purpose of Loan
		'------------------------------------------------
		Case "02B" 
			objApplication.Add "02B-020", Mid(record_line,4,2)	' [ReservedFutureUse]
			objApplication.Add "02B-030", Mid(record_line,6,2)	'Purpose of Loan [PurposeOfLoan]
			objApplication.Add "02B-040", Mid(record_line,8,80)	'Purpose of Loan (Other) [PurposeOfLoanOther]
			objApplication.Add "02B-050", Mid(record_line,88,1)	'Property will be [PropWillBe]
			objApplication.Add "02B-060", Mid(record_line,89,60)'Manner in which Title will be held [MannerTitleWillBeHeld] 
			objApplication.Add "02B-070", Mid(record_line,149,1)'Estate will be held in [EstateWillBeHeldIn]
			objApplication.Add "02B-080", Mid(record_line,150,8)'(Estate will be held in) Leasehold expiration date [LeaseholdExpirationDate]
			
			ReservedFutureUse = Trim(Mid(record_line,4,2))
			PurposeOfLoan = Trim(Mid(record_line,6,2))
			PurposeOfLoanOther = Trim(Mid(record_line,8,80))
			PropWillBe = Trim(Mid(record_line,88,1))
			MannerTitleWillBeHeld = Trim(Mid(record_line,89,60))
			EstateWillBeHeldIn = Trim(Mid(record_line,149,1))
			LeaseholdExpirationDate = Trim(Mid(record_line,150,8))
			'------------------------------------------------
			'- [02B] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("02B",3,"F")
			'Reserved for Future Use
			result_str = result_str & WriteEDI("",2,"F")
			'Purpose of Loan
			result_str = result_str & WriteEDI(PurposeOfLoan,2,"F")
			'Purpose of Loan (Other)
			result_str = result_str & WriteEDI(PurposeOfLoanOther,80,"F")
			'Property will be
			result_str = result_str & WriteEDI(PropWillBe,1,"F")
			'Manner in which Title will be held
			result_str = result_str & WriteEDI(MannerTitleWillBeHeld,60,"F")
			'Estate will be held in
			result_str = result_str & WriteEDI(EstateWillBeHeldIn,1,"F")
			'(Estate will be held in) Leasehold expiration date
			result_str = result_str & WriteEDI(LeaseholdExpirationDate,8,"E")
			
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- [02C]	II	 Title Holder
		'------------------------------------------------
		Case "02C" 
			titleName = Mid(record_line,4,60)
			Set objTitleHolder = GetDuplicateData(objApplication("Title Holder"),titleName)
			objTitleHolder.Add "02C-020", Mid(record_line,4,60) 'Titleholder Name [TitleholderName]
			
			titleName = Trim(Mid(record_line,4,60))
			'------------------------------------------------
			'- [02C] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("02C",3,"F")
			'Titleholder Name
			result_str = result_str & WriteEDI(titleName,60,"F")
			
			Response.write result_str & "<br>"
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
			
			YearAcquired = Trim(Mid(record_line,4,4))
			OriginalCost = Trim(Mid(record_line,8,15))
			AmtExistingLiens = Trim(Mid(record_line,23,15))
			PresentValueOfLot = Trim(Mid(record_line,38,15))
			CostOfImprovements = Trim(Mid(record_line,53,15))
			PurposeOfRefinance = Trim(Mid(record_line,68,2))
			DescribeImprovements = Trim(Mid(record_line,70,80))
			DescImporvMadeToBeMade = Trim(Mid(record_line,150,1))
			DescImporvCost = Trim(Mid(record_line,151,15))
			'------------------------------------------------
			'- [02D] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("02D",3,"F")
			'Year Lot Acquired (Construction) or Year Acquired (Refinance)
			result_str = result_str & WriteEDI(YearAcquired,4,"F")
			'Original Cost (Construction or Refinance)
			result_str = result_str & WriteEDI(OriginalCost,15,"E")
			'Amount Existing Liens (Construction or Refinance)
			result_str = result_str & WriteEDI(AmtExistingLiens,15,"E")
			'(a) Present Value of Lot
			result_str = result_str & WriteEDI(PresentValueOfLot,15,"E")
			'(b) Cost of Improvements
			result_str = result_str & WriteEDI(CostOfImprovements,15,"E")
			'Purpose of Refinance
			result_str = result_str & WriteEDI(PurposeOfRefinance,2,"F")
			'Describe Improvements
			result_str = result_str & WriteEDI(DescribeImprovements,80,"F")
			'(Describe Improvements) made/to be made
			result_str = result_str & WriteEDI(DescImporvMadeToBeMade,1,"F")
			'(Describe Improvements) Cost
			result_str = result_str & WriteEDI(DescImporvCost,15,"E")
			
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- [02E]	II	 Down Payment
		'------------------------------------------------
		Case "02E" 
			typeCode = Mid(record_line,4,2)
			Set objDownPayment = GetDuplicateData(objApplication("Down Payment"),typeCode)
			objDownPayment.Add "02E-020", Mid(record_line,4,2) 	'Down Payment Type Code [DownPaymentTypeCode]
			objDownPayment.Add "02E-030", Mid(record_line,6,15) 'Down Payment Amount [DownPaymentAmt]
			objDownPayment.Add "02E-040", Mid(record_line,21,80)'Down Payment Explanation [DownPaymentExplanation]
			
			DownPaymentTypeCode = Trim(Mid(record_line,4,2))
			DownPaymentAmt = Trim(Mid(record_line,6,15))
			DownPaymentExplanation = Trim(Mid(record_line,21,80))
			'------------------------------------------------
			'- [02E] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("02E",3,"F")
			'Down Payment Type Code
			result_str = result_str & WriteEDI(DownPaymentTypeCode,2,"F")
			'Down Payment Amount
			result_str = result_str & WriteEDI(DownPaymentAmt,15,"F")
			'Down Payment Explanation
			result_str = result_str & WriteEDI(DownPaymentExplanation,80,"F")
			
			Response.write result_str & "<br>"
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
			
			PurchasePrice = Trim(Mid(record_line,4,15))
			AlterationsImprovRepair = Trim(Mid(record_line,19,15))
			Land = Trim(Mid(record_line,34,15))
			Refinance = Trim(Mid(record_line,49,15))
			EstimatedPrepaidItems = Trim(Mid(record_line,64,15))
			EstimatedClosingCosts = Trim(Mid(record_line,79,15))
			PMIMIPFundingFee = Trim(Mid(record_line,94,15))
			Discount = Trim(Mid(record_line,109,15))
			SubordinateFinancing = Trim(Mid(record_line,124,15))
			ClosingCostPaidBySeller = Trim(Mid(record_line,139,15))
			PMIMIPFundingFeeFinan = Trim(Mid(record_line,154,15))
			'------------------------------------------------
			'- [07A] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("07A",3,"F")
			'a. Purchase price
			result_str = result_str & WriteEDI(PurchasePrice,15,"E")
			'b. Alterations, improvements, repairs
			result_str = result_str & WriteEDI(AlterationsImprovRepair,15,"E")
			'c. Land
			result_str = result_str & WriteEDI(Land,15,"E")
			'd. Refinance (Inc. debts to be paid off)
			result_str = result_str & WriteEDI(Refinance,15,"E")
			'e. Estimated prepaid items
			result_str = result_str & WriteEDI(EstimatedPrepaidItems,15,"E")
			'f. Estimated closing costs
			result_str = result_str & WriteEDI(EstimatedClosingCosts,15,"E")
			'g. PMI MIP, Funding Fee 
			result_str = result_str & WriteEDI(PMIMIPFundingFee,15,"E")
			'h. Discount
			result_str = result_str & WriteEDI(Discount,15,"E")
			'j. Subordinate financing
			result_str = result_str & WriteEDI(SubordinateFinancing,15,"E")
			'k. Applicant's closing costs paid by Seller
			result_str = result_str & WriteEDI(ClosingCostPaidBySeller,15,"E")
			'n. PMI, MIP, Funding Fee financed
			result_str = result_str & WriteEDI(PMIMIPFundingFeeFinan,15,"E")
			
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- [07B]	VII	 Other Credits
		'------------------------------------------------
		Case "07B" 
			typeCode = Mid(record_line,4,2)
			Set objOtherCredit = GetDuplicateData(objApplication("Other Credit"),typeCode)
			objOtherCredit.Add "07B-020", Mid(record_line,4,2) 		'Other Credit Type Code [OtherCreditTypeCode]
			objOtherCredit.Add "07B-030", Mid(record_line,6,15) 	'Amount of Other Credit [AmtOfOtherCredit]
			
			OtherCreditTypeCode = Trim(Mid(record_line,4,2))
			AmtOfOtherCredit = Trim(Mid(record_line,6,15))
			'------------------------------------------------
			'- [07B] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("07B",3,"F")
			'Other Credit Type Code
			result_str = result_str & WriteEDI(OtherCreditTypeCode,2,"F")
			'Amount of Other Credit
			result_str = result_str & WriteEDI(AmtOfOtherCredit,15,"E")
			
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- [10B]	X	 Loan Originator InFormation
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
			
			ThisAppWasTakenBy = Trim(Mid(record_line,4,1))
			LoanOriginatorName = Trim(Mid(record_line,5,60))
			InterviewDate = Trim(Mid(record_line,65,8))
			LOPhoneNo = Trim(Mid(record_line,73,10))
			LOCompanyName = Trim(Mid(record_line,83,35))
			LOCompanyStAddr = Trim(Mid(record_line,118,35))
			LOCompanyStAddr2 = Trim(Mid(record_line,153,35))
			LOCompanyCity = Trim(Mid(record_line,188,35))
			LOCompanyStateCode = Trim(Mid(record_line,223,2))
			LOCompanyZip = Trim(Mid(record_line,225,5))
			LOCompanyZipFour = Trim(Mid(record_line,230,4))
			'------------------------------------------------
			'- [10B] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("10B",3,"F")
			'This application was taken by
			result_str = result_str & WriteEDI(ThisAppWasTakenBy,1,"F")
			'Loan Originator's Name
			result_str = result_str & WriteEDI(LoanOriginatorName,60,"F")
			'Interview Date
			result_str = result_str & WriteEDI(InterviewDate,8,"F")
			'Loan Originator's Phone Number
			result_str = result_str & WriteEDI(LOPhoneNo,10,"F")
			'Loan Origination Company's Name
			result_str = result_str & WriteEDI(LOCompanyName,35,"F")
			'Loan Origination Company's Street Address
			result_str = result_str & WriteEDI(LOCompanyStAddr,35,"F")
			'Loan Origination Company's Street Address 2
			result_str = result_str & WriteEDI(LOCompanyStAddr2,35,"F")
			'Loan Origination Company's City 
			result_str = result_str & WriteEDI(LOCompanyCity,35,"F")
			'Loan Origination Company's State Code
			result_str = result_str & WriteEDI(LOCompanyStateCode,2,"F")
			'Loan Origination Company's Zip Code
			result_str = result_str & WriteEDI(LOCompanyZip,5,"F")
			'Loan Origination Company's Zip Code Plus Four 
			result_str = result_str & WriteEDI(LOCompanyZipFour,4,"F")
			
			Response.write result_str & "<br>"
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
			objApplicant.Add "03A-040", Mid(record_line,15,35) 		'Applicant First Name [ApplicantFirstName]
			objApplicant.Add "03A-050", Mid(record_line,50,35) 		'Applicant Middle Name [ApplicantMidName]
			objApplicant.Add "03A-060", Mid(record_line,85,35) 		'Applicant Last Name [ApplicantLastName]
			objApplicant.Add "03A-070", Mid(record_line,120,4) 		'Applicant Generation [ApplicantGeneration]
			objApplicant.Add "03A-080", Mid(record_line,124,10)		'Home Phone [HomePhone]
			objApplicant.Add "03A-090", Mid(record_line,134,3)		'Age [Age]
			objApplicant.Add "03A-100", Mid(record_line,137,2)		'Yrs. School [YrsSchool]
			objApplicant.Add "03A-110", Mid(record_line,139,1)		'Marital Status [MaritalStatus]
			objApplicant.Add "03A-120", Mid(record_line,140,2)		'Dependents (no.) [DependantsNo]
			objApplicant.Add "03A-130", Mid(record_line,142,1)	    'Completed Jointly/Not Jointly [CompletedJoinNotJoin]
			objApplicant.Add "03A-140", Mid(record_line,143,9)		'Cross-Reference Number [CrossRefNumber]
			objApplicant.Add "03A-150", Mid(record_line,152,8)		'Date of Birth [DateOfBirth]
			objApplicant.Add "03A-160", Mid(record_line,160,80)		'Email Address [EmailAddr]
			
			ApplicantIndicator = Trim(Mid(record_line,4,2))
			ApplicantFirstName = Trim(Mid(record_line,15,35))
			ApplicantMidName  = Trim(Mid(record_line,50,35))
			ApplicantLastName = Trim( Mid(record_line,85,35))
			ApplicantGeneration = Trim(Mid(record_line,120,4))
			HomePhone = Trim(Mid(record_line,124,10))
			Age = Trim(Mid(record_line,134,3))
			YrsSchool = Trim(Mid(record_line,137,2))
			MaritalStatus = Trim(Mid(record_line,139,1))
			DependantsNo = Trim(Mid(record_line,140,2))
			CompletedJoinNotJoin = Trim(Mid(record_line,142,1))
			CrossRefNumber = Trim(Mid(record_line,143,9))
			DateOfBirth = Trim(Mid(record_line,152,8))
			EmailAddr = Trim(Mid(record_line,160,80))
			'------------------------------------------------
			'- [03A] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("03A",3,"F")
			result_str = result_str & WriteEDI(ApplicantIndicator,2,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(ApplicantFirstName,35,"F")
			result_str = result_str & WriteEDI(ApplicantMidName,35,"F")
			result_str = result_str & WriteEDI(ApplicantLastName,35,"F")
			result_str = result_str & WriteEDI(ApplicantGeneration,4,"F")
			result_str = result_str & WriteEDI(HomePhone,10,"F")
			result_str = result_str & WriteEDI(Age,3,"F")
			result_str = result_str & WriteEDI(YrsSchool,2,"F")
			result_str = result_str & WriteEDI(MaritalStatus,1,"F")
			result_str = result_str & WriteEDI(DependantsNo,2,"F")
			result_str = result_str & WriteEDI(CompletedJoinNotJoin,1,"F")
			result_str = result_str & WriteEDI(CrossRefNumber,9,"F")
			result_str = result_str & WriteEDI(DateOfBirth,8,"F")
			result_str = result_str & WriteEDI(EmailAddr,80,"F")
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- 03B	III	 Dependent's Age.
		'------------------------------------------------
		Case "03B"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "03B-030", Mid(record_line,13,3) 		'Dependant's age [DependantsAge]
			
			DependantsAge = Trim(Mid(record_line,13,3))
			'------------------------------------------------
			'- [03B] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("03B",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(DependantsAge,3,"F")
			Response.write result_str & "<br>"
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
			objItem(strKey).Add "03C-120", Mid(record_line,116,50)	'Country [ApplicantAddrCountry]
			
			PresentFormer = Trim(Mid(record_line,13,2))	
			ResidenceStAddr = Trim(Mid(record_line,15,50))	
			ResidenceCity = Trim(Mid(record_line,65,35))	
			ResidenceState = Trim(Mid(record_line,100,2))	
			ResidenceZip = Trim(Mid(record_line,102,5))	
			ResidenceZipFour = Trim(Mid(record_line,107,4))	
			OwnRentLivingRentFree = Trim(Mid(record_line,111,1))	
			AddrNoYrs = Trim(Mid(record_line,112,2))	
			AddrNoMonth = Trim(Mid(record_line,114,2))
			ApplicantAddrCountry = Trim(Mid(record_line,116,50))	
			'------------------------------------------------
			'- [03C] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("03C",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(PresentFormer,2,"F")
			result_str = result_str & WriteEDI(ResidenceStAddr,50,"F")
			result_str = result_str & WriteEDI(ResidenceCity,35,"F")
			result_str = result_str & WriteEDI(ResidenceState,2,"F")
			result_str = result_str & WriteEDI(ResidenceZip,5,"F")
			result_str = result_str & WriteEDI(ResidenceZipFour,4,"F")
			result_str = result_str & WriteEDI(OwnRentLivingRentFree,1,"F")
			result_str = result_str & WriteEDI(AddrNoYrs,2,"F")
			result_str = result_str & WriteEDI(AddrNoMonth,2,"F")
			result_str = result_str & WriteEDI(ApplicantAddrCountry,50,"F")
			Response.write result_str & "<br>"
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
			
			EmpName = Trim(Mid(record_line,13,35))
			EmpStAddr = Trim(Mid(record_line,48,35))
			EmpCity =Trim(Mid(record_line,83,35))
			EmpState =Trim(Mid(record_line,118,2))
			EmpZip =Trim(Mid(record_line,120,5))
			EmpZipFour =Trim(Mid(record_line,125,4))
			SelfEmployed =Trim(Mid(record_line,129,1))
			YrsOnThisJob =Trim(Mid(record_line,130,2))
			MonthsOnThisJob =Trim(Mid(record_line,132,2))
			YrsEmpInThisLineWork =Trim(Mid(record_line,134,2))
			PositionTitleTypeBiz =Trim(Mid(record_line,136,25))
			BizPhone =Trim(Mid(record_line,161,10))
			'------------------------------------------------
			'- [04A] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("04A",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(EmpName,35,"F")
			result_str = result_str & WriteEDI(EmpStAddr,35,"F")
			result_str = result_str & WriteEDI(EmpCity,35,"F")
			result_str = result_str & WriteEDI(EmpState,2,"F")
			result_str = result_str & WriteEDI(EmpZip,5,"F")
			result_str = result_str & WriteEDI(EmpZipFour,4,"F")
			result_str = result_str & WriteEDI(SelfEmployed,1,"F")
			result_str = result_str & WriteEDI(YrsOnThisJob,2,"F")
			result_str = result_str & WriteEDI(MonthsOnThisJob,2,"F")
			result_str = result_str & WriteEDI(YrsEmpInThisLineWork,2,"F")
			result_str = result_str & WriteEDI(PositionTitleTypeBiz,25,"F")
			result_str = result_str & WriteEDI(BizPhone,10,"F")
			Response.write result_str & "<br>"
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
			objItem(strKey).Add "04B-100", Mid(record_line,130,1)	'Current Employment Flag [SPrevCurrentEmpFlag]
			objItem(strKey).Add "04B-110", Mid(record_line,131,8)	'From Date [SPrevFromDate]
			objItem(strKey).Add "04B-120", Mid(record_line,139,8)	'To Date [SPrevToDate]
			objItem(strKey).Add "04B-130", Mid(record_line,147,15)	'Monthly Income [SPrevMonthlyIncome]
			objItem(strKey).Add "04B-140", Mid(record_line,162,25)	'Position / Title / Type of Business [SPrevPositionTitleTypeBiz]
			objItem(strKey).Add "04B-150", Mid(record_line,187,10)	'Business Phone [SPrevBizPhone]
			
			SPrevEmpName = Trim(Mid(record_line,13,35))
			SPrevEmpStAddr = Trim(Mid(record_line,48,35))
			SPrevEmpCity = Trim(Mid(record_line,83,35))
			SPrevEmpState = Trim(Mid(record_line,118,2))
			SPrevEmpZip = Trim(Mid(record_line,120,5))
			SPrevEmpZipFour = Trim(Mid(record_line,125,4))
			SPrevSelfEmployed = Trim(Mid(record_line,129,1))
			SPrevCurrentEmpFlag = Trim(Mid(record_line,130,1))
			SPrevFromDate = Trim(Mid(record_line,131,8))
			SPrevToDate = Trim(Mid(record_line,139,8))
			SPrevMonthlyIncome = Trim(Mid(record_line,147,15))
			SPrevPositionTitleTypeBiz = Trim(Mid(record_line,162,25))
			SPrevBizPhone = Trim(Mid(record_line,187,10))
			'------------------------------------------------
			'- [04B] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("04B",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(SPrevEmpName,35,"F")
			result_str = result_str & WriteEDI(SPrevEmpStAddr,35,"F")
			result_str = result_str & WriteEDI(SPrevEmpCity,35,"F")
			result_str = result_str & WriteEDI(SPrevEmpState,2,"F")
			result_str = result_str & WriteEDI(SPrevEmpZip,5,"F")
			result_str = result_str & WriteEDI(SPrevEmpZipFour,4,"F")
			result_str = result_str & WriteEDI(SPrevSelfEmployed,1,"F")
			result_str = result_str & WriteEDI(SPrevCurrentEmpFlag,1,"F")
			result_str = result_str & WriteEDI(SPrevFromDate,8,"F")
			result_str = result_str & WriteEDI(SPrevToDate,8,"F")
			result_str = result_str & WriteEDI(SPrevMonthlyIncome,15,"E")
			result_str = result_str & WriteEDI(SPrevPositionTitleTypeBiz,25,"F")
			result_str = result_str & WriteEDI(SPrevBizPhone,10,"F")
			Response.write result_str & "<br>"
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
			
			HousingExpensePresentIndicator = Trim(Mid(record_line,13,1))
			HousingPaymentTypeCode = Trim(Mid(record_line,14,2))
			HousingPaymentAmt = Trim(Mid(record_line,16,15))
			'------------------------------------------------
			'- [05H] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("05H",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(HousingExpensePresentIndicator,1,"F")
			result_str = result_str & WriteEDI(HousingPaymentTypeCode,2,"F")
			result_str = result_str & WriteEDI(HousingPaymentAmt,15,"E")
			Response.write result_str & "<br>"
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
			
			TypeOfIncomeCode = Trim(Mid(record_line,13,2))
			IncomeAmt = Trim(Mid(record_line,15,15))
			'------------------------------------------------
			'- [05I] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("05I",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(TypeOfIncomeCode,2,"F")
			result_str = result_str & WriteEDI(IncomeAmt,15,"E")
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- 06A	VI	 For all asset types, enter data in the 06C assets segment.
		'------------------------------------------------	
		Case "06A"		
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "06A-030", Mid(record_line,13,35)	'Cash deposit toward purchase held by [CashDepositPurcHeldBy]
			objApplicant.Add "06A-040", Mid(record_line,48,15)	'Cash or Market Value [CashOrMarketValue]
			
			CashDepositPurcHeldBy = Trim(Mid(record_line,13,35))
			CashOrMarketValue = Trim(Mid(record_line,48,15))
			'------------------------------------------------
			'- [06A] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("06A",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(CashDepositPurcHeldBy,35,"F")
			result_str = result_str & WriteEDI(CashOrMarketValue,15,"E")
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- 06B	VI	 Life Insurance
		'------------------------------------------------
		Case "06B"		
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "06B-030", Mid(record_line,13,30)	'Acct. no. [LifeInsurAcctNo]
			objApplicant.Add "06B-040", Mid(record_line,43,15)	'Life Insurance Cash or Market Value [LifeInsurCashMarketVal]
			objApplicant.Add "06B-050", Mid(record_line,58,15)	'Life insurance Face Amount	[LifeInsurFaceAmt]
			
			LifeInsurAcctNo = Trim(Mid(record_line,13,30))
			LifeInsurCashMarketVal = Trim(Mid(record_line,43,15))
			LifeInsurFaceAmt = Trim(Mid(record_line,58,15))
			'------------------------------------------------
			'- [06B] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("06B",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(LifeInsurAcctNo,30,"F")
			result_str = result_str & WriteEDI(LifeInsurCashMarketVal,15,"E")
			result_str = result_str & WriteEDI(LifeInsurFaceAmt,15,"E")
			Response.write result_str & "<br>"
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
			
			AccountAssetType = Trim(Mid(record_line,13,3))	
			DepositoryStockBondName = Trim(Mid(record_line,16,35))					
			DepositoryStAddr = Trim(Mid(record_line,51,35))						
			DepositoryCity = Trim(Mid(record_line,86,35))
			DepositoryState = Trim(Mid(record_line,121,2))
			DepositoryZip = Trim(Mid(record_line,123,5))
			DepositoryZipFour = Trim(Mid(record_line,128,4))
			AssetAcctNo = Trim(Mid(record_line,132,30))
			AssetCashMarketVal = Trim( Mid(record_line,162,15))
			NumberOfStockBondShares = Trim(Mid(record_line,177,7))
			AssetDesc = Trim(Mid(record_line,184,80))
			'------------------------------------------------
			'- [06C] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("06C",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(AccountAssetType,3,"F")
			result_str = result_str & WriteEDI(DepositoryStockBondName,35,"F")
			result_str = result_str & WriteEDI(DepositoryStAddr,35,"F")
			result_str = result_str & WriteEDI(DepositoryCity,35,"F")
			result_str = result_str & WriteEDI(DepositoryState,2,"F")
			result_str = result_str & WriteEDI(DepositoryZip,5,"F")
			result_str = result_str & WriteEDI(DepositoryZipFour,4,"F")
			result_str = result_str & WriteEDI(AssetAcctNo,30,"F")
			result_str = result_str & WriteEDI(AssetCashMarketVal,15,"E")
			result_str = result_str & WriteEDI(NumberOfStockBondShares,7,"E")
			result_str = result_str & WriteEDI(AssetDesc,80,"F")
			Response.write result_str & "<br>"
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
			
			AutomobileMakeModel = Trim(Mid(record_line,13,30))
			AutomobileYear = Trim(Mid(record_line,43,4))
			AutomobileCashMarketVal = Trim(Mid(record_line,47,15))
			'------------------------------------------------
			'- [06D] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("06D",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(AutomobileMakeModel,30,"F")
			result_str = result_str & WriteEDI(AutomobileYear,4,"F")
			result_str = result_str & WriteEDI(AutomobileCashMarketVal,15,"E")
			Response.write result_str & "<br>"
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
			
			ExpenseTypeCode = Trim(Mid(record_line,13,3))
			MonthlyPaymentAmt = Trim(Mid(record_line,16,15))
			MonthsLeftToPay = Trim(Mid(record_line,31,3))
			AlimonyCSSperateOwedTo = Trim(Mid(record_line,34,60))
			'------------------------------------------------
			'- [06F] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("06F",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(ExpenseTypeCode,3,"F")
			result_str = result_str & WriteEDI(MonthlyPaymentAmt,15,"F")
			result_str = result_str & WriteEDI(MonthsLeftToPay,3,"E")
			result_str = result_str & WriteEDI(AlimonyCSSperateOwedTo,60,"E")
			Response.write result_str & "<br>"
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
			objItem(strKey).Add "06G-090", Mid(record_line,95,2)	'Type of Property [REOTypeOfProp]
			objItem(strKey).Add "06G-100", Mid(record_line,97,15)	'Present Market Value [REOPresentMarketValue]
			objItem(strKey).Add "06G-110", Mid(record_line,112,15)	'Amount of Mortgages & Liens [REOAmtMortgageLiens]
			objItem(strKey).Add "06G-120", Mid(record_line,127,15)	'Gross Rental Income [REOGrossRentalIncome]
			objItem(strKey).Add "06G-130", Mid(record_line,142,15)	'Mortgage Payments [REOMortgagePayment]
			objItem(strKey).Add "06G-140", Mid(record_line,157,15)	'Insurance, Maintenance Taxes & Misc. [InsurMaintenanceTaxMisc]
			objItem(strKey).Add "06G-150", Mid(record_line,172,15)	'Net Rental Income [REONetRentalIncome]
			objItem(strKey).Add "06G-160", Mid(record_line,187,1)	'Current Residence Indicator [REOCurResidenceIndicator]
			objItem(strKey).Add "06G-170", Mid(record_line,188,1)	'Subject Property Indicator [REOSubjectPropIndicator]
			objItem(strKey).Add "06G-180", Mid(record_line,189,2)	'REO Asset ID [REOAssetID]
			objItem(strKey).Add "06G-190", Mid(record_line,191,15)	'Reserved For Future Use [REOReservedFutureUse]
			
			REOPropStAddr = Trim(Mid(record_line,13,35))	
			REOPropCity = Trim(Mid(record_line,48,35))	
			REOPropState = Trim(Mid(record_line,83,2))	
			REOPropZip = Trim(Mid(record_line,85,5))	
			REOPropZipFour = Trim(Mid(record_line,90,4))	
			REOPropDisposition = Trim(Mid(record_line,94,1))	
			REOTypeOfProp = Trim(Mid(record_line,95,2))	
			REOPresentMarketValue = Trim(Mid(record_line,97,15))	
			REOAmtMortgageLiens = Trim(Mid(record_line,112,15))	
			REOGrossRentalIncome = Trim(Mid(record_line,127,15))	
			REOMortgagePayment = Trim(Mid(record_line,142,15))	
			InsurMaintenanceTaxMisc = Trim(Mid(record_line,157,15))	
			REONetRentalIncome = Trim( Mid(record_line,172,15))	
			REOCurResidenceIndicator = Trim(Mid(record_line,187,1))	
			REOSubjectPropIndicator = Trim(Mid(record_line,188,1))	
			REOAssetID = Trim(Mid(record_line,189,2))	
			REOReservedFutureUse = Trim(Mid(record_line,191,15))	
			
			'------------------------------------------------
			'- [06G] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("06G",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(REOPropStAddr,35,"F")
			result_str = result_str & WriteEDI(REOPropCity,35,"F")
			result_str = result_str & WriteEDI(REOPropState,2,"F")
			result_str = result_str & WriteEDI(REOPropZip,5,"F")
			result_str = result_str & WriteEDI(REOPropZipFour,4,"F")
			result_str = result_str & WriteEDI(REOPropDisposition,1,"F")
			result_str = result_str & WriteEDI(REOTypeOfProp,2,"F")
			result_str = result_str & WriteEDI(REOPresentMarketValue,15,"F")
			result_str = result_str & WriteEDI(REOAmtMortgageLiens,15,"F")
			result_str = result_str & WriteEDI(REOGrossRentalIncome,15,"F")
			result_str = result_str & WriteEDI(REOMortgagePayment,15,"F")
			result_str = result_str & WriteEDI(InsurMaintenanceTaxMisc,15,"F")
			result_str = result_str & WriteEDI(REONetRentalIncome,15,"F")
			result_str = result_str & WriteEDI(REOCurResidenceIndicator,1,"F")
			result_str = result_str & WriteEDI(REOSubjectPropIndicator,1,"F")
			result_str = result_str & WriteEDI(REOAssetID,2,"F")
			result_str = result_str & WriteEDI(REOReservedFutureUse,15,"F")
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- 06H	VI	 Alias
		'------------------------------------------------
		Case "06H"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "06H-030", Mid(record_line,13,35) 'Alternate First Name [AliasFirstName]
			objApplicant.Add "06H-040", Mid(record_line,48,35) 'Alternate Middle Name [AliasMidName]
			objApplicant.Add "06H-050", Mid(record_line,83,35) 'Alternate Last Name [AliasLastName]
			objApplicant.Add "06H-060", Mid(record_line,118,15)'Reserved For Future Use [ReservedFutureUse6_1]
			objApplicant.Add "06H-070", Mid(record_line,153,15)'Reserved For Future Use [ReservedFutureUse6_2]
			
			AliasFirstName = Trim(Mid(record_line,13,35))
			AliasMidName = Trim(Mid(record_line,48,35))
			AliasLastName = Trim(Mid(record_line,83,35))
			ReservedFutureUse6_1 = Trim(Mid(record_line,118,15))
			ReservedFutureUse6_2 = Trim(Mid(record_line,153,15))
			'------------------------------------------------
			'- [06H] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("06H",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(AliasFirstName,35,"F")
			result_str = result_str & WriteEDI(AliasMidName,35,"F")
			result_str = result_str & WriteEDI(AliasLastName,35,"F")
			result_str = result_str & WriteEDI(ReservedFutureUse6_1,15,"E")
			result_str = result_str & WriteEDI(ReservedFutureUse6_2,15,"E")
			Response.write result_str & "<br>"
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
			
			LiabilityType = Trim(Mid(record_line,13,2))
			CreditorName = Trim(Mid(record_line,15,35))
			CreditorStAddr = Trim(Mid(record_line,50,35))
			CreditorCity = Trim(Mid(record_line,85,35))
			CreditorState = Trim(Mid(record_line,120,2))
			CreditorZip = Trim(Mid(record_line,122,5))
			CreditorZipFour = Trim(Mid(record_line,127,4))
			LiabilityAcctNo = Trim(Mid(record_line,131,30))
			LiabilityMonPaymentAmt = Trim(Mid(record_line,161,15))
			LiabilityMonLeftToPay = Trim(Mid(record_line,176,3))
			UnpaidBalance = Trim(Mid(record_line,179,15))
			LiabilityPaidClosing = Trim(Mid(record_line,194,1))
			REOAssetID = Trim(Mid(record_line,195,2))
			ResubordinatedIndicator = Trim(Mid(record_line,197,1))
			OmittedIndicator = Trim(Mid(record_line,198,1))
			SubjectPropIndicator = Trim(Mid(record_line,199,1))
			RentalPropIndicator = Trim(Mid(record_line,200,1))
			'------------------------------------------------
			'- [06L] Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("06L",3,"F")
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(LiabilityType,2,"F")
			result_str = result_str & WriteEDI(CreditorName,35,"F")
			result_str = result_str & WriteEDI(CreditorStAddr,35,"F")
			result_str = result_str & WriteEDI(CreditorCity,35,"F")
			result_str = result_str & WriteEDI(CreditorState,2,"F")
			result_str = result_str & WriteEDI(CreditorZip,5,"F")
			result_str = result_str & WriteEDI(CreditorZipFour,4,"F")
			result_str = result_str & WriteEDI(LiabilityAcctNo,30,"F")
			result_str = result_str & WriteEDI(LiabilityMonPaymentAmt,15,"E")
			result_str = result_str & WriteEDI(LiabilityMonLeftToPay,3,"E")
			result_str = result_str & WriteEDI(UnpaidBalance,15,"E")
			result_str = result_str & WriteEDI(LiabilityPaidClosing,1,"F")
			result_str = result_str & WriteEDI(REOAssetID,2,"F")
			result_str = result_str & WriteEDI(ResubordinatedIndicator,1,"F")
			result_str = result_str & WriteEDI(OmittedIndicator,1,"F")
			result_str = result_str & WriteEDI(SubjectPropIndicator,1,"F")
			result_str = result_str & WriteEDI(RentalPropIndicator,1,"F")
			Response.write result_str & "<br>"
		'------------------------------------------------ 
		'- 06S	VI	 Undrawn HELOC and IPCs
		'------------------------------------------------
		Case "06S"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "06S-030", Mid(record_line,13,3) 'Summary Amount Type Code [HELOCSummaryAmtTypeCode]
			objApplicant.Add "06S-040", Mid(record_line,16,15)'Amount [HELEOCAmt]
			
			HELOCSummaryAmtTypeCode = Trim(Mid(record_line,13,3))
			HELEOCAmt = Trim(Mid(record_line,16,15))
			
			For i=0 to 3-len(HELOCSummaryAmtTypeCode)-1
			 HELOCSummaryAmtTypeCode = "&nbsp;" & HELOCSummaryAmtTypeCode
			Next
			For i=0 to 15-len(HELEOCAmt)-1
			 HELEOCAmt = "&nbsp;" & HELEOCAmt
			Next
			
			Response.write record_id & ssn & HELOCSummaryAmtTypeCode & HELEOCAmt & "<br>"
		'------------------------------------------------
		'- 08A	VIII Declarations
		'------------------------------------------------
		Case "08A"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "08A-030", Mid(record_line,13,1)   'a. Are there any outstanding judgments against you? [DeclarationsA]
			objApplicant.Add "08A-040", Mid(record_line,14,1)   'b. Have you been declared bankrupt within the past 7 years?[DeclarationsB]
			objApplicant.Add "08A-050", Mid(record_line,15,1)   'c. Have you had property Foreclosed upon or given title or deed in lieu thereof in the last 7 years? [DeclarationsC]
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
			
			DeclarationsA = Trim(Mid(record_line,13,1))
			DeclarationsB = Trim(Mid(record_line,14,1))
			DeclarationsC = Trim(Mid(record_line,15,1))
			DeclarationsD = Trim(Mid(record_line,16,1))
			DeclarationsE = Trim(Mid(record_line,17,1))
			DeclarationsF = Trim(Mid(record_line,18,1))
			DeclarationsG = Trim(Mid(record_line,19,1))
			DeclarationsH = Trim(Mid(record_line,20,1))
			DeclarationsI = Trim(Mid(record_line,21,1))
			DeclarationsJ = Trim(Mid(record_line,22,2))
			DeclarationsL = Trim(Mid(record_line,24,1))
			DeclarationsM = Trim(Mid(record_line,25,1))
			DeclarationsM1 = Trim(Mid(record_line,26,1))
			DeclarationsM2 = Trim(Mid(record_line,27,2))
			
			'------------------------------------------------
			'- 08A	VIII	Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("08A",3,"F")
			'Borrower SSN
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(DeclarationsA,1,"F")
			result_str = result_str & WriteEDI(DeclarationsB,1,"F")
			result_str = result_str & WriteEDI(DeclarationsC,1,"F")
			result_str = result_str & WriteEDI(DeclarationsD,1,"F")
			result_str = result_str & WriteEDI(DeclarationsE,1,"F")
			result_str = result_str & WriteEDI(DeclarationsF,1,"F")
			result_str = result_str & WriteEDI(DeclarationsG,1,"F")
			result_str = result_str & WriteEDI(DeclarationsH,1,"F")
			result_str = result_str & WriteEDI(DeclarationsI,1,"F")
			result_str = result_str & WriteEDI(DeclarationsJ,2,"F")
			result_str = result_str & WriteEDI(DeclarationsL,1,"F")
			result_str = result_str & WriteEDI(DeclarationsM,1,"F")
			result_str = result_str & WriteEDI(DeclarationsM1,1,"F")
			result_str = result_str & WriteEDI(DeclarationsM2,2,"F")
			Response.write result_str & "<br>"
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
			
			DeclarationTypeCode = Trim(Mid(record_line,13,2))
			DeclarationExplanation = Trim(Mid(record_line,15,255))
			'------------------------------------------------
			'- 08B	VIII	Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("08B",3,"F")
			'Borrower SSN
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(DeclarationTypeCode,2,"F")
			result_str = result_str & WriteEDI(DeclarationExplanation,255,"F")
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- 09A	IX	 Acknowledgment and Agreement
		'------------------------------------------------
		Case "09A"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "09A-030", Mid(record_line,13,8) 'Signature Date [SignatureDate]
			
			SignatureDate = Trim(Mid(record_line,13,8))
			'------------------------------------------------
			'- 09A	IX	 Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("09A",3,"F")
			'Borrower SSN
			result_str = result_str & WriteEDI(ssn,9,"F")
			result_str = result_str & WriteEDI(SignatureDate,8,"F")
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- 10A	X	 InFormation For Government Monitoring Purposes
		'------------------------------------------------
		Case "10A"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "10A-030", Mid(record_line,13,1) 	'I do not wish to furnish this inFormation [IDoNotFurnishMyInfo]
			objApplicant.Add "10A-040", Mid(record_line,14,1)	'Ethnicity [Ethnicity]
			objApplicant.Add "10A-050", Mid(record_line,15,30)	'Filler [Filler]
			objApplicant.Add "10A-060", Mid(record_line,45,1)	'Sex [Sex]
			
			IDoNotFurnishMyInfo = Trim(Mid(record_line,13,1))
			Ethnicity = Trim(Mid(record_line,14,1))
			Filler = Trim( Mid(record_line,15,30))
			Sex = Trim(Mid(record_line,45,1))
			'------------------------------------------------
			'- 10A	X	 Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("10A",3,"F")
			'Borrower SSN
			result_str = result_str & WriteEDI(ssn,9,"F")
			'I do not wish to furnish this inFormation
			result_str = result_str & WriteEDI(IDoNotFurnishMyInfo,1,"F")
			'Ethnicity
			result_str = result_str & WriteEDI(Ethnicity,1,"F")
			'Filler
			result_str = result_str & WriteEDI(Filler,30,"F")
			'Sex
			result_str = result_str & WriteEDI(Sex,1,"F")
			
			Response.write result_str & "<br>"
		'------------------------------------------------
		'- 10R	X	 InFormation For Government Monitoring Purposes
		'------------------------------------------------
		Case "10R"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "10R-030", Mid(record_line,13,2)	'Race type [RaceType]
			
			RaceType = Trim(Mid(record_line,13,2))
			'------------------------------------------------
			'- 10R	X	 Printing EDI
			'------------------------------------------------
			result_str = ""
			'Header
			result_str = result_str & WriteEDI("10R",3,"F")
			'Borrower SSN
			result_str = result_str & WriteEDI(ssn,9,"F")
			'Race type
			result_str = result_str & WriteEDI(RaceType,2,"F")
			
			Response.write result_str & "<br>"
	End Select

Loop
'================================================
'= Printing EDI Footer
'================================================
	'------------------------------------------------
	'- 000 line
	'------------------------------------------------
	result_str = ""
	'Header
	result_str = result_str & WriteEDI("000",3,"F")
	result_str = result_str & WriteEDI("11",3,"F")
	result_str = result_str & WriteEDI("3.20",5,"F")
	Response.write result_str & "<br>"
	'------------------------------------------------
	'- TT line
	'------------------------------------------------
	result_str = ""
	result_str = result_str & WriteEDI("TT",3,"F")
	result_str = result_str & WriteEDI("1",9,"F")
	Response.write result_str & "<br>"
	'------------------------------------------------
	'- ET line
	'------------------------------------------------
	result_str = ""
	result_str = result_str & WriteEDI("ET",3,"F")
	result_str = result_str & WriteEDI("0",9,"F")
	Response.write result_str & "<br>"
'================================================
'= Printing Application
'================================================
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
				print_dupicate_code("Other Credit")
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
'= WriteEDI
'==================================================================================================
Function WriteEDI(str, length, direction)
	Dim space : space = ""
	For i=len(str) To length-1
		space = space & "&nbsp;"
	Next
	If direction = "F" Then
		str = str & space
	ElseIf direction = "E" Then
		str = space & str
	End If
	WriteEDI = str
End Function
'==================================================================================================
'= LeftCut
'==================================================================================================
Function LeftCut(strString, intCut)
    Dim intPos, chrTemp, strCut, intLength
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