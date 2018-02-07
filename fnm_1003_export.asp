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

Dim i

Dim LoanQualificationUsed '00A
Dim LoanQualificationNotUsed
Dim MortgageAppliedFor '01A
Dim MortgageAppliedForOther
Dim AgencyCaseNumber
Dim CaseNumber
Dim LoanAmt
Dim InterestRate
Dim NoOfMonth
Dim AmortizationType
Dim AmortizationTypeOtherExplain
Dim ARMTrxtualDesc
Dim PropStAddress '02A
Dim PropCity
Dim PropState
Dim PropZip
Dim PropZipPlusFour
Dim NoOfUnits
Dim LegalDescSubjPropCode
Dim LegalDescSubjProp
Dim YearBuilt
Dim ReservedFutureUse '02B
Dim PurposeOfLoan
Dim PurposeOfLoanOther
Dim PropWillBe
Dim MannerTitleWillBeHeld
Dim EstateWillBeHeldIn
Dim LeaseholdExpirationDate
Dim TitleholderName '02C
Dim YearAcquired '02D
Dim OriginalCost
Dim AmtExistingLiens
Dim PresentValueOfLot
Dim CostOfImprovements
Dim PurposeOfRefinance
Dim DescribeImprovements
Dim DescImporvMadeToBeMade
Dim DescImporvCost
Dim DownPaymentTypeCode '02E
Dim DownPaymentamt
Dim DownPaymentExplanation
Dim PurchasePrice '07A
Dim AlterationsImprovRepair
Dim Land
Dim Refinance
Dim EstimatedPrepaidItems
Dim EstimatedClosingCosts
Dim PMIMIPFundingFee
Dim Discount
Dim SubordinateFinancing
Dim ClosingCostPaidBySeller
Dim PMIMIPFundingFeeFinan
Dim OtherCreditTypeCode '07B
Dim AmtOfOtherCredit
Dim LoanOriginatorName '10B
Dim InterviewDate
Dim LOPhoneNo
Dim LOCompanyName
Dim LOCompanyStAddr
Dim LOCompanyStAddr2
Dim LOCompanyCity
Dim LOCompanyStateCode
Dim LOCompanyZip
Dim LOCompanyZipFour
Dim ApplicantIndicator '03A
Dim ApplicantFirstName
Dim ApplicantMidName
Dim ApplicantLastName
Dim ApplicantGeneration
Dim HomePhone
Dim Age
Dim YrsSchool
Dim MaritalStatus
Dim DependantsNo
Dim CompletedJoinNotJoin
Dim CrossRefNumber
Dim DateOfBirth
Dim EmailAddr
Dim DependantsAge '03B
Dim PresentFormer '03C
Dim ResidenceStAddr
Dim ResidenceCity
Dim ResidenceState
Dim ResidenceZip
Dim ResidenceZipFour
Dim OwnRentLivingRentFree
Dim AddrNoYrs
Dim AddrNoMonth
Dim ApplicantAddrCounrtry
Dim EmpName '04A
Dim EmpStAddr
Dim EmpCity
Dim EmpState
Dim EmpZip
Dim EmpZipFour
Dim SelfEmployed
Dim YrsOnThisJob
Dim MonthsOnThisJob
Dim YrsEmpInThisLineWork
Dim PositionTitleTypeBiz
Dim BizPhone
Dim SPrevEmpName '04B
Dim SPrevEmpStAddr
Dim SPrevEmpCity
Dim SPrevEmpState
Dim SPrevEmpZip
Dim SPrevEmpZipFour
Dim SPrevSelfEmployed
Dim SPrevCurrentEmpFlag
Dim SPrevFromDate
Dim SPrevToDate
Dim SPrevMonthlyIncome
Dim SPrevPositionTitleTypeBiz
Dim SPrevBizPhone
Dim HousingExpensePresentIndicator '05H
Dim HousingPaymentTypeCode
Dim HousingPaymentAmt
Dim TypeOfIncomeCode '05I
Dim IncomeAmt
Dim CashDepositPurcHeldBy '06A
Dim CashOrMarketValue
Dim LifeInsurAcctNo '06B
Dim LifeInsurCashMarketVal
Dim LifeInsurFaceAmt
Dim AccountAssetType '06C
Dim DepositoryStockBondName
Dim DepositoryStAddr
Dim DepositoryCity
Dim DepositoryState
Dim DepositoryZip
Dim DepositoryZipFour
Dim AssetAcctNo
Dim AssetCashMarketVal
Dim NumberOfStockBondShares
Dim AssetDesc
Dim AutomobileMakeModel '06D
Dim AutomobileYear
Dim AutomobileCashMarketVal
Dim MonthlyPaymentAmt '06F
Dim MonthsLeftToPay
Dim AlimonyCSSperateOwedTo
Dim REOPropStAddr '06G
Dim REOPropCity
Dim REOPropState
Dim REOPropZip
Dim REOPropZipFour
Dim REOPropDisposition
Dim REOTypeOfProp
Dim REOPresentMarketValue
Dim REOAmtMortgageLiens
Dim REOGrossRentalIncome
Dim REOMortgagePayment
Dim InsurMaintenanceTaxMisc
Dim REONetRentalIncome
Dim REOCurResidenceIndicator
Dim REOSubjectPropIndicator
Dim REOAssetID
Dim REOReservedFutureUse
Dim AliasMidNam '06H
Dim AliasLastName
Dim ReservedFutureUse6_1
Dim ReservedFutureUse6_2
Dim LiabilityType '06L
Dim CreditorName
Dim CreditorStAddr
Dim CreditorCity
Dim CreditorState
Dim CreditorZip
Dim CreditorZipFour
Dim LiabilityAcctNo
Dim LiabilityMonPaymentAmt
Dim LiabilityMonLeftToPay
Dim UnpaidBalance
Dim LiabilityPaidClosing
Dim ResubordinatedIndicator
Dim OmittedIndicator
Dim SubjectPropIndicator
Dim RentalPropIndicator
Dim HELOCSummaryAmtTypeCode '06S
Dim HELEOCAmt
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
Dim DeclarationTypeCode '08B
Dim DeclarationExplanation
Dim SignatureDate
Dim IDoNotFurnishMyInfo '10A
Dim Ethnicity
Dim Filler
Dim Sex
Dim ThisAppWasTakenBy
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
objApplication.Add "Other Credit Type"	, Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Title Holder"		, Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Down Payment"		, Server.CreateObject("Scripting.Dictionary")
'================================================
'= Reading EDI File
'================================================
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/C0101904_1.txt"),1,true)
Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test2.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test3.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test4.txt"),1,true)
'Set objEDI = objFS.OpenTextFile(Server.MapPath("edi_test/test5.txt"),1,true)
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
			
			Response.write record_id & LoanQualificationUsed & LoanQualificationNotUsed & "<br>"
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
			objApplication.Add "01A-110", Mid(record_line,238,80)  'ARM Textual Description [ARMTrxtualDesc]
			
			MortgageAppliedFor = Trim(Mid(record_line,4,2))
			MortgageAppliedForOther = Trim(Mid(record_line,6,80))
			AgencyCaseNumber =Trim( Mid(record_line,86,30))
			CaseNumber = Trim(Mid(record_line,116,15))
			LoanAmt = Trim(Mid(record_line,131,15)) 
			InterestRate = Trim(Mid(record_line,146,7))
			NoOfMonth = Trim(Mid(record_line,153,3))
			AmortizationType = Trim(Mid(record_line,156,2))
			AmortizationTypeOtherExplain = Trim(Mid(record_line,158,80))
			ARMTrxtualDesc = Trim(Mid(record_line,238,80))
			
			For i=0 to 80-len(MortgageAppliedForOther)-1
			 MortgageAppliedForOther = "&nbsp;" & MortgageAppliedForOther 
			Next	
			For i=0 to 30-len(AgencyCaseNumber)-1
			 AgencyCaseNumber = AgencyCaseNumber & "&nbsp;"
			Next
			For i=0 to 15-len(CaseNumber)-1
			 CaseNumber = CaseNumber & "&nbsp;"
			Next
			For i=0 to 15-len(LoanAmt)-1
			 LoanAmt = "&nbsp;" & LoanAmt
			Next
			For i=0 to 7-len(InterestRate)-1
			 InterestRate = "&nbsp;" & InterestRate
			Next
			For i=0 to 3-len(NoOfMonth)-1
			 NoOfMonth = "&nbsp;" & NoOfMonth
			Next
			For i=0 to 2-len(AmortizationType)-1
			 AmortizationType = "&nbsp;" & AmortizationType
			Next
			For i=0 to 80-len(AmortizationTypeOtherExplain)-1
			 AmortizationTypeOtherExplain = "&nbsp;" & AmortizationTypeOtherExplain
			Next
			
			Response.write record_id & MortgageAppliedFor & MortgageAppliedForOther & AgencyCaseNumber &_
			CaseNumber & LoanAmt & InterestRate & NoOfMonth & AmortizationType & AmortizationTypeOtherExplain & "<br>"
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
			
			For i=0 to 50-len(PropStAddress)-1
			 PropStAddress = PropStAddress & "&nbsp;"
			Next
			For i=0 to 35-len(PropCity)-1
			 PropCity = PropCity & "&nbsp;"
			Next
			For i=0 to 2-len(PropState)-1
			 PropState = PropState & "&nbsp;"
			Next
			For i=0 to 5-len(PropZip)-1
			 PropZip = PropZip & "&nbsp;"
			Next
			For i=0 to 4-len(PropZipPlusFour)-1
			 PropZipPlusFour = PropZipPlusFour & "&nbsp;"
			Next
			For i=0 to 3-len(NoOfUnits)-1
			 NoOfUnits = NoOfUnits & "&nbsp;"
			Next
			For i=0 to 2-len(LegalDescSubjPropCode)-1
			 LegalDescSubjPropCode = LegalDescSubjPropCode & "&nbsp;"
			Next
			For i=0 to 80-len(LegalDescSubjProp)-1
			 LegalDescSubjProp = LegalDescSubjProp & "&nbsp;"
			Next
			For i=0 to 4-len(YearBuilt)-1
			 YearBuilt = YearBuilt & "&nbsp;"
			Next
			
			Response.write record_id & PropStAddress & PropCity & PropState & PropZip & PropZipPlusFour & NoOfUnits &_
			LegalDescSubjPropCode & LegalDescSubjProp & YearBuilt & "<br>"
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
			
			For i=0 to 2-len(ReservedFutureUse)-1
			 ReservedFutureUse = "&nbsp;" & ReservedFutureUse 
			Next
			For i=0 to 2-len(PurposeOfLoan)-1
			 PurposeOfLoan = "&nbsp;" & PurposeOfLoan
			Next
			For i=0 to 80-len(PurposeOfLoanOther)-1
			 PurposeOfLoanOther = "&nbsp;" & PurposeOfLoanOther 
			Next
			For i=0 to 1-len(PropWillBe)-1
			 PropWillBe = "&nbsp;" & PropWillBe
			Next
			For i=0 to 60-len(MannerTitleWillBeHeld)-1
			 MannerTitleWillBeHeld = MannerTitleWillBeHeld & "&nbsp;"
			Next
			For i=0 to 1-len(EstateWillBeHeldIn)-1
			 EstateWillBeHeldIn = "&nbsp;" & EstateWillBeHeldIn 
			Next
			For i=0 to 8-len(LeaseholdExpirationDate)-1
			 LeaseholdExpirationDate = "&nbsp;" & LeaseholdExpirationDate 
			Next
			
			Response.write record_id & ReservedFutureUse & PurposeOfLoan & PurposeOfLoanOther & PropWillBe & MannerTitleWillBeHeld &_
			EstateWillBeHeldIn & LeaseholdExpirationDate & "<br>"
		'------------------------------------------------
		'- [02C]	II	 Title Holder
		'------------------------------------------------
		Case "02C" 
			titleName = Mid(record_line,4,60)
			Set objTitleHolder = GetDuplicateData(objApplication("Title Holder"),titleName)
			objTitleHolder.Add "02C-020", Mid(record_line,4,60) 'Titleholder Name [TitleholderName]
			
			Response.write record_id & titleName & "<br>"
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
			
			For i=0 to 4-len(YearAcquired)-1
			 YearAcquired = "&nbsp;" & YearAcquired 
			Next
			For i=0 to 15-len(OriginalCost)-1
			 OriginalCost = "&nbsp;" & OriginalCost
			Next
			For i=0 to 15-len(AmtExistingLiens)-1
			 AmtExistingLiens = "&nbsp;" & AmtExistingLiens
			Next
			For i=0 to 15-len(PresentValueOfLot)-1
			 PresentValueOfLot = "&nbsp;" & PresentValueOfLot
			Next
			For i=0 to 15-len(CostOfImprovements)-1
			 CostOfImprovements = "&nbsp;" & CostOfImprovements 
			Next
			For i=0 to 2-len(PurposeOfRefinance)-1
			 PurposeOfRefinance = "&nbsp;" & PurposeOfRefinance
			Next
			For i=0 to 80-len(DescribeImprovements)-1
			 DescribeImprovements = "&nbsp;" & DescribeImprovements
			Next
			For i=0 to 1-len(DescImporvMadeToBeMade)-1
			 DescImporvMadeToBeMade = "&nbsp;" & DescImporvMadeToBeMade 
			Next
			For i=0 to 15-len(DescImporvCost)-1
			 DescImporvCost = "&nbsp;" & DescImporvCost
			Next
			
			Response.write record_id & YearAcquired & OriginalCost & AmtExistingLiens & PresentValueOfLot & CostOfImprovements &_
			PurposeOfRefinance & DescribeImprovements & DescImporvMadeToBeMade & DescImporvCost & "<br>"
		'------------------------------------------------
		'- [02E]	II	 Down Payment
		'------------------------------------------------
		Case "02E" 
			typeCode = Mid(record_line,4,2)
			Set objDownPayment = GetDuplicateData(objApplication("Down Payment"),typeCode)
			objDownPayment.Add "02E-020", Mid(record_line,4,2) 	'Down Payment Type Code [DownPaymentTypeCode]
			objDownPayment.Add "02E-030", Mid(record_line,6,15) 'Down Payment Amount [DownPaymentamt]
			objDownPayment.Add "02E-040", Mid(record_line,21,80)'Down Payment Explanation [DownPaymentExplanation]
			
			DownPaymentTypeCode = Trim(Mid(record_line,4,2))
			DownPaymentamt = Trim(Mid(record_line,6,15))
			DownPaymentExplanation = Trim(Mid(record_line,21,80))
			
			For i=0 to 2-len(DownPaymentTypeCode)-1
			 DownPaymentTypeCode = "&nbsp;" & DownPaymentTypeCode
			Next
			For i=0 to 15-len(DownPaymentamt)-1
			 DownPaymentamt = "&nbsp;" & DownPaymentamt
			Next
			For i=0 to 80-len(DownPaymentExplanation)-1
			 DownPaymentExplanation = "&nbsp;" & DownPaymentExplanation
			Next
			
			Response.write record_id & DownPaymentTypeCode & DownPaymentamt & DownPaymentExplanation & "<br>"
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
			
			For i=0 to 15-len(PurchasePrice)-1
			 PurchasePrice = "&nbsp;" & PurchasePrice
			Next
			For i=0 to 15-len(AlterationsImprovRepair)-1
			 AlterationsImprovRepair = "&nbsp;" & AlterationsImprovRepair
			Next
			For i=0 to 15-len(Land)-1
			 Land = "&nbsp;" & Land
			Next
			For i=0 to 15-len(Refinance)-1
			 Refinance = "&nbsp;" & Refinance
			Next
			For i=0 to 15-len(EstimatedPrepaidItems)-1
			 EstimatedPrepaidItems = "&nbsp;" & EstimatedPrepaidItems
			Next
			For i=0 to 15-len(EstimatedClosingCosts)-1
			 EstimatedClosingCosts = "&nbsp;" & EstimatedClosingCosts
			Next
			For i=0 to 15-len(PMIMIPFundingFee)-1
			 PMIMIPFundingFee = "&nbsp;" & PMIMIPFundingFee
			Next
			For i=0 to 15-len(Discount)-1
			 Discount = "&nbsp;" & Discount
			Next
			For i=0 to 15-len(SubordinateFinancing)-1
			 SubordinateFinancing = "&nbsp;" & SubordinateFinancing
			Next
			For i=0 to 15-len(ClosingCostPaidBySeller)-1
			 ClosingCostPaidBySeller = "&nbsp;" & ClosingCostPaidBySeller
			Next
			For i=0 to 15-len(PMIMIPFundingFeeFinan)-1
			 PMIMIPFundingFeeFinan = "&nbsp;" & PMIMIPFundingFeeFinan
			Next
			
			Response.write record_id & PurchasePrice & AlterationsImprovRepair & Land & Refinance &EstimatedPrepaidItems &_
			EstimatedClosingCosts & PMIMIPFundingFee & Discount & SubordinateFinancing & ClosingCostPaidBySeller & PMIMIPFundingFeeFinan & "<br>"
			
		'------------------------------------------------
		'- [07B]	VII	 Other Credits
		'------------------------------------------------
		Case "07B" 
			typeCode = Mid(record_line,4,2)
			Set objOtherCredit = GetDuplicateData(objApplication("Other Credit Type"),typeCode)
			objOtherCredit.Add "07B-020", Mid(record_line,4,2) 		'Other Credit Type Code [OtherCreditTypeCode]
			objOtherCredit.Add "07B-030", Mid(record_line,6,15) 	'Amount of Other Credit [AmtOfOtherCredit]
			
			OtherCreditTypeCode = Trim(Mid(record_line,4,2))
			AmtOfOtherCredit = Trim(Mid(record_line,6,15))
			
			For i=0 to 2-len(OtherCreditTypeCode)-1
			 OtherCreditTypeCode = "&nbsp;" & OtherCreditTypeCode 
			Next
			For i=0 to 15-len(AmtOfOtherCredit)-1
			 AmtOfOtherCredit = "&nbsp;" & AmtOfOtherCredit 
			Next
			
			Response.write record_id & OtherCreditTypeCode & AmtOfOtherCredit & "<br>"
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
			
			For i=0 to 1-len(ThisAppWasTakenBy)-1
			 ThisAppWasTakenBy = "&nbsp;" & ThisAppWasTakenBy
			Next
			For i=0 to 60-len(LoanOriginatorName)-1
			 LoanOriginatorName =  LoanOriginatorName & "&nbsp;"
			Next
			For i=0 to 8-len(InterviewDate)-1
			 InterviewDate = "&nbsp;" & InterviewDate 
			Next
			For i=0 to 10-len(LOPhoneNo)-1
			 LOPhoneNo = "&nbsp;" & LOPhoneNo 
			Next
			For i=0 to 35-len(LOCompanyName)-1
			 LOCompanyName = LOCompanyName & "&nbsp;"
			Next
			For i=0 to 35-len(LOCompanyStAddr)-1
			 LOCompanyStAddr = LOCompanyStAddr & "&nbsp;"
			Next
			For i=0 to 35-len(LOCompanyStAddr2)-1
			 LOCompanyStAddr2 = "&nbsp;" & LOCompanyStAddr2 
			Next
			For i=0 to 35-len(LOCompanyCity)-1
			 LOCompanyCity = LOCompanyCity & "&nbsp;"
			Next
			For i=0 to 2-len(LOCompanyStateCode)-1
			 LOCompanyStateCode = "&nbsp;" & LOCompanyStateCode 
			Next
			For i=0 to 5-len(LOCompanyZip)-1
			 LOCompanyZip = "&nbsp;" & LOCompanyZip 
			Next
			For i=0 to 4-len(LOCompanyZipFour)-1
			 LOCompanyZipFour = "&nbsp;" & LOCompanyZipFour 
			Next
			
			Response.write record_id & ThisAppWasTakenBy & LoanOriginatorName & InterviewDate & LOPhoneNo & LOCompanyName & LOCompanyStAddr &_
			LOCompanyStAddr2 & LOCompanyCity & LOCompanyStateCode & LOCompanyZip & LOCompanyZipFour & "<br>"
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
			
			For i=0 to 2-len(ApplicantIndicator)-1
			 ApplicantIndicator = "&nbsp;" & ApplicantIndicator 
			Next
			For i=0 to 35-len(ApplicantFirstName)-1
			 ApplicantFirstName = ApplicantFirstName & "&nbsp;"
			Next
			For i=0 to 35-len(ApplicantMidName)-1
			 ApplicantMidName = ApplicantMidName & "&nbsp;"
			Next
			For i=0 to 35-len(ApplicantLastName)-1
			 ApplicantLastName = ApplicantLastName & "&nbsp;"
			Next
			For i=0 to 4-len(ApplicantGeneration)-1
			 ApplicantGeneration = "&nbsp;" & ApplicantGeneration 
			Next
			For i=0 to 10-len(HomePhone)-1
			 HomePhone = HomePhone & "&nbsp;"
			Next
			For i=0 to 3-len(Age)-1
			 Age = Age & "&nbsp;"
			Next
			For i=0 to 2-len(YrsSchool)-1
			 YrsSchool = YrsSchool  & "&nbsp;"
			Next
			For i=0 to 1-len(MaritalStatus)-1
			 MaritalStatus =  MaritalStatus & "&nbsp;"
			Next
			For i=0 to 2-len(DependantsNo)-1
			 DependantsNo = DependantsNo & "&nbsp;"
			Next
			For i=0 to 1-len(CompletedJoinNotJoin)-1
			 CompletedJoinNotJoin = ompletedJoinNotJoin & "&nbsp;"
			Next
			For i=0 to 9-len(CrossRefNumber)-1
			 CrossRefNumber =CrossRefNumber & "&nbsp;"
			Next
			For i=0 to 8-len(DateOfBirth)-1
			 DateOfBirth = DateOfBirth & "&nbsp;"
			Next
			For i=0 to 80-len(EmailAddr)-1
			 EmailAddr = EmailAddr & "&nbsp;"
			Next
			
			Response.write record_id  & ApplicantIndicator & ssn & ApplicantFirstName & ApplicantMidName & ApplicantLastName &_
			ApplicantGeneration & HomePhone & Age & YrsSchool & MaritalStatus & DependantsNo & CompletedJoinNotJoin & CrossRefNumber & DateOfBirth & EmailAddr & "<br>"
		'------------------------------------------------
		'- 03B	III	 Dependent's Age.
		'------------------------------------------------
		Case "03B"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "03B-030", Mid(record_line,13,3) 		'Dependant's age [DependantsAge]
			
			DependantsAge = Trim(Mid(record_line,13,3))
			
			For i=0 to 3-len(DependantsAge)-1
			 DependantsAge = DependantsAge & "&nbsp;" 
			Next
			
			Response.write record_id & ssn & DependantsAge & "<br>"
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
			
			PresentFormer = Trim(Mid(record_line,13,2))	
			ResidenceStAddr = Trim(Mid(record_line,15,50))	
			ResidenceCity = Trim(Mid(record_line,65,35))	
			ResidenceState = Trim(Mid(record_line,100,2))	
			ResidenceZip = Trim(Mid(record_line,102,5))	
			ResidenceZipFour = Trim(Mid(record_line,107,4))	
			OwnRentLivingRentFree = Trim(Mid(record_line,111,1))	
			AddrNoYrs = Trim(Mid(record_line,112,2))	
			AddrNoMonth = Trim(Mid(record_line,114,2))	
			ApplicantAddrCounrtry = Trim(Mid(record_line,116,50))	
			
			For i=0 to 2-len(PresentFormer)-1
			 PresentFormer = PresentFormer & "&nbsp;" 
			Next
			For i=0 to 50-len(ResidenceStAddr)-1
			 ResidenceStAddr = ResidenceStAddr & "&nbsp;" 
			Next
			For i=0 to 35-len(ResidenceCity)-1
			 ResidenceCity = ResidenceCity & "&nbsp;" 
			Next
			For i=0 to 2-len(ResidenceState)-1
			 ResidenceState = ResidenceState & "&nbsp;" 
			Next
			For i=0 to 5-len(ResidenceZip)-1
			 ResidenceZip = ResidenceZip & "&nbsp;" 
			Next
			For i=0 to 4-len(ResidenceZipFour)-1
			 ResidenceZipFour = ResidenceZipFour & "&nbsp;" 
			Next
			For i=0 to 1-len(OwnRentLivingRentFree)-1
			 OwnRentLivingRentFree = OwnRentLivingRentFree & "&nbsp;" 
			Next
			For i=0 to 2-len(AddrNoYrs)-1
			 AddrNoYrs = AddrNoYrs & "&nbsp;" 
			Next
			For i=0 to 2-len(AddrNoMonth)-1
			 AddrNoMonth = AddrNoMonth & "&nbsp;" 
			Next
			For i=0 to 50-len(ApplicantAddrCounrtry)-1
			 ApplicantAddrCounrtry = ApplicantAddrCounrtry & "&nbsp;" 
			Next
			
			Response.write record_id & ssn & PresentFormer & ResidenceStAddr & ResidenceCity & ResidenceState & ResidenceZip & ResidenceZipFour &_
			OwnRentLivingRentFree & AddrNoYrs & AddrNoMonth & ApplicantAddrCounrtry & "<br>"
			
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
			
			For i=0 to 35-len(EmpName)-1
			 EmpName = EmpName & "&nbsp;" 
			Next
			For i=0 to 35-len(EmpStAddr)-1
			 EmpStAddr = EmpStAddr & "&nbsp;" 
			Next
			For i=0 to 35-len(EmpCity)-1
			 EmpCity = EmpCity & "&nbsp;"
			Next
			For i=0 to 2-len(EmpState)-1
			 EmpState = "&nbsp;" & EmpState
			Next
			For i=0 to 5-len(EmpZip)-1
			 EmpZip = "&nbsp;" & EmpZip
			Next
			For i=0 to 4-len(EmpZipFour)-1
			 EmpZipFour = "&nbsp;" & EmpZipFour
			Next
			For i=0 to 1-len(SelfEmployed)-1
			 SelfEmployed = "&nbsp;" & SelfEmployed
			Next
			For i=0 to 2-len(YrsOnThisJob)-1
			 YrsOnThisJob = "&nbsp;" & YrsOnThisJob
			Next
			For i=0 to 2-len(MonthsOnThisJob)-1
			 MonthsOnThisJob =  MonthsOnThisJob & "&nbsp;" 
			Next
			For i=0 to 2-len(YrsEmpInThisLineWork)-1
			 YrsEmpInThisLineWork = "&nbsp;" & YrsEmpInThisLineWork
			Next
			For i=0 to 25-len(PositionTitleTypeBiz)-1
			 PositionTitleTypeBiz = PositionTitleTypeBiz & "&nbsp;" 
			Next
			For i=0 to 10-len(BizPhone)-1
			 BizPhone = "&nbsp;" & BizPhone
			Next
			
			Response.write record_id & ssn & EmpName & EmpStAddr & EmpCity & EmpState & EmpZip & EmpZipFour & SelfEmployed &_
			YrsOnThisJob & MonthsOnThisJob & YrsEmpInThisLineWork & PositionTitleTypeBiz & BizPhone & "<br>"
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
			
			For i=0 to 35-len(SPrevEmpName)-1
			 SPrevEmpName = SPrevEmpName & "&nbsp;" 
			Next
			For i=0 to 35-len(SPrevEmpStAddr)-1
			 SPrevEmpStAddr = SPrevEmpStAddr & "&nbsp;" 
			Next
			For i=0 to 35-len(SPrevEmpCity)-1
			 SPrevEmpCity = SPrevEmpCity & "&nbsp;"
			Next
			For i=0 to 2-len(SPrevEmpState)-1
			 SPrevEmpState = "&nbsp;" & SPrevEmpState
			Next
			For i=0 to 5-len(SPrevEmpZip)-1
			 SPrevEmpZip = "&nbsp;" & SPrevEmpZip
			Next
			For i=0 to 4-len(SPrevEmpZipFour)-1
			 SPrevEmpZipFour = "&nbsp;" & SPrevEmpZipFour
			Next
			For i=0 to 1-len(SPrevSelfEmployed)-1
			 SPrevSelfEmployed = "&nbsp;" & SPrevSelfEmployed
			Next
			For i=0 to 1-len(SPrevCurrentEmpFlag)-1
			 SPrevCurrentEmpFlag = "&nbsp;" & SPrevCurrentEmpFlag
			Next
			For i=0 to 8-len(SPrevFromDate)-1
			 SPrevFromDate = "&nbsp;" & SPrevFromDate
			Next
			For i=0 to 8-len(SPrevToDate)-1
			 SPrevToDate = "&nbsp;" & SPrevToDate
			Next
			For i=0 to 15-len(SPrevMonthlyIncome)-1
			 SPrevMonthlyIncome = "&nbsp;" & SPrevMonthlyIncome
			Next
			For i=0 to 25-len(SPrevPositionTitleTypeBiz)-1
			 SPrevPositionTitleTypeBiz = SPrevPositionTitleTypeBiz & "&nbsp;" 
			Next
			For i=0 to 10-len(SPrevBizPhone)-1
			 SPrevBizPhone= "&nbsp;" & SPrevBizPhone
			Next
			
			Response.write record_id & ssn & SPrevEmpName & SPrevEmpStAddr & SPrevEmpCity & SPrevEmpState & SPrevEmpZip & SPrevEmpZipFour & SPrevSelfEmployed &_
			SPrevCurrentEmpFlag & SPrevFromDate & SPrevToDate & SPrevMonthlyIncome & SPrevPositionTitleTypeBiz & SPrevBizPhone & "<br>"
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

			For i=0 to 1-len(HousingExpensePresentIndicator)-1
			 HousingExpensePresentIndicator = "&nbsp;" & HousingExpensePresentIndicator
			Next
			For i=0 to 2-len(HousingPaymentTypeCode)-1
			 HousingPaymentTypeCode = "&nbsp;" & HousingPaymentTypeCode
			Next
			For i=0 to 15-len(HousingPaymentAmt)-1
			 HousingPaymentAmt = "&nbsp;" & HousingPaymentAmt
			Next
			
			Response.write record_id & ssn & HousingExpensePresentIndicator & HousingPaymentTypeCode & HousingPaymentAmt & "<br>"
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
			
			For i=0 to 2-len(TypeOfIncomeCode)-1
			 TypeOfIncomeCode = "&nbsp;" & TypeOfIncomeCode
			Next
			For i=0 to 15-len(IncomeAmt)-1
			 IncomeAmt = "&nbsp;" & IncomeAmt
			Next
			
			Response.write record_id & ssn & TypeOfIncomeCode & IncomeAmt & "<br>"
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
			
			For i=0 to 35-len(CashDepositPurcHeldBy)-1
			 CashDepositPurcHeldBy = "&nbsp;" & CashDepositPurcHeldBy
			Next
			For i=0 to 15-len(CashOrMarketValue)-1
			 CashOrMarketValue = "&nbsp;" & CashOrMarketValue
			Next
			
			Response.write record_id & ssn & CashDepositPurcHeldBy & CashOrMarketValue & "<br>"
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
			
			For i=0 to 30-len(LifeInsurAcctNo)-1
			 LifeInsurAcctNo = "&nbsp;" & LifeInsurAcctNo
			Next
			For i=0 to 15-len(LifeInsurCashMarketVal)-1
			 LifeInsurCashMarketVal = "&nbsp;" & LifeInsurCashMarketVal
			Next
			For i=0 to 15-len(LifeInsurFaceAmt)-1
			 LifeInsurFaceAmt = "&nbsp;" & LifeInsurFaceAmt
			Next
			
			Response.write record_id & ssn & LifeInsurAcctNo  & LifeInsurCashMarketVal & LifeInsurFaceAmt & "<br>"
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
			
			For i=0 to 3-len(AccountAssetType)-1
			 AccountAssetType =  AccountAssetType & "&nbsp;"
			Next
			For i=0 to 35-len(DepositoryStockBondName)-1
			 DepositoryStockBondName = DepositoryStockBondName & "&nbsp;"
			Next
			For i=0 to 35-len(DepositoryStAddr)-1
			 DepositoryStAddr = "&nbsp;" & DepositoryStAddr
			Next
			For i=0 to 35-len(DepositoryCity)-1
			 DepositoryCity = "&nbsp;" & DepositoryCity
			Next
			For i=0 to 2-len(DepositoryState)-1
			 DepositoryState = "&nbsp;" & DepositoryState
			Next
			For i=0 to 5-len(DepositoryZip)-1
			 DepositoryZip = "&nbsp;" & DepositoryZip
			Next
			For i=0 to 4-len(DepositoryZipFour)-1
			 DepositoryZipFour = "&nbsp;" & DepositoryZipFour
			Next
			For i=0 to 30-len(AssetAcctNo)-1
			 AssetAcctNo = AssetAcctNo & "&nbsp;"
			Next
			For i=0 to 15-len(AssetCashMarketVal)-1
			 AssetCashMarketVal = "&nbsp;" & AssetCashMarketVal
			Next
			For i=0 to 7-len(NumberOfStockBondShares)-1
			 NumberOfStockBondShares = NumberOfStockBondShares & "&nbsp;"
			Next
			For i=0 to 80-len(AssetDesc)-1
			 AssetDesc = "&nbsp;" & AssetDesc
			Next
			
			Response.write record_id & ssn & AccountAssetType & DepositoryStockBondName & DepositoryStAddr & DepositoryCity & DepositoryState &_
			DepositoryZip & DepositoryZipFour & AssetAcctNo & AssetCashMarketVal & NumberOfStockBondShares & AssetDesc & "<br>"
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
			
			For i=0 to 30-len(AutomobileMakeModel)-1
			 AutomobileMakeModel = AutomobileMakeModel & "&nbsp;"
			Next
			For i=0 to 4-len(AutomobileYear)-1
			 AutomobileYear = "&nbsp;" & AutomobileYear
			Next
			For i=0 to 15-len(AutomobileCashMarketVal)-1
			 AutomobileCashMarketVal = "&nbsp;" & AutomobileCashMarketVal
			Next
			
			Response.write record_id & ssn & AutomobileMakeModel & AutomobileYear & AutomobileCashMarketVal & "<br>"
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
			
			For i=0 to 3-len(ExpenseTypeCode)-1
			 ExpenseTypeCode = ExpenseTypeCode & "&nbsp;"
			Next
			For i=0 to 15-len(MonthlyPaymentAmt)-1
			 MonthlyPaymentAmt = "&nbsp;" & MonthlyPaymentAmt
			Next
			For i=0 to 3-len(MonthsLeftToPay)-1
			 MonthsLeftToPay = "&nbsp;" & MonthsLeftToPay
			Next
			For i=0 to 60-len(AlimonyCSSperateOwedTo)-1
			 AlimonyCSSperateOwedTo = "&nbsp;" & AlimonyCSSperateOwedTo
			Next
			
			Response.write record_id & ssn & ExpenseTypeCode & MonthlyPaymentAmt & MonthsLeftToPay & AlimonyCSSperateOwedTo & "<br>"
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
			objItem(strKey).Add "06G-150", Mid(record_line,172,25)	'Net Rental Income [REONetRentalIncome]
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
			REONetRentalIncome = Trim( Mid(record_line,172,25))	
			REOCurResidenceIndicator = Trim(Mid(record_line,187,1))	
			REOSubjectPropIndicator = Trim(Mid(record_line,188,1))	
			REOAssetID = Trim(Mid(record_line,189,2))	
			REOReservedFutureUse = Trim(Mid(record_line,191,15))	
			
			For i=0 to 35-len(REOPropStAddr)-1
			 REOPropStAddr = REOPropStAddr & "&nbsp;" 
			Next
			For i=0 to 35-len(REOPropCity)-1
			 REOPropCity = REOPropCity & "&nbsp;"
			Next
			For i=0 to 2-len(REOPropState)-1
			 REOPropState = "&nbsp;" & REOPropState
			Next
			For i=0 to 5-len(REOPropZip)-1
			 REOPropZip = "&nbsp;" & REOPropZip
			Next
			For i=0 to 4-len(REOPropZipFour)-1
			 REOPropZipFour = "&nbsp;" & REOPropZipFour
			Next
			For i=0 to 1-len(REOPropDisposition)-1
			 REOPropDisposition = "&nbsp;" & REOPropDisposition
			Next
			For i=0 to 2-len(REOTypeOfProp)-1
			 REOTypeOfProp = "&nbsp;" & REOTypeOfProp
			Next
			For i=0 to 15-len(REOPresentMarketValue)-1
			 REOPresentMarketValue = REOPresentMarketValue & "&nbsp;"
			Next
			For i=0 to 15-len(REOAmtMortgageLiens)-1
			 REOAmtMortgageLiens = REOAmtMortgageLiens & "&nbsp;"
			Next
			For i=0 to 15-len(REOGrossRentalIncome)-1
			 REOGrossRentalIncome = REOGrossRentalIncome & "&nbsp;"
			Next
			For i=0 to 15-len(REOMortgagePayment)-1
			 REOMortgagePayment = REOMortgagePayment & "&nbsp;"
			Next
			For i=0 to 15-len(InsurMaintenanceTaxMisc)-1
			 InsurMaintenanceTaxMisc = InsurMaintenanceTaxMisc & "&nbsp;"
			Next
			For i=0 to 25-len(REONetRentalIncome)-1
			 REONetRentalIncome = REONetRentalIncome & "&nbsp;"
			Next
			For i=0 to 1-len(REOCurResidenceIndicator)-1
			 REOCurResidenceIndicator = "&nbsp;" & REOCurResidenceIndicator
			Next
			For i=0 to 1-len(REOSubjectPropIndicator)-1
			 REOSubjectPropIndicator =  "&nbsp;" & REOSubjectPropIndicator
			Next
			For i=0 to 2-len(REOAssetID)-1
			 REOAssetID = "&nbsp;" & REOAssetID
			Next
			For i=0 to 15-len(REOReservedFutureUse)-1
			 REOReservedFutureUse = "&nbsp;" & REOReservedFutureUse
			Next
			
			Response.write record_id & ssn & REOPropStAddr & REOPropCity & REOPropState & REOPropZip & REOPropZipFour & REOPropDisposition &_
			REOTypeOfProp & REOPresentMarketValue & REOAmtMortgageLiens & REOGrossRentalIncome & REOMortgagePayment & InsurMaintenanceTaxMisc &_
			REONetRentalIncome & REOCurResidenceIndicator & REOSubjectPropIndicator & REOAssetID & REOReservedFutureUse & "<br>"
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
			
			For i=0 to 35-len(AliasFirstName)-1
			 AliasFirstName = "&nbsp;" & AliasFirstName
			Next
			For i=0 to 35-len(AliasMidName)-1
			 AliasMidName = "&nbsp;" & AliasMidName
			Next
			For i=0 to 35-len(AliasLastName)-1
			 AliasLastName = "&nbsp;" & AliasLastName
			Next
			For i=0 to 15-len(ReservedFutureUse6_1)-1
			 ReservedFutureUse6_1 = "&nbsp;" & ReservedFutureUse6_1
			Next
			For i=0 to 15-len(ReservedFutureUse6_2)-1
			 ReservedFutureUse6_2 = "&nbsp;" & ReservedFutureUse6_2
			Next
			
			Response.write record_id & ssn & AliasFirstName & AliasMidName & AliasLastName &ReservedFutureUse6_1 &  ReservedFutureUse6_2 & "<br>"
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
			
			For i=0 to 2-len(LiabilityType)-1
			 LiabilityType = LiabilityType & "&nbsp;" 
			Next
			For i=0 to 35-len(CreditorName)-1
			 CreditorName = CreditorName & "&nbsp;" 
			Next
			For i=0 to 35-len(CreditorStAddr)-1
			 CreditorStAddr = CreditorStAddr & "&nbsp;" 
			Next
			For i=0 to 35-len(CreditorCity)-1
			 CreditorCity = CreditorCity & "&nbsp;" 
			Next
			For i=0 to 2-len(CreditorState)-1
			 CreditorState = "&nbsp;" & CreditorState
			Next
			For i=0 to 5-len(CreditorZip)-1
			 CreditorZip = "&nbsp;" & CreditorZip
			Next
			For i=0 to 4-len(CreditorZipFour)-1
			 CreditorZipFour = "&nbsp;" & CreditorZipFour
			Next
			For i=0 to 30-len(LiabilityAcctNo)-1
			 LiabilityAcctNo = LiabilityAcctNo & "&nbsp;" 
			Next
			For i=0 to 15-len(LiabilityMonPaymentAmt)-1
			 LiabilityMonPaymentAmt = "&nbsp;" & LiabilityMonPaymentAmt
			Next
			For i=0 to 3-len(LiabilityMonLeftToPay)-1
			 LiabilityMonLeftToPay = "&nbsp;" & LiabilityMonLeftToPay
			Next
			For i=0 to 15-len(UnpaidBalance)-1
			 UnpaidBalance = "&nbsp;" &UnpaidBalance
			Next
			For i=0 to 1-len(LiabilityPaidClosing)-1
			 LiabilityPaidClosing = "&nbsp;" & LiabilityPaidClosing 
			Next
			For i=0 to 2-len(REOAssetID)-1
			 REOAssetID = "&nbsp;" & REOAssetID 
			Next
			For i=0 to 1-len(ResubordinatedIndicator)-1
			 ResubordinatedIndicator = "&nbsp;" &ResubordinatedIndicator 
			Next
			For i=0 to 1-len(OmittedIndicator)-1
			 OmittedIndicator = OmittedIndicator & "&nbsp;" 
			Next
			For i=0 to 1-len(SubjectPropIndicator)-1
			 SubjectPropIndicator  = "&nbsp;" & SubjectPropIndicator 
			Next
			For i=0 to 1-len(RentalPropIndicator)-1
			 RentalPropIndicator = "&nbsp;" &RentalPropIndicator
			Next
			
			Response.write record_id & ssn & LiabilityType & CreditorName & CreditorStAddr & CreditorCity & CreditorState & CreditorZip & CreditorZipFour &_
			LiabilityAcctNo & LiabilityMonPaymentAmt & LiabilityMonLeftToPay & UnpaidBalance & LiabilityPaidClosing & REOAssetID & ResubordinatedIndicator &_
			OmittedIndicator & SubjectPropIndicator & RentalPropIndicator & "<br>"
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
			
			For i=0 to 1-len(DeclarationsA)-1
			 DeclarationsA = "&nbsp;" & DeclarationsA
			Next
			For i=0 to 1-len(DeclarationsB)-1
			 DeclarationsB = "&nbsp;" & DeclarationsB
			Next
			For i=0 to 1-len(DeclarationsC)-1
			 DeclarationsC = "&nbsp;" & DeclarationsC
			Next
			For i=0 to 1-len(DeclarationsD)-1
			 DeclarationsD = "&nbsp;" & DeclarationsD
			Next
			For i=0 to 1-len(DeclarationsE)-1
			 DeclarationsE = "&nbsp;" & DeclarationsE
			Next
			For i=0 to 1-len(DeclarationsF)-1
			 DeclarationsF = "&nbsp;" & DeclarationsF
			Next
			For i=0 to 1-len(DeclarationsG)-1
			 DeclarationsG = "&nbsp;" & DeclarationsG
			Next
			For i=0 to 1-len(DeclarationsH)-1
			 DeclarationsH = "&nbsp;" & DeclarationsH
			Next
			For i=0 to 1-len(DeclarationsI)-1
			 DeclarationsI = "&nbsp;" & DeclarationsI
			Next
			For i=0 to 2-len(DeclarationsJ)-1
			 DeclarationsJ = "&nbsp;" & DeclarationsJ
			Next
			For i=0 to 1-len(DeclarationsL)-1
			 DeclarationsL = "&nbsp;" & DeclarationsL
			Next
			For i=0 to 1-len(DeclarationsM)-1
			 DeclarationsM = "&nbsp;" & DeclarationsM
			Next
			For i=0 to 1-len(DeclarationsM1)-1
			 DeclarationsM1 = "&nbsp;" & DeclarationsM1
			Next
			For i=0 to 2-len(DeclarationsM2)-1
			 DeclarationsM2 = "&nbsp;" & DeclarationsM2
			Next
			
			Response.write record_id & ssn & DeclarationsA & DeclarationsB & DeclarationsC & DeclarationsD & DeclarationsE &_
			DeclarationsF & DeclarationsG & DeclarationsH & DeclarationsI & DeclarationsJ &DeclarationsL & DeclarationsM &_
			DeclarationsM1 & DeclarationsM2 & "<br>"
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
			
			For i=0 to 2-len(DeclarationTypeCode)-1
			 DeclarationTypeCode = DeclarationTypeCode & "&nbsp;"
			Next
			For i=0 to 255-len(DeclarationExplanation)-1
			 DeclarationExplanation = DeclarationExplanation & "&nbsp;"
			Next
			
			Response.write record_id & ssn & DeclarationTypeCode & "&nbsp;" & DeclarationExplanation & "<br>"
		'------------------------------------------------
		'- 09A	IX	 Acknowledgment and Agreement
		'------------------------------------------------
		Case "09A"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "09A-030", Mid(record_line,13,8) 'Signature Date [SignatureDate]
			
			SignatureDate = Trim(Mid(record_line,13,8))
			
			For i=0 to 8-len(SignatureDate)-1
			 SignatureDate = "&nbsp;" & SignatureDate
			Next
			
			Response.write record_id & ssn & SignatureDate & "<br>"
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
			
			For i=0 to 1-len(IDoNotFurnishMyInfo)-1
			 IDoNotFurnishMyInfo = "&nbsp;" & IDoNotFurnishMyInfo
			Next
			For i=0 to 1-len(Ethnicity)-1
			 Ethnicity = "&nbsp;" & Ethnicity
			Next
			For i=0 to 30-len(Filler)-1
			 Filler = "&nbsp;" & Filler
			Next
			For i=0 to 1-len(Sex)-1
			 Sex = "&nbsp;" & Sex
			Next
			
			Response.write record_id & ssn & IDoNotFurnishMyInfo & Ethnicity & Filler & Sex &"<br>"
		'------------------------------------------------
		'- 10R	X	 InFormation For Government Monitoring Purposes
		'------------------------------------------------
		Case "10R"
			ssn = Mid(record_line,4,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "10R-030", Mid(record_line,13,2)	'Race type [RaceType]
			
			RaceType = Trim(Mid(record_line,13,2))
			
			For i=0 to 1-len(RaceType)-1
			 RaceType = "&nbsp;" & RaceType
			Next
			
			Response.write record_id & ssn & RaceType &"<br>"
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