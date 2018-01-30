<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
'##################################################################################################
'# Declaration of constants
'##################################################################################################

'##################################################################################################
'# Declaration of variables
'##################################################################################################
Dim objFS
Dim objEDI
Dim objFieldIdName

Dim record_line
Dim record_id
Dim field_id

Dim objApplication
Dim objApplicant
Dim objItem
Dim strKey

' Applicant(s)
Dim ssn
'##################################################################################################
'# Initializing Page
'##################################################################################################

'##################################################################################################
'# Loading Page
'##################################################################################################
Set objFieldIdName = SetFieldIdName()

Set objFS	= Server.CreateObject("Scripting.FileSystemObject")
'================================================
'= Application
'================================================
Set objApplication = Server.CreateObject("Scripting.Dictionary")
objApplication.Add "Applicant(s)", Server.CreateObject("Scripting.Dictionary")

'================================================
'= Reading EDI File
'================================================
Set objEDI = objFS.OpenTextFile(Server.MapPath("test.txt"),1,true)

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
			objApplication.Add "00A-020", Mid(record_line,4,1)
			objApplication.Add "00A-030", Mid(record_line,5,1)
		'------------------------------------------------
		'- [02A] II	Property Information
		'------------------------------------------------
		Case "02A" 
			objApplication.Add "02A-020", Mid(record_line,4,50) 	'Property Street Address
			objApplication.Add "02A-030", Mid(record_line,54,35)	'Property City
			objApplication.Add "02A-040", Mid(record_line,89,2) 	'Property State
			objApplication.Add "02A-050", Mid(record_line,91,5) 	'Property Zip Code
			objApplication.Add "02A-060", Mid(record_line,96,4)		'Property Zip Code Plus Four
			objApplication.Add "02A-070", Mid(record_line,100,3)	'No. of Units
			objApplication.Add "02A-080", Mid(record_line,103,2)	'Legal Description of Subject Property – Code
			objApplication.Add "02A-090", Mid(record_line,105,80)	'Legal Description of Subject Property
			objApplication.Add "02A-100", Mid(record_line,185,4)	'Year Built
		'------------------------------------------------
		'- PAI Property Address Information
		'------------------------------------------------
		Case "PAI"
		'------------------------------------------------
		'- 02B	II	Purpose of Loan
		'------------------------------------------------
		Case "02B" 
			objApplication.Add "02B-030", Mid(record_line,6,2)	'Purpose of Loan
			objApplication.Add "02B-040", Mid(record_line,8,80)	'Purpose of Loan (Other)
			objApplication.Add "02B-050", Mid(record_line,88,1)	'Property will be
			objApplication.Add "02B-060", Mid(record_line,89,60)'Manner in which Title will be held
			objApplication.Add "02B-070", Mid(record_line,149,1)'Estate will be held in
			objApplication.Add "02B-080", Mid(record_line,150,8)'(Estate will be held in) Leasehold expiration date
		'------------------------------------------------
		'- 02C	II	 Title Holder
		'------------------------------------------------

			objApplication.Add "02C-020", Mid(record_line,4,60) 'Titleholder Name
		'------------------------------------------------
		'- 02D	II	 Construction or Refinance Data
		'------------------------------------------------
		Case "02D" 
			objApplication.Add "02D-020", Mid(record_line,4,4)	'Year Lot Acquired (Construction) or Year Acquired (Refinance)
			objApplication.Add "02D-030", Mid(record_line,8,15)	'Original Cost (Construction or Refinance)
			objApplication.Add "02D-040", Mid(record_line,23,15)'Amount Existing Liens (Construction or Refinance)
			objApplication.Add "02D-050", Mid(record_line,38,15)'(a) Present Value of Lot
			objApplication.Add "02D-060", Mid(record_line,53,15)'(b) Cost of Improvements
			objApplication.Add "02D-070", Mid(record_line,68,2) 'Purpose of Refinance
			objApplication.Add "02D-080", Mid(record_line,70,80)'Describe Improvements
			objApplication.Add "02D-090", Mid(record_line,150,1)'(Describe Improvements) mad/tobe made
			objApplication.Add "02D-100", Mid(record_line,151,15)'(Describe Improvements) Cost
		'------------------------------------------------
		'- 02E	II	 Down Payment
		'------------------------------------------------
		Case "02E" 
			objApplication.Add "02E-020", Mid(record_line,4,2) 	'Down Payment Type Code
			objApplication.Add "02E-030", Mid(record_line,6,15) 'Down Payment Amount
			objApplication.Add "02E-040", Mid(record_line,21,80)'Down Payment Explanation
		'------------------------------------------------
		'- 10B	X	 Loan Originator Information
		'------------------------------------------------
		Case "10B" 
			objApplication.Add "10B-020", Mid(record_line,4,1)		'This application was taken by
			objApplication.Add "10B-030", Mid(record_line,5,60)		'Loan Originator's Name
			objApplication.Add "10B-040", Mid(record_line,65,8)		'Interview Date
			objApplication.Add "10B-050", Mid(record_line,73,10)	'Loan Originator's Phone Number
			objApplication.Add "10B-060", Mid(record_line,83,35)	'Loan Origination Company's Name
			objApplication.Add "10B-070", Mid(record_line,118,35)	'Loan Origination Company's Street Address
			objApplication.Add "10B-080", Mid(record_line,153,35)	'Loan Origination Company's Street Address 2
			objApplication.Add "10B-090", Mid(record_line,188,35)	'Loan Origination Company's City 
			objApplication.Add "10B-100", Mid(record_line,223,2)	'Loan Origination Company's State Code
			objApplication.Add "10B-110", Mid(record_line,225,5)	'Loan Origination Company's Zip Code
			objApplication.Add "10B-120", Mid(record_line,230,4)	'Loan Origination Company's Zip Code Plus Four
		'------------------------------------------------
		'- 10R	X	 Information for Government Monitoring Purposes
		'------------------------------------------------
		'Case "10R" 
			objApplication.Add "10R-030", Mid(record_line,13,2)		'Race type 
		'--------------------------------------------------------------------------------------------------
		'- [Applicant]
		'--------------------------------------------------------------------------------------------------
		'------------------------------------------------
		'- 03A	III	 Applicant(s) Data
		'------------------------------------------------
		Case "03A"
			ssn = Mid(record_line,6,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "03A-020", Mid(record_line,4,2) 		'Applicant / Co-Applicant Indicator
			objApplicant.Add "03A-040", Mid(record_line,15,35) 		'Applicant First Name
			objApplicant.Add "03A-050", Mid(record_line,50,35) 		'Applicant Middle Name
			objApplicant.Add "03A-060", Mid(record_line,85,35) 		'Applicant Last Name
			objApplicant.Add "03A-070", Mid(record_line,120,4) 		'Applicant Generation
			objApplicant.Add "03A-080", Mid(record_line,124,10)		'Home Phone
			objApplicant.Add "03A-090", Mid(record_line,134,3)		'Age
			objApplicant.Add "03A-100", Mid(record_line,137,2)		'Yrs. School
			objApplicant.Add "03A-110", Mid(record_line,139,1)		'Marital Status
			objApplicant.Add "03A-120", Mid(record_line,140,2)		'Dependents (no.)
			objApplicant.Add "03A-130", Mid(record_line,142,1)		'Completed Jointly/Not Jointly
			objApplicant.Add "03A-140", Mid(record_line,143,9)		'Cross-Reference Number
			objApplicant.Add "03A-150", Mid(record_line,152,8)		'Date of Birth
			objApplicant.Add "03A-160", Mid(record_line,160,80)		'Email Address
		'------------------------------------------------
		'- 05H	V	 Present/Proposed Housing Expense 
		'------------------------------------------------
		Case "05H"
			ssn = Mid(record_line,4,9)
			Set objItem = objApplication("Applicant(s)")(ssn)("Present/Proposed Housing Expences")
			strKey = Mid(record_line,13,1) & "-" & Mid(record_line,14,2)
			objItem.Add strKey, Server.CreateObject("Scripting.Dictionary")
			
			objItem(strKey).Add "05H-030", Mid(record_line,13,1)	'Present/Proposed Indicator
			objItem(strKey).Add "05H-040", Mid(record_line,14,2)	'Housing Payment Type Code
			objItem(strKey).Add "05H-050", Mid(record_line,16,15)	'Housing Payment Amount (Monthly Housing Exp.)
		
	End Select
	
Loop
'------------------------------------------------
'- Printing Application
'------------------------------------------------
Call PrintApplication(objApplication)
 
'##################################################################################################
'# Unloading Page
'##################################################################################################
Set objFieldIdName = Nothing
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

	Dim objItem
	Dim objFields
	Dim str_key
	Dim fld_item
	'------------------------------------------------
	'- Application
	'------------------------------------------------
	For Each fld_application In objApplication.Keys
		Select Case fld_application
			Case "Applicant(s)"
				For Each ssn In objApplication("Applicant(s)")
					Response.Write "<hr>"
					Response.Write "<h1>Applicant : " & ssn & "</h1>"
					Set objApplicant = objApplication("Applicant(s)")(ssn)
					'------------------------------------------------
					'- Applicant
					'------------------------------------------------
					For Each fld_applicant In objApplicant.Keys
						Select Case fld_applicant
							Case "Present/Proposed Housing Expences"
								Response.Write "<strong>" & fld_applicant & "</storng><br>"
								Set objItem = objApplicant(fld_applicant)
								For Each str_key In objItem.Keys
									Response.Write "- <strong>" & str_key & "</storng><br>"
									Set objFields = objItem(str_key)
									For Each fld_item In objFields.Keys
										Response.Write fld_item & "(" & objFieldIdName(fld_item) & "): <strong>" & objFields(fld_item) & "</strong><br>"
										
									Next
								Next
								Response.Write "<p>"
							Case "Liability(s)"
							Case "Income(s)"
							Case Else
								Response.Write fld_applicant & "(" & objFieldIdName(fld_applicant) & "): <strong>" & objApplicant(fld_applicant) & "</strong><br>"
						End Select
					Next
				Next
				Response.Write ssn & "<hr>"
			Case Else
		Response.Write fld_application & "(" & objFieldIdName(fld_application) & "): <strong>" & objApplication(fld_application) & "</strong><br>"
		End Select
	Next
End Sub
'==================================================================================================
'= GetApplicant
'==================================================================================================
Function GetApplicant(ByRef obj_applicants, ByVal ssn)
	Dim obj_applicant
	If obj_applicants.Exists(ssn) = FALSE Then
		obj_applicants.Add ssn, Server.CreateObject("Scripting.Dictionary")
		'------------------------------------------------
		'-
		'------------------------------------------------
		obj_applicants(ssn).Add "Present/Proposed Housing Expences",Server.CreateObject("Scripting.Dictionary")
		obj_applicants(ssn).Add "Liability(s)",Server.CreateObject("Scripting.Dictionary")
		obj_applicants(ssn).Add "Income(s)",Server.CreateObject("Scripting.Dictionary")
	End If
	Set GetApplicant = obj_applicants(ssn)
End Function
'==================================================================================================
'= LeftCut
'==================================================================================================
function LeftCut(strString, intCut)
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
 end function	
'==================================================================================================
'= SetFiledIdName
'==================================================================================================			
Function SetFieldIdName()

	Set objFieldName = Server.CreateObject("Scripting.Dictionary")
		objFieldName.Add "00A-020", "The income or assets of a person other than the borrower (including the borrower's spouse) will be used as a basis for loan qualification."
		objFieldName.Add "00A-030", "The income or assets of the borrower's spouse will not be used as a basis for loan qualification, but his or her liabilities must be considered because the borrower resides in a community property state, the security property is located in a community property state, or the borrower is relying on other property located in a community property state as a basis for repayment of the loan."
		objFieldName.Add "01A-020", "Mortgage Applied For"
		objFieldName.Add "01A-030", "Mortgage Applied For (Other)"
		objFieldName.Add "01A-040", "Agency Case Number"
		objFieldName.Add "01A-050", "Case Number"
		objFieldName.Add "01A-060", "Loan Amount"
		objFieldName.Add "01A-070", "Interest Rate"
		objFieldName.Add "01A-080", "No. of Months"
		objFieldName.Add "01A-090", "Amortization Type"
		objFieldName.Add "01A-100", "Amortization Type Other Explanation"
		objFieldName.Add "01A-110", "ARM Textual Description"
		objFieldName.Add "02A-020", "Property Street Address"
		objFieldName.Add "02A-030", "Property City"
		objFieldName.Add "02A-040", "Property State"
		objFieldName.Add "02A-050 ", "Property Zip Code"
		objFieldName.Add "02A-060", "Property Zip Code Plus Four"
		objFieldName.Add "02A-070", "No. of Units"
		objFieldName.Add "02A-080", "Legal Description of Subject Property-Code"
		objFieldName.Add "02A-090", "Legal Description of Subject Property"
		objFieldName.Add "02A-100", "Year Built"
		objFieldName.Add "02B-030", "Purpose of Loan"
		objFieldName.Add "02B-040", "Purpose of Loan (Other)"
		objFieldName.Add "02B-050", "Property will be"
		objFieldName.Add "02B-060", "Manner in which Title will be held"
		objFieldName.Add "02B-070", "Estate will be held in"
		objFieldName.Add "02B-080", "(Estate will be held in) Leasehold expiration date"
		objFieldName.Add "02C-020", "Titleholder Name"
		objFieldName.Add "02D", ""
		objFieldName.Add "02D-020", "Year Lot Acquired (Construction) or Year Acquired (Refinance)"
		objFieldName.Add "02D-030", "Original Cost (Construction or Refinance)"
		objFieldName.Add "02D-040", "Amount Existing Liens (Construction or Refinance)"
		objFieldName.Add "02D-050", "(a) Present Value of Lot"
		objFieldName.Add "02D-060", "(b) Cost of Improvements"
		objFieldName.Add "02D-070", "Purpose of Refinance"
		objFieldName.Add "02D-080", "Describe Improvements"
		objFieldName.Add "02D-090", "(Describe Improvements) mad/tobe made"
		objFieldName.Add "02D-100", "(Describe Improvements) Cost"
		objFieldName.Add "02E-020", "Down Payment Type Code"
		objFieldName.Add "02E-030", "Down Payment Amount"
		objFieldName.Add "02E-040", "Down Payment Explanation"
		objFieldName.Add "03A-020", "Applicant / Co-Applicant Indicator"
		objFieldName.Add "03A-040", "Applicant First Name"
		objFieldName.Add "03A-050", "Applicant Middle Name"
		objFieldName.Add "03A-060", "Applicant Last Name"
		objFieldName.Add "03A-070", "Applicant Generation"
		objFieldName.Add "03A-080", "Home Phone"
		objFieldName.Add "03A-090", "Age"
		objFieldName.Add "03A-100", "Yrs. School"
		objFieldName.Add "03A-110", "Marital Status"
		objFieldName.Add "03A-120", "Dependents (no.)"
		objFieldName.Add "03A-130", "Completed Jointly/Not Jointly"
		objFieldName.Add "03A-140", "Cross-Reference Number"
		objFieldName.Add "03A-150", "Date of Birth"
		objFieldName.Add "03A-160", "Email Address"
		objFieldName.Add "03B-030", "Dependent's Age"
		objFieldName.Add "03C-030", "(Present/Former)"
		objFieldName.Add "03C-040", "Residence Street Address"
		objFieldName.Add "03C-050", "Residence City"
		objFieldName.Add "03C-060", "Residence State"
		objFieldName.Add "03C-070", "Residence Zip Code"
		objFieldName.Add "03C-080", "Residence Zip Code Plus Four"
		objFieldName.Add "03C-090", "Own/Rent/Living Rent Free"
		objFieldName.Add "03C-100", "No. Yrs."
		objFieldName.Add "03C-110", "No. Months"
		objFieldName.Add "03C-120", "Country"
		objFieldName.Add "04A-030", "Employer Name"
		objFieldName.Add "04A-040", "Employer Street Address"
		objFieldName.Add "04A-050", "Employer City"
		objFieldName.Add "04A-060", "Employer State"
		objFieldName.Add "04A-070", "Employer Zip Code"
		objFieldName.Add "04A-080", "Employer Zip Code Plus Four"
		objFieldName.Add "04A-090", "Self Employed"
		objFieldName.Add "04A-100", "Yrs. on this job"
		objFieldName.Add "04A-110", "Months on this job"
		objFieldName.Add "04A-120", "Yrs. employed in this line of work/profession"
		objFieldName.Add "04A-130", "Position / Title / Type of Business"
		objFieldName.Add "04A-140", "Business Phone"
		objFieldName.Add "04B-030", "Employer Name"
		objFieldName.Add "04B-040", "Employer Street Address"
		objFieldName.Add "04B-050", "Employer City"
		objFieldName.Add "04B-060", "Employer State"
		objFieldName.Add "04B-070", "Employer Zip Code"
		objFieldName.Add "04B-080", "Employer Zip Code Plus Four"
		objFieldName.Add "04B-090", "Self Employed"
		objFieldName.Add "04B-100", "Current Employment Flag"
		objFieldName.Add "04B-110", "From Date"
		objFieldName.Add "04B-120", "To Date"
		objFieldName.Add "04B-130", "Monthly Income"
		objFieldName.Add "04B-140", "Position / Title / Type of Business"
		objFieldName.Add "04B-150", "Business Phone"
		objFieldName.Add "05H-030", "Present/Proposed Indicator"
		objFieldName.Add "05H-040", "Housing Payment Type Code"
		objFieldName.Add "05H-050", "Housing Payment Amount (Monthly Housing Exp.)"
		objFieldName.Add "05I-030", "Type of Income Code"
		objFieldName.Add "05I-040", "Income Amount (Monthly Income)"
		objFieldName.Add "06A-030", "Cash deposit toward purchase held by"
		objFieldName.Add "06A-040", "Cash or Market Value"
		objFieldName.Add "06B-030", "Acct. no."
		objFieldName.Add "06B-040", "Life Insurance Cash or Market Value"
		objFieldName.Add "06B-050", "Life insurance Face Amount"
		objFieldName.Add "06C-030", "Account/Asset Type"
		objFieldName.Add "06C-040", "Depository/Stock/Bond Institution Name"
		objFieldName.Add "06C-050", "Depository Street Address"
		objFieldName.Add "06C-060", "Depository City"
		objFieldName.Add "06C-070", "Depository State"
		objFieldName.Add "06C-080", "Depository Zip Code"
		objFieldName.Add "06C-090", "Depository Zip Code Plus Four"
		objFieldName.Add "06C-100", "Acct. no."
		objFieldName.Add "06C-110", "Cash or Market Value"
		objFieldName.Add "06C-120", "Number of Stock/Bond Shares"
		objFieldName.Add "06C-130", "Asset Description"
		objFieldName.Add "06C-140", "Reserved for Future Use"
		objFieldName.Add "06C-150", "Reserved for Future Use"
		objFieldName.Add "06D-030", "Automobile Make/ Model"
		objFieldName.Add "06D-040", "Automobile Year"
		objFieldName.Add "06D-050", "Cash or Market Value"
		objFieldName.Add "06F-030", "Expense Type Code"
		objFieldName.Add "06F-040 ", "Monthly Payment Amount"
		objFieldName.Add "06F-050", "Months Left to Pay"
		objFieldName.Add "06F-060", "Alimony/ Child Support/ Separate Maintenance Owed To"
		objFieldName.Add "06G-030", "Property Street Address"
		objFieldName.Add "06G-040", "Property City"
		objFieldName.Add "06G-050", "Property State"
		objFieldName.Add "06G-060", "Property Zip Code"
		objFieldName.Add "06G-070", "Property Zip Code Plus Four"
		objFieldName.Add "06G-080", "Property Disposition"
		objFieldName.Add "06G-090", "Type of Property"
		objFieldName.Add "06G-100", "Present Market Value"
		objFieldName.Add "06G-110", "Amount of Mortgages & Liens"
		objFieldName.Add "06G-120", "Gross Rental Income"
		objFieldName.Add "06G-130", "Mortgage Payments"
		objFieldName.Add "06G-140", "Insurance, Maintenance Taxes & Misc."
		objFieldName.Add "06G-150", "Net Rental Income"
		objFieldName.Add "06G-160", "Current Residence Indicator"
		objFieldName.Add "06G-170", "Subject Property Indicator"
		objFieldName.Add "06G-180", "REO Asset ID"
		objFieldName.Add "06G-190", "Reserved for Future Use"
		objFieldName.Add "06H-030", "Alternate First Name"
		objFieldName.Add "06H-040", "Alternate Middle Name"
		objFieldName.Add "06H-050", "Alternate Last Name"
		objFieldName.Add "06H-060", "Reserved for Future Use"
		objFieldName.Add "06H-070", "Reserved for Future Use"
		objFieldName.Add "06L-030", "Liability Type"
		objFieldName.Add "06L-040", "Creditor Name"
		objFieldName.Add "06L-050", "Creditor Street Address"
		objFieldName.Add "06L-060", "Creditor City"
		objFieldName.Add "06L-070", "Creditor State"
		objFieldName.Add "06L-080", "Creditor Zip Code"
		objFieldName.Add "06L-090", "Creditor Zip Code Plus Four"
		objFieldName.Add "06L-100", "Acct. no."
		objFieldName.Add "06L-110", "Monthly Payment Amount"
		objFieldName.Add "06L-120", "Months Left to Pay"
		objFieldName.Add "06L-130", "Unpaid Balance"
		objFieldName.Add "06L-140", "Liability will be paid prior to closing"
		objFieldName.Add "06L-150", "REO Asset ID"
		objFieldName.Add "06L-160", "Resubordinated Indicator"
		objFieldName.Add "06L-170", "Omitted Indicator"
		objFieldName.Add "06L-180", "Subject Property Indicator"
		objFieldName.Add "06L-190", "Rental Property Indicator"
		objFieldName.Add "06S-030", "Summary Amount Type Code"
		objFieldName.Add "06S-040", "Amount"
		objFieldName.Add "07A-020", "a. Purchase price"
		objFieldName.Add "07A-030", "b. Alterations, improvements, repairs"
		objFieldName.Add "07A-040", "c. Land"
		objFieldName.Add "07A-050", "d. Refinance (Inc. debts to be paid off)"
		objFieldName.Add "07A-060", "e. Estimated prepaid items"
		objFieldName.Add "07A-070", "f. Estimated closing costs"
		objFieldName.Add "07A-080", "g. PMI MIP, Funding Fee"
		objFieldName.Add "07A-090", "h. Discount"
		objFieldName.Add "07A-100", "j. Subordinate financing"
		objFieldName.Add "07A-110", "k. Applicant's closing costs paid by Seller"
		objFieldName.Add "07A-120", "n. PMI, MIP, Funding Fee financed"
		objFieldName.Add "07B-020", "Other Credit Type Code"
		objFieldName.Add "07B-030", "Amount of Other Credit"
		objFieldName.Add "08A-030", "a. Are there any outstanding judgments against you?"
		objFieldName.Add "08A-040", "b. Have you been declared bankrupt within the past 7 years?"
		objFieldName.Add "08A-050", "c. Have you had property foreclosed upon or given title or deed in lieu thereof in the last 7 years?"
		objFieldName.Add "08A-060", "d. Are you a party to a lawsuit?"
		objFieldName.Add "08A-070", "e. Have you directly or indirectly been obligated on any loan"
		objFieldName.Add "08A-080", "f. Are you presently delinquent or in default on any Federal debt"
		objFieldName.Add "08A-090", "g. Are you obligated to pay alimony child support or separate maintenance?"
		objFieldName.Add "08A-100", "h. Is any part of the down payment borrowed?"
		objFieldName.Add "08A-110", "i. Are you a co-maker or"
		objFieldName.Add "08A-120", "j. Are you a U.S. citizen?k. Are you a permanent resident alien?"
		objFieldName.Add "08A-130", "l. Do you intend to occupy"
		objFieldName.Add "08A-140", "m. Have you had an ownership interest"
		objFieldName.Add "08A-150", "m. (1) What type of property"
		objFieldName.Add "08A-160", "m. (2) How did you hold title"
		objFieldName.Add "08B-030", "Declaration Type Code"
		objFieldName.Add "08B-040", "Declaration Explanation"
		objFieldName.Add "09A-030", "Signature Date"
		objFieldName.Add "10A-030", "I do not wish to furnish this information"
		objFieldName.Add "10A-040", "Ethnicity"
		objFieldName.Add "10A-050", "Filler"
		objFieldName.Add "10A-060", "Sex"
		objFieldName.Add "10B-020", "This application was taken by"
		objFieldName.Add "10B-030", "Loan Originator's Name"
		objFieldName.Add "10B-040", "Interview Date"
		objFieldName.Add "10B-050", "Loan Originator's Phone Number"
		objFieldName.Add "10B-060", "Loan Origination Company's Name"
		objFieldName.Add "10B-070", "Loan Origination Company's Street Address"
		objFieldName.Add "10B-080", "Loan Origination Company's Street Address 2"
		objFieldName.Add "10B-090", "Loan Origination Company's City "
		objFieldName.Add "10B-100", "Loan Origination Company's State Code"
		objFieldName.Add "10B-110", "Loan Origination Company's Zip Code"
		objFieldName.Add "10B-120", "Loan Origination Company's Zip Code Plus Four"
		objFieldName.Add "10R-030", "RACE"
	Set SetFieldIdName = objFieldName
End Function
%>