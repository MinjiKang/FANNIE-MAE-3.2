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
Set objEDI = objFS.OpenTextFile(Server.MapPath("C0101904_1.txt"),1,true)


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
		'------------------------------------------------
		'- 02C	II	 Title Holder
		'------------------------------------------------
		'------------------------------------------------
		'- 02D	II	 Construction or Refinance Data
		'------------------------------------------------
		'------------------------------------------------
		'- 02E	II	 Down Payment
		'------------------------------------------------

		'------------------------------------------------
		'- 10B	X	 Loan Originator Information
		'------------------------------------------------
		'------------------------------------------------
		'- 10R	X	 Information for Government Monitoring Purposes
		'------------------------------------------------

		'--------------------------------------------------------------------------------------------------
		'- [Applicant]
		'--------------------------------------------------------------------------------------------------
		'------------------------------------------------
		'- 03A	III	 Applicant(s) Data
		'------------------------------------------------
		Case "03A"
			ssn = Mid(record_line,6,9)
			Set objApplicant = GetApplicant(objApplication("Applicant(s)"),ssn)
			objApplicant.Add "03A-040", Mid(record_line,15,35) 'Applicant First Name
			objApplicant.Add "03A-050", Mid(record_line,50,35) 'Applicant Middle Name
			objApplicant.Add "03A-060", Mid(record_line,85,35) 'Applicant Last Name
			objApplicant.Add "03A-070", Mid(record_line,120,4) 'Applicant Generation
		'------------------------------------------------
		'- 03B	III	 Dependent’s Age.
		'------------------------------------------------
		'------------------------------------------------
		'- 03C	III	 Applicant(s) Address
		'------------------------------------------------
		'------------------------------------------------
		'- 04A	IV	 Primary Current Employer(s)
		'------------------------------------------------
		'------------------------------------------------
		'- 04B	IV	 Secondary/Previous Employer(s)
		'------------------------------------------------
		'------------------------------------------------
		'- 05H	V	 Present/Proposed Housing Expense 
		'------------------------------------------------
		'------------------------------------------------
		'- 05I	V	 Income
		'------------------------------------------------
		'------------------------------------------------
		'- 06A	VI	 For all asset types, enter data in the 06C assets segment.
		'------------------------------------------------
		'------------------------------------------------
		'- 06B	VI	 Life Insurance
		'------------------------------------------------
		'------------------------------------------------
		'- 06C	VI	 Assets
		'------------------------------------------------
		'------------------------------------------------
		'- 06D	VI	 Automobile(s)
		'------------------------------------------------
		'------------------------------------------------
		'- 06F	VI	 Alimony, Child Support/ Separate Maintenance and/or Job Related Expense(s)
		'------------------------------------------------
		'------------------------------------------------
		'- 06G	VI	 Real Estate Owned
		'------------------------------------------------
		'------------------------------------------------
		'- 06H	VI	 Alias
		'------------------------------------------------
		'------------------------------------------------
		'- 06L	VI	 Liabilities
		'------------------------------------------------
		'------------------------------------------------
		'- 06S	VI	 Undrawn HELOC and IPCs
		'------------------------------------------------
		'------------------------------------------------
		'- 07A	VII	 Details of Transaction
		'------------------------------------------------
		'------------------------------------------------
		'- 07B	VII	 Other Credits
		'------------------------------------------------
		'------------------------------------------------
		'- 08A	VIII Declarations
		'------------------------------------------------
		'------------------------------------------------
		'- 08B	VIII Declaration Explanations
		'------------------------------------------------

		'------------------------------------------------
		'- 09A	IX	 Acknowledgment and Agreement
		'------------------------------------------------
		'------------------------------------------------
		'- 10A	X	 Information for Government Monitoring Purposes.
		'------------------------------------------------


		
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

	'------------------------------------------------
	'- Application
	'------------------------------------------------
	For Each fld_application In objApplication.Keys
		Select Case fld_application
			Case "Applicant(s)"
				For Each ssn In objApplication("Applicant(s)")
					Response.Write ssn & "<br>"
					Set objApplicant = objApplication("Applicant(s)")(ssn)
					

					'------------------------------------------------
					'- Applicant
					'------------------------------------------------
					For Each fld_applicant In objApplicant.Keys
						Select Case fld_applicant
							Case "Liability(s)"
							Case "Income(s)"
							Case Else
								Response.Write fld_applicant & ": <strong>" & objApplicant(fld_applicant) & "</strong><br>"

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

	objFieldName.Add "02A-020", "Property Street Address"
	objFieldName.Add "02A-030", "Property City"
	objFieldName.Add "02A-040", "Property State"
	objFieldName.Add "02A-050", "Property Zip Code"
	objFieldName.Add "02A-060", "Property Zip Code Plus Four"
	objFieldName.Add "02A-070", "No. of Units"
	objFieldName.Add "02A-080", "Legal Description of Subject Property-Code"
	objFieldName.Add "02A-090", "Legal Description of Subject Property"
	objFieldName.Add "02A-100", "Year Built"
	Set SetFieldIdName = objFieldName
End Function
%>