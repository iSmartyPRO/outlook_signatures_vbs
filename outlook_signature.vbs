On Error Resume Next
logoImg = "\\path\to\yout\logo.png"
Set objSysInfo = CreateObject("ADSystemInfo")
Set WshShell = CreateObject("WScript.Shell")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
Set objSelection = objWord.Selection 


strName 		= objUser.description
strCompany 		= objUser.physicalDeliveryOfficeName
strTitle 		= objUser.Description
strCred 		= objUser.info
strCountry 		= objUser.co
strCity 		= objUser.l
strStreet 		= objUser.StreetAddress
strLocation 	= objUser.l
strJobtitle 	= objUser.title
strDepartment	= objUser.department
strPostCode		= objUser.PostalCode
strPhone 		= objUser.TelephoneNumber
strExtNum		= objUser.pager
strFax 			= objUser.facsimileTelephoneNumber
strMobile 		= objUser.Mobile
strFax 			= objUser.FacsimileTelephoneNumber
strEmail 		= objUser.mail
strWebsiteLink	= ObjDoc.Hyperlinks.Add(objSelection.Range, "http://www.example.com",,, "www.example.com")


Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
objSelection.Font.Name = "Arial"
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(144,140,140) 
objSelection.TypeParagraph()
objSelection.TypeText "� ���������,"
objSelection.TypeText Chr(11)
objSelection.TypeText strName & ", " & strDepartment & ", " & strJobtitle
objselection.TypeText Chr(11)
objselection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture logoImg
objSelection.TypeText Chr(11) & Chr(11)
objSelection.Font.Bold = true
objSelection.TypeText strCompany
objSelection.Font.Bold = false
objSelection.TypeText Chr(11)
objSelection.TypeText strCity & ", " & strStreet
objSelection.TypeText Chr(11)
'objSelection.TypeText "T: : +7 (495) 916 6868, F: +7 (495) 916 6868"
if (strExtNum) Then typeExtNum = " (���. " & strExtNum & ") "
if (strPhone) Then objSelection.TypeText "T: " & strPhone Else objSelection.TypeText "T: +7 (495) 123-45-67" & typeExtNum End If
if (strFax) Then objSelection.TypeText ", F: " & strFax Else objSelection.TypeText ", F: +7 (495) 123-45-67" End If
if (strMobile) Then objSelection.TypeText ", M: " & strMobile
objSelection.TypeText Chr(11)
Set objCell = objTable.Cell(11, 1) 
Set objCellRange = objCell.Range 
objCell.Select 
objselection.typeText strEmailTEXT 
Set objLink = objSelection.Hyperlinks.Add(objSelection.Range, "mailto: " & strEmail, , , strEmail) 
  objLink.Range.Font.Bold = false 
objselection.typeText ", "
objSelection.TypeText 
Set objLink = objSelection.Hyperlinks.Add(objSelection.Range, "http://www.example.com", , , "http://www.example.com") 
  objLink.Range.Font.Bold = false 
  objSelection.Font.Color = RGB (000,045,154) 

objSelection.Font.Color = RGB(144,140,140) 
objSelection.TypeText Chr(11) & Chr(11)
objSelection.TypeText "========================================================================" & Chr(11)
objSelection.TypeText "������������������" & Chr(11)
objSelection.TypeText "========================================================================" & Chr(11)
objSelection.Font.italic = true 
objSelection.TypeText "��������� ����������� ������ � ���������� � ���� �������� ����������, ������������ ������������ �����." & Chr(11)
objSelection.TypeText "��������� ���������� �� ����� ���� ������������, ����������� ��� ���������� ����, " & Chr(11)
objSelection.TypeText "���� �������� �� ���������� ����� �������� ����� �� ���� ������������� ��� ����������� ����� ����������. " & Chr(11) & Chr(11)
objSelection.TypeText "���� �� �������� ��������� ����������� ������ �� ������ ���� ��� �� ��� ����� ������������ ������ � ����������, " & Chr(11)
objSelection.TypeText "������������ � ��������� ����������� ������ � ����������� � ����, ����������, ���������� ���������" & Chr(11)
objSelection.TypeText "� ����������� ����������� � ������� ������ ����������� ������ � ���������� � ����." & Chr(11)
objSelection.Font.italic = false
Set objSelection = objDoc.Range()
objSignatureEntries.Add "Signature - New", objSelection
objSignatureObject.NewMessageSignature = "Signature - New"
objDoc.Saved = True
objWord.Quit
Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
objSelection.Font.Name = "Arial"
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(144,140,140) 
objSelection.TypeParagraph()
objSelection.TypeText "C ���������,"
objSelection.TypeText Chr(11)
objSelection.Font.Bold = true
objSelection.TypeText strName
objSelection.Font.Bold = false
objSelection.TypeText Chr(11)
objSelection.TypeText strJobtitle & ", " & strDepartment
Set objSelection = objDoc.Range()
objSignatureEntries.Add "Signature - Reply", objSelection
objSignatureObject.ReplyMessageSignature = "Signature - Reply"
objDoc.Saved = True
objWord.Quit