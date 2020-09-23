strPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") & "\"

Set TK = CreateObject("APToolkit.Object")



'maken van een pdf pagina en de afbeeldingen toevoegen 

r = TK.OpenOutputFile("_Stationery.pdf")
TK.SetInfo "stationery for e-papers ", "Documentation", "Kurt Koenig   http://kurtkoenig.homeunix.net", "Created with ActivePDF Toolkit COM Object and VBScript"
r = TK.PrintJPEG("Logo.jpg", 495, 710, 200, 64, True)
r = TK.PrintJPEG("BGLogo.jpg", 50, 50, 0, 0, False)

' hier stellen we de dikte van de lijn in.

TK.LineWidth 0.35

' coordinaten voor de bovenste lijn

TK.MoveTo 464, 754
TK.DrawTo 20, 754

' coordinaten voor de zijlijn

TK.MoveTo 20, 25
TK.DrawTo 20, 754

' coordinaten voor de onderste lijn

TK.MoveTo 555, 25
TK.DrawTo 20, 25
TK.SetTextColor 29, 80, 52, 0

' instellen van het te gebruiken lettertype en de puntgrote van de letters.

TK.SetFont "Helvetica", 10
TK.PrintText 30, 64, "Test Ltd                                                                                                                                                                      DeptX"

TK.SetFont "Helvetica", 7.5

' alle tekst van het briefpapier toevoegen

TK.PrintText 30, 767, "Recognised by the Ministry  & Accredited by SomeOrg" 
TK.PrintText 30, 736, "Wallstreet 164" 
TK.PrintText 30, 727, "B-2000 Antwerp" 
TK.PrintText 30, 718, "Belgium"
TK.PrintText 30, 708, "Tel +32 4 666 60"
TK.PrintText 30, 698, "Fax +32 4 666 61"
TK.PrintText 30, 689, "http://kurtkoenig.homeunix.net"
TK.PrintText 30, 679, "info@kurtkoenig.homeunix.net" 

TK.PrintText 30, 55, "Certifiërings- en controle-organisatie voor landbouw en voeding                                                                                               Afdeling voor biologische productie"
TK.PrintText 30, 45, "Organisme de certification et de côntrole du secteur agro-alimentaire                                                                                  Division  pour la production biologique"
TK.PrintText 30, 35, "Certification and inspection body for agriculture and food processing                                                                                               Division for organic production"

TK.CloseOutputFile
Set TK = Nothing
