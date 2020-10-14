Attribute VB_Name = "ABI_PartGroups"

Public Function SynthArea(part As Variant) As String
Dim Area As String
On Error Resume Next

Select Case part
    Case 360122: Area = "Other Amidites"
    Case 360180: Area = "Other Amidites"
    Case 360317: Area = "Dye NHS Esters"
    Case 360318: Area = "Dye NHS Esters"
    Case 360336: Area = "Dye NHS Esters"
    Case 360337: Area = "Dye NHS Esters"
    Case 360340: Area = "Other Amidites"
    Case 360343: Area = "Other Amidites"
    Case 360346: Area = "Other Amidites"
    Case 360367: Area = "Dye Terminators"
    Case 360368: Area = "Dye Terminators"
    Case 360510: Area = "Dye NHS Esters"
    Case 360511: Area = "Dye NHS Esters"
    Case 360512: Area = "Dye NHS Esters"
    Case 360513: Area = "Dye NHS Esters"
    Case 360648: Area = "Obsolete"
    Case 360649: Area = "Obsolete"
    Case 360666: Area = "Obsolete"
    Case 360667: Area = "Obsolete"
    Case 360677: Area = "Obsolete"
    Case 360683: Area = "Dye Terminators"
    Case 360685: Area = "Dye Terminators"
    Case 360686: Area = "Dye Terminators"
    Case 360687: Area = "Dye NHS Esters"
    Case 360740: Area = "Dye Amidites"
    Case 360741: Area = "Dye Amidites"
    Case 360742: Area = "Dye Amidites"
    Case 360743: Area = "Dye Amidites"
    Case 360744: Area = "Dye Amidites"
    Case 360745: Area = "Dye Amidites"
    Case 360746: Area = "Dye Amidites"
    Case 360747: Area = "Dye NHS Esters"
    Case 360818: Area = "Obsolete"
    Case 360819: Area = "Obsolete"
    Case 360895: Area = "Obsolete"
    Case 360908: Area = "Dye Amidites"
    Case 360909: Area = "Dye Amidites"
    Case 360910: Area = "Dye Amidites"
    Case 360928: Area = "Other Amidites"
    Case 360964: Area = "Other Triphosphates"
    Case 360965: Area = "Other Triphosphates"
    Case 360966: Area = "Other Triphosphates"
    Case 360967: Area = "Other Triphosphates"
    Case 360971: Area = "Terminator Nucleosides"
    Case 360975: Area = "Obsolete"
    Case 361011: Area = "Obsolete"
    Case 361038: Area = "Dye Amidites"
    Case 361040: Area = "Terminator Nucleosides"
    Case 361065: Area = "Terminator Nucleosides"
    Case 361340: Area = "Miscellaneous"
    Case 361377: Area = "Dye Terminators"
    Case 361452: Area = "Dye Terminators"
    Case 361453: Area = "Dye Terminators"
    Case 361454: Area = "Dye Terminators"
    Case 361455: Area = "Dye Terminators"
    Case 361464: Area = "Dye Terminators"
    Case 361511: Area = "Obsolete"
    Case 361792: Area = "Other Amidites"
    Case 361828: Area = "Dye Amidites"
    Case 361891: Area = "Dye NHS Esters"
    Case 361894: Area = "Dye Amidites"
    Case 361895: Area = "Dye Amidites"
    Case 361899: Area = "Dye Amidites"
    Case 361907: Area = "Dye Amidites"
    Case 361908: Area = "Dye Amidites"
    Case 361909: Area = "Dye Amidites"
    Case 362000: Area = "Dye Terminators"
    Case 362001: Area = "Dye Terminators"
    Case 362002: Area = "Dye Terminators"
    Case 362003: Area = "Dye Terminators"
    Case 362088: Area = "Terminator Nucleosides"
    Case 362089: Area = "Terminator Nucleosides"
    Case 362092: Area = "Dye NHS Esters"
    Case 362098: Area = "Dye NHS Esters"
    Case 362099: Area = "Dye NHS Esters"
    Case 362102: Area = "Dye NHS Esters"
    Case 362104: Area = "Dye NHS Esters"
    Case 362107: Area = "Dye NHS Esters"
    Case 362109: Area = "Dye NHS Esters"
    Case 362116: Area = "Dye NHS Esters"
    Case 362122: Area = "Dye NHS Esters"
    Case 362137: Area = "Dye Terminators"
    Case 362138: Area = "Dye Terminators"
    Case 362139: Area = "Dye Terminators"
    Case 362140: Area = "Dye Terminators"
    Case 362141: Area = "Dye Terminators"
    Case 362142: Area = "Dye Terminators"
    Case 362143: Area = "Dye Terminators"
    Case 362144: Area = "Dye Terminators"
    Case 362145: Area = "Dye Terminators"
    Case 362146: Area = "Dye Terminators"
    Case 362147: Area = "Dye Terminators"
    Case 362148: Area = "Dye Terminators"
    Case 362149: Area = "Dye Terminators"
    Case 362150: Area = "Dye Terminators"
    Case 4304303: Area = "Dye NHS Esters"
    Case 4304742: Area = "Dye Amidites"
    Case 4304744: Area = "Dye Amidites"
    Case 4304746: Area = "Dye Amidites"
    Case 4304748: Area = "Dye Amidites"
    Case 4304750: Area = "Dye Amidites"
    Case 4304752: Area = "Dye Amidites"
    Case 4306029: Area = "Other Triphosphates"
    Case Else:  Area = "Not Found"
End Select
If Right(part, 1) = ("C") Then Area = "Custom"
SynthArea = Area
End Function

Public Function SynthDesc(part As Variant) As String
Dim Desc_T As String
On Error Resume Next

Select Case part
    Case 360122: Desc_T = "AMIDITE, CE INOSINE BULK"
    Case 360180: Desc_T = "DMT-dI NUCLEOSIDE"
    Case 360317: Desc_T = "BULK,ROX-NHS/DMSO"
    Case 360318: Desc_T = "BULK,TAMRA-NHS/DMSO"
    Case 360336: Desc_T = "BULK,FAM NHS/DMSO"
    Case 360337: Desc_T = "BULK,JOE NHS/DMSO"
    Case 360340: Desc_T = "dG(dmf)NUCLEOSIDE"
    Case 360343: Desc_T = "DMT dG(dmf) NUCLEOSIDE"
    Case 360346: Desc_T = "BULK,dG(dmf)PHOSPHORAMIDITE"
    Case 360367: Desc_T = "BULK, C DYETERM"
    Case 360368: Desc_T = "BULK, T DYETERM"
    Case 360510: Desc_T = "2-CHLOROISOVANILLIN"
    Case 360511: Desc_T = "2-CHLORO-4-METHOXYRESORCINOL"
    Case 360512: Desc_T = "JOE DIACETATE"
    Case 360513: Desc_T = "6-JOE ACID"
    Case 360648: Desc_T = "OBS,BULK,RNA G DMF"
    Case 360649: Desc_T = "OBS,BULK,RNA C iBu"
    Case 360666: Desc_T = "BULK,T7 'A' TERMINATOR"
    Case 360667: Desc_T = "BULK,T7 'G' TERMINATOR"
    Case 360677: Desc_T = "NHS-ESTER, 5-NAN"
    Case 360683: Desc_T = "A-TERM AMINOTRIPHOSPHATE"
    Case 360685: Desc_T = "G-TERM AMINOTRIPHOSPHATE"
    Case 360686: Desc_T = "C-TERM AMINOTRIPHOSPHATE"
    Case 360687: Desc_T = "5-R6G NHS ESTER"
    Case 360740: Desc_T = "BULK, HEX-1 AMIDITE"
    Case 360741: Desc_T = "HEX-1 HHA"
    Case 360742: Desc_T = "BULK,HEX-1 DYE ACID"
    Case 360743: Desc_T = "BULK, TET-1 DYE ACID"
    Case 360744: Desc_T = "BULK, 6-FAM AMIDITE"
    Case 360745: Desc_T = "BULK,6-FAM-HHA"
    Case 360746: Desc_T = "BULK,6-FAM DIPIVALATE"
    Case 360747: Desc_T = "BULK,6-FAM DYE ACID       ^"
    Case 360818: Desc_T = "OBS,CRYSTAL OF PMP"
    Case 360819: Desc_T = "BULK,5/6-FAM ACID"
    Case 360895: Desc_T = "OBS-BULK,QPCR CELL CLEANER"
    Case 360908: Desc_T = "TET-1 DIPIVALATE"
    Case 360909: Desc_T = "TET-1 HHA"
    Case 360910: Desc_T = "BULK, TET-1 AMIDITE"
    Case 360928: Desc_T = "BULK,PHOSPHALINK IN ACN"
    Case 360964: Desc_T = "BULK,dUTP-PROPARGYLAMINE"
    Case 360965: Desc_T = "BULK,dUTP R110"
    Case 360966: Desc_T = "BULK,dUTP 5R6G"
    Case 360967: Desc_T = "BULK,dUTP TAMRA"
    Case 360971: Desc_T = "BULK,IODO-2',3'-DIDEOXYURIDI"
    Case 360975: Desc_T = "QPCR,RUTHENIUM AMIDITE,BULK"
    Case 361011: Desc_T = "OBS,BULK,RNA A-PAC AMIDITE"
    Case 361038: Desc_T = "HEX-1 DIPIVALATE"
    Case 361040: Desc_T = "N-PROPARGYLTRIFLUOROACETAMID"
    Case 361065: Desc_T = "5-IODO-2',3'-DIDEOXYCYTIDINE"
    Case 361340: Desc_T = "BULK,PMTC STD"
    Case 361377: Desc_T = "BULK, G  DYETERM STOCK SOLN"
    Case 361452: Desc_T = "BULK,FS T-DYE TERM"
    Case 361453: Desc_T = "BULK,FS G-DYE TERM"
    Case 361454: Desc_T = "BULK,FS C-DYE TERM"
    Case 361455: Desc_T = "BULK,FS A-DYE TERM"
    Case 361464: Desc_T = "BULK, A  DYETERM STOCK SOLN"
    Case 361511: Desc_T = "TUBE,FSP RR-G -21M13 24^"
    Case 361792: Desc_T = "BULK,CE-AMINOLINK TFA"
    Case 361828: Desc_T = "2-FLUORONAPHTHALENE-1,3-DIOL"
    Case 361891: Desc_T = "BULK,NED 2-DYE NHS"
    Case 361894: Desc_T = "NED 2 DYE ACID"
    Case 361895: Desc_T = "AMINOHEXANOL PROCESSED,BULK"
    Case 361899: Desc_T = "2-BROMOPHENYL AA ORTHO ESTER"
    Case 361907: Desc_T = "BULK,NED-2 AMIDITE"
    Case 361908: Desc_T = "NED-2 HHA"
    Case 361909: Desc_T = "NED-2 DIPIVALATE"
    Case 362000: Desc_T = "ddA dR DYE TERM, 100uM"
    Case 362001: Desc_T = "ddC dR DYE TERM, 100uM"
    Case 362002: Desc_T = "ddG dR DYE TERM, 100uM"
    Case 362003: Desc_T = "ddU dR DYE TERM, 100uM"
    Case 362088: Desc_T = "ddG(aep) Terminator Nucleosi"
    Case 362089: Desc_T = "BULK,ddU(aep)TERMINATOR NUCL"
    Case 362092: Desc_T = "bis-TFA d-R110-2 Dye Acid"
    Case 362098: Desc_T = "d-Rox-1 Dye Acid"
    Case 362099: Desc_T = "d-Rox-2 Dye Acid"
    Case 362102: Desc_T = "bis-TFA d-R110-2 NHS Ester"
    Case 362104: Desc_T = "d-R6G-2 Dye NHS Ester"
    Case 362107: Desc_T = "d-Tamra-2 Dye NHS Ester"
    Case 362109: Desc_T = "d-Rox-2 Dye NHS Ester"
    Case 362116: Desc_T = "FMOC 5-AMBAMFAM NHS Ester"
    Case 362122: Desc_T = "FMOC 6-AMBAMFAM NHS Ester"
    Case 362137: Desc_T = "BULK,ddGTP(aep) 30mM"
    Case 362138: Desc_T = "BULK,ddUTP(aep) 30mM"
    Case 362139: Desc_T = "ddA dR DYE TERM, 1.465uM"
    Case 362140: Desc_T = "ddC dR DYE TERM, 4.43uM"
    Case 362141: Desc_T = "ddG dR DYE TERM, 0.552 uM"
    Case 362142: Desc_T = "ddU dR DYE TERM, 6.877 uM"
    Case 362143: Desc_T = "ddA Big Dye Term, 1 mM"
    Case 362144: Desc_T = "ddC Big Dye Term, 1 mM"
    Case 362145: Desc_T = "ddG Big Dye Term, 1 mM"
    Case 362146: Desc_T = "ddU Big Dye Term, 1 mM"
    Case 362147: Desc_T = "ddA Big Dye TERM, 4.3 uM"
    Case 362148: Desc_T = "ddC Big Dye Term, 6.6 uM"
    Case 362149: Desc_T = "ddG Big Dye Term, 4.2 uM"
    Case 362150: Desc_T = "ddU Big Dye Term, 45 uM"
    Case 4304303: Desc_T = "Joe NHS Ester"
    Case 4304742: Desc_T = "VIC-2 Amidite Bulk"
    Case 4304744: Desc_T = "VIC-2 HHA"
    Case 4304746: Desc_T = "VIC-2 Dipivalate"
    Case 4304748: Desc_T = "VIC-2 Acid"
    Case 4304750: Desc_T = "4-Phenyl Resorcinol"
    Case 4304752: Desc_T = "1-Pheny-2,4-Dimethoxy benzen"
    Case 4306029: Desc_T = "BULK,ddA HEX-2, 1 mM"
    Case Else: Desc_T = "Not Found"
End Select
SynthDesc = Desc_T
End Function

Public Function MaxCost(p1 As Variant, p2 As Variant, p3 As Variant) As String
'inputs in the order US GB JP
On Error Resume Next
If p1 = "" Then p1 = -99999
If p2 = "" Then p2 = -99999
If p3 = "" Then p3 = -99999

If p1 >= p2 Then
    If p1 >= p3 Then
        MaxCost = "US"
    ElseIf p3 > p1 Then
        MaxCost = "JP"
    End If
ElseIf p2 > p1 Then
    If p2 >= p3 Then
        MaxCost = "GB"
    ElseIf p3 > p2 Then
        MaxCost = "JP"
    End If
End If
End Function
Public Function MinCost(p1 As Variant, p2 As Variant, p3 As Variant) As String
'inputs in the order US GB JP
On Error Resume Next
If p1 = "" Then p1 = 999999
If p2 = "" Then p2 = 999999
If p3 = "" Then p3 = 999999

If p1 <= p2 Then
    If p1 <= p3 Then
        MinCost = "US"
    ElseIf p3 < p1 Then
        MinCost = "JP"
    End If
ElseIf p2 < p1 Then
    If p2 <= p3 Then
        MinCost = "GB"
    ElseIf p3 < p2 Then
        MinCost = "JP"
    End If
End If
End Function

