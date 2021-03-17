option explicit
Dim C,D,P,U,Umax,Borne
Dim a,x,j,k,tempD,tempC,tempP
Dim I,R,L,Lp,Bp,Hp,PPMC
Dim Cas
Cas = 0

'Ti(C,D,P)
C = Array(1,2,3)		'Temps de Calcul
D = Array(4,5,7) 	'Deadline
P = Array(4,6,8)		'Periode

'RM: priorité aux periodes les plus courtes : ordonnacement à priorité fixe
'DM: priorité aux deadline les plus courtes : ordonnancement optimal à priorité fixe
'EDF: priorité à la prochaine échéance : ordonnacement optimal à priorité variable
'BP ou L: longueur de la periode d'activité (busy period) : utile pour EDF quand D<P
'PPMC: plus petit multiplicateur commun = Hp: Hyperperiode
'Tr=R=Temps de réponse maximal

Call SortD
Call CalcPPMC
'MsgBox PPMC
Call CalculU
MsgBox "U= "&U
Call CalculUmax
Call CalculBorne 'pour RM

For k=0 to uBound(D)
	If D(k)<P(k) Then
		Cas = 1
	End If
Next


'If Cas = 0 Then
	MsgBox "U = "&U&" -> OK si <1 en EDF(D=P) ou <0.69 en RM"& vbCrLf & "RM OK si U<" & Borne

	'------Si D=P------'
	If U<Borne Then
	MsgBox "EDF(D=P) et RM OK"
	elseIf U<=1 Then
	MsgBox "EDF(D=P) Ok / RM indefini"
	elseif U>1 Then
	MsgBox "EDF(D=P) et RM NOK"
	End If
'Else
	'-----Si D<=P------'
	If Umax<Borne Then	'Test DM grossier
	'MsgBox "DM OK"
	End If
	Call OrdoDM 'Version D<=P de RM, equivalent si D=P 'test fin
	
	Call CalcBp
	Call OrdoEDF

	'MsgBox PPMC
	'MsgBox Bp
'End If



'_______________________________________________________________'

Sub OrdoEDF
Dim Max, Liste, CurrL, xxx,zzz,Cii
Liste = D
xxx=2
zzz=0
	if Bp > PPMC then
		Max = Bp
	Else
		Max = PPMC
	End If

	For zzz=0 to uBound(D)
		Do While CurrL<=Max
		
			If D(zzz)*xxx > Max Then
			Else
				ReDim Preserve Liste(UBound(Liste) + 1)
				Liste(UBound(Liste)) = D(zzz)*xxx
			End If
			
			CurrL=D(zzz)*xxx
			xxx=xxx+1
			
			If xxx>50 Then 'au cas ou :D /fear
				Exit Do
			End If
		Loop
	Next
	
	For zzz=0 to uBound(Liste)
	
		For xxx=0 to uBound(C)
			Cii=Cii+((Int((Liste(zzz)-D(xxx))/P(xxx))+1)*C(xxx))
		Next
		
		if Cii>Liste(zzz) Then
			MsgBox "Non ordonancable en EDF(D<P) : L = "&Liste(zzz)&" < "&Cii&" = Sum L-D/P +1 *C"
			
			Exit For
		End If
		
		Cii=0
	Next
	
			MsgBox "Ordonancable en EDF(D<P) sauf avis contraire"
	
End Sub




Sub OrdoDM

For x=0 To uBound(C) 
	If Tr(x)>D(x) Then
		MsgBox "Pas Ordonancable en DM : T"& x &"  ne peut pas s'executer"
		Exit For
	End If
	If x=uBound(C) Then
		MsgBox "Ordonancable en DM et RM : toutes les deadlines sont respectees"
	End If
Next

End Sub


Function Tr(num)
Dim count
I=0
R=C(num)
	If num=0 Then
		Tr=R
		Exit Function
	End If

	For count = 0 to num-1
		I=I+(roundUp(R/P(count))*C(count))
	Next

	Do While (I+C(num)) > R
		R=I+C(num)
		I=0
		For count = 0 to num-1
			I=I+(roundUp(R/P(count))*C(count))
		Next
	Loop
'MsgBox R
Tr=R
End Function

Function LCM(a,b)
Dim zz,tempa,tempb
tempa=a
tempb=b
	zz=tempa*tempb
	while tempa<>tempb
		if tempa<tempb Then
		tempb= tempb-tempa
		else tempa=tempa-tempb
		end if
	Wend
LCM = zz/tempa
End Function

Sub CalcPPMC
Dim z
	PPMC = P(0)
		For z=1 to uBound(P)
		PPMC = LCM(PPMC,P(z))
		Next
End Sub


Function W(L)
Dim res,xx
res=0
	For xx=0 to uBound(C)
	res = res+(roundUp(L/P(xx))*C(xx))
	Next
W = res
End Function

Sub CalcBp
For x=0 to uBound(C)
 L=L + C(x)
Next

Lp=W(L)
Hp=PPMC

	Do While Lp<=Hp
	L=Lp
	Lp=W(L)
		If Lp=L Then
			Exit Do
		End If
	Loop
	
	If(Lp<=Hp) Then
	Bp=L
	else Bp=9999999
	End If
End Sub


Sub CalculU
	For x=0 To uBound(C)	
		U=U+(C(x)/D(x))	
	Next
End Sub

Sub CalculUmax
	For x=0 To uBound(C)	
		Umax=Umax+(C(x)/D(x))	
	Next
End Sub

Sub CalculBorne
	
	Borne = (uBound(C)+1)*((2^(1/(uBound(C)+1)))-1)

End Sub

Sub SortD
	for a = UBound(D) - 1 To 0 Step -1
		for j= 0 to a
			if D(j)>D(j+1) then
				tempD=D(j+1)
				D(j+1)=D(j)
				D(j)=tempD
				
				tempC=C(j+1)
				C(j+1)=C(j)
				C(j)=tempC
				
				tempP=P(j+1)
				P(j+1)=P(j)
				P(j)=tempP
			end if
		next
	next 
End Sub

function roundUp(numToRound)
  roundUp = Round(numToRound + .499)
end function