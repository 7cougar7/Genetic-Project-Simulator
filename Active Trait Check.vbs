Option Explicit
dim objFSO
dim shell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set shell=CreateObject("Wscript.Shell")


dim dadtrait
dim momtrait
dim curactive

'Dominant traits are represented in the form of positive numbers
'Use only if trait is either active or not active

dadtrait=-3
momtrait=-3

Sub activeCheck(x,y)
	if x<0 then 'not dominant
		if momtrait<0 then 'not dominant
			curactive=x
		else				'if dominant
			curactive=y
		end if
	else
		curactive=x
	end if
	
End Sub

Call activeCheck(dadtrait,momtrait)
msgbox(curactive)