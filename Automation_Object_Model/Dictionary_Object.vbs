
Option Explicit

Dim dobj
Dim i

set dobj=CreateObject("Scripting.Dictionary")

Msgbox dobj.Count

For i=0 to 3 

	dobj.Add i,"Aniket"&i 

Next

For each i in dobj.keys()
	Msgbox dobj(i)
Next


dobj(2)="Aniket4" 'override the value


For each i in dobj.items()
	Msgbox i
Next


Msgbox dobj.Exists(1)


dobj.Remove(1)

Msgbox dobj.Exists(1)

Msgbox dobj.Count

dobj.RemoveALl()

Msgbox dobj.Count