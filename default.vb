option explicit
response.expiresAbsolute = now() - 1
response.buffer = true
session.LCID = 1053
session.timeOut = 60
response.charset="ISO-8859-1"


'skapar instans till projektet
dim myObj
Set myObj = New clsObj

Dim objFso
Dim f
dim ts

dim dblCase
dim x : x = 0
dim Y : y = 0

dim arrValues

dim strKonsol
dim strValues
dim strTextArea
dim strContainer

set objFso = server.createObject("scripting.fileSystemObject")

dblCase = request("case")
if not isNumeric(dblCase) then
	dblCase = 0
end if


class clsObj

	'case 1
	public function callHelloWorld(byVal dblCase, byRef strKonsol)
		strKonsol = "Hello World!"
	end function

	'case 2
	public function callUser(byVal dblCase, strValues, byRef strKonsol)

		arrValues = split(strValues, "¤¤¤", -1, 1)

		for x = 0 to Ubound(arrValues)
			strKonsol = strKonsol &" "& arrValues(x)
		next

	end function

	'case 3
	public function callColor(byVal dblCase, byVal strValues, byRef strKonsol)

		if strValues = 1 then
			strKonsol = "<span style=""color:#CCCCCC;""><h7>Ändra färg...</h7></span>"
		else
			strKonsol = "Ändra färg..."
		end if

	end function

	'case 4
	public function callDate(byVal dblCase, byRef strKonsol)
		strKonsol = formatDateTime(now, 2)
	end function

	'case 5
	public function callHighest(byVal dblCase, strValues, byRef strKonsol)

		arrValues = split(strValues, "¤¤¤", -1, 1)

		if arrValues(0) < arrValues(1) then
			strKonsol = "Högsta värdet: "& arrValues(1)

		elseIf  arrValues(0) > arrValues(1) then
			strKonsol = "Högsta värdet: "& arrValues(0)

		elseIf  arrValues(1) = arrValues(0) then
			strKonsol = "Lika..."

		elseIf  CInt(arrValues(1)) = 0 and CInt(arrValues(0)) = 0 then
			strKonsol = ""

		end if

	end function


	'case 6 - hämta nummer 1-100
	public function callRandomNr(byVal dblCase, byVal strValues, byRef strKonsol)

		dim intUpper : intUpper = 100
		dim intLower : intLower = 1

		arrValues = split(strValues, "¤¤¤", -1, 1)

		if arrValues(0) = 1 then 'createRandomNr
			randomize
			session("strKonsol") = "Gissa på talet: "
			session("dblSessionNr") = int((intUpper - intLower + 1) * rnd + intLower)
			session("dblCounter") = 0
		end if

		strKonsol = session("strKonsol") & session("dblSessionNr")

		if CInt(session("dblSessionNr")) < CInt(arrValues(1)) and CInt(arrValues(1)) <> 0 then
			strKonsol = strKonsol &", Talet är för stort!"
			session("dblCounter") = session("dblCounter") + 1

		elseIf CInt(session("dblSessionNr")) > CInt(arrValues(1)) and CInt(arrValues(1)) <> 0 then
			strKonsol = strKonsol &", Talet är fär litet!"
			session("dblCounter") = session("dblCounter") + 1

		elseIf CInt(session("dblSessionNr")) = CInt(arrValues(1)) and CInt(arrValues(1)) <> 0  then
			strKonsol = strKonsol &", Talet är rätt!"

		end if

	end function


	'case 7 - 8
	public function callEditFile(byVal dblCase, byVal strValues, byRef strKonsol, byRef strTextArea)

		arrValues = split(strValues, "¤¤¤", -1, 1)

		set f = objFso.GetFile(server.mapPath("file/lexicon.txt"))
		set ts = f.OpenAsTextStream(2)
			ts.write  arrValues(1) &" "
		ts.close

		strKonsol = server.mapPath(".") &"\file\lexicon.txt"

		set f = objFso.openTextFile(server.mapPath("file/lexicon.txt"), 1)
			strTextArea = f.readAll
		f.close

	end function

	'case 10
	public function callMulti(byVal dblCase, byRef strContainer)

	strContainer = "<table>"

		for y = 1 to 10
			strContainer = strContainer &"<tr>"
				for x = 1 to 10
					strContainer = strContainer &"<td>"& x * y &"</td>"
				next
			strContainer = strContainer &"</tr>"
		next

		strContainer = strContainer &"</table>"

	end function

	'case 11
	public function callArray(byVal dblCase, byRef strKonsol)

		redim intArray(1)

		randomize
		intArray(0) = int((100 - 1 + 1) * rnd + 1)
		intArray(1) = int((99 - 1 + 1) * rnd + 1)

		if intArray(0) > intArray(1) then
			strKonsol = intArray(0) &" "& intArray(1)
		else
			strKonsol = intArray(0) &" "& intArray(1)
		end if

	end function

end class

select case dblCase

	case 1
		myObj.callHelloWorld dblCase, strKonsol

	case 2
		strValues = request("firstname") &"¤¤¤"& request("lastname") &"¤¤¤"& request("age")
		myObj.callUser dblCase, strValues, strKonsol

	case 3
		strValues = request("color")
		myObj.callColor dblCase, strValues, strKonsol

	case 4
		myObj.callDate dblCase, strKonsol

	case 5
		strValues = request("nr1") &"¤¤¤"& request("nr2")
		myObj.callHighest dblCase, strValues, strKonsol

	case 6
		strValues = CInt(request("createRandomNr")) &"¤¤¤"& CInt(request("compareNr"))
		myObj.callRandomNr dblCase, strValues, strKonsol

	case 7
		strValues = "lexicon.txt"&"¤¤¤"& request("updFile")
		myObj.callEditFile dblCase, strValues, strKonsol, strTextArea

	case 10
		myObj.callMulti dblCase, strContainer

	case 11
		myObj.callArray dblCase, strKonsol



end select

set myObj = nothing
set objFso = nothing
