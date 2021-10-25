<%

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

%>

<!doctype html>
<html lang="en">
<head>

	<meta name="viewport" content="width=device-width, initial-scale=1">

	<!-- Bootstrap CSS -->
	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">

	<style>
		html, body {
			background: #ffffff;
			height:100%;
			margin: 0;
		}

		body, div, p {
			font-family: Arial;
			color:#333333;
			font-size:13px;
		}

		p {
			margin: 0 0 5px 0;
			line-height:20px;
		}

		hr {
			border:1px;
			height:1px;
			color:#CCCCCC;
			background-color: #CCCCCC;
			margin: 20px 0 20px 0;
			clear:both;
		}

		table {
			border: 1px solid #cccccc;
			padding:10px;
			margin:10px;
		}

		td {
			border: 1px solid #cccccc;
			padding:10px;
			text-align:right;
		}

	/* debug
	-------------------------------------------
	div {border: 1px solid #000; }*/

	#page {
		position: relative;
		margin: 50px auto;
		padding: 0;
		width: 500px;
		padding: 0;
		text-align: left;
	}

	#container {
		position:relative;
		top:100px;
	}

	#konsol {
		position: fixed;
		margin:0px;
		padding:20px;
		width:500px;
		height:100px;
		border: 1px solid #CCCCCC;
		background-color:#ffffff;
		z-index:10;
	}

	</style>

	<title>Hello, world!</title>

</head>


<body>

<div id="page">

		<div class="dropdown" style="margin:0 0 15px 0;">

			<button class="btn btn-secondary dropdown-toggle" type="button" id="dropdownMenu2" data-bs-toggle="dropdown" aria-expanded="false">
				Meny
			</button>

			<ul class="dropdown-menu" aria-labelledby="dropdownMenu2">
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=0';">0. Start</button></li>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=1';">1. Hello World</button></li>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=2';">2. Användare</button></li>
				<% if request("color") = 0 then %>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=3&color=1';">3. Ändra färg</button></li>
				<% else %>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=3';">3. Ändra färg</button></li>
				<% end if %>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=4';">4. Dagens datum</button></li>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=5';">5. Högsta talet</button></li>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=6&createRandomNr=1';">6. Gissa på talet</button></li>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=7';">7 - 8. Skapa textfil</button></li>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=10';">10. Multiplikations tabellen</button></li>
				<li><button class="dropdown-item" type="button" onClick="location.href='default.asp?case=11';">11. Array</button></li>
			</ul>

		</div>

	<div id="konsol">
		<h3>Konsol</h3>
			<%= strKonsol %>
	</div>

	<hr />

	<div id="container">

		<!-- Användare -->
		<% if dblCase = 2 then %>
		<form method="post" action="default.asp" id="frm2">
		<input type="hidden" name="case" value="2">
			<div class="row">
				<div class="col">
					<input type="text" name="firstname" class="form-control" placeholder="Förnamn" aria-label="Förnamn" required />
				</div>
				<div class="col">
					<input type="text" name="lastname" class="form-control" placeholder="Efternamn" aria-label="Efternamn" required />
				</div>
				<div class="col">
					<input type="number" name="age" class="form-control" placeholder="Ålder" aria-label="Ålder" required />
				</div>
			</div>

			<br />
			<button type="submit" class="btn btn-primary">Sänd...</button>
		</form>
		<% end if %>

		<!-- Ändra färg -->
		<!-- // -->

		<!-- Högsta talet -->
		<% if dblCase = 5 then %>
		<form method="post" action="default.asp" id="frm5">
		<input type="hidden" name="case" value="5">
			<div class="row">
				<div class="col">
					<input type="number" name="nr1" value="<%= request("nr1") %>" class="form-control" placeholder="0" aria-label="Tal 1" required />
				</div>
				<div class="col">
					<input type="number" name="nr2" value="<%= request("nr2") %>" class="form-control" placeholder="0" aria-label="Tal 2" required />
				</div>
			</div>
			<br />
			<button type="submit" class="btn btn-primary">Kolla...</button>
		</form>
		<% end if %>

		<!-- Gissa på talet -->
		<% if dblCase = 6 then %>
		<form method="post" action="default.asp" id="frm6">
		<input type="hidden" name="case" value="6">
		<input type="hidden" name="createRandomNr" value="0">
			<div class="row">
				<div class="col">
					<input type="number" name="compareNr" value="<%= request("compareNr") %>" class="form-control" placeholder="Ange ett tal" aria-label="Ange ett tal" required />
				</div>
			</div>
			<br />

			<button type="button" class="btn btn-primary" onclick="document.location.href=('default.asp?case=6&createRandomNr=1');">Hämta ett nytt tal</button>
			<button type="submit" class="btn btn-primary">Jämför (<%= session("dblCounter") %>)</button>

		</form>
		<% end if %>


		<!-- Spara en fil -->
		<% if dblCase = 7 then %>
		<form method="post" action="default.asp" id="frm7">
		<input type="hidden" name="case" value="7">
		<div class="container">
			<div class="form-group">
			<textarea name="updFile" class="form-control" rows="5" placeholder="Skapa/Uppdatera en textfil..."><%= strTextArea %></textarea>
		</div>

		<br />
		<button type="submit" class="btn btn-primary">Skapa/Uppdatera en textfil... (lexicon.txt) fil</button>

		</form>
		<% end if %>

		<!-- Multiplikation / tabell-->
		<% if dblCase = 10 then %>

			<%= strContainer %>

		<% end if %>
		</div><!-- //container -->
</div><!-- //page -->


<!-- Option 1: Bootstrap Bundle with Popper -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>

</body>
</html>

<%

set myObj = nothing
set objFso = nothing

%>