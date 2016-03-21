Sub errchk(r, desc, o)
	if (r = 0) then
		MsgBox("Error " & desc & ": " & o.sLastErrorMsg)
		oss.nCleanup()
		oss.DropObject()
		Set oss = Nothing
		WScript.Quit
	end if
end sub

function splitstring(data)
	r=""
    for i=1 to Len(data)
		m = Mid(data, i, 1)
		if Asc(m) > 128 then
			r = r & Chr(255)
		else
			r = r & Mid(data, i, 1)
		end if
	next

	splitstring=split(r, CStr(Chr(255)))
end function

function splitstring_crlf(data)
	r=""
    i=1
	do while true
		m = Mid(data, i, 1)
		if (Asc(m) > 128) then
			r = r & Chr(255)
		elseif (Asc(m) = 13) then
			r = r & Chr(255)
			i = i + 1
			
			if (i <= Len(data) and Asc(m) = 10) then
				i = i - 1
			end if
		else
			r = r & Mid(data, i, 1)
		end if

		i = i + 1
		if (i > Len(data)) then
			Exit do
		end if
	loop

	splitstring_crlf=split(r, CStr(Chr(255)))
end function

function findindex(arr(), val)
	findindex=-1
	
	for i=LBound(arr) to UBound(arr)
		if (arr(i) = val) then
			findindex=i
			exit for
		end if
	next
end function

function getfieldvalue(name, fields, vals)
	getfieldvalue=""
	ind=findindex(splitstring(fields), name)
	if (ind >=0) then
		getfieldvalue=(splitstring(vals))(ind)
	end if
end function

function getfieldnames(data)
	r=""
	data=mid(data, instr(data, Chr(2)) + 1, Len(data))
	for i=1 to Len(data)
		m = Mid(data, i, 1)
		if (m = "%") then
			r = r & ","
		elseif (Asc(m) < Asc(" ")) then
			r = r & " "
		elseif (Asc(m) = "?") then
			r = r & " "
		elseif (Asc(m) > Asc("z")) then
			r = r & " "
		else
			r = r & m
		end if
	next
	r = Replace(r, " ", "")
	r = Replace(r, ",", Chr(255))
	getfieldnames = r
end function

function wSplitString(data)
   	r = ""
    i = 1
    if (InStr(data, Chr(138)) > 0) then
        while i <= Len(data)
            char=Mid(data, i, 1)
            if (char = Chr(138)) then
                r = r & Chr(255)
            else
                r = r & char
            end if
            i = i + 1
        wend
    else
        while i <= Len(data)
            char=Mid(data, i, 1)
            if (Asc(char) > 128) then
                r = r & Chr(255)
            else
                r = r & char
            end if
            i = i + 1
        wend
    end if

    r = Replace(r, Chr(13), " ")
    r = Replace(r, Chr(10), " ")
    r = Replace(r, Chr(255), vbCrLf)
    wSplitString = r
end function

'_ParseFieldValuesList
function PreparedString(data, removeEmptyStrings)
	data = wSplitString(data)
    if (removeEmptyStrings) then
        while InStr(data, vbCrLf & vbCrLf) > 0
            data = Replace(data, vbCrLf & vbCrLf, vbCrLf)
            if (InStr(data, vbCrLf) = 1) then
                data = Mid(data, Len(vbCrLf) + 1)
            end if
        wend

        if (Mid(data, Len(data) - Len(vbCrLf) + 1, Len(vbCrLf)) = vbCrLf) then
            data = Mid(data, 1, Len(data) - Len(vbCrLf))
        end if
    end if

    PreparedString = data
end function

Function strDup(dup, c)
  Dim res, i

  res = ""
  For i = 1 To c
    res = res & dup
  Next
  strDup = res
End Function
