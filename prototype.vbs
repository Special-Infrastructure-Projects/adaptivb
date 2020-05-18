dim fso
set fso = CreateObject("Scripting.FileSystemObject")
function read_csv(file,delimiter)
	set f = fso.OpenTextFile(file)
	redim arr(-1)
	do until f.AtEndOfStream
		str = f.ReadLine
		cells = Split(str,delimiter)
		redim preserve arr(UBound(arr)+1)
		arr(UBound(arr)) = cells
	loop
	read_csv = arr
end function
function LTrimEX(str)
	dim re
	set re = New RegExp
	re.Pattern = "^\s*"
	re.Multiline = False
	LTrimEX = re.Replace(str,"")
end function
function filter(df, column, value)
	x = 0
	for each col in df(0)
		if col = column then
			column = x
		else 
			x = x + 1
		end if
	next
	redim f(-1)
	x = 0
	for each row in df
		if x = 0 then
			redim preserve f(UBound(f)+1)
			f(UBound(f)) = row
			x = x + 1
		elseif row(column) = value then
			redim preserve f(UBound(f)+1)
			f(UBound(f)) = row
		end if
	next	
	filter = f
end function
function len(df)
	len = UBound(df)+1	
end function
function inArray(arr, obj)
	On Error Resume Next
		dim x: x = -1
		if isArray(arr) then	
			for i = 0 to UBound(arr)
				if arr(i) = obj then
					x = i
					exit for
				end if
			next
		end if
		call err_report()
		inArray = x		
end function
function append(arr,value)
	redim preserve arr(UBound(arr)+1)
	arr(UBound(arr)) = value
	append = arr
end function
function to_csv(df,file,delimiter)
	Set File = FSO.CreateTextFile(file,True)
	for each line in df
		File.WriteLine Join(line,delimiter)
	next
	File.Close
end function
function columns(df)
	columns = df(0)
end function
function print(value)
	vt = TypeName(value)
	if vt = "String" or vt = "Integer" or vt = "Long" then
		wscript.echo value
	else 
		if vt = "Variant()" then
			if TypeName(value(0)) = "String" or TypeName(value(0)) = "Integer" then
				wscript.echo "[" & Join(value,",") & "]"
			else
				if len(value(0)) > 6 then
					x = 0
					do while x < 5 
						wscript.echo value(x)(0)&"	" & value(x)(1) & "	"& value(x)(2) & "	...	" & value(x)(len(value(x))-2) & "	" & value(x)(len(value(x))-1)
						x = x + 1
					loop
					wscript.echo "			...			"
					y = 5
					do while y > 0
						v = len(value) - y
						wscript.echo value(v)(0)&"	" & value(v)(1) & "	"& value(v)(2) & "	...	" & value(v)(len(value(v))-2) & "	" & value(v)(len(value(v))-1)
						y = y-1
					loop
				else if len(value(0)) < 6 and len(value) > 10 then
					x = 0
					do while x < 5
						wscript.echo join(value(x),"	")
						x = x + 1
					loop
					y = 5
					do while y > 0
						v = len(value) - y
						wscript.echo join(value(v),"	")
						y = y - 1
					loop
					end if
				end if
			end if	
		end if	
	end if
end function
function unique(df, column)
	cols = df(0)	
	x = 0
	y = 0
	for each col in cols
		if col = column then
			y = x
		end if	
		x = x + 1
	next
	redim uniqueArray(-1)
	for each row in df
		target = row(y)
		if target <> column then
			if inArray(uniqueArray,target) = -1 then
				v = append(uniqueArray,target)
			end if
		end if
	next
	unique = uniqueArray
end function
