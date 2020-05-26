function print(df)
	vt = typename(df)
	if vt = "Variant()" then
		c = typename(df(0))
		if c = "Variant()" then
			if array_len(df(0)) > 6 then
				p(df)
			else
				l(df)
			end if
		else 
			wscript.echo "['" & join(df,"','") & "']"
		end if
	else
		wscript.echo df
	end if
end function
function l(df)
	redim arr(-1)
	do while y < array_len(df(0))
		redim preserve arr(ubound(arr)+1)
		arr(ubound(arr)) = 0
		y = y + 1
	loop
	x = 0
	do while x < 5
		y = 0
		do while y < array_len(df(0))
			string_value_len = str_len(df(x)(y))
			if string_value_len > arr(y) then
				arr(y) = string_value_len
			end if
			y = y + 1
		loop
		x = x + 1
	loop
	x = array_len(df)-5
	do while x < array_len(df)
		y = 0
		do while y < array_len(df(0))
			string_value_len = str_len(df(x)(y))
			if string_value_len > arr(y) then
				arr(y) = string_value_len
			end if
			y = y + 1
		loop
		x = x + 1
	loop
	x = 0
	do while x < array_len(df(0))
		y = 0
		do while y < 5
			if str_len(df(y)(x)) < arr(x) then
				difference = arr(x) - str_len(df(y)(x))
				x = 0
				padding = ""
				do while z < difference
					padding = padding + " "
					z = z + 1
				loop
				df(y)(x) = padding & df(y)(x)
			end if
			y = y + 1
		loop
		x = x + 1
	loop
	x = 0
	redim filler_line_arr(-1)
	do while x < array_len(df(0))
		y = array_len(df)-4
		do while y < array_len(df)
			if str_len(df(y)(x)) < arr(x) then
				difference = arr(x) - str_len(df(y)(x))
				z = 0
				padding = ""
				do while z < difference
					padding = padding & " "
					z = z + 1
				loop
				df(y)(x) = padding & df(y)(x)
			end if
			y = y + 1
		loop
		filler_line = padding & filler_line
		redim preserve filler_line_arr(ubound(filler_line_arr)+1)
		filler_line_arr(ubound(filler_line_arr)) = filler_line
		x = x + 1
	loop
	wscript.echo " " & chr(9) & join(df(0),chr(9))
	wscript.echo "0" & chr(9) & join(df(1),chr(9))
	wscript.echo "1" & chr(9) & join(df(2),chr(9))
	wscript.echo "2" & chr(9) & join(df(3),chr(9))
	wscript.echo "3" & chr(9) & join(df(4),chr(9))
	wscript.echo "..." & chr(9) & join(filler_line_arr,chr(9))
	wscript.echo array_len(df)-3 & chr(9) & join(df(array_len(df)-3),chr(9))
	wscript.echo array_len(df)-2 & chr(9) & join(df(array_len(df)-2),chr(9))
	wscript.echo array_len(df)-1 & chr(9) & join(df(array_len(df)-1),chr(9))
end function
function p(df)
	redim arr(-1)
	do while y < array_len(df(0))
		redim preserve arr(ubound(arr)+1)
		arr(ubound(arr)) = 0
		y = y + 1
	loop
	x = 0
	do while x < 5
		y = 0
		do while y < array_len(df(0))
			string_value_len = str_len(df(x)(y))
			if string_value_len > arr(y) then
				arr(y) = string_value_len
			end if
			y = y + 1
		loop
		x = x + 1
	loop
	x = array_len(df)-5
	do while x < array_len(df)
		y = 0
		do while y < array_len(df(0))
			if string_value_len > arr(y) then
				arr(y) = string_value_len
			end if
			y = y + 1
		loop
		x = x + 1
	loop
	x = 0
	do while x < array_len(df(0))
		y = 0
		do while y < 5
			if str_len(df(y)(x)) < arr(x) then
				difference = arr(x) - str_len(df(y)(x))
				z = 0
				padding = ""
				do while z < difference
					padding = padding & " "
					z = z + 1
				loop
				df(y)(x) = padding & df(y)(x)
			end if
			y = y + 1
		loop
		x = x + 1
	loop
	
	x = 0
	do while x < array_len(df(0))
		y = array_len(df)-3
		do while y < array_len(df)
			if str_len(df(y)(x)) < arr(x) then
				difference = arr(x) - str_len(df(y)(x))
				z = 0
				padding = ""
				do while z < difference
					padding = padding & " "
					z = z + 1
				loop
				df(y)(x) = padding & df(y)(x)
			end if
			y = y + 1
		loop
		x = x + 1
	loop
	x = 0
	redim padded_array(-1)
	do while x < 3
		difference = arr(x) - 3
		y = 0
		padded_line = ""
		do while y < difference
			padded_line = padded_line + " "
			y = y + 1
		loop
		redim preserve padded_array(ubound(padded_array)+1)
		padded_array(ubound(padded_array)) = padded_line & "..."
		x = x + 1
	loop
	x = array_len(arr)-3
	redim padded_array2(-1)
	do while x < array_len(arr)
		difference = arr(x)-3
		y = 0
		padded_line = ""
		do while y < difference
			padded_line = padded_line + " "
			y = y + 1
		loop
		redim preserve padded_array2(ubound(padded_array2)+1)
		padded_array2(ubound(padded_array2)) = padded_line & "..."
		x = x + 1
	loop
	wscript.echo " " & chr(9) & df(0)(0) & chr(9) & df(0)(1) & chr(9) & df(0)(2) & chr(9) & "..." & chr(9) & df(0)(array_len(df(0))-3) & chr(9)  & df(0)(array_len(df(0))-2) & chr(9) & df(0)(array_len(df(0))-1)  
	wscript.echo "0" & chr(9) & df(1)(0) & chr(9) & df(1)(1) & chr(9) & df(1)(2) & chr(9) & "..." & chr(9) & df(1)(array_len(df(1))-3) & chr(9)  & df(1)(array_len(df(1))-2) & chr(9) & df(1)(array_len(df(1))-1)  
	wscript.echo "1" & chr(9) & df(2)(0) & chr(9) & df(2)(1) & chr(9) & df(2)(2) & chr(9) & "..." & chr(9) & df(2)(array_len(df(2))-3) & chr(9)  & df(2)(array_len(df(2))-2) & chr(9) & df(2)(array_len(df(2))-1)  
	wscript.echo "2" & chr(9) & df(3)(0) & chr(9) & df(3)(1) & chr(9) & df(3)(2) & chr(9) & "..." & chr(9) & df(3)(array_len(df(3))-3) & chr(9)  & df(3)(array_len(df(3))-2) & chr(9) & df(3)(array_len(df(3))-1)  
	wscript.echo "..." & chr(9) & join(padded_array,chr(9))  & chr(9) & "..." & join(padded_array2,chr(9))	
	wscript.echo array_len(df)-3 & chr(9) & df(array_len(df)-3)(0) & chr(9) & df(array_len(df)-3)(1) & chr(9) & df(array_len(df)-3)(2) & chr(9) & "..." & chr(9) & df(array_len(df)-3)(array_len(df(0))-3) & chr(9) & df(array_len(df)-3)(array_len(df(0))-2) & chr(9) & df(array_len(df)-3)(array_len(df(0))-1)
	wscript.echo array_len(df)-2 & chr(9) & df(array_len(df)-2)(0) & chr(9) & df(array_len(df)-2)(1) & chr(9) & df(array_len(df)-2)(2) & chr(9) & "..." & chr(9) & df(array_len(df)-2)(array_len(df(0))-2) & chr(9) & df(array_len(df)-2)(array_len(df(0))-2) & chr(9) & df(array_len(df)-2)(array_len(df(0))-1)
	wscript.echo array_len(df)-1 & chr(9) & df(array_len(df)-1)(0) & chr(9) & df(array_len(df)-1)(1) & chr(9) & df(array_len(df)-1)(2) & chr(9) & "..." & chr(9) & df(array_len(df)-1)(array_len(df(0))-1) & chr(9) & df(array_len(df)-1)(array_len(df(0))-2) & chr(9) & df(array_len(df)-1)(array_len(df(0))-1)
end function
function append(arr,value)
	redim preserve arr(ubound(arr)+1)
	arr(ubound(arr)) = value
	append = arr
end function
function alphabet()
	alphabet = array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")
end function
