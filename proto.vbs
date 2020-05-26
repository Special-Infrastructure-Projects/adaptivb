set fso = CreateObject("Scripting.FileSystemObject")
function read_csv(file,delimiter)
	set f = fso.OpenTextFile(file)
	redim arr(-1)	
	if delimiter = "\t" then
		delimiter = chr(9)
	end if
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
function filter(df, column,value,operator_type)
	ty = typename(value)
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
			redim preserve f(ubound(f)+1)
			f(ubound(f)) = row
			x = x + 1
		else if ty = "String" then
			if row(column) = value then
				redim preserve f(ubound(f)+1)
				f(ubound(f)) = row
			end if
		end if
		end if
	next
	filter = f
end function
function array_len(df)
	array_len = ubound(df)+1
end function
function alphabet()
	alphabet = array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")
end function
function str_len(df)
	str_len = len(df)
end function
function inArray(arr,obj)
	on error resume next
	dim x: x = -1
	if isArray(arr) then
		for i = 0 to ubound(arr)
			if arr(i) = obj then
				x = i
				exit for
			end if
		next
	end if
	call err_report()
	inArray = x
end function
function to_csv_index(df,file,delimiter)
	set File = fso.CreateTextFile(file,true)
	x = 0
	for each line in df
		if x > 0 then
			File.WriteLine x & delimiter & join(line,delimiter)
		else
			File.WriteLine "" & delimiter & join(line,delimiter)
		end if
		x = x + 1
	next
	File.Close
end function
function to_csv(df,file,delimiter,index)
	if delimiter = "\t" then
		delimiter = chr(9)
	end if
	if index = 0 then
		output_file = to_csv_index(df,file,delimiter)
	else
		set File = fso.CreateTextFile(file,True)
		for each line in df
			File.WriteLine Join(line,delimiter)
		next
		File.Close
	end if
	to_csv = df
end function
function columns(df)
	columns = df(0)
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
				redim preserve uniqueArray(ubound(uniqueArray)+1)
				uniqueArray(ubound(uniqueArray)) = target
			end if
		end if
	next
	unique = uniqueArray
end function
function slice(df,columns)
	redim col_locs(-1)
	for each column in columns
		x = 0
		for each col in df(0)
			if col = column then
				redim preserve col_locs(ubound(col_locs)+1)
				col_locs(ubound(col_locs)) = x
			else
				x = x + 1
			end if
		next
	next
	redim new_frame(-1)
	for each row in df
		redim new_row(-1)
		for each col in col_locs
			redim preserve new_row(ubound(new_row)+1)
			new_row(ubound(new_row)) = new_row
		next
		redim preserve new_frame(ubound(new_frame)+1)
		new_frame(ubound(new_frame)) = new_row
	next
	slice = new_frame
end function		
function fillna(df)
	redim filled_frame(-1)
	for each row in df
		redim filled_row(-1)
		for each col in row
			redim preserve filled_row(ubound(filled_row)+1)
			if col = "" then
				filled_row(ubound(filled_row)) = "N/A"
			else
				filled_row(ubound(filled_row)) = col
			end if
		next
		redim preserve filled_frame(ubound(filled_frame)+1)
		filled_frame(ubound(filled_frame)) = filled_row
	next
	fillna = filled_frame
end function			
function merge(left,lo,right,ro)	
	left_cols = left(0)
	right_cols = right(0)
	x = 1
	do while x < array_len(left_cols)
		if left_cols(x) = lo then
			left_target = x
		end if
		x = x + 1
	loop	
	x = 1
	do while x < array_len(right_cols)
		if right_cols(x) = ro then
			right_target = x
		end if
		x = x + 1
	loop
	right_cols_str = join(right_cols,"@@@@@@@###@")
	right_cols_str = replace(right_cols_str,right_cols(right_target)&"@@@@@@@###@","")
	left_cols_str = join(left_cols,"@@@@@@@###@")
	left_cols_str = replace(left_cols_str,left_cols(left_target)&"@@@@@@@###@","")
	new_cols_str = left_cols_str & "@@@@@@@###@" & right_cols_str
	new_cols = split(new_cols_str,"@@@@@@@###@")
	redim merged(-1)
	redim preserve merged(ubound(merged)+1)
	merged(ubound(merged)) = new_cols
	x = 1
	do while x < array_len(left)
		tar = left(x)(left_target)
		left_row_str = join(left(x),"@@@@@@@###@")
		for each row2 in right
			if row2(right_target) = tar then
				right_row_str = join(row2,"@@@@@@@###@")
				right_row_str = replace(right_row_str,tar & "@@@@@@@###@","")
				merged_row_str = left_row_str & "@@@@@@@###@"
				merged_row = split(merged_row_str,"@@@@@@@###@")
				redim preserve merged(ubound(merged)+1)
				merged(ubound(merged)) = merged_row
			end if
		next
		x = x + 1
	loop
	merge = merged
end function
function sum(df,column)
	df = slice(df,array(column))
	total = 0
	for each value in df
		if value(0) <> column then
			if value(0) > 0 then
				on error resume next
					total = total + value(0)
			end if
		end if
	next
	sum = total
end function
function drop_duplicates(df,column)
	cols = df(0)
	x = 0
	do while x < array_len(cols)
		if cols(x) = column then
			target_column = x
			x = array_len(cols)
		end if
		x = x + 1
	loop
	redim arr(-1)
	redim out_arr(-1)
	for each row in df
		if inArray(arr,row(target_column)) = -1 then
			redim preserve arr(ubound(arr)+1)
			arr(ubound(arr)) = row(target_column)
			redim preserve out_arr(ubound(out_arr)+1)
			out_arr(ubound(out_arr)) = row
		end if
	next
	drop_duplicates = out_arr
end function
function to_date(df,column)
	cols = df(0)
	x = 0
	do while x < array_len(cols)
		if cols(x) = column then
			target_column = x
			x = array_len(cols)
		end if
		x = x + 1
	loop
	x = 1
	redim new_arr(-1)
	redim preserve new_arr(ubound(new_arr)+1)
	new_arr(ubound(new_arr)) = cols
	do while x < array_len(df)
		row = df(x)
		date_column = cdate(row(target_column))
		row_str = join(row,"@@@@@@@###@")
		row_str = replace(row_str,row(target_column),date_column)
		row = split(row_str,"@@@@@@@###@")
		redim preserve new_arr(ubound(new_arr)+1)
		new_arr(ubound(new_arr)) = row
		x = x + 1
	loop
	to_date = new_arr
end function
function index_of(df,value)
	redim arr(-1)
	y = 0
	do while x < array_len(df)
		if df(x) = value then
			io = x
			y = 1
		end if
		x = x + 1
	loop
	if y <> 0 then
		index_of = io
	else
		print(value & " not found")
	end if
end function
function group_by_count(df,column)
	redim arr(-1)
	unique_values = unique(df,column)
	for each value in unique_values
		dt = array_len(filter(df,column,value,"="))
		row = array(value,dt)
		redim preserve arr(ubound(arr)+1)
		arr(ubound(arr)) = row
	next
	group_by_count = arr	
end function
function shape(df)
	x = array_len(df(0))
	y = array_len(df)
	shape = array(x,y)
end function
function read_excel(file,sheet)
	currentdir = Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))
	set objExcel = CreateObject("Excel.Application")
	set objWorkbook = objExcel.Workbooks.Open(currentdir & file)
	set s = objWorkbook.Worksheets	
	r = alphabet()
	for each sh in s
		if sh.Name = sheet then
			x = 0
			panel_width = 0
			panel_height = 0
			do while x < array_len(r)
				cell = sh.Range(r(x)&"1")
				if cell  = "" then
					panel_width = x
					x = array_len(r)
				end if
				x = x + 1
			loop
			x = 1
			do while x < 100000
				a1 = sh.Range("A"&x)
				b1 = sh.Range("B"&x)
				c1 = sh.Range("B"&x)
				a2 = sh.Range("A"&x+1)
				b2 = sh.Range("B"&x+1)
				c2 = sh.Range("B"&x+1)		
				if a1 = "" and b2 = "" and c2 = "" and a2 = "" and b2 = "" and c2 = "" then
					panel_height = x-2
					x = 100001
				end if
				x = x +1
			loop
			panel_dimensions = array(panel_width,panel_height)
			x = 1
			redim arr(-1)
			do while x < panel_height
				y = 0
				redim row(-1)
				do while y < panel_width
					redim preserve row(ubound(row)+1)
					row(ubound(row)) = sh.Range(r(y) & x)
					y = y + 1
				loop
				redim preserve arr(ubound(arr)+1)
				arr(ubound(arr)) = row
				x = x + 1
			loop
			read_excel = arr
		end if
	next
end function
df = read_excel("../data\a.xlsx","POWERBALL")
unique_values = unique(df,"Ball 2")
for each value in unique_values
	df = filter(df,"Ball 2",value, "=")
	print(df)
next
' fix the unique values function
