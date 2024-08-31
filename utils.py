def col_to_ind(column:str, start:int=0) -> int:
	''' Converts a column name (e.g. 'A', 'AF', 'CK') to an index '''
	index = start - 1
	for i, letter in enumerate(column[::-1]): # Run through reversed column name
		index += (ord(letter) - ord('A') + 1) * 26**i
	# 'ABC'  ->  'CBA'  ->  'A'*(26^2) + 'B'*(26^1) + 'C'*(26^0)  ->  1*676 + 2*26 + 3*1  ->  731
	return index
