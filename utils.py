def col_to_ind(column:str, start:int=0) -> int:
	''' Converts a column name (e.g. 'A', 'AF', 'CK') to an index '''
	index = start - 1
	for i, letter in enumerate(column[::-1]): # Run through reversed column name
		index += (ord(letter) - ord('A') + 1) * 26**i
	# 'ABC'  ->  'CBA'  ->  'A'*(26^2) + 'B'*(26^1) + 'C'*(26^0)  ->  1*676 + 2*26 + 3*1  ->  731
	return index


def try_save_wb(workbook, title, filename):
	while True:
		try:
			workbook.save(filename)
		except Exception as e:
			print(f"[‼] Failed to save {title}. See the error below and close the \"{filename}\" file if it is open.")
			print(type(e), e)
			print("\nPlease press Enter to try again or enter other filename:")
			new_filename = input("» ")
			if new_filename:
				filename = new_filename
		else:
			print(f"[¤] Successfully saved {title} into \"{filename}\"")
			break
