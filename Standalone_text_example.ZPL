! Retrieve the lens file path and add a text file to write to
FULLPATH$ = $PATHNAME() + "\Standalone_output_text.txt"

! Select the text file as the output for the PRINT keyword
OUTPUT FULLPATH$

! Print a dummy statement
PRINT "This is a text written through a macro solve with a ZOS-API standalone application"

! Return a dummy value (mandatory for a macro solve)
SOLVERETURN 0
