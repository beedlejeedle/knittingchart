PURPOSE
In the first version, there are two functions:
1. generate a colorwork knitting chart given an image
2. generate a knitting chart given Instructions

FUTURE WORK
I hope to work on adding parsing ability so that it can directly parse a pattern to an Instruction.

USE
In this version, you have to initialize all of the Instructions, then you can generate a chart from them. You will need PIL and openpyxl.

INFORMATION
Class Colorwork_chart (img, style = "intarsia")
Parameters: img : Image
	    width : int
	    length: int
	    style : str
		stranded/fair isle or intarsia
	    pixel_map : 2d list
		Each item is a hex value
Methods: img_to_map() : from Image generate pixel map
	 export_to_excel(): generate chart to Excel
	 calc_yarn_chart(needlesize): calculates all yarn needed from pixel_map given needle size in mm

Class Yarn (name, material, yardage, price, weightperball, weight='aran', gauge=None)
Parameters: name, weight, material : str
	    yardage, price, weightperball : int
	    gauge: tuple of (row, column). Default gauge based on weight unless specified.

Class Instruction (input_str="", num_repeat=1, sts="", side="RS", name='NA', dec_alt=1)
Parameters: input_str : str
		Is just the original instruction from the pattern
	    num_repeat : int
		How many times to repeat
	    sts : str
		Stitches for this Instruction separated by space, i.e. "K1 SSK PAT". The PAT
		signifies pattern
	    side : str
		"RS" or "WS" or "NA" if in the round
	    name : str
	    dec_alt : int
		If the instruction is to be performed on alternate rows.
	    dec_alt_working : int
		For chart() to keep track of alternating Instructions
	    connected : Instruction
		If another Instruction is to be performed at the same time
	    next : Instruction
		The next Instruction.

chart(first_inst, total_st = 10) : create a chart given the beginning Instruction

