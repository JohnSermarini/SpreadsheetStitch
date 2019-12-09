"""
Name: TBD Cross Stich application 
Desc: ~
Auth: John Sermarini
"""

import sys
from PIL import Image
import os
from openpyxl import styles
from openpyxl import Workbook
from openpyxl import load_workbook
import string
import numpy as np


resize_width = 30
resize_height = 30
num_colors = 16
column_size = 2.8 # This number is about 20 pixels, same as the default height
cell_fill_type = 'solid'
legend_buffer = 1
use_dmc = True


# TODO check out openpyxl.utils.cell.get_column_letter(idx)
# from openpyxl.utils import get_column_letter

# TODO add window for importing files and settings

# TODO add pixelizing to decrease image complexity

# TODO get dmc name

# TODO trim white space in image

# TODO give option for outer buffer in image after trimming

# TODO saturation slider?

# TODO custom symbols/colors

def main(argv):
	# Get file
	input_directory = get_image_directory()
	file_name = get_image_name()

	# Get image from file
	image = read_image(input_directory + file_name)
	image = adjust_image_size(image, resize_width, resize_height)
	image = trim_image(image)

	# Get colors from image
	image = reduce_color_palette(image, num_colors)
	colors, color_map = get_colors(image)
	print(color_map[0])
	if use_dmc:
		colors = convert_colors_to_dmc(colors, color_map)

	# Create worksheet
	wb = Workbook()
	ws = wb.create_sheet(file_name, index=0)

	# Fill worksheet
	#fill_type = 'solid'
	for x in range(0, len(colors)):
		print("Converting - " +  str(x) + "/" + str(len(colors)) + " to Excel")
		for y in range(0, len(colors[x])):
			cell_color = rgb_to_hex(colors[x][y])
			cell_symbol = color_map[x][y]
			cell_alignment = styles.Alignment(horizontal='center')
			cell_fill = styles.PatternFill(fill_type=cell_fill_type, start_color=cell_color, end_color=cell_color)
			cell_border = styles.Border(left=styles.Side(style='thin'), right=styles.Side(style='thin'), top=styles.Side(style='thin'), bottom=styles.Side(style='thin'))
			cell_name = get_cell_name(x, y)
			ws[cell_name].alignment  = cell_alignment
			ws[cell_name].value = cell_symbol
			ws[cell_name].fill = cell_fill
			ws[cell_name].border = cell_border
		ws.column_dimensions[get_column(x + 1)].width = column_size # Set column size
	print("Conversion complete")

	# Add legend
	used_colors, used_map = get_used_color_palette(colors, color_map)
	for c in range(-1, len(used_colors)):
		if(c == -1):
			ws[get_cell_name(image.width + legend_buffer, 0)].value = "Color"
			ws[get_cell_name(image.width + legend_buffer + 1, 0)].value = "DMC Name"			
			ws[get_cell_name(image.width + legend_buffer + 2, 0)].value = "HEX"
			ws[get_cell_name(image.width + legend_buffer + 3, 0)].value = "RGB - R"
			ws[get_cell_name(image.width + legend_buffer + 4, 0)].value = "RGB - G"
			ws[get_cell_name(image.width + legend_buffer + 5, 0)].value = "RGB - B"
			continue		
		color_rgb = used_colors[c]
		color_symbol = used_map[c]
		color_hex = rgb_to_hex(color_rgb)
		ws[get_cell_name(image.width + legend_buffer, c + 1)].fill = styles.PatternFill(fill_type=cell_fill_type, start_color=color_hex, end_color=color_hex)
		ws[get_cell_name(image.width + legend_buffer, c + 1)].value = str(color_symbol)
		ws[get_cell_name(image.width + legend_buffer + 1, c + 1)].value = get_dmc_name(color_rgb)
		ws[get_cell_name(image.width + legend_buffer + 2, c + 1)].value = str(color_hex)
		ws[get_cell_name(image.width + legend_buffer + 3, c + 1)].value = str(color_rgb[0])
		ws[get_cell_name(image.width + legend_buffer + 4, c + 1)].value = str(color_rgb[1])
		ws[get_cell_name(image.width + legend_buffer + 5, c + 1)].value = str(color_rgb[2])

	# Save the file
	output_directory = get_output_directory()
	output_file_name = get_output_file_name(file_name)
	wb.save(output_directory + output_file_name)
	print(output_file_name + " created")


def get_image_directory():
	#TODO
	return "images\\"


def get_image_name():
	#TODO get filename from user

	#return "kirby.png"
	return "fsu.jpg"


def get_output_directory():
	#TODO 
	#TODO make directory if it doesn't exist
	return "out\\"


def get_output_file_name(file_name):
	file_name = file_name[:-4]

	return file_name + ".xlsx"


def read_image(file_name):
	try:
		image = Image.open(file_name)

		return image
	except Exception as e:
		raise e
		#TODO image not found handling

		return None


#############################################
# IN: PIL image
# OUT: 2D array with each value containing a rgb tuple
# OUT: 2D array with each value containing an int representing the color map value
#############################################
def get_colors(image):
	used_colors = []
	colors = []
	color_map = []

	for x in range(0, image.size[0]): # Left column to right column
		column_colors = []
		column_map = []
		for y in range(0, image.size[1]): # Top row to bottom row
			pixel_color = image.getpixel((x,y))
			if(pixel_color not in used_colors):
				used_colors.append(pixel_color)
			pixel_map = used_colors.index(pixel_color)

			column_colors.append(pixel_color)
			column_map.append(pixel_map)
		colors.append(column_colors)
		color_map.append(column_map)

	return colors, color_map


def rgb_to_hex(color):
	# Note: Have color[3] for alpha for future expansion.
	return '%02x%02x%02x' % (color[0], color[1], color[2])
	#return '000000' # all black


def get_cell_name(x, y):
	col = get_column(x + 1)
	row = get_row(y)

	return col + row


def get_column(num):
	# Lifted from : https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
	# All credit to user 'poke'

	def divmod_excel(n):
	    a, b = divmod(n, 26)
	    if b == 0:
	        return a - 1, b + 26
	    return a, b

	chars = []
	while num > 0:
		num, d = divmod_excel(num)
		chars.append(string.ascii_uppercase[d - 1])
	return ''.join(reversed(chars)).upper()


def get_row(y):
	return str(y + 1)


#############################################
# IN:
# OUT: 2D array with each value containing a rgb tuple
#############################################
def reduce_color_palette(image, num_colors):
	#TODO use machine learning to do this instead? Current version kind of jank...

	pixel_image = image.convert("P", palette=Image.ADAPTIVE, colors=num_colors, dither=0)

	return pixel_image.convert("RGB") # convert back to RGB mode


#############################################
# DESC: Convert RGB colors to closest DMC color
# IN:
# OUT: 2D array with each value containing a rgb tuple
#############################################
def convert_colors_to_dmc(colors, color_map):
	#converted_colors = np.full(len(colors), -1, dtype=object)
	converted_colors = []
	for i in range(0, len(color_map)):
		converted_colors.append((-1, -1, -1))

	for x in range(0, len(colors)):
		print("Converting - " +  str(x) + "/" + str(len(colors)) + " to DMC color palette")
		for y in range(0, len(colors[x])):
			map_value = color_map[x][y]
			if(converted_colors[map_value] == (-1, -1, -1)): # converted color not set
				converted_colors[map_value] = find_closest_dmc_color(colors[x][y])
				print("Converted " + str(colors[x][y]) + " to " + str(converted_colors[map_value]))
			colors[x][y] = converted_colors[map_value]

	return colors


def find_closest_dmc_color(color):
	# dmc_colors
	closest_distance = 1000000.0
	closest_index = -1

	dmc_colors = get_dmc_colors()

	for d in range(0, len(dmc_colors)):
		r = color[0] - dmc_colors[d][0]
		g = color[1] - dmc_colors[d][1]
		b = color[2] - dmc_colors[d][2]
		euclidean_distance = np.linalg.norm([r, g, b])
		if(euclidean_distance < closest_distance):
			closest_distance = euclidean_distance
			closest_index = d

	return dmc_colors[closest_index]


def get_dmc_colors():
	#TODO error handling
	ws = load_workbook('color_chart.xlsx').worksheets[0]
	dmc_colors = []
	for row in ws.rows:
		r = row[2].value
		g = row[3].value
		b = row[4].value
		dmc_colors.append((r, g, b))

	return dmc_colors


def adjust_image_size(image, width, height):
	return image.resize((width, height))


def get_used_color_palette(colors, color_map):
	used_colors = []
	used_map = []

	# Get list of used colors
	for x in range(0, len(colors)):
		for y in range(0, len(colors[x])):
			color = colors[x][y]
			color_map_value = color_map[x][y]
			if color not in used_colors:
				used_colors.append(color)
				used_map.append(color_map_value)

	return used_colors, used_map


def get_dmc_name(color_rgb):
	if use_dmc:
		#TODO
		return "TODO"
	else:
		return "N/A"


def trim_image(image):
	# TODO
	# if row or col are entirely transparent... delete row/col
	return image


def get_worksheet_name():
	# TODO
	return "TODO"


if __name__ == "__main__":
	main(sys.argv[1:])