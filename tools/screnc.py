#!/usr/bin/python3
import base64
import struct
import argparse

PICK_ENCODING = (
    1, 2, 0, 1, 2, 0, 2, 0,
    0, 2, 0, 2, 1, 0, 2, 0, 
    1, 0, 2, 0, 1, 1, 2, 0,
    0, 2, 1, 0, 2, 0, 0, 2, 
    1, 1, 0, 2, 0, 2, 0, 1, 
    0, 1, 1, 2, 0, 1, 0, 2, 
    1, 0, 2, 0, 1, 1, 2, 0, 
    0, 1, 1, 2, 0, 1, 0, 2
)

ESCAPE_CHARS = {
	"\r": "#",
	"\n": "&",
	"<":  "!",
	">":  "*",
	"@": "$",
}

ENCODING_TABLE = {
    '\t': (0x64,0x37,0x69), ' ' : (0x50,0x7e,0x2c), '!' : (0x22,0x5a,0x65),
    '"' : (0x4a,0x45,0x72), '#' : (0x61,0x3a,0x5b), '$' : (0x5e,0x79,0x66),
    '%' : (0x5d,0x59,0x75), '&' : (0x5b,0x27,0x4c), "'" : (0x42,0x76,0x45),
    '(' : (0x60,0x63,0x76), ')' : (0x23,0x62,0x2a), '*' : (0x65,0x4d,0x43),
    '+' : (0x5f,0x51,0x33), ',' : (0x7e,0x53,0x42), '-' : (0x4f,0x52,0x20),
    '.' : (0x52,0x20,0x63), '/' : (0x7a,0x26,0x4a), '0' : (0x21,0x54,0x5a),
    '1' : (0x46,0x71,0x38), '2' : (0x20,0x2b,0x79), '3' : (0x26,0x66,0x32),
    '4' : (0x63,0x2a,0x57), '5' : (0x2a,0x58,0x6c), '6' : (0x76,0x7f,0x2b),
    '7' : (0x47,0x7b,0x46), '8' : (0x25,0x30,0x52), '9' : (0x2c,0x31,0x4f),
    ':' : (0x29,0x6c,0x3d), ';' : (0x69,0x49,0x70), '<' : (0x3f,0x3f,0x3f),
    '=' : (0x27,0x78,0x7b), '>' : (0x3f,0x3f,0x3f), '?' : (0x67,0x5f,0x51),
    '@' : (0x3f,0x3f,0x3f), 'A' : (0x62,0x29,0x7a), 'B' : (0x41,0x24,0x7e),
    'C' : (0x5a,0x2f,0x3b), 'D' : (0x66,0x39,0x47), 'E' : (0x32,0x33,0x41),
    'F' : (0x73,0x6f,0x77), 'G' : (0x4d,0x21,0x56), 'H' : (0x43,0x75,0x5f),
    'I' : (0x71,0x28,0x26), 'J' : (0x39,0x42,0x78), 'K' : (0x7c,0x46,0x6e),
    'L' : (0x53,0x4a,0x64), 'M' : (0x48,0x5c,0x74), 'N' : (0x31,0x48,0x67),
    'O' : (0x72,0x36,0x7d), 'P' : (0x6e,0x4b,0x68), 'Q' : (0x70,0x7d,0x35),
    'R' : (0x49,0x5d,0x22), 'S' : (0x3f,0x6a,0x55), 'T' : (0x4b,0x50,0x3a),
    'U' : (0x6a,0x69,0x60), 'V' : (0x2e,0x23,0x6a), 'W' : (0x7f,0x09,0x71),
    'X' : (0x28,0x70,0x6f), 'Y' : (0x35,0x65,0x49), 'Z' : (0x7d,0x74,0x5c),
    '[' : (0x24,0x2c,0x5d), '\\': (0x2d,0x77,0x27), ']' : (0x54,0x44,0x59),
    '^' : (0x37,0x3f,0x25), '_' : (0x7b,0x6d,0x7c), '`' : (0x3d,0x7c,0x23),
    'a' : (0x6c,0x43,0x6d), 'b' : (0x34,0x38,0x28), 'c' : (0x6d,0x5e,0x31),
    'd' : (0x4e,0x5b,0x39), 'e' : (0x2b,0x6e,0x7f), 'f' : (0x30,0x57,0x36),
    'g' : (0x6f,0x4c,0x54), 'h' : (0x74,0x34,0x34), 'i' : (0x6b,0x72,0x62),
    'j' : (0x4c,0x25,0x4e), 'k' : (0x33,0x56,0x30), 'l' : (0x56,0x73,0x5e),
    'm' : (0x3a,0x68,0x73), 'n' : (0x78,0x55,0x09), 'o' : (0x57,0x47,0x4b),
    'p' : (0x77,0x32,0x61), 'q' : (0x3b,0x35,0x24), 'r' : (0x44,0x2e,0x4d),
    's' : (0x2f,0x64,0x6b), 't' : (0x59,0x4f,0x44), 'u' : (0x45,0x3b,0x21),
    'v' : (0x5c,0x2d,0x37), 'w' : (0x68,0x41,0x53), 'x' : (0x36,0x61,0x58),
    'y' : (0x58,0x7a,0x48), 'z' : (0x79,0x22,0x2e), '{' : (0x09,0x60,0x50),
    '|' : (0x75,0x6b,0x2d), '}' : (0x38,0x4e,0x29), '~' : (0x55,0x3d,0x3f)
}

def main():
	parser = argparse.ArgumentParser(description='Encodes scripts encoded by Windows Script Editor')
	parser.add_argument("-i", "--input", help='Input file')
	parser.add_argument("-o", "--output", help='Output file')

	args = parser.parse_args()

	f = open(args.input)
	data = f.read()
	f.close()

	out = ""
	chk = 0
	for i in range(len(data)):
		c = data[i]

		chk += ord(c)

		if c in ESCAPE_CHARS.keys():
			out += "@" + ESCAPE_CHARS[c]
		else:
			out += chr(ENCODING_TABLE[c][PICK_ENCODING[i%64]])

	size = struct.pack("<I", len(out))
	chk = struct.pack("<I", chk)
	out = "#@~^" + base64.b64encode(size).decode() + out + base64.b64encode(chk).decode() + "^#~@"

	f = open(args.output, "w")
	f.write(out)
	f.close()


main()


