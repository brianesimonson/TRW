"""Shared legacy address and source helpers; no Office COM dependencies."""
import re
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries


def vb_code_only(text):
    lines=[]
    for line in text.splitlines():
        quoted=False;out=[];i=0
        while i<len(line):
            c=line[i]
            if c=='"':
                if quoted and i+1<len(line) and line[i+1]=='"':
                    out.extend('  ');i+=2;continue
                quoted=not quoted;out.append(' ')
            elif c=="'" and not quoted:
                break
            else:
                out.append(' ' if quoted else c)
            i+=1
        if not re.match(r'^\s*Rem\b', ''.join(out),re.I):
            lines.append(''.join(out))
    return '\n'.join(lines)


def column_number(label):
    return column_index_from_string(label.upper())


def column_label(index):
    return get_column_letter(index)


def range_coordinates(address):
    return range_boundaries(address)


def range_moves(address):
    c1,r1,c2,r2=range_boundaries(address)
    return c1,r1,c2-c1,r2-r1
