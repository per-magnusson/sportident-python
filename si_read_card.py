#!/usr/bin/env python3
#
#    Copyright (C)    2019  Per Magnusson <per.magnusson@gmail.com>
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""
Script to read out the data from an SI card.

To select a specific serial port, provide it's name as the first command line
parameter to the program.
"""

from sireader2 import SIReader, SIReaderException, SIReaderControl, SIReaderReadout
from time import sleep
from datetime import datetime
import sys


try:
    if len(sys.argv) > 1:
        # Use command line argument as serial port name
        si = SIReaderReadout(port = sys.argv[1])
    else:
        # Find serial port automatically
        si = SIReaderReadout()
    print('Connected to station on port ' + si.port)
except:
    print('Failed to connect to an SI station on any of the available serial ports.')
    exit()
    


print('Insert SI-card to be read')
# wait for a card to be inserted into the reader
while not si.poll_sicard():
    sleep(1)

# some properties are now set
card_number = si.sicard
card_type = si.cardtype

# read out card data
card_data = si.read_sicard()

# beep
si.ack_sicard()

print('Number: ' + str(card_number))
print('Type:   ' + card_type)
print('Data:')
for key, val in card_data.items():
    if key == 'punches':
        for p in val:
            code = p[0]
            tim = p[1]
            print(str(code) + ' ' + tim.isoformat(sep = ' '))
    else:
        print(key + ' : ', end='')
        print(val)


