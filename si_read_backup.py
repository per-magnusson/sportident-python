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
Script to efficiently read out the backup memories of one or more Sportident stations.
The script automatically connects to one of the stations connected to the computer.
To select a specific serial port, provide it's name as the first command line
parameter to the program:

si_read_backup.py COM4
"""

from sireader2 import SIReader, SIReaderException
import sys


try:
    if len(sys.argv) > 1:
        # Use command line argument as serial port name
        si = SIReader(port = sys.argv[1])
    else:
        # Find serial port automatically
        si = SIReader()
    print('Connected to station on port ' + si.port)
except:
    print('Failed to connect to an SI station on any of the available serial ports.')
    exit()
    

# Set station in remote mode
ok = False
errmsg = ''
for ii in range(0,3):
    try:
        si.set_remote()
        ok = True
        break;
    except SIReaderException as msg:
        errmsg = msg
if not ok:
    print('ERROR: Failed to set station in remote mode: %s' % errmsg)
    exit()


print('Ready to read backup memory of SI station.')
maxretries = 5
while True:
    inp = input('    Press <Enter> to read remote station, d to read direct station or q to quit: ')
    if inp == 'q':
        break
    elif inp == 'd':
        if not si.direct:
            si.set_direct()
    elif inp == '':
        if si.direct:
            si.set_remote()
    else:
        print('    Unrecognized input')
        continue
        
    ok = False
    errmsg = ''
    for ii in range(0, maxretries):
        try:
            si._update_proto_config()
            ok = True
            break;
        except SIReaderException as msg:
            errmsg = msg
    if not ok:
        print('ERROR: Failed to talk to the station: %s' % errmsg)
        print('Maybe the station is not connected or not awake?')
        continue
        
    if not si.proto_config['mode'] in si.SUPPORTED_READ_BACKUP_MODES:
        print("ERROR: Station is in mode %s, which is not supported for backup readout" % 
              si.MODE2NAME[si.proto_config['mode']])
        continue

    ok = False
    for ii in range(0, maxretries):
        try:
            print('    Trying to read backup memory of station: ' + str(si._station_code) + ' ', end='')
            sys.stdout.flush()
            backup = si.read_backup(progress=1)
            csvfilename = si.write_backup_csv(backup)
            print(csvfilename + ' was created')
            ok = True
            si.beep()
            break
        except SIReaderException as msg:
            print('')
            errmsg = msg
    if not ok:
        print('ERROR: Failed to talk to the station: %s' % errmsg)
        print('Maybe the station is not connected, not awake or not in a supported mode?')
        continue

