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
si_normalize_station.py
Script to set up remote stations to "standard" values, typically to 
prepare them to be used at an event.

The script automatically connects to one of the stations connected to the computer.
To select a specific serial port, provide it's name as the first command line
parameter to the program:

si_normalize_station.py COM4

The following settings are made:
- Clear the backup memory.
- Set the active time.
- Set the time to that of the computer.
- Disable "SI-Card with 192 punches", except for clear stations.
- Enable optical feedback.
- Enable audible feedback.
- Disable autosend on readout stations.
- Turn the station off.

The progress is saved to a CSV file. Each station that is processed 
gets two lines in the file; one line with information about the state
before the changes and a second line with information about the state
after the changes. 
"""

from sireader2 import SIReader, SIReaderException, SIReaderControl, SIReaderReadout
from time import sleep
from datetime import datetime, timedelta
import sys
import csv


#####################################################
# Edit this if another active time is desired.
active_minutes = 4*60
#####################################################


def get_station_status(si):
    """Read out some status information from an SI station.
    @param si: SIReader object
    @return:   A list with data suitable for writing to a CSV file
    """
    now_time = datetime.now()
    si_time = si.get_time()
    time_delta = si_time-now_time
    ms = timedelta(microseconds=1000)
    if time_delta >= timedelta(0):
        time_delta_str = str(time_delta)
    else:
        time_delta = -time_delta
        time_delta_str = '-' + str(time_delta)
    time_delta_str = time_delta_str[0:-3]  # Truncate to ms
    si.refresh_sysval()

    serno = si.sysval_serno()
    model = si.sysval_model_str()
    build_date = si.sysval_build_date()
    fwver = si.sysval_fwver()
    mem_size = si.sysval_mem_size()
    mem_size_str = "%d K" % mem_size
    bat_date = si.sysval_battery_date()
    bat_cap = si.sysval_battery_capacity()
    bat_cap_str = "%d" % int(round(bat_cap, 0))
    bat_use = si.sysval_used_battery()
    bat_use_str = "%4.1f %%" % bat_use
    volt = si.sysval_volt()
    volt_str = "%4.2f" % volt
    code = si.sysval_code()
    mode = si.sysval_mode_str()
    active_time = si.sysval_active_time()
    active_hours = active_time//60
    active_minutes = active_time%60
    active_str = "%02d:%02d:00" % (active_hours, active_minutes)
    protocol = si.sysval_protocol()
    si6_192 = si.sysval_192_punches()
    feedback = si.sysval_feedback()

    autosend_str = '-'
    if protocol & 0b10:
        autosend_str = 'Autosend'

    legacy_str = 'LegacyProtocol!'
    if protocol & 0b1:
        legacy_str = '-'

    if si6_192 is True:
        si6_192_str = 'Card6With192Records!'
    elif si6_192 is False:
        si6_192_str = '-'
    else:
        si6_192_str = '0x%02x' % si6_192
        
    optical_str = ''
    if feedback & 0b1:
        optical_str = 'OpticalSignal1'
    audible_str = ''
    if feedback & 0b100:
        audible_str = 'AcousticSignal'

    date = now_time.strftime("%Y-%m-%d")
    time = now_time.strftime("%H:%M:%S")
    
    row = [date, time, serno, model, fwver, bat_date, bat_use_str, volt_str,
           code, mode, time_delta_str, active_str, autosend_str, legacy_str,
           si6_192_str, audible_str, optical_str, build_date, mem_size_str,
           bat_cap_str ]

    return row


try:
    if len(sys.argv) > 1:
        # Use command line argument as serial port name
        si = SIReader(port = sys.argv[1])
    else:
        # Find serial port automatically
        si = SIReader()
    print('Connected to station on port ' + si.port)
except SIReaderException as e:
    print('ERROR: ' + str(e))
    print('Failed to connect to an SI station on any of the available serial ports.')
    exit()
except Exception as e:
    print('Error: ' + str(e))
    exit()
    

# Set station in remote mode
ok = False
errmsg = ''
for ii in range(0,3):
    try:
        si.set_direct()
        sleep(.1)
        si.set_baud_rate_38400()
        sleep(.1)
        si.set_remote()
        ok = True
        break;
    except SIReaderException as msg:
        print(str(msg))
        errmsg = msg
if not ok:
    print('ERROR: Failed to set station in remote mode: %s' % errmsg)
    exit()


# Create a csv log file 
datestr = datetime.now().strftime("%Y-%m-%d_%H.%M.%S")
csv_filename = 'Log_SI_normalize_' + datestr + '.csv'
with open(csv_filename, 'w', newline='') as csvfile:
    csvwriter = csv.writer(csvfile, delimiter=';',
                           quotechar='"', quoting=csv.QUOTE_MINIMAL)
    # This structure is similar to the log from SI Config+, but lacks some fields,
    # has some fields in a different order (to put more interesting information 
    # more to the left on the long rows) and has an extra TimeDiff field.
    header = ['Date', 'Time', 'SerialNo', 'Hardware', 'Software', 'BatteryDate',
              'BattUsage', 'Voltage', 'CodeNo', 'Mode', 'TimeDiff', 'OpTime', 'Autosend',
              'LegacyProtocol', 'Card6with192punches', 'AcousticSignal', 
              'OpticalSignal1', 'ProductionDate', 'MemorySize', 'BatteryCapacity']
    csvwriter.writerow(header)

    # Loop over stations
    print('Ready to normalize remote station.')
    maxretries = 5
    while True:
        inp = input('Press <Enter> to process remote station. Press q to quit: ')
        if inp == 'q':
            break
        elif inp == '':
            pass
        else:
            print('    Unrecognized input')
            continue

        # Try a few times if it does not work the first time
        ok = False
        errmsg = ''
        for ii in range(0, maxretries):
            try:
                # Read status information from the station
                row = get_station_status(si)
                #row = [date, time, serno, model, fwver, bat_date, bat_use_str, volt_str,
                #       code, mode, time_delta_str, active_str, autosend_str, legacy_str,
                #       si6_192_str, audible_str, optical_str, build_date, mem_size_str,
                #       bat_cap_str ]
                fwver = row[4]
                volt = float(row[7])
                code = row[8]
                mode = row[9]
                
                # Write old status to CSV file

                csvwriter.writerow(row)

                print('code: %3d, volt: %4.2f V, mode: %s' % (code, volt, mode))
                if volt < 3.1:
                    print('WARNING: VERY low battery: %4.2f V' % volt)
                    si.beep(2)
                elif volt < 3.2:
                    print('Warning: low battery: %4.2f V' % volt)
                    si.beep
                if fwver < "656":
                    print('WARNING: Old firmware: %s' % fwver)
                    si.beep

    
                # Normalize the station's settings
                si.set_time(datetime.now())
                si.erase_backup()
                si.set_feedback(True, True)
                si.set_active_time(active_minutes)
                if mode == "Clear":
                    si.set_si6_192(True)
                else:
                    si.set_si6_192(False)
                if mode == "Readout":
                    si.set_autosend(False)

                # Reead the station's updated settings and write it to the CSV file
                row = get_station_status(si)
                csvwriter.writerow(row)

                # Turn the remote station off
                si.poweroff()
                
                ok = True
                break;
            except SIReaderException as msg:
                print(str(msg))
        if not ok:
            print('ERROR: Failed to talk to the station: %s' % errmsg)
            print('Maybe the station is not connected or not awake?')
            continue

print('Log file is: %s' % csv_filename)
