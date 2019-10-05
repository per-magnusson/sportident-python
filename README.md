# sportident-python
Python code for talking to Sportident stations. Used e.g. in the sport of orienteering.

The core functionality is in the file sireader2.py, which can be used as a library when
developing other Python programs that interfaces to Sportident stations. 
It is mostly an extended version of sireader.py developed by Gaudenz Steinlin, 
Simon Harston and Jan Vorwerk.

si_read_backup.py is useful when reading out the backup memories of several stations.

si_normalize_station.py is useful when preparing several stations for an event.

Additions and modifications in sireader2.py compared to sireader.py:
- A few more parts of the SYS_VAL structure were worked out and described.
- The format of the data when reading out the backup memory was reverse 
  engineered and documented, both for stations in legacy and extended 
  protocol modes.
- Added function to read out backup memory of stations in legacy and 
  extended protocol modes.
- Added function to save the backup data to a CSV file of the same 
  format as that used by Sportident Config+.
- Added functions to set the station in direct and remote mode.
- A wakeup byte is sent as default before packets to stations. 
  This seems to make the communication more robust.
- Made the serial port search smarter for Windows.
- Fixed an issue with "raise StopIteration" in the _crc routine,
  which no longer works in Python 3.7 (PEP 479).
- Compatibility with Python 2 is probably no longer preserved.
- Added sysval_ functions to access configuration data in SYS_VAL.
  
Per Magnusson, 2019-10-05
