Ecograph OPC Server v0.11

overview:
---------
Classic OPC Server on Ecograph. Included 486 tags (122 rVal tags).
Include all data,information and settings tags. 
Working with device with number 01-99. 
Allow include only needed tags, by configurate type and number of channal in
config file. Support configuration from .ini file. Server based on lightopc v0.88

installation notes:
-------------------
After install software you must register server with key /r. You also may unregistered server with key /u.
For example opc.exe /r or opc.exe /u.
Show help key /?.

version history:
----------------
v0.11 build 59
0 fixed bug with show incorrect status on all tags if last tag in list is uncertain.
+ add remaining settings tags.
+ add data channal tags.
+ add "Integrator" option in config file, include/exclude integration data tags.
+ add "Digital" option in config file, include/exclude digital channal data tags.
+ add "Analog" option in config file, include/exclude analog channal data tags.
0 fixed unknown bug, when server wrong count tags.

v0.10 build 43
+ add ability to switch intellect from .ini file.
+ add digital input settings tag.
+ add intellect system on digital inputs tags.
+ add ability switch on/off rVal tags.
+ add new section .ini file "Server".
+ add ability write comment to .ini file
= total device scanned on bus increased to 99.

v0.09 build 103
- remove not used temporary data equation.
+ add intelegence analyze analog data system. check analog inputs only when
inputs has been used (active). 
+ add prioritet system. each tag now have a attribute named "priority", which
show frequency request this tag. maximum priority is 1, minimum 65000.
+ significantly increase frequency of processing tag.
+ add new 280 tags on another 5 channal.
+ add ability to configurate connection speed from .ini file.

v0.08 build 118
+ quantity of tags increased to 112 (included 24 rVal and 5 unknown).
= fixed possible bug with memory overflow on find token in rVal tags.
0 fixed bug when opc quality is good, when answer on command do not recieve.
0 fixed bug opc quality is still good, when device is temporary not answer.
= possible quantity tags increased to 500.
0 fixed bug on show bad status on tags recivied on version commands.
0 fixed bug with wrong access rights on some tags. 

v0.07 build 119
+ add new type of tags (rVal) - real value tags. readable tags is data from
correspondly tags throw enumerated table with links between real value and protocol
data.
+ add full support rVal tags. add and register tags. read and form tag data.
+ add 18 new basic settings tags and correspondly 8 rVal tags.

v0.06 build 57
+ added 11 settings tag recieved on various read command.
+ add support variant type unsigned int;
+ add intelligent conversion and round from type (unsigned) speed on server
to enumerated string in device;
+ add intelligent conversion from type (string) unit feed on server
to enumerated string in device;
0 fixed wrong "write success" text in output log;
0 fixed bug with device data buffer overflow;
0 fixed wrong "add command" text in output log;
0 fixed bug with use wrong buffer on "no scan" command;
+ add new log events on error in read, write and version command;
0 fixed possible problem in define success result on recieve answer from device;
+ add tag hierarchy;

v0.05 build 36
+ added 12 no information tag recieved on version command.
+ global change write function.
+ added device identificator tag (read/write tag).

[tag list]:
----------
Information/Programme
Information/Version
Information/CPU number
Settings/Identificator
Settings/Current date
Settings/Current time
Settings/Date WT>ST
Settings/Time WT>ST
Settings/Date ST>WT
Settings/Time ST>WT
Settings/Access Code
Settings/Feed Unit
Settings/Normal feed speed
Settings/Active feed speed
Settings/Channal Identificator
Settings/Channal Identificator (rVal)
Settings/Group Identificator

Settings/Switch time
Settings/Switch time (rVal)
Settings/Temperature unit
Settings/Temperature unit (rVal)

Settings/View/Grid Lines
Settings/View/Grid Lines (rVal)
Settings/View/Pen
Settings/View/Pen (rVal)
Settings/View/Show Pen
Settings/View/Show Pen (rVal)

Settings/Floppy/Warning
Settings/Floppy/Relay
Settings/Floppy/Relay (rVal)
Settings/Floppy/Quota
Settings/Floppy/Quota (rVal)

Settings/Screensaver/Screensaver
Settings/Screensaver/On time
Settings/Screensaver/Off time

Other/Serial interface/Device address
Other/Serial interface/Interface type
Other/Serial interface/Interface type (rVal)
Other/Serial interface/Speed
Other/Serial interface/Speed (rVal)
Other/Serial interface/Parity
Other/Serial interface/Parity (rVal)
Other/Serial interface/Stop bits
Other/Serial interface/Stop bits (rVal)
Other/Serial interface/Data bits
Other/Serial interface/Data bits (rVal)
Other/Memory/Work regim
Other/Memory/Work regim (rVal)

Service/Common/Software version
Service/Common/Last switch on
Service/Common/Last C-instruction
Service/Common/Setup
Service/Common/Show address
Service/Common/CPU number
Service/Common/Work time total
Service/Common/Work time LCD

Service/Running Costs/Money unit
Service/Running Costs/Metre cost
Service/Running Costs/Pen cost
Service/Running Costs/Reset

Information/Module board 1
Information/Module board 2
Information/Digital IO
Information/RS485
Information/RS485-Profibus
Information/Data memory
Information/Internal memory
Information/Integration
Information/Digital board 1
Information/Digital board 2
Information/Math channel
Information/Data interface

tags for analog input when x [1...6]
------------------------------------
Analog channal x
Analog channal x (intermediate)
Analog channal x (daily)
Analog channal x (monthly)
Analog channal x (yearly)

Analog input/x/Type
Analog input/x/Type (rVal)
Analog input/x/Channal ID
Analog input/x/Measure Unit
Analog input/x/Decimal point
Analog input/x/Decimal point (rVal)
Analog input/x/Begin diapason
Analog input/x/End diapason
Analog input/x/Begin undiapason
Analog input/x/End undiapason
Analog input/x/Damping filter
Analog input/x/Termo compensation value
Analog input/x/Termo compensation
Analog input/x/Termo compensation (rVal)
Analog input/x/Copy settings
Analog input/x/Registration type
Analog input/x/Registration type (rVal)
Analog input/x/Wire type
Analog input/x/Wire type (rVal)
Analog input/x/Diapason indetion/Begin
Analog input/x/Diapason indetion/End
Analog input/x/Integration/Integration
Analog input/x/Integration/Integration (rVal)
Analog input/x/Integration/Time basis
Analog input/x/Integration/Time basis (rVal)
Analog input/x/Integration/Integration unit
Analog input/x/Integration/Show sequence
Analog input/x/Integration/Show sequence (rVal)
Analog input/x/Integration/Maximum value
Analog input/x/Integration/Coefficient

Analog input/x/Valuepoint 1/Valuepoint type
Analog input/x/Valuepoint 1/Valuepoint type (rVal)
Analog input/x/Valuepoint 1/Value
Analog input/x/Valuepoint 1/Hysteresis
Analog input/x/Valuepoint 1/Pause time
Analog input/x/Valuepoint 1/Relay on
Analog input/x/Valuepoint 1/Relay on (rVal)
Analog input/x/Valuepoint 1/Switching on message
Analog input/x/Valuepoint 1/Switching off message
Analog input/x/Valuepoint 1/Show message
Analog input/x/Valuepoint 1/Show message (rVal)
Analog input/x/Valuepoint 1/Feeding change
Analog input/x/Valuepoint 1/Feeding change (rVal)

Analog input/x/Valuepoint 2/Valuepoint type
Analog input/x/Valuepoint 2/Valuepoint type (rVal)
Analog input/x/Valuepoint 2/Value
Analog input/x/Valuepoint 2/Hysteresis
Analog input/x/Valuepoint 2/Pause time
Analog input/x/Valuepoint 2/Relay on
Analog input/x/Valuepoint 2/Relay on (rVal)
Analog input/x/Valuepoint 2/Switching on message
Analog input/x/Valuepoint 2/Switching off message
Analog input/x/Valuepoint 2/Show message
Analog input/x/Valuepoint 2/Show message (rVal)
Analog input/x/Valuepoint 2/Feeding change
Analog input/x/Valuepoint 2/Feeding change (rVal)

tags for digital input when x [1...4] 
-------------------------------------
Digital channal x

Digital input/x/Input function
Digital input/x/Input function (rVal)
Digital input/x/Identificator
Digital input/x/Action
Digital input/x/Action (rVal)
Digital input/x/Label log 1
Digital input/x/Label log 0
Digital input/x/Message on 0->1
Digital input/x/Message on 1->0
Digital input/x/Message text
Digital input/x/Message text (rVal)
