# ARSM

machinecontrol.py
this library is designed to control the shockspot application in a multi-monitor setup. The application is launched or detected then moved to the monitor to the right of the main one.

create a mControl object using m = mControl()
Then the sliders can be controlled with:
m.depth.setValue(5)

depth sliders have values 0-8
speed / smoothness values 0-1
