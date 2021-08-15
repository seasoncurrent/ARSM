import logging, threading, time, concurrent.futures, psutil, sys, numpy, win32api, os
import pygame as pg
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
import pyautogui, pywinauto
import commandQueue as cq
from PIL import ImageGrab
from functools import partial
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
'''
a threaded machinecontrol program
'''

APPDATA = os.getenv('LOCALAPPDATA')
APPPATH = APPDATA + r"\Apps\2.0\C75B96K0.H2D\HN3DO2PP.HKG\shoc..tion_92aab741509f25fe_0005.0000_2eefa70506f6a3c9\Shockspot-Control.exe"
PROCNAME = "Shockspot-Control.exe"
SCRWID = win32api.GetSystemMetrics(0)

def processCommands():
    '''take in commands from the queue in dict form'''
    while True:
        command = cq.q.get()
        logging.info(f"received command {command}")
        time.sleep(3)
        logging.info("done")
        cq.q.task_done()

class mControl:
    def __init__(self, toyDim = [7,4,5]):
        # get app PID  or False if unable to find
        self.PID = self.getPID()
        if self.PID:
            logging.debug(f"existing PID found: {self.PID}")
            self.app = pywinauto.Application().connect(process=self.PID,backend="uia")
        else:
            logging.debug("no PID found, starting app")
            self.app = self.startApp()
        
        self.dlg = self.app.top_window()
        self.dlg.move_window(SCRWID + 50,50,1820,980)
        self.dlg.set_focus()
        
        self.maxDepth = self.slider(self,"maxDepth", [0,8], 5)
        self.depth = self.slider(self,"depth", [0,8], 5)
        self.stroke = self.slider(self,"stroke", [0,8], 5)
        self.speed = self.slider(self,"speed", [0,1], 10)
        self.roughness = self.slider(self,"roughness", [0,1], 10)
        
        # self.toy = toy(*toyDim,self.depth.getValue(),self.stroke.getValue())    
    
    def startApp(self):
        app = pywinauto.Application().start(APPPATH)
        app.top_window()['I have read and agree with the terms above'].click()
        return app
        
    def restartMachine(self):
        pass
        # kill & relaunch if problem detected
        
    def getPID(self):
        PID = False
        for proc in psutil.process_iter(['pid','name']):
            if proc.info['name'] == PROCNAME:
                PID = proc.info['pid']
        return PID
    
    class slider:
        """ an individual slider with pixel & real values """
        sliderDict = {
            "maxDepth"  : "Max_Pos_Slider"
            ,"depth"     : "Pos_Slider1"
            ,"stroke"    : "pullback_pos_slider"
            ,"speed"     : "Vel_Slider2"
            ,"roughness" : "Accel_Slider3"
        }
        def __init__(self,parent,name,valMap,minMove):
            self.name = name
            self.valMap = valMap
            self.minMove = minMove
            self.parent = parent
            self.r = self.parent.dlg.child_window(auto_id=self.sliderDict[name], control_type="System.Windows.Forms.TrackBar").rectangle()
            self.bounds = (self.r.left,self.r.top,self.r.width(),self.r.height())
            
            self.getValue()
        
        def getLocation(self):
            pywinauto.mouse.move((100,100))
            return pyautogui.locateCenterOnScreen('needle.png',
                confidence=0.6,
                region=self.bounds,
                grayscale=False
            )
        
        def getValue(self):
            # return real value (not pixel value)
            self.loc = self.getLocation()
            logging.debug(f"slider getLocation {self.name} result: {self.loc}")
            if self.loc:
                self.lastValue = numpy.interp(self.loc[0], (self.r.left,self.r.right), self.valMap)
                return self.lastValue
            else:
                return False
            
        def setValue(self, val):
            logging.debug(f"setValue for {self.name} to {val}")
            if self.getValue() is not False:
                desiredPos = int(numpy.interp(val, self.valMap, (self.r.left,self.r.right)))
                dist = abs(desiredPos - self.loc[0])
                attempts = 1
                while dist >= self.minMove and attempts <= 5:
                    logging.debug("attempt",attempts,"dist:",dist)
                    pyautogui.moveTo(self.loc)
                    pywinauto.mouse.press(coords=(self.loc[0],self.loc[1]))
                    pywinauto.mouse.release(coords=(desiredPos,self.loc[1]))
                    self.getValue()
                    dist = abs(desiredPos - self.loc[0])
                    attempts += 1
                if dist <= self.minMove:
                    logging.debug("already at desired position, skipped")
        
        
class toy:
    def __init__(self, length, knotStart, knotMid, knotDepth, initialStroke):
        # physical characteristics
        self.length = length
        self.knotStart = knotStart
        self.knotMid = knotMid
        
        # adjusted for machine reading
        self.knotDepth = knotDepth # true knot start depth
        self.strokeOffset = max(0,knotDepth - initialStroke) # stroke must always be this much less than depth (may be 0)
        self.midDepth = knotDepth + knotMid - knotStart # true depth setting for knot midpoint
        self.totalDepth = knotDepth + length - knotStart # true max depth / length
        self.progressDepth = knotDepth # for incrementing til midDepth
    
    def increment(self, percent):
        # move this % towards knot, max out at knot mid 
        self.progressDepth = min(self.progressDepth + (self.midDepth - self.knotDepth) * percent/100, self.midDepth)        
    

if __name__ == "__main__":
    m = mControl()
    