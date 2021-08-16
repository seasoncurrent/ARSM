import logging, numpy, win32api, os , psutil, sys, pyautogui
import threading, time, concurrent.futures
from screeninfo import get_monitors
from PIL import ImageGrab
from functools import partial
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
import pywinauto
'''
controls a shockspot machine with an optional process to tace a command queue
'''

APPDATA = os.getenv('LOCALAPPDATA')
PROCNAME = "Shockspot-Control.exe"
SCRWID = win32api.GetSystemMetrics(0)                               # current screen width
SCRH = win32api.GetSystemMetrics(1)                                 # current screen height
MONITORS = win32api.GetSystemMetrics(80)                            # number of current monitors
LEFTEXTENT = win32api.GetSystemMetrics(76)                          # coordinate of leftmost extent. (negative if monitor to the left of main)
TOTALWIDTH = win32api.GetSystemMetrics(78)                          # total width, e.g. 7680 for 3x 1440p monitors
NEEDLE = os.path.join(os.path.dirname(__file__), 'needle.png')      # picture of the slider thumb to find current location
# try to find the shockspot EXE in appdata
APPPATH = False
for dirpath, dirs, files in os.walk(os.path.join(APPDATA,"Apps","2.0")):
    for filename in files:
        if filename == PROCNAME:
            APPPATH = os.path.join(dirpath,filename)

if not APPPATH:
    logging.error("ERROR: Shockspot EXE not found in %APPDATA%")


# get monitor info
monitors = get_monitors()
# get total monitor coordinate system canvas
xmin, xmax, ymin, ymax = 0,0,0,0
for monitor in monitors:
    xmin = min(xmin,monitor.x)
    ymin = min(ymin,monitor.y)
    xmax = max(xmax,monitor.x+monitor.width)
    ymax = max(ymax,monitor.y+monitor.width)

# margin between edge of slider object and it's maximum/minimum value
SLIDERMARGIN = (14,-14)




# this should really exist in the main file, not implemented here
def processCommands():
    '''take in commands from the queue in dict form'''
    while True:
        command = cq.q.get()
        logging.info(f"received command {command}")
        time.sleep(3)
        logging.info("done")
        cq.q.task_done()

class shockspot:
    def __init__(self, monitor = MONITORS, length = 8, maxDepthValue = 6, toyDim = [7,4,5]):
        # get app PID  or False if unable to find
        self.monitor = monitor # by default use rightmost monitor
        self.maxDepthValue = maxDepthValue
        
        self.PID = self.getPID()
        if self.PID:
            logging.debug(f"existing PID found: {self.PID}")
            self.app = pywinauto.Application().connect(process=self.PID,backend="uia")
        else:
            logging.debug("no PID found, starting app")
            self.app = self.startApp()
        
        self.dlg = self.getDialog(self.app)
        
        self.maxDepth = self.slider(self,"maxDepth", [0,length], 5)
        self.depth = self.slider(self,"depth", [0,length], 5)
        self.stroke = self.slider(self,"stroke", [0,length], 5)
        self.speed = self.slider(self,"speed", [0,1], 10)
        self.roughness = self.slider(self,"roughness", [0,1], 10)
        
        self.maxDepth.setValue(self.maxDepthValue)
        # self.toy = toy(*toyDim,self.depth.getValue(),self.stroke.getValue())    
    
    def getDialog(self,app):
        dlg = app.top_window()
        dlg.move_window(SCRWID + 50,50,1820,980)
        dlg.set_focus()
        return dlg
    
    def startApp(self):
        app = pywinauto.Application().start(APPPATH)
        app.top_window()['I have read and agree with the terms above'].click()
        return app
        
    def restart(self):
        # kill & relaunch if problem detected
        self.dlg.close()
        self.app = self.startApp()
        self.dlg = self.getDialog(self.app)
        self.maxDepth.setValue(self.maxDepthValue)
        # todo: set all sliders to previous value... need to keep some sort of state
        
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
            self.bounds = (self.r.left-xmin,self.r.top-ymin,self.r.width(),self.r.height())
            
            self.getValue()
        
        def getLocation(self):
            pywinauto.mouse.move((100,100))
            loc = pyautogui.locateCenterOnScreen(NEEDLE,
                confidence=0.6,
                region=self.bounds,
                grayscale=True
            )
            if loc:
                return xmin+loc[0],ymin+loc[1]
            else:
                return False
                
        def getValue(self):
            # return real value (not pixel value)
            self.loc = self.getLocation()
            logging.debug(f"slider getLocation {self.name} result: {self.loc}")
            if self.loc:
                self.lastValue = numpy.interp(self.loc[0], (self.r.left+SLIDERMARGIN[0],self.r.right-SLIDERMARGIN[1]), self.valMap)
                return self.lastValue
            else:
                return False
            
        def setValue(self, val):
            logging.debug(f"setValue for {self.name} to {val}")
            if self.getValue() is not False:
                desiredPos = int(numpy.interp(val, self.valMap, (self.r.left+SLIDERMARGIN[0],self.r.right+SLIDERMARGIN[1])))
                dist = abs(desiredPos - self.loc[0])
                attempts = 1
                while dist >= self.minMove and attempts <= 5:
                    logging.debug(f"attempt {attempts}, dist: {dist}")
                    pyautogui.moveTo(self.loc)
                    pywinauto.mouse.press(coords=(self.loc[0],self.loc[1]))
                    pywinauto.mouse.release(coords=(desiredPos,self.loc[1]))
                    self.getValue()
                    dist = abs(desiredPos - self.loc[0])
                    attempts += 1
                if dist <= self.minMove:
                    logging.debug("already at desired position, skipped")
            else:
                logging.error("setValue abort due to getValue false")
        
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
    format = "%(asctime)s: %(message)s"
    logging.basicConfig(format=format, level=logging.DEBUG,
                        datefmt="%H:%M:%S")
    m = shockspot()
    