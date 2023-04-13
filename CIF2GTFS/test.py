import time
import wx

class Mywin(wx.Frame): 
            
   def __init__(self, parent, title): 
      super(Mywin, self).__init__(parent, title = title,size = (300,200))  
      self.InitUI() 
         
   def InitUI(self):    
      self.count = 0 
      pnl = wx.Panel(self)
		
      self.gauge = wx.Gauge(pnl, range = 20, size = (300, 25), style =  wx.GA_HORIZONTAL) 
         
      self.SetSize((300, 100)) 
      self.Centre() 
      self.Show(True)
				
ex = wx.App() 
prog = Mywin(None, 'wx.Gauge')

for x in range(20):
    time.sleep(1)
    prog.gauge.SetValue(x)