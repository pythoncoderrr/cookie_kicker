import copy
import json
import os
import time
import win32api
import win32com.client
import win32con
import win32gui
import win32process

from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from threading import RLock

index = """
<html>
  <head>
  </head>
  <body>
    Hi. Welcome to Ronaldinho Kicker&trade; (Patent Pending).<p>
    Unforunately, you did not arrive here with your special link.<p>
    Please use your super special link that was given to you.
  </body>
</html>
"""

kick_main = """
<html>
  <head>
    <style>
      .column{
        float: left;
        width: 15%%;
        padding: 0px;
      }
      .row:after{
        content: "";
        display: table;
        clear: both;
      }
    </style>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <script>
      $(document).ready(function() {
        // Load history
        function load_history() {
          $('#history').empty();
          var str = '';
          $.ajax({url: '/history',
                  success: function(data) {
                    $.each(data, function(i,line){
                      str = '<div class="row">';
                      $.each(line, function(i,e){
                        str += '<div class="column">'+e+'</div>';
                     }); // end each
                      str += '</div>';
                      $('#history').append(str);
                    }); // end each
                  }
          });  // ajax close
        };  // load_history close
        
        // Ajax call on clicking icon
        $('#kick_img').on("click",function(e){
          $.ajax({url: '%s/kick',
                  success: function() {
                    $('#kickable').hide();
                    $('#not_kickable').show();
                    load_history();
                  } // success function close
          });  // ajax close
        }); // click close

        // Hide kicked div
        $('#not_kickable').hide();
        
        // Load the history
        load_history();
      });  // document close
    </script>
  </head>
  <body>
    Hi %s.<p>
    Hi. Welcome to Ronaldinho Kicker&trade; (Patent Pending).<p>
    <div id='kickable'>
      Click
        <img id="kick_img" src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT3PGrRWbH0bropL9c_XHiEGs9DSLJRuOo5rg&usqp=CAU" width="15" height="15" />
        to kick.<p>
    </div>
    <div id = 'not_kickable'>
      You have already kicked.<p>
    </div>
    History (20 most recent)<p>
    <div id="history">
    </div>
  </body>
</html>
"""

error_page = """
<html>
  <head>
  </head>
  <body>
    Uhh what?
  </body>
</html>
"""


class KickingManager(object):
    DELAY_BETWEEN_KICKS = 5
    USERS = ['ambagalord','bloodgaylord','dwad','nomi','s2k','korngalord']
    history = []
    last_kick = 0
    kicking_lock = RLock()
    d3_hwnd = None
    shell = None

    def __init__(self):
        pass
    
    @classmethod
    def add(cls,l):
        cls.history.append(l)
        
    @classmethod
    def kick(cls,user):
        try:
            ctime = time.time()
            if ctime > cls.last_kick + cls.DELAY_BETWEEN_KICKS:
                cls.kicking_lock.acquire()
                cls.last_kick = ctime
                # Sending command to kick
                cls.send()
                # Log it
                cls.add([user,ctime])
                cls.last_kick = ctime
                cls.kicking_lock.release()
                return True
        except Exception as e:
            print(e)
        return False
    

    @classmethod
    def send(cls):
        try:
            cls.d3_hwnd = win32gui.FindWindow(None,'Diablo III')
            print(cls.d3_hwnd)
            if cls.d3_hwnd:
                _,pid = win32process.GetWindowThreadProcessId(cls.d3_hwnd)
                #os.kill(pid,9)
                os.system("taskkill /F /pid %s" %(pid))
            #print(cls.d3_hwnd)
            #cls.shell = win32com.client.Dispatch("WScript.Shell")
            #if cls.shell.AppActivate('Diablo III'):
                #win32gui.SetForegroundWindow(cls.d3_hwnd)
                #win32api.keybd_event(0x32, 0,0,0) # 2 key
                #time.sleep(.01)
                #win32api.keybd_event(0x32,0 ,win32con.KEYEVENTF_KEYUP ,0)
                #win32api.keybd_event(VK_CODE['F9'], 0,0,0) # F9 key
                #time.sleep(.01)
                #win32api.keybd_event(VK_CODE['F9'],0 ,win32con.KEYEVENTF_KEYUP ,0)
                #cls.shell.SendKeys("{F9}",0)
        except Exception as e:
            print(e)
        return

    @classmethod
    def check(cls,u):
        if u in cls.USERS:
            return True
        else:
            return False
        
    @classmethod
    def recent_history(cls):
        try:
            now = datetime.now()
            recent = copy.deepcopy(cls.history)
            recent.reverse()
            recent = recent[0:20]
            converted = []
            for user,t in recent:
                dt_t = datetime.fromtimestamp(t)
                now = datetime.now()
                difference = now-dt_t
                time_string = time.strftime("%m/%d/%Y, %H:%M:%S", time.localtime(t))
                converted.append([user,time_string,cls.time_ago(difference)])
            return json.dumps(converted)
        except Exception as e:
            print(e)
    
    @classmethod
    def time_ago(cls,difference):
        if difference.total_seconds() < 3600:
            return '%s minutes ago' %(int(difference.total_seconds()/60))
        elif difference.total_seconds() < 3600*2:
            return '%s hour ago' %(int(difference.total_seconds()/3600))
        elif difference.total_seconds() < 3600*24:
            return '%s hours ago' %(int(difference.total_seconds()/3600))
        elif difference.total_seconds() < 3600*24*2:
            return '%s day ago' %(int(difference.total_seconds()/(3600*24)))
        else:
            return '%s days ago' %(int(difference.total_seconds()/(3600*24)))
        
    
class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            self.send_response(200)
            
            print(self.path)
            if self.path == '/':
                self.send_header('Content-type','text/html')
                self.end_headers()            
                self.wfile.write(bytes(index, "utf8"))
                return
            elif self.path == '/history':
                self.send_header('Content-type','application/json')
                self.end_headers()
                self.wfile.write(bytes(KickingManager.recent_history(), "utf8"))
                return

            path_split = self.path.split('/')
            while '' in path_split:
                path_split.remove('')

            if len(path_split) == 1:
                current_user = path_split[0]
                if KickingManager.check(current_user):
                    print(current_user)
                    self.send_header('Content-type','text/html')
                    self.end_headers() 
                    self.wfile.write(bytes(kick_main%(current_user,current_user), "utf8"))
                    return
            elif len(path_split) == 2:
                if 'kick' in path_split:
                    path_split.remove('kick')
                    kicker = path_split[0]
                    # Kick it and log it
                    if KickingManager.kick(kicker):
                        self.send_header('Content-type','text/html')
                        self.end_headers()
                        return
        except Exception as e:
            print(e)
            
        # Error
        self.send_header('Content-type','text/html')
        self.end_headers() 
        self.wfile.write(bytes(error_page, "utf8"))
        return
            
    def do_POST(self):
        self.send_response(200)
        self.send_header('Content-type','text/html')
        self.end_headers()

        message = "Whattt? Posts arent allowed."
        self.wfile.write(bytes(message, "utf8"))

def main():
    while True:
        print('Starting')
        try:
            with HTTPServer(('0.0.0.0', 8000), handler) as server:
                server.serve_forever()
        except Exception as e:
            print(e)

if __name__ == '__main__':
    main()
