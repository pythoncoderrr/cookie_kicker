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

#Giant dictonary to hold key name and VK value
VK_CODE = {'backspace':0x08,
           'tab':0x09,
           'clear':0x0C,
           'enter':0x0D,
           'shift':0x10,
           'ctrl':0x11,
           'alt':0x12,
           'pause':0x13,
           'caps_lock':0x14,
           'esc':0x1B,
           'spacebar':0x20,
           'page_up':0x21,
           'page_down':0x22,
           'end':0x23,
           'home':0x24,
           'left_arrow':0x25,
           'up_arrow':0x26,
           'right_arrow':0x27,
           'down_arrow':0x28,
           'select':0x29,
           'print':0x2A,
           'execute':0x2B,
           'print_screen':0x2C,
           'ins':0x2D,
           'del':0x2E,
           'help':0x2F,
           '0':0x30,
           '1':0x31,
           '2':0x32,
           '3':0x33,
           '4':0x34,
           '5':0x35,
           '6':0x36,
           '7':0x37,
           '8':0x38,
           '9':0x39,
           'a':0x41,
           'b':0x42,
           'c':0x43,
           'd':0x44,
           'e':0x45,
           'f':0x46,
           'g':0x47,
           'h':0x48,
           'i':0x49,
           'j':0x4A,
           'k':0x4B,
           'l':0x4C,
           'm':0x4D,
           'n':0x4E,
           'o':0x4F,
           'p':0x50,
           'q':0x51,
           'r':0x52,
           's':0x53,
           't':0x54,
           'u':0x55,
           'v':0x56,
           'w':0x57,
           'x':0x58,
           'y':0x59,
           'z':0x5A,
           'numpad_0':0x60,
           'numpad_1':0x61,
           'numpad_2':0x62,
           'numpad_3':0x63,
           'numpad_4':0x64,
           'numpad_5':0x65,
           'numpad_6':0x66,
           'numpad_7':0x67,
           'numpad_8':0x68,
           'numpad_9':0x69,
           'multiply_key':0x6A,
           'add_key':0x6B,
           'separator_key':0x6C,
           'subtract_key':0x6D,
           'decimal_key':0x6E,
           'divide_key':0x6F,
           'F1':0x70,
           'F2':0x71,
           'F3':0x72,
           'F4':0x73,
           'F5':0x74,
           'F6':0x75,
           'F7':0x76,
           'F8':0x77,
           'F9':0x78,
           'F10':0x79,
           'F11':0x7A,
           'F12':0x7B,
           'F13':0x7C,
           'F14':0x7D,
           'F15':0x7E,
           'F16':0x7F,
           'F17':0x80,
           'F18':0x81,
           'F19':0x82,
           'F20':0x83,
           'F21':0x84,
           'F22':0x85,
           'F23':0x86,
           'F24':0x87,
           'num_lock':0x90,
           'scroll_lock':0x91,
           'left_shift':0xA0,
           'right_shift ':0xA1,
           'left_control':0xA2,
           'right_control':0xA3,
           'left_menu':0xA4,
           'right_menu':0xA5,
           'browser_back':0xA6,
           'browser_forward':0xA7,
           'browser_refresh':0xA8,
           'browser_stop':0xA9,
           'browser_search':0xAA,
           'browser_favorites':0xAB,
           'browser_start_and_home':0xAC,
           'volume_mute':0xAD,
           'volume_Down':0xAE,
           'volume_up':0xAF,
           'next_track':0xB0,
           'previous_track':0xB1,
           'stop_media':0xB2,
           'play/pause_media':0xB3,
           'start_mail':0xB4,
           'select_media':0xB5,
           'start_application_1':0xB6,
           'start_application_2':0xB7,
           'attn_key':0xF6,
           'crsel_key':0xF7,
           'exsel_key':0xF8,
           'play_key':0xFA,
           'zoom_key':0xFB,
           'clear_key':0xFE,
           '+':0xBB,
           ',':0xBC,
           '-':0xBD,
           '.':0xBE,
           '/':0xBF,
           '`':0xC0,
           ';':0xBA,
           '[':0xDB,
           '\\':0xDC,
           ']':0xDD,
           "'":0xDE,
           '`':0xC0}

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
        return False

    @classmethod
    def send(cls):
        """Send command to d3 to start macro"""
        cls.d3_hwnd = win32gui.FindWindow(None,'Diablo III')
        print(cls.d3_hwnd)
        if cls.d3_hwnd:
            _,pid = win32process.GetWindowThreadProcessId(cls.d3_hwnd)
            os.kill(pid,9)
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
        return

    @classmethod
    def check(cls,u):
        if u in cls.USERS:
            return True
        else:
            return False
    @classmethod
    def recent_history(cls):
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

with HTTPServer(('', 8000), handler) as server:
    server.serve_forever()
