import os
import win32api
import win32com.client
import win32con
import win32gui
import win32process
import socket

OFFICE_HOST = 'dwadcookie.mooo.com'
PORT = 5555
TIMEOUT = 30


class check_kick(object):
    d3_hwnd = None
    
    def __init__(self):
        pass

    @classmethod
    def kick(cls):
        """Send command to d3 to start macro"""
        cls.d3_hwnd = win32gui.FindWindow(None,'Diablo III')
        #print(cls.d3_hwnd)
        if cls.d3_hwnd:
            _,pid = win32process.GetWindowThreadProcessId(cls.d3_hwnd)
            #os.kill(pid,9)
            os.system("taskkill /F /pid %s" %(pid))

    @classmethod  
    def run(cls):
        """Main loop""" 
        count = 0
        while True:
            try:
                # Connect to server
                client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                client_socket.connect((OFFICE_HOST, PORT))
                print("Connected")
            except:
                print("Couldn't connect")
                time.sleep(TIMEOUT)
                continue
            
            while True:
                try:
                    data = client_socket.recv(1024).strip()
                    print(data)
                    if data.lower() == b'kick':
                        print("Kicking")
                        cls.kick()
                except:
                    break
        

          
if __name__ == '__main__':        
  check_kick.run()


