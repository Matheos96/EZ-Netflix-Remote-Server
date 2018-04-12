import win32com.client as comclt, socket, threading,os, traceback

wsh = comclt.Dispatch("WScript.shell")

port = 50096

s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
s.bind(("", port))
s.listen(1)


def shutDown():  
    connect.shutdown(socket.SHUT_RDWR)
    connect.close()
    print ("Closed port on "+adress[0]+"\n")
    os._exit(0)


def doAction(action):
    print action
    
    wsh.AppActivate("Netflix")
    if action == "actionPause":
       wsh.SendKeys("{ENTER}")
    elif action == "actionFullscreen":
        wsh.SendKeys("{F}")
    elif action == "actionNoFullScreen":
        wsh.SendKeys("{Esc}")
    elif action == "actionForward":
        wsh.SendKeys("{RIGHT}")
    elif action == "actionBack":
        wsh.SendKeys("{LEFT}")
    elif action == "actionVolUp":
        wsh.SendKeys("{UP}")
    elif action == "actionVolDown":
        wsh.SendKeys("{DOWN}")
    elif action == "actionMute":
        wsh.SendKeys("{M}")

def server(connect, adress):
    while True:
        try:
            data = (connect.recv(1024)).strip().decode()

            if data.startswith("action"):
                doAction(data)
            elif data.lower()=="test":
                answer = "ok\n"
                connect.send(answer.encode())
            elif data.lower()=="exit":
                raise Exception("Haha")
            elif data.lower()=="shutdown":
                shutDown()
            elif len(data)>0:
                print "Unknown command '"+data+"'"
        except:
            connect.shutdown(socket.SHUT_RDWR)
            connect.close()
            print ("Closed port on "+adress[0]+"\n")
            break

while True:
    try:
        
        connect, adress = s.accept()
        print adress[0]+" connected"
        thread = threading.Thread(target=server, args=(connect, adress))
        thread.daemon = True
        thread.start()
        print "TJO"
        
    except:
        print "Server Crashed. Restarting.."
    
s.close()
print "Server stopped"
