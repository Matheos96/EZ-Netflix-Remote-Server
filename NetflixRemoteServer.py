import win32com.client as comclt, socket, threading,os, traceback

wsh = comclt.Dispatch("WScript.shell")

#Port for the server.
port = 50096

#Setting up ipv4 TCP Socket
s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
s.bind(("", port))
s.listen(1)
print "Server Started on port " + str(port)

#Shuts down the whole server
def shutDown():  
    connect.shutdown(socket.SHUT_RDWR)
    connect.close()
    print ("Closed port on "+adress[0]+"\n")
    os._exit(0)


#Emulates keypresses
def doAction(action):
    print(action)
    wsh.AppActivate("Netflix") #Look for Window with name "Netflix"
    #wsh.AppActivate("Netflix -") #Look for Window with name "Netflix"
    if action == "actionPause":
        wsh.SendKeys(" ")
       #wsh.SendKeys("{ENTER}")
    elif action == "actionFullscreen":
        wsh.SendKeys("{F}")
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

#The Actual Server
def server(connect, adress):
    while True:
        try:
            data = (connect.recv(1024)).strip().decode()

            if data.startswith("action"):
                doAction(data)
            elif data.lower()=="test": #Test statement one can send to server. If ok is received, is working
                answer = "ok\n"
                connect.send(answer.encode())
            elif data.lower()=="shutdown":
                shutDown()
            elif len(data)>0:
                print("Unknown command '"+data+"'")
        except:
            connect.shutdown(socket.SHUT_RDWR)
            connect.close()
            print("Closed port on "+adress[0]+"\n")
            break



#Infinite Loop for running the server using Threads
while True:
    try:
        connect, adress = s.accept()
        print(adress[0]+":")
        thread = threading.Thread(target=server, args=(connect, adress))
        thread.daemon = True
        thread.start()
		
		
        
    except:
        print("Server Crashed. Restarting...")
    
s.close()
print("Server stopped")
