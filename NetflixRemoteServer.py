import win32com.client as comclt, socket, threading, os, sys

wsh = comclt.Dispatch("WScript.shell")

window_name = "Netflix -"

if len(sys.argv) >= 2:
    if sys.argv[1].lower() == "app":
        window_name = "Netflix"
        print "Started in Netflix app mode"
else:
    print "Started in Netflix Browser mode"

# Port for the server.
port = 50096

# Setting up ipv4 TCP Socket
s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
s.bind(("", port))
s.listen(1)
print "Server Started on port " + str(port)


# Shuts down the whole server
def shut_down():
    connect.shutdown(socket.SHUT_RDWR)
    connect.close()
    print ("Closed port on " + address[0] + "\n")
    os._exit(0)


# Emulates keypresses
def do_action(action):
    print(action)
    wsh.AppActivate(window_name)  # Look for Window with name "Netflix"
    # wsh.AppActivate("Netflix -")  # Look for Window with name "Netflix"
    if action == "actionPause":
        wsh.SendKeys(" ")
    elif action == "actionFullscreen" and window_name == "Netflix -":
        wsh.SendKeys("{F}")
    elif action == "actionForward":
        wsh.SendKeys("{RIGHT}")
    elif action == "actionBack":
        wsh.SendKeys("{LEFT}")
    elif action == "actionVolUp" and window_name == "Netflix -":
        wsh.SendKeys("{UP}")
    elif action == "actionVolDown" and window_name == "Netflix -":
        wsh.SendKeys("{DOWN}")
    elif action == "actionMute" and window_name == "Netflix -":
        wsh.SendKeys("{M}")


# The Actual Server
def server(connection, addrs):
    while True:
        try:
            data = (connection.recv(1024)).strip().decode()

            if data.startswith("action"):
                do_action(data)
            elif data.lower() == "test":  # Test statement one can send to server. If ok is received, is working
                answer = "ok\n"
                connection.send(answer.encode())
            elif data.lower() == "shutdown":
                shut_down()
            elif len(data) > 0:
                print("Unknown command '" + data + "'")
        except:
            connection.shutdown(socket.SHUT_RDWR)
            connection.close()
            print("Closed port on " + addrs[0] + "\n")
            break


# Infinite Loop for running the server using Threads
while True:
    try:
        connect, address = s.accept()
        print(address[0] + ":")
        thread = threading.Thread(target=server, args=(connect, address))
        thread.daemon = True
        thread.start()



    except:
        print("Server Crashed. Restarting...")

s.close()
print("Server stopped")
