import tkinter as tk
from tkinter import scrolledtext, messagebox, Menu, ttk
import socket
import threading
import time
from PIL import Image, ImageTk
import openpyxl

class SocketGUI:
    def __init__(self, master):
        
        #============================窗口設置=================================
        #====================================================================
        self.master = master
        master.title("Socket GUI")
        master.geometry("1480x820")  
        master.resizable(False, False)  
        bg_image = Image.open("C:\\Users\\A0592\\MainGUI_IMG\\bg2.jpg")
        self.bg_photo = ImageTk.PhotoImage(bg_image)
        self.canvas = tk.Canvas(master, width=bg_image.width, height=bg_image.height)
        self.canvas.pack()
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.bg_photo)
        
        self.label = tk.Label(master, text="Server IP:")
        self.label.place(relx=0.01, rely=0.01, relheight=0.03)
        self.server_ip_entry = tk.Entry(master)
        self.server_ip_entry.insert(tk.END, "192.168.10.182")  
        self.server_ip_entry.place(relx=0.06, rely=0.01, relwidth=0.1, relheight=0.03)

        self.label2 = tk.Label(master, text="Port:")
        self.label2.place(relx=0.01, rely=0.06, relheight=0.03)
        self.port_entry = tk.Entry(master)
        self.port_entry.insert(tk.END, "1300")  
        self.port_entry.place(relx=0.06, rely=0.06, relwidth=0.1, relheight=0.03)

        self.connect_button = tk.Button(master, text="Connect", command=self.connect)
        self.connect_button.place(relx=0.2, rely=0.01, relwidth=0.1, relheight=0.03)

        self.disconnect_button = tk.Button(master, text="Disconnect", command=self.disconnect, state=tk.DISABLED)
        self.disconnect_button.place(relx=0.2, rely=0.06, relwidth=0.1, relheight=0.03)

        self.remote_button = tk.Button(master, text="Remote Mode", command=self.remote_mode, state=tk.DISABLED)
        self.remote_button.place(relx=0.2, rely=0.12, relwidth=0.1, relheight=0.03)
        
        self.local_button = tk.Button(master, text="Local Mode", command=self.local_mode, state=tk.NORMAL)
        self.local_button.place(relx=0.2, rely=0.18, relwidth=0.1, relheight=0.03)

        self.status_label = tk.Label(master, text="Status: Disconnected", fg="red")
        self.status_label.place(relx=0.07, rely=0.12, relheight=0.03)

        self.command_label = tk.Label(master, text="Command:")
        self.command_label.place(relx=0.01, rely=0.18)
        self.command_area = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=40, height=5)
        self.command_area.place(relx=0.01, rely=0.22, relwidth=0.18,relheight=0.1)

        self.send_label = tk.Label(master, text="Send History:")
        self.send_label.place(relx=0.35, rely=0.01)
        self.send_area = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=40, height=15)
        self.send_area.place(relx=0.35, rely=0.06, relwidth=0.22)

        self.recv_label = tk.Label(master, text="Receive History:")
        self.recv_label.place(relx=0.60, rely=0.01)
        self.recv_area = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=40, height=15)
        self.recv_area.place(relx=0.60, rely=0.06, relwidth=0.22)

        self.send_button = tk.Button(master, text="Send", command=self.send, state=tk.DISABLED)
        self.send_button.place(relx=0.2, rely=0.22, relwidth=0.1,relheight=0.1)

        self.clear_send_button = tk.Button(master, text="Clear Send", command=self.clear_send)
        self.clear_send_button.place(relx=0.5, rely=0.32, relwidth=0.1,relheight=0.1)

        self.clear_recv_button = tk.Button(master, text="Clear Receive", command=self.clear_receive)
        self.clear_recv_button.place(relx=0.75, rely=0.32, relwidth=0.1,relheight=0.1)

        self.sock = None
        self.start_time = None

        #============================選單建立=================================
        #====================================================================
        self.menu = Menu(master)
        master.config(menu=self.menu)
        self.function_menu = Menu(self.menu, tearoff=0)
        self.function_menu.add_command(label="Continuous Command", command=self.show_continuous_command_window)
        
        self.function_menu.add_command(label="Export to Excel", command=self.export_to_excel) 
        self.menu.add_cascade(label="Function", menu=self.function_menu)

        self.option_menu = Menu(self.menu, tearoff=0)
        self.option_menu.add_command(label="RS232", command=self.open_rs232_window)
        self.menu.add_cascade(label="Option", menu=self.option_menu)

    def connect(self):
        server_ip = self.server_ip_entry.get()
        port = int(self.port_entry.get())

        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            self.sock.connect((server_ip, port))
            self.status_label.config(text="Status: Connected", fg="green")
            self.connect_button.config(state=tk.DISABLED) 
            self.disconnect_button.config(state=tk.NORMAL)  
            self.remote_button.config(state=tk.NORMAL)  
            self.send_button.config(state=tk.NORMAL)  
            threading.Thread(target=self.receive).start()
        except Exception as e:
            self.status_label.config(text="Status: Connection failed", fg="red")
            print("Connection failed:", e)

    def disconnect(self):
        if self.sock:
            self.sock.close()
            self.status_label.config(text="Status: Disconnected", fg="red")
            self.connect_button.config(state=tk.NORMAL)  
            self.disconnect_button.config(state=tk.DISABLED)  
            self.remote_button.config(state=tk.DISABLED)  
            self.send_button.config(state=tk.DISABLED)  
            self.sock = None

    def remote_mode(self):
        if self.sock:
            self.send_message("SYST:REM\n")
        else:
            self.command_area.insert(tk.END, "Not connected\n")

    def local_mode(self):
        if self.sock:
            self.send_message("SYST:LOC\n")
        else:
            self.command_area.insert(tk.END, "Not connected\n")

    def show_continuous_command_window(self):
        continuous_command_window = tk.Toplevel(self.master)
        continuous_command_window.title("Continuous Command")
        tk.Label(continuous_command_window, text="Interval (ms):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.interval_entry = tk.Entry(continuous_command_window)
        self.interval_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(continuous_command_window, text="Times:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.times_entry = tk.Entry(continuous_command_window)
        self.times_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(continuous_command_window, text="Number of Commands:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.num_commands_var = tk.StringVar()  
        num_commands_dropdown = ttk.Combobox(continuous_command_window, textvariable=self.num_commands_var, state="readonly")
        num_commands_dropdown["values"] = tuple(range(1, 21))  
        num_commands_dropdown.current(0)  
        num_commands_dropdown.grid(row=2, column=1, padx=5, pady=5)

        #============================命令輸入=================================
        #====================================================================
        
        
        for i in range(20): 
            tk.Label(continuous_command_window, text=f"Command {i+1}:").grid(row=3+i, column=0, padx=5, pady=5, sticky="w")
            setattr(self, f"command_entry_{i+1}", tk.Entry(continuous_command_window))
            getattr(self, f"command_entry_{i+1}").grid(row=3+i, column=1, padx=5, pady=5)
            
        tk.Button(continuous_command_window, text="Send", command=self.send_multiple).grid(row=2, column=3, columnspan=3, padx=5, pady=5) 

    def open_rs232_window(self):
        rs232_window = tk.Toplevel(self.master)
        rs232_window.title("RS232 Communication")
        tk.Label(rs232_window, text="RS232 Settings").pack()

    def receive(self):
        while True:
            if self.sock:
                try:
                    data = self.sock.recv(1024)
                    if not data:
                        break
                    decoded_data = data.decode()
                    recv_time = time.time()
                    elapsed_time = recv_time - self.start_time
                    self.recv_area.insert(tk.END, "Received: {} (Time: {:.2f}ms)\n".format(decoded_data, elapsed_time*1000))
                    self.recv_area.yview(tk.END)  
                except Exception as e:
                    print("Error receiving message:", e)
                    self.status_label.config(text="Status: Disconnected", fg="red")
                    self.connect_button.config(state=tk.NORMAL)  
                    self.disconnect_button.config(state=tk.DISABLED)  
                    self.remote_button.config(state=tk.DISABLED)  
                    self.send_button.config(state=tk.DISABLED)  
                    break

    def send(self):
        if self.sock:
            command = self.command_area.get("1.0", tk.END)
            self.start_time = time.time()  
            self.send_message(command)
            self.command_area.delete("1.0", tk.END)
        else:
            self.command_area.insert(tk.END, "Not connected\n")

    def send_message(self, command):
        if self.sock:
            try:
                self.sock.sendall(command.encode())
                self.send_area.insert(tk.END, "Sent: {}\n".format(command))
                self.send_area.yview(tk.END)  
            except Exception as e:
                print("Error sending message:", e)
                self.status_label.config(text="Status: Disconnected", fg="red")
                self.connect_button.config(state=tk.NORMAL)  
                self.disconnect_button.config(state=tk.DISABLED)  
                self.remote_button.config(state=tk.DISABLED)  
                self.send_button.config(state=tk.DISABLED)  

    def send_multiple(self):
        interval = int(self.interval_entry.get())
        times = int(self.times_entry.get())
        num_commands = int(self.num_commands_var.get())  
        commands = [getattr(self, f"command_entry_{j+1}").get() + '\n' for j in range(num_commands)]    
        
        threading.Thread(target=self.receive).start()   
        
        for _ in range(times):
            self.start_time = time.time()  
            for command in commands:
                if command:
                    #print("Sending command:", command)  
                    self.send_message(command) 
                    time.sleep(interval / 1000)

                    
                    
    #============================Excel 輸出=================================
    #====================================================================                
    def export_to_excel(self):
        wb = openpyxl.Workbook()
        
        send_ws = wb.active
        send_ws.title = "Send History"
        send_ws.append(["Send Time", "Message"])
        for line in self.send_area.get("1.0", tk.END).splitlines():
            send_ws.append([time.strftime("%Y-%m-%d %H:%M:%S"), line])
        recv_ws = wb.create_sheet(title="Receive History")
        recv_ws.append(["Receive Time", "Message"])

        for line in self.recv_area.get("1.0", tk.END).splitlines():
            recv_ws.append([time.strftime("%Y-%m-%d %H:%M:%S"), line])
        excel_file = "Socket_History.xlsx"
        wb.save(excel_file)
        messagebox.showinfo("Export Successful", f"Socket history exported to {excel_file}")

    def clear_send(self):
        self.send_area.delete("1.0", tk.END)

    def clear_receive(self):
        self.recv_area.delete("1.0", tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = SocketGUI(root)
    root.mainloop()
