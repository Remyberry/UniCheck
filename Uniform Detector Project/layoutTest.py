import customtkinter as tk
from ttkbootstrap import Style
import cv2 
from PIL import Image, ImageTk 

# Create the main window
root = tk.CTk()
root.title("Live Camera Feed")
root.geometry("1920x1080")  # Set the window size
root.attributes('-fullscreen', True)

# Create a frame to hold the grid layout
mainframe = tk.CTkFrame(root)
mainframe.pack(fill="both", expand=True)  # Make the frame fill the entire window
# Configure the grid layout
for col in range(8): 
    mainframe.columnconfigure(col, weight=1, uniform=True)
    for row in range(8):
        mainframe.rowconfigure(row, weight=1, uniform=True)
        control_label = tk.CTkFrame(mainframe, border_width=1, border_color="#000000")
        control_label.grid(row=row, column=col, sticky="nsew")
    

# Create and place widgets in the grid

# left_frame = tk.CTkFrame(mainframe)
# left_frame.grid(row=0, column=0, rowspan=8, columnspan=6, sticky="nsew")

#Canvas is where Camera 1 outputs
canvas_1 = tk.CTkCanvas(mainframe)
canvas_1.grid(row=0, column=0, rowspan=6, columnspan=6, sticky="nsew")

# #Canvas_2 is where Camera 2 outputs
canvas_2 = tk.CTkCanvas(mainframe)
canvas_2.grid(row=0, column=6, rowspan=2, sticky="nsew")

# #Canvas_2 is where snapshot outputs
canvas_3 = tk.CTkCanvas(mainframe)
canvas_3.grid(row=0, column=7, rowspan=2, sticky="nsew")


# # for i in range(6):
# #     control_label.columnconfigure(i, weight=1)
# # control_label.rowconfigure(0, weight=1)
# # control_label.rowconfigure(1, weight=1)

button_open_1 = tk.CTkButton(mainframe, text="Open Camera 1")
button_open_1.grid(column=0, row=6, padx=10, pady=10)

button_close_1 = tk.CTkButton(mainframe, text="Close Camera 1")
button_close_1.grid(column=0, row=7, padx=10, pady=10)

button_open_2 = tk.CTkButton(mainframe, text="Open Camera 2")
button_open_2.grid(column=1, row=6, padx=10, pady=10)

button_close_2 = tk.CTkButton(mainframe, text="Close Camera 2")
button_close_2.grid(column=1, row=7, padx=10, pady=10)

text_field = tk.CTkLabel(mainframe, text="Text Ouput here")
text_field.grid(column=2, row=6, columnspan=2, rowspan=2, sticky="nsew")

button_snaps = tk.CTkButton(mainframe, text="View Recent Snaps")
button_snaps.grid(column=4, row=6, padx=10, pady=10)

button_logs = tk.CTkButton(mainframe, text="View System Logs")
button_logs.grid(column=4, row=7, padx=10, pady=10)

button_clr = tk.CTkButton(mainframe, text="Clear Fields")
button_clr.grid(column=5, row=6, padx=10, pady=10)

button_generate_ticket = tk.CTkButton(mainframe, text="Generate Ticket")
button_generate_ticket.grid(column=5, row=7, padx=10, pady=10)

student_info_field = tk.CTkLabel(mainframe, text="Student Info")
student_info_field.grid(column=6, row=2, columnspan=2, rowspan=6, sticky="nsew")


# # Global variable to store the PhotoImage object reference
# image_id_1 = canvas_1.create_image(0, 0, anchor="nw")
# # Global variable to store the PhotoImage object reference
# image_id_2 = canvas_2.create_image(0, 0, anchor="nw")


# # Load an image and display it on the canvas
# snapshot = ImageTk.PhotoImage(Image.open("qwe.jpg"))
# canvas_3.create_image(0, 0, anchor="nw", image=snapshot)
root.mainloop()
