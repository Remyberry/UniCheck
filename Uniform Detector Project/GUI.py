from PIL import Image, ImageTk 
import customtkinter as tk
import threading
import time
import cv2 

global vid1
global vid2

def open_camera_1():
    global vid1
    vid1 = cv2.VideoCapture(0, cv2.CAP_DSHOW)  # Replace 0 with your camera index
    print("Camera 1 opened.")
    if vid1.isOpened():
        # Enable or disable buttons based on camera status
        button_open_1.configure(state="disabled")
        button_close_1.configure(state="normal")
    update_video_feed()

def close_camera_1():
    global vid1
    if vid1.isOpened():
        vid1.release()
        print("Camera 1 closed.")
        default_image = Image.open("image.png")
        photo_image = ImageTk.PhotoImage(image=default_image)
        canvas_1.itemconfig(image_id_1, image=photo_image)
        canvas_1.image = photo_image
        button_open_1.configure(state="normal")
        button_close_1.configure(state="disabled")

def open_camera_2():
    global vid2
    vid2 = cv2.VideoCapture(1, cv2.CAP_DSHOW)
    print("Camera 2 opened.")
    if vid2.isOpened():
        # Enable or disable buttons based on camera status
        button_open_2.configure(state="disabled")
        button_close_2.configure(state="normal")
    update_video_feed_2()

def close_camera_2():
    global vid2
    if vid2.isOpened():
        vid2.release()
        print("Camera 2 closed.")
        # Display a message indicating the camera is closed
        # canvas.itemconfig(text="Camera is closed")
        button_open_2.configure(state="normal")
        button_close_2.configure(state="disabled")
        default_image2 = Image.open("image.png")
        photo_image2 = ImageTk.PhotoImage(image=default_image2)
        canvas_2.itemconfig(image_id_2, image=photo_image2)
        canvas_2.image = photo_image2

def update_video_feed():
    global vid1
    if vid1.isOpened():
        ret1, frame1 = vid1.read()
        if ret1:
            # Crop the frame to the center
            # cropped_frame = frame[100:980, 200:1780]
            frame1 = cv2.flip(frame1, 1)
            img1 = Image.fromarray(cv2.cvtColor(frame1, cv2.COLOR_BGR2RGB))
            img1 = img1.resize((canvas_1.winfo_width(), canvas_1.winfo_height()), Image.NEAREST)
            # img = img.thumbnail((canvas.winfo_width(), canvas.winfo_height()))
            # ctk_image = ImageTk.PhotoImage(img, size=(canvas.winfo_width(), canvas.winfo_height()))
            # img = img.resize((canvas.winfo_width(), canvas.winfo_height()))
            ctk_image1 = ImageTk.PhotoImage(image=img1)
            canvas_1.itemconfig(image_id_1, image=ctk_image1)
            canvas_1.image = ctk_image1  # Keep a reference
            root.after(10, update_video_feed)

def update_video_feed_2():
    global vid2
    if vid2.isOpened():
        ret2, frame2 = vid2.read()
        if ret2:
            # Crop the frame to the center
            # cropped_frame = frame[100:980, 200:1780]
            frame2 = cv2.flip(frame2, 1)
            img2 = Image.fromarray(cv2.cvtColor(frame2, cv2.COLOR_BGR2RGB))
            # img2 = img2.resize((1920, 1080), Image.NEAREST)
            # img = img.thumbnail((canvas.winfo_width(), canvas.winfo_height()))
            # ctk_image = ImageTk.PhotoImage(img, size=(canvas.winfo_width(), canvas.winfo_height()))
            # img = img.resize((canvas.winfo_width(), canvas.winfo_height()))
            ctk_image2 = ImageTk.PhotoImage(image=img2)
            canvas_2.itemconfig(image_id_2, image=ctk_image2)
            canvas_2.image = ctk_image2  # Keep a reference
            root.after(10, update_video_feed_2)



####################################################################################################################
#------------------------------------------------------------------------------------------------------------------#
####################################################################################################################

# Create the main window
root = tk.CTk()
root.title("Live Camera Feed")
root.geometry("1920x1080")  # Set the window size
root.attributes('-fullscreen', True)

# Create a frame to hold the grid layout
mainframe = tk.CTkFrame(root)
mainframe.pack(fill="both", expand=True, padx=20, pady=(80, 20))  # Make the frame fill the entire window
# Configure the grid layout
for i in range(8):
    mainframe.rowconfigure(i, weight=1, uniform=True)
    mainframe.columnconfigure(i, weight=1, uniform=True)


# Create and place widgets in the grid

#Canvas_1 is where Camera 1 outputs
canvas_1 = tk.CTkCanvas(mainframe, highlightthickness = 1, highlightbackground = "black", background = "grey")
canvas_1.grid(row=0, column=0, rowspan=6, columnspan=6, sticky="nsew")

#Canvas_2 is where Camera 2 outputs
canvas_2 = tk.CTkCanvas(mainframe, highlightthickness = 1, highlightbackground = "black", background = "grey")
canvas_2.grid(row=4, column=4, rowspan=2, columnspan=2, sticky="nsew")

#Canvas_3 is where snapshot outputs
canvas_3 = tk.CTkCanvas(mainframe, highlightthickness = 1, highlightbackground = "black", background = "grey")
canvas_3.grid(row=0, column=6, rowspan=2, columnspan=2, sticky="nsew")

button_open_1 = tk.CTkButton(mainframe, text="Open Camera 1", command= open_camera_1)
button_open_1.grid(column=0, row=6, padx=10, pady=10)

button_close_1 = tk.CTkButton(mainframe, text="Close Camera 1", command= close_camera_1)
button_close_1.grid(column=0, row=7, padx=10, pady=10)

button_open_2 = tk.CTkButton(mainframe, text="Open Camera 2", command= open_camera_2)
button_open_2.grid(column=1, row=6, padx=10, pady=10)

button_close_2 = tk.CTkButton(mainframe, text="Close Camera 2", command= close_camera_2)
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

student_info_field = tk.CTkFrame(mainframe)
student_info_field.grid(column=6, row=2, columnspan=2, rowspan=6, sticky="nsew")

student_info_field.rowconfigure(0, uniform=True)
student_info_field.rowconfigure(1, uniform=True)
student_info_field.rowconfigure(2, uniform=True)
student_info_field.rowconfigure(3, uniform=True)

student_number = tk.CTkLabel(student_info_field, text="Student No. : ", font=("Arial", 18))
student_number.grid(row=0, sticky="w", ipadx= 10, ipady= 10)

student_name = tk.CTkLabel(student_info_field, text="Student Name : ", font=("Arial", 18))
student_name.grid(row=1, sticky="w", ipadx= 10, ipady= 10)

student_course = tk.CTkLabel(student_info_field, text="Course-Section : ", font=("Arial", 18))
student_course.grid(row=2, sticky="w", ipadx= 10,ipady= 10)

student_gmail = tk.CTkLabel(student_info_field, text="GMail : ", font=("Arial", 18))
student_gmail.grid(row=3, sticky="w", ipadx= 10, ipady= 10)

root.update()

# Global variable to store the PhotoImage object reference
image_id_1 = canvas_1.create_image(canvas_1.winfo_width()//2, canvas_1.winfo_height()//2, anchor="center")
image_id_2 = canvas_2.create_image(canvas_2.winfo_width()//2, canvas_2.winfo_height()//2, anchor="center")
# image_id_3 = canvas_3.create_image(0, 0, anchor="nw", image=snapshot)


####################################################################################################################
#------------------------------------------------------------------------------------------------------------------#
####################################################################################################################

# Start the main loop
print("[INFO] starting...")

default_image = Image.open("image.png")
photo_image = ImageTk.PhotoImage(image=default_image)
canvas_1.itemconfig(image_id_1, image=photo_image)
canvas_1.image = photo_image
canvas_2.itemconfig(image_id_2, image=photo_image)
canvas_2.image = photo_image

root.mainloop()
