import tkinter as tk


# Initialize tkinter.
root = tk.Tk()

# Stopwatch label
stopwatch = tk.Label(root, text="Test")
stopwatch.pack()

minutes = 0
seconds = 0

def update_stopwatch():
    global minutes
    global seconds

    if seconds < 59:
        seconds += 1
    elif seconds == 59:
        seconds = 0
        minutes +=1

    # Update Label.
    time_string = "{:02d}:{:02d}".format(minutes, seconds)
    stopwatch.config(text=time_string)

    root.after(1000, update_stopwatch)  # Call again in 1000 millisecs.


update_stopwatch()  # Start stopwatch updating.
root.mainloop()