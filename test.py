import tkinter as tk

# Create tkinter window
root = tk.Tk()
root.title("Pack Propagate Example")

# Create a frame with pack propagation enabled (default)
frame1 = tk.Frame(root, bg="red", width=200, height=100)
frame1.pack_propagate(True)  # Default behavior
frame1.pack()

# Create a label inside frame1
label1 = tk.Label(frame1, text="Frame 1", bg="blue", fg="white")
label1.pack(fill=tk.BOTH, expand=True)

# Create another frame with pack propagation disabled
frame2 = tk.Frame(root, bg="green", width=200, height=100)
frame2.pack_propagate(False)  # Disable pack propagation
frame2.pack()

# Create a label inside frame2
label2 = tk.Label(frame2, text="Frame 2", bg="yellow", fg="black")
label2.pack(fill=tk.BOTH, expand=True)

root.mainloop()
