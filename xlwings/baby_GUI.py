from tkinter import *

# window, container for all widgets
window = Tk()
window.geometry("420x420") # width x height
window.title("Window Title")
window.config(background="#5cfcff")

# widgets
#1 - label: holds text/image
label = Label(window, 
              text="Hello :)", 
              font=('Arial', 40, 'bold'), 
              fg='blue', 
              bg='#00ff00',
              relief=RAISED, # border style
              bd=10,
              padx=20,
              pady=20)
label.pack() # zet het in uw window
# label.place(x=100, y=100) # zet het in window op een bepaalde plaats

#2 - button
def click():
    print("CLICKED :D")

button = Button(window, text="click me!")
button.config(command=click)
button.config(font=('Ink Free', 50, 'bold'))
button.config(bg="#ff6200")
button.config(fg="#fffb1f")
button.config(activebackground="red")
button.config(activeforeground="#fffb1f")
button.pack()

window.mainloop() # place window on screen and listen for events