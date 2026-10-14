
import customtkinter as ctk

app = ctk.CTk()
app.title("Placeholder Test")
app.geometry("400x180")

entry = ctk.CTkEntry(
    app,
    width=300,
    placeholder_text="Enter programme name"
)
entry.pack(padx=30, pady=40)

app.mainloop()
