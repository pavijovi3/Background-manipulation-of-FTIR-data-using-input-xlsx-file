import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd


def choose_input_file():
    file_path = filedialog.askopenfilename(title="Select Input XLSX File", filetypes=[("XLSX Files", "*.xlsx")])
    return file_path


def choose_column(root, df):
    column_window = tk.Toplevel(root)
    column_window.title("Select Column")
    column_window.geometry("300x100")

    column_label = ttk.Label(column_window, text="Choose a column:")
    column_label.pack()

    column_combobox = ttk.Combobox(column_window, values=df.columns.tolist())
    column_combobox.pack()

    chosen_column = tk.StringVar()  # StringVar to store the chosen column

    def on_confirm():
        chosen_column.set(column_combobox.get())  # Update the StringVar with the chosen column
        if chosen_column.get():
            column_window.quit()  # Break out of the main loop when the "Confirm" button is clicked

    confirm_button = ttk.Button(column_window, text="Confirm", command=on_confirm)
    confirm_button.pack()

    column_window.mainloop()

    column_window.destroy()  # Destroy the window after the main loop is finished
    return chosen_column.get()  # Return the chosen column after the window is closed


def process_and_save_data(file_path, chosen_column, output_dir):
    xlsx = pd.ExcelFile(file_path)
    df = xlsx.parse(xlsx.sheet_names[0])

    # Create a new DataFrame with the same values as the original DataFrame
    # but subtract the values of the chosen column from all other columns
    processed_sheet = df.copy()
    processed_sheet.iloc[:, 1:] -= df[chosen_column].values.reshape(-1, 1)

    # Extract the input file name and create the output file name
    input_filename = os.path.splitext(os.path.basename(file_path))[0]
    output_file_name = f"{input_filename}_{chosen_column}.xlsx"
    output_file_path = os.path.join(output_dir, output_file_name)

    # Save the processed DataFrame to a new XLSX file
    with pd.ExcelWriter(output_file_path) as writer:
        processed_sheet.to_excel(writer, sheet_name="Sheet1", index=False)

    return output_file_path


def bg_processing():
    root = tk.Tk()
    root.withdraw()

    file_path = choose_input_file()

    if not file_path:
        messagebox.showinfo("Info", "No input file selected. Exiting the program.")
        root.destroy()
        return

    df = pd.read_excel(file_path)

    chosen_column = choose_column(root, df)

    if not chosen_column:
        messagebox.showinfo("Error", "No column selected.")
        root.destroy()
        return

    output_dir = filedialog.askdirectory(title="Select Output Directory")

    if not output_dir:
        messagebox.showinfo("Error", "No output directory selected.")
        root.destroy()
        return

    output_file_path = process_and_save_data(file_path, chosen_column, output_dir)

    message = f"Processing completed! Output saved as:\n{output_file_path}"
    # Open the folder where the output files are saved
    os.startfile(output_file_path)
    messagebox.showinfo("Processing Complete", message)
    root.destroy()


if __name__ == "__main__":
    window = tk.Tk()
    window.title("Background reprocessing tool")
    window.geometry("400x130")  # Adjusted width to accommodate the label and button

    # Create a frame to contain the label and button
    content_frame = tk.Frame(window, padx=20, pady=20)
    content_frame.pack(anchor="w")

    # Create a label with your desired text
    label = tk.Label(content_frame, text="This program is to manipulate the background of the FTIR data series.\nPlease click to start.")
    label.pack(anchor="w")  # Justify label to the left

    # Create the "Reprocess background now" button with styling
    process_background_data_button = tk.Button(content_frame, text="Reprocess background now", command=bg_processing, bg="light blue")
    process_background_data_button.pack(pady=10, anchor="center")

    window.mainloop()
