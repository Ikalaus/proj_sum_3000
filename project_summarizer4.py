import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk, filedialog
from pathlib import Path
import os
import re
import csv
from datetime import date

today = date.today()
csv_combo = None  

def change_directory():
    selected_directory = filedialog.askdirectory()
    if selected_directory:
        os.chdir(selected_directory)

def extract_numeric_value(value, delimiter):
    match = re.search(rf"(\d+)\s*{re.escape(delimiter)}(\s*A)?", value, re.IGNORECASE)
    return int(match.group(1)) if match else 0

def read_csv_to_dict(csv_file_path):
    data = []
    csv_file_path = Path(csv_file_path)  # Convert the path to a Path object
    with csv_file_path.open('r', encoding='utf-8-sig') as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            data.append(row)
    return data

def analyze_columns(data):
    zeros = (0, ) * 19
    (sum_a, sum_u, count_p, 
     value_rc, count_power, count_image, count_tapfaded, count_r_plus, 
     count_c_star, count_mdu, count_g_star, count_vl, max_e_number, 
     max_a_number, sum_lots, total_rows, count_le, count_amp, 
     count_drop, 
     
     ) = zeros


    e_number_pattern = r'E(\d+)'
    a_number_pattern = r'A(\d+)'
    

    for row in data:
        subj = row.get('Subject', '')
        comments = row.get('Comments', '')
        color = row.get('Color', '')

        value_a1 = extract_numeric_value(comments, "'A")
        value_a2 = extract_numeric_value(comments, "' A")
        value_a3 = extract_numeric_value(comments, "' R")
        value_a4 = extract_numeric_value(comments, "'R")
        value_u1 = extract_numeric_value(comments, "'U")
        value_u2 = extract_numeric_value(comments, "' U")
        sum_a += value_a1 + value_a2 + value_a3 + value_a4
        sum_u += value_u1 + value_u2
        sum_lots = count_c_star + count_g_star + count_mdu + count_r_plus + count_vl
        total_rows = len(data)
            

        value_rc += 1 if re.search(r"#00FFFF", color, re.IGNORECASE) else 0
        count_p += 1 if re.search(r"[ ]*P\"[ ]*", comments, re.IGNORECASE) else 0        
        count_r_plus += 1 if re.search(r"^[ ]*R\+[ ]*$", comments, re.IGNORECASE) else 0
        count_c_star += 1 if re.search(r"^[ ]*C\*[ ]*$", comments, re.IGNORECASE) else 0
        count_mdu += 1 if re.search(r"^[ ]*mdu\%[ ]*$", comments, re.IGNORECASE) else 0
        count_g_star += 1 if re.search(r"^[ ]*G\*[ ]*$", comments, re.IGNORECASE) else 0
        count_vl += 1 if re.search(r"\bVL\b", comments, re.IGNORECASE) else 0
        count_image += 1 if re.search(r"\bimage\b", subj, re.IGNORECASE) else 0
        count_tapfaded += 1 if re.search(r"\btap faded\b", subj, re.IGNORECASE) else 0
        count_power += 1 if re.search(r"\bpower\b", subj, re.IGNORECASE) else 0
        count_le += 1 if re.search(r"[ ]*\#LE[ ]*", comments, re.IGNORECASE) else 0
        count_amp += 1 if re.search(r"[ ]*\#AMP[ ]*", comments, re.IGNORECASE) else 0  
        count_drop += 1 if re.search(r"^drop$", subj, re.IGNORECASE) else 0



        if not subj.strip():
            matches = re.findall(e_number_pattern, comments)
            for match in matches:
                number = int(match)
                if number > max_e_number:
                    max_e_number = number
                    
        if not subj.strip():
            matches = re.findall(a_number_pattern, comments)
            for match in matches:
                number = int(match)
                if number > max_a_number:
                    max_a_number = number    

    unique_values = set()

    for row in data:
        subj = row.get('Subject', '')
        comments = row.get('Comments', '')

        if subj == 'MFG':
            unique_values.add(comments)    

    return (
        sum_a, sum_u, count_p, value_rc, count_power, 
        count_r_plus, count_c_star, count_mdu, count_g_star, count_vl, 
        max_e_number, max_a_number, sum_lots, total_rows, count_le, 
        count_amp, count_drop, unique_values
    )

def check_text_box(data):
    return any(re.search(r"\btext box\b", row.get('Subject', ''), re.IGNORECASE) for row in data)

def check_groups(data):
    return any(re.search(r"\bgroup\b", row.get('Subject', ''), re.IGNORECASE) for row in data)

def check_spans(data):
    spans_rows = [row for row in data if re.search(r"\bspans\b", row.get('Subject', ''), re.IGNORECASE)]
    
    for row in spans_rows:
        comments = row.get('Comments', '')

        if re.search(r"'", comments, re.IGNORECASE):
            return True
    
    return False 

def check_sequence_column(data):
    return any('Sequence' in row for row in data)

def check_sequence_values(data):
    return any(row.get('Sequence') for row in data if 'Sequence' in row and row['Sequence'])

def logo_click(result_text_widget, csv_combo_widget):
    selected_directory = filedialog.askdirectory()  # Open a directory selection dialog

    if selected_directory:
        global current_directory, csv_files
        current_directory = Path(selected_directory)
        
        result_text_widget.configure(state="normal")
        result_text_widget.delete(1.0, tk.END)
        result_text_widget.insert(tk.END, f"Current Directory: {current_directory}\n")
        result_text_widget.configure(state="disabled")

        result_text_widget.configure(state="normal")


def open_gui():
    def run_script():
        csv_name = csv_combo.get()
        csv_name = csv_name.replace(".csv", "")
        csv_file_path = Path(current_directory) / f"{csv_name}.csv"  # Use Path object to create file path
        print("CSV File Path:", csv_file_path)
        
        if csv_file_path.exists():
            print("CSV File Exists")
            data = read_csv_to_dict(csv_file_path)

            (
                sum_a, sum_u, count_p, value_rc, count_power, count_r_plus, 
                count_c_star, count_mdu, count_g_star, count_vl, max_e_number, max_a_number,
                sum_lots, total_rows, count_le, count_amp, count_drops, unique_values
            ) = analyze_columns(data)
            
            allsum = sum_a + sum_u 
            
            
            output_text = f"Project Summary for {csv_name.upper()}"
            remaining_spaces = 100 - len("Project Summary for {csv_name.upper()}")
            output_text += f"Today's Date: {today.strftime('%Y-%m-%d')}\n\n".rjust(remaining_spaces)
            output_text += f"Total Markups: {total_rows}\n"
            output_text += f"Total Markups/Mile: {(total_rows/allsum)*5280}\n\n"
            
            output_text += f"Aerial Footage - including Risers: {sum_a}\n"
            output_text += f"Underground Footage: {sum_u}\n"
            
            output_text += f"Total Footage: {allsum}\n\n"
            
            output_text += f"Poles: {count_p}\n"
            output_text += f"LEs: {count_le}\n"
            output_text += f"AMPs: {count_amp}\n"
            output_text += f"R+: {count_r_plus}\n"
            output_text += f"C*: {count_c_star}\n"
            output_text += f"MDU: {count_mdu}\n"
            output_text += f"G*: {count_g_star}\n"
            output_text += f"VL: {count_vl}\n\n"
            
            output_text += f"Total Lots: {sum_lots}\n\n"
            
            output_text += f"RC spots: {value_rc}\n"
            output_text += f"Actives: {max_a_number}\n"
            output_text += f"EoL: {max_e_number}\n\n"
            
            output_text += f"MFG's in this Job: {unique_values}\n"
            output_text += f"Total Drops: {count_drops}\nTotal lots - VL: {sum_lots - count_vl}\n\n"
        
            if check_sequence_column(data):
                if check_sequence_values(data):
                    output_text += f"Sequences found in Job please double check!\n"
                else:
                    output_text += f"Sequence column exists and no sequences found. Nice.\n"
            else:
                output_text += f"Missing sequence column, please add this to your markup list.\n"


            if check_groups(data):
                output_text += f"Groups found, please sort these"
            else:
                output_text += f"No Groups found. Nice."
            
        
            if count_power < 1:
                output_text += f"\nNo Power Supplies in job. Please make sure there is a viable way for this Node to get power!\n"
            else:
                output_text += f"\nPower Supplies: {count_power}. Nice.\n"

                
            if check_spans(data):
                spans_rows = [row for row in data if re.search(r"\bspans\b", row.get('Subject', ''), re.IGNORECASE)]
                missing_apostrophe_rows = []
                missing_000_rows = []
                
                for row in spans_rows:
                    comments = row.get('Comments', '')
                    if "'" not in comments and "’" not in comments:
                        missing_apostrophe_rows.append(row)
                    if "000" in comments:
                        missing_000_rows.append(row)
                
                if missing_apostrophe_rows:
                    output_text += "Some spans are missing apostrophes, please double check these spans:\n"
                    for row in missing_apostrophe_rows:
                        comments_value = row.get('Comments', '')
                        output_text += f"'{comments_value}'\n"
                
                if missing_000_rows:
                    output_text += "Some spans have '000' in their comments:\n"
                    for row in missing_000_rows:
                        comments_value = row.get('Comments', '')
                        output_text += f"'{comments_value}'\n"
                
                if not missing_apostrophe_rows and not missing_000_rows:
                    output_text += "All spans have an apostrophe. Nice.\n"
            else:
                output_text += "All spans are good. Nice.\n"
                    
            if check_text_box(data):
                text_box_rows = [row for row in data if re.search(r"\btext box\b", row.get('Subject', ''), re.IGNORECASE)]
                remaining_text_box_rows = []
                
                for row in text_box_rows:
                    comments = row.get('Comments', '')
                    if not any(phrase in comments.lower() for phrase in ['complete', 'done', 'finished']):
                        remaining_text_box_rows.append(row)
                
                if remaining_text_box_rows:
                    output_text += "\nSome text boxes are still unsorted, please make sure they are not important:\n"
                    for row in remaining_text_box_rows:
                        comments_value = row.get('Comments', '')
                        output_text += f"'{comments_value}'\n"
                else:
                    output_text += "\nAll text boxes are complete. Nice.\n"
            else:
                output_text += "\nNo text boxes found. Nice.\n"
                
            output_text += "\nTN drops should be BLUE\nMA drops should be RED\n"

            result_text.configure(state="normal")  # Set to normal to insert text
            result_text.delete(1.0, tk.END)  # Clear previous results
            result_text.insert(tk.END, output_text)
            result_text.configure(state="disabled")  # Set back to disabled

        else:
            messagebox.showerror("Error", "The specified CSV file does not exist or cannot be accessed.")

    def reset():
        result_text.configure(state="normal")  # Set to normal to clear text
        result_text.delete(1.0, tk.END)
        result_text.configure(state="disabled")  # Set back to disabled

        # Update the list of available CSV files
        csv_files = [file.stem for file in current_directory.glob("*.csv")]
        csv_combo["values"] = csv_files

    root = tk.Tk()
    root.title("Project Summarizer")

    version = "1.5"

    s = ttk.Style()
    s.theme_use('clam')

    version_label = tk.Label(root, text=f"Version: {version}\n")
    version_label.pack()
   
    csv_label = tk.Label(root, text="Select the Folder with your CSV files:")
    csv_label.pack()
    
    root.geometry("750x950")
    canvas = tk.Canvas(root, width=200, height=50)
    canvas.pack()

    # change dir
    change_dir_button = tk.Button(canvas, text="Change Directory", command=lambda: logo_click(result_text, csv_combo))
    change_dir_button.pack()

    csv_label = tk.Label(root, text="\nSelect a CSV file:")
    csv_label.pack()

    global csv_combo  # Declare csv_combo as a global variable
    csv_combo = ttk.Combobox(root, values=csv_files)
    csv_combo.pack()

    button_frame = tk.Frame(root)
    button_frame.pack(side="top", padx=10, pady=5)

    run_button = tk.Button(button_frame, text="Run Script", command=run_script)
    run_button.pack(side="left", padx=5)

    reset_button = tk.Button(button_frame, text="Reset", command=reset)
    reset_button.pack(side="left", padx=5)

    result_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=30)
    result_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    root.mainloop()

if __name__ == '__main__':
    current_directory = Path.cwd()
    print("Current working directory:", current_directory)

    csv_files = [file.stem for file in current_directory.glob("*.csv")]

    open_gui()
