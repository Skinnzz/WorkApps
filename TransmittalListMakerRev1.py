import os
import json
import pandas as pd
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox

def get_pdf_file_list(directory_path):
    """
    Get list of PDF files in a directory without extensions.
    
    Args:
        directory_path (str): Path to directory with PDF files
        
    Returns:
        list: List of filenames without .pdf extension
    """
    file_list = []
    
    try:
        # Loop through all items in directory
        for item in os.listdir(directory_path):
            # Check if item is a file and has a .pdf extension
            if os.path.isfile(os.path.join(directory_path, item)) and item.lower().endswith('.pdf'):
                # Remove the .pdf extension and add to list
                file_name_without_extension = os.path.splitext(item)[0]
                file_list.append(file_name_without_extension)
                
        return file_list
    except Exception as e:
        print(f"Error reading directory: {str(e)}")
        return []

def filter_bom_with_pdf_list(excel_file_path, pdf_directory_path):
    """
    Main function that combines PDF listing and Excel filtering.
    
    Args:
        excel_file_path (str): Path to the Excel BOM file
        pdf_directory_path (str): Path to directory with PDF files
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Step 1: Get list of PDF files without extensions
        print(f"Reading PDF files from: {pdf_directory_path}")
        pdf_files = get_pdf_file_list(pdf_directory_path)
        
        if not pdf_files:
            print("No PDF files found in the specified directory.")
            return False, []
            
        print(f"Found {len(pdf_files)} PDF files")
        
        # Create temporary JSON file in the same directory as the Excel file
        temp_dir = os.path.dirname(excel_file_path)
        temp_json_path = os.path.join(temp_dir, "temp_pdf_list.json")
        
        # Save the PDF list to temporary JSON
        json_data = {
            "filenames": pdf_files,
            "total_count": len(pdf_files),
            "directory": pdf_directory_path
        }
        
        with open(temp_json_path, 'w') as json_file:
            json.dump(json_data, json_file, indent=4)
            
        print(f"Temporarily saved PDF list to: {temp_json_path}")
        
        # Step 2: Filter the Excel file using the PDF list
        print(f"\nProcessing Excel file: {excel_file_path}")
        
        # Read the Excel file
        df = pd.read_excel(excel_file_path)
        
        # Record original row count
        original_row_count = len(df)
        print(f"Original row count: {original_row_count}")
        
        # Create a set of part numbers to keep (both with and without VG- prefix)
        all_parts_to_keep = set()
        for part in pdf_files:
            all_parts_to_keep.add(part)
            # If part starts with VG-, add version without prefix
            if part.startswith('VG-'):
                all_parts_to_keep.add(part[3:])
            # If part doesn't have prefix, add version with VG- prefix
            else:
                all_parts_to_keep.add(f'VG-{part}')
        
        # Helper function to normalize part number for comparison
        def normalize_part_number(part_no):
            if pd.isna(part_no):
                return ""
            return str(part_no).replace('\r', '').replace('\n', '').strip()
        
        # Add normalized part number column for easier processing
        df['Normalized_Part_No'] = df['PART No.'].apply(normalize_part_number)
        
        # Get all unique part numbers in the Excel file
        excel_part_numbers = set(df['Normalized_Part_No'].dropna().unique())
        
        # Check which PDF part numbers don't exist in the Excel file
        missing_parts = []
        for part in pdf_files:
            # Check if the part (with or without VG- prefix) exists in Excel
            if part not in excel_part_numbers and (
                not part.startswith('VG-') or part[3:] not in excel_part_numbers
            ) and (
                not part.startswith('VG-') or f'VG-{part}' not in excel_part_numbers
            ):
                missing_parts.append(part)
        
        # Step 3: Remove duplicate part numbers
        print("Removing duplicate part numbers...")
        df_no_dupes = df.drop_duplicates(subset=['Normalized_Part_No'], keep='first')
        dupes_removed = original_row_count - len(df_no_dupes)
        print(f"Removed {dupes_removed} duplicate rows")
        
        # Step 4: Filter based on part numbers from PDF files
        def should_keep_row(row):
            part_no = row['Normalized_Part_No']
            if part_no == "":
                return False
                
            # Check if part number matches any in our list (with or without VG- prefix)
            for keep_part in all_parts_to_keep:
                # Direct match
                if part_no == keep_part:
                    return True
                
                # Check without the "VG-" prefix if present
                if part_no.startswith('VG-') and part_no[3:] == keep_part:
                    return True
            
            return False
        
        # Filter the DataFrame
        filtered_df = df_no_dupes[df_no_dupes.apply(should_keep_row, axis=1)]
        
        # Remove the temporary column we added
        filtered_df = filtered_df.drop(columns=['Normalized_Part_No'])
        
        # Record filtered row count
        filtered_row_count = len(filtered_df)
        print(f"Filtered row count: {filtered_row_count}")
        
        # STEP 5: Restructure the DataFrame - keep only PART No. and DESCRIPTION
        simplified_df = filtered_df[['PART No.', 'DESCRIPTION']].copy()
        print("Simplified DataFrame to keep only PART No. and DESCRIPTION")
        
        # Set output file path (same directory as input Excel file)
        output_dir = os.path.dirname(excel_file_path)
        output_filename = "FILTERED_" + os.path.basename(excel_file_path)
        output_file = os.path.join(output_dir, output_filename)
        
        # Save the filtered DataFrame to a new Excel file
        simplified_df.to_excel(output_file, index=False)
        print(f"Filtered BOM saved to: {output_file}")
        
        # Clean up: Delete the temporary JSON file
        if os.path.exists(temp_json_path):
            os.remove(temp_json_path)
            print(f"Temporary JSON file deleted")
        
        # Print summary
        print("\nSummary:")
        print(f"Original rows: {original_row_count}")
        print(f"Duplicate rows removed: {dupes_removed}")
        print(f"Filtered rows: {filtered_row_count}")
        print(f"Total rows removed: {original_row_count - filtered_row_count}")
        print(f"Percentage kept: {filtered_row_count/original_row_count*100:.2f}%")
        print(f"Columns kept: PART No., DESCRIPTION")
        
        # Return information about missing parts
        return True, missing_parts
        
    except Exception as e:
        print(f"Error in processing: {str(e)}")
        return False, []

def main():
    # Create a simple GUI
    root = tk.Tk()
    root.title("BOM Filter Tool")
    root.geometry("500x300")
    
    # Add some instructions
    tk.Label(root, text="BOM Filter Tool", font=("Arial", 16, "bold")).pack(pady=10)
    tk.Label(root, text="This tool filters an Excel BOM file based on PDF filenames.").pack()
    tk.Label(root, text="Step 1: Select the Excel BOM file").pack(pady=5)
    tk.Label(root, text="Step 2: Select the folder containing PDF files").pack(pady=5)
    
    # Variables to store file paths
    excel_path = tk.StringVar()
    pdf_dir = tk.StringVar()
    
    # Function to browse for Excel file
    def browse_excel():
        file_path = filedialog.askopenfilename(
            title="Select Excel BOM file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            excel_path.set(file_path)
            excel_entry.delete(0, tk.END)
            excel_entry.insert(0, file_path)
    
    # Function to browse for PDF directory
    def browse_pdf_dir():
        dir_path = filedialog.askdirectory(
            title="Select Folder Containing PDF Files"
        )
        if dir_path:
            pdf_dir.set(dir_path)
            pdf_entry.delete(0, tk.END)
            pdf_entry.insert(0, dir_path)
    
    # Function to run the process
    def run_process():
        if not excel_path.get() or not pdf_dir.get():
            messagebox.showerror("Error", "Please select both an Excel file and a PDF directory")
            return
            
        status_label.config(text="Processing... Please wait.")
        root.update()
        
        success, missing_parts = filter_bom_with_pdf_list(excel_path.get(), pdf_dir.get())
        
        if success:
            messagebox.showinfo("Success", "BOM filtering completed successfully!")
            status_label.config(text="Done! Filtered BOM saved in the same folder as the original.")
            
            # Show warning if there are missing parts
            if missing_parts:
                # Create a nicely formatted list of missing parts
                missing_parts_text = "\n".join(missing_parts[:50])  # Limit to first 50 to avoid huge dialog
                
                if len(missing_parts) > 50:
                    missing_parts_text += f"\n\n...and {len(missing_parts) - 50} more"
                
                # Create a custom dialog to display missing parts
                missing_dialog = tk.Toplevel(root)
                missing_dialog.title("Warning: Missing Parts")
                missing_dialog.geometry("500x400")
                
                tk.Label(
                    missing_dialog, 
                    text="The following part numbers from PDF files were not found in the BOM:",
                    font=("Arial", 12, "bold"),
                    fg="red"
                ).pack(pady=10)
                
                # Create a scrollable text area
                frame = tk.Frame(missing_dialog)
                frame.pack(fill="both", expand=True, padx=10, pady=10)
                
                scrollbar = tk.Scrollbar(frame)
                scrollbar.pack(side="right", fill="y")
                
                text_area = tk.Text(frame, yscrollcommand=scrollbar.set)
                text_area.pack(side="left", fill="both", expand=True)
                text_area.insert(tk.END, missing_parts_text)
                text_area.config(state="disabled")  # Make it read-only
                
                scrollbar.config(command=text_area.yview)
                
                # Close button
                tk.Button(
                    missing_dialog, 
                    text="Close", 
                    command=missing_dialog.destroy
                ).pack(pady=10)
                
                # Save button
                def save_missing_parts():
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".txt",
                        filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                        title="Save Missing Parts List"
                    )
                    if save_path:
                        with open(save_path, 'w') as f:
                            for part in missing_parts:
                                f.write(f"{part}\n")
                        messagebox.showinfo("Saved", f"Missing parts list saved to {save_path}")
                
                tk.Button(
                    missing_dialog, 
                    text="Save List", 
                    command=save_missing_parts
                ).pack(pady=10)
                
        else:
            messagebox.showerror("Error", "An error occurred during processing. Check the console for details.")
            status_label.config(text="Error occurred. Please try again.")
    
    # Create input fields and buttons
    frame1 = tk.Frame(root)
    frame1.pack(fill="x", padx=20, pady=5)
    tk.Label(frame1, text="Excel file:").pack(side="left")
    excel_entry = tk.Entry(frame1, textvariable=excel_path, width=50)
    excel_entry.pack(side="left", expand=True, fill="x", padx=5)
    tk.Button(frame1, text="Browse", command=browse_excel).pack(side="right")
    
    frame2 = tk.Frame(root)
    frame2.pack(fill="x", padx=20, pady=5)
    tk.Label(frame2, text="PDF folder:").pack(side="left")
    pdf_entry = tk.Entry(frame2, textvariable=pdf_dir, width=50)
    pdf_entry.pack(side="left", expand=True, fill="x", padx=5)
    tk.Button(frame2, text="Browse", command=browse_pdf_dir).pack(side="right")
    
    # Add run button
    tk.Button(
        root, 
        text="Run", 
        command=run_process,
        font=("Arial", 12),
        bg="#4CAF50",
        fg="white",
        width=20,
        height=2
    ).pack(pady=20)
    
    # Status label
    status_label = tk.Label(root, text="Ready", font=("Arial", 10, "italic"))
    status_label.pack(pady=10)
    
    # Run the GUI
    root.mainloop()

if __name__ == "__main__":
    main()