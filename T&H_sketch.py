#!/usr/bin/env python3
"""
Line Diagram Generator for ENCRYPTED DOCX Files with Merged Date Cells
Carries forward the date from morning to afternoon rows.
"""

import os
import re
from datetime import datetime
import matplotlib.pyplot as plt
import win32com.client
from win32com.client import constants
import time
import tkinter as tk
from tkinter import filedialog

def get_chart_title(filename):
    """
    Determine the chart title based on keywords in the filename.
    """
    filename_lower = filename.lower()
    
    if "检修层" in filename_lower:
        return "Temperature & Humidity for Level Maintenance"
    elif "治疗层" in filename_lower:
        return "Temperature & Humidity for Level Treatment"
    elif "设备层" in filename_lower:
        return "Temperature & Humidity for Level Equipment"
    else:
        # Default fallback - use the filename without extension
        base_name = os.path.splitext(filename)[0]
        return f"{base_name} - Temperature & Humidity Over Time"

def get_month_from_data(data):
    """
    Determine the month (English full name) from the timestamps in the data.
    Assumes a single file only contains records from one month.
    Returns e.g. 'September'. If multiple months are found, they are joined.
    Returns None if no valid timestamp is available.
    """
    months = []
    for record in data:
        dt = record[0]
        if isinstance(dt, datetime):
            month_name = dt.strftime("%B")
            if month_name not in months:
                months.append(month_name)
    if not months:
        return None
    return ", ".join(months)


def extract_table_from_word(docx_path):
    """
    Extract data - handles merged date cells by carrying forward the date.
    """
    data = []
    word = None
    doc = None
    
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        time.sleep(0.5)
        
        if doc.Tables.Count == 0:
            print(f"   ⚠️ No tables found")
            return data
        
        table = doc.Tables(1)
        rows = table.Rows.Count
        
        print(f"   📊 Found table with {rows} rows")
        
        # Keep track of the current date
        current_date = None
        
        # Start from row 2 (skip header)
        for row_idx in range(2, rows + 1):
            try:
                # Read all cells in this row
                row_cells = []
                for col_idx in range(1, 13):  # 12 columns as shown in your data
                    try:
                        cell_text = table.Cell(row_idx, col_idx).Range.Text.strip()
                        cell_text = cell_text.replace('\r\x07', '').strip()
                        # If cell is empty or contains only the end-of-cell marker, treat as empty
                        if cell_text == '\x07' or not cell_text:
                            cell_text = ""
                        row_cells.append(cell_text)
                    except:
                        row_cells.append("")
                
                # Extract the key fields
                # Col 0: Date (or empty if merged)
                # Col 1: Time
                # Col 2: Temperature
                # Col 3: Humidity
                # Col 4: Signature
                
                date_str = row_cells[0] if len(row_cells) > 0 else ""
                time_str = row_cells[1] if len(row_cells) > 1 else ""
                temp_str = row_cells[2] if len(row_cells) > 2 else ""
                humid_str = row_cells[3] if len(row_cells) > 3 else ""
                
                # Skip completely empty rows
                if not date_str and not time_str and not temp_str and not humid_str:
                    continue
                
                # If there's a date in this row, update current_date
                if date_str and date_str != "" and date_str != "(error)":
                    try:
                        # Try to parse the date
                        current_date = datetime.strptime(date_str, "%b %d, %Y")
                    except:
                        # If date parsing fails, keep the previous date
                        pass
                
                # If we don't have a time, skip this row
                if not time_str:
                    continue
                
                # Use the current_date (carried forward from morning row)
                if current_date is None:
                    # No date available yet - skip this row
                    continue
                
                # Parse time
                try:
                    time_clean = time_str.lower().replace(".", ":")
                    if "am" in time_clean or "pm" in time_clean:
                        time_clean = time_clean.replace("am", " AM").replace("pm", " PM")
                        time_obj = datetime.strptime(time_clean, "%I:%M %p")
                    else:
                        time_obj = datetime.strptime(time_clean, "%H:%M")
                    
                    full_dt = current_date.replace(hour=time_obj.hour, minute=time_obj.minute)
                    
                    # Parse temperature and humidity
                    temp = None
                    humidity = None
                    
                    if temp_str:
                        try:
                            temp = float(temp_str.replace(',', '.'))
                        except:
                            pass
                    
                    if humid_str:
                        try:
                            humidity = float(humid_str.replace(',', '.'))
                        except:
                            pass
                    
                    # Only add if we have at least temperature or humidity
                    if temp is not None or humidity is not None:
                        data.append((full_dt, temp, humidity))
                        
                except Exception as e:
                    # Skip rows that can't be parsed
                    continue
                    
            except Exception as e:
                # Skip problem rows
                continue
        
        print(f"   ✅ Extracted {len(data)} data points")
        
    except Exception as e:
        print(f"   ❌ Error processing document: {e}")
        
    finally:
        try:
            if doc:
                doc.Close(SaveChanges=False)
            if word:
                word.Quit()
        except:
            pass
    
    return data

def create_line_diagram(data, file_path, output_path, title):
    """Create a line diagram with Temperature and Humidity on the same chart."""
    if not data:
        print(f"⚠️ No valid data for {os.path.basename(file_path)}")
        return
    
    # Sort data by datetime
    data.sort(key=lambda x: x[0])
    
    # Extract values
    dates = [d[0] for d in data]
    temps = [d[1] for d in data]
    humidities = [d[2] for d in data]
    
    # Create the plot
    fig, ax1 = plt.subplots(figsize=(12, 6))
    
    # Plot Temperature (left Y-axis)
    color1 = 'tab:red'
    ax1.set_xlabel('Date & Time')
    ax1.set_ylabel('Temperature (°C)', color=color1)
    ax1.plot(dates, temps, color=color1, marker='o', linestyle='-', linewidth=2, markersize=4, label='Temperature')
    ax1.tick_params(axis='y', labelcolor=color1)
    ax1.grid(True, alpha=0.3)
    
    # Create second Y-axis for Humidity
    ax2 = ax1.twinx()
    color2 = 'tab:blue'
    ax2.set_ylabel('Humidity (%)', color=color2)
    ax2.plot(dates, humidities, color=color2, marker='s', linestyle='-', linewidth=2, markersize=4, label='Humidity')
    ax2.tick_params(axis='y', labelcolor=color2)
    
    # Format x-axis dates
    plt.xticks(rotation=45, ha='right')
    
    # Title (now passed as parameter)
    plt.title(title, fontsize=14, fontweight='bold')
    
    # Add legend
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', bbox_to_anchor=(1.02, 1))
    
    # Adjust layout
    fig.tight_layout()
    
    # Save the figure
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close(fig)
    
    print(f"   ✅ Diagram saved: {output_path}")

def select_files():
    """Open a file dialog for the user to select .docx files."""
    root = tk.Tk()
    root.withdraw()
    root.lift()
    root.attributes('-topmost', True)
    
    file_paths = filedialog.askopenfilenames(
        title="Select your encrypted .docx files (up to 3)",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
    )
    
    root.destroy()
    return list(file_paths)

def main():
    """Main execution with Word automation for encrypted files."""
    print("=" * 60)
    print("📊 Encrypted DOCX to Line Diagram Generator")
    print("   (Handles merged date cells & conditional titles)")
    print("=" * 60)
    print()
    print("⚠️ IMPORTANT: Close Microsoft Word before running this script.")
    print()
    
    # Check if the required package is installed
    try:
        import win32com.client
    except ImportError:
        print("❌ Missing required package: pywin32")
        print("   Install it with: pip install pywin32 --user")
        return
    
    # Ask user for input method
    # print("How do you want to select the files?")
    # print("   [1] Auto-detect .docx files in current folder")
    # print("   [2] Browse and select files manually")
    # choice = input("Your choice (1/2): ").strip()

    choice = '1'
    
    file_paths = []
    
    if choice == '1':
        current_dir = os.getcwd()
        docx_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.docx')]
        
        if not docx_files:
            print("❌ No .docx files found in current folder.")
            return
        
        if len(docx_files) > 3:
            print(f"⚠️ Found {len(docx_files)} files. Processing only the first 3.")
            docx_files = docx_files[:3]
        
        file_paths = [os.path.join(current_dir, f) for f in docx_files]
    else:
        file_paths = select_files()
    
    if not file_paths:
        print("❌ No files selected. Exiting.")
        return
    
    print(f"\n📁 Selected {len(file_paths)} file(s):")
    for f in file_paths:
        print(f"   - {os.path.basename(f)}")
    print()
    
    # Process each file
    for i, docx_path in enumerate(file_paths, 1):
        print(f"\n📄 [{i}/{len(file_paths)}] Processing: {os.path.basename(docx_path)}")
        
        if not os.path.exists(docx_path):
            print(f"   ❌ File not found: {docx_path}")
            continue
        
        # Extract data using Word
        data = extract_table_from_word(docx_path)
        
        if not data:
            print(f"   ⚠️ No data extracted. Skipping chart generation.")
            continue
        
        # Determine the chart title based on filename
        filename = os.path.basename(docx_path)
        chart_title = get_chart_title(filename)
        # Append the month (from timestamps) to the title
        month = get_month_from_data(data)
        if month:
            chart_title = f"{chart_title} - {month}"
        print(f"   🏷️ Title: {chart_title}")
        
        # Generate output filename
        folder = os.path.dirname(docx_path)
        base_name = os.path.splitext(filename)[0]
        output_path = os.path.join(folder, f"{base_name}_chart.png")
        
        # Create diagram with the determined title
        create_line_diagram(data, docx_path, output_path, chart_title)
    
    print("\n" + "=" * 60)
    print("🎉 All diagrams generated successfully!")
    print("=" * 60)

if __name__ == "__main__":
    main()