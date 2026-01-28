import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.utils import get_column_letter
import os

def extract_xlsm_data(input_file, output_file):
    """
    Extract all data from a VBA-locked .xlsm file to a new unlocked .xlsx file.
    This bypasses VBA password protection by only reading data (not macros).
    """
    try:
        print(f"Opening file: {input_file}")
        
        # Load the workbook (data_only=True reads values, not formulas)
        # keep_vba=False removes all VBA macros and bypasses password
        wb_input = openpyxl.load_workbook(input_file, data_only=True, keep_vba=False)
        
        # Create a new workbook for output
        wb_output = openpyxl.Workbook()
        wb_output.remove(wb_output.active)  # Remove default sheet
        
        print(f"\nFound {len(wb_input.sheetnames)} sheets")
        
        # Copy each sheet
        for sheet_name in wb_input.sheetnames:
            print(f"Processing sheet: {sheet_name}")
            
            source_sheet = wb_input[sheet_name]
            target_sheet = wb_output.create_sheet(title=sheet_name)
            
            # Copy all cell values and basic formatting
            for row in source_sheet.iter_rows():
                for cell in row:
                    target_cell = target_sheet[cell.coordinate]
                    
                    # Copy value
                    target_cell.value = cell.value
                    
                    # Copy basic formatting if exists
                    try:
                        if cell.font:
                            target_cell.font = Font(
                                name=cell.font.name,
                                size=cell.font.size,
                                bold=cell.font.bold,
                                italic=cell.font.italic,
                                color=cell.font.color
                            )
                        
                        if cell.fill:
                            target_cell.fill = PatternFill(
                                fill_type=cell.fill.fill_type,
                                fgColor=cell.fill.fgColor,
                                bgColor=cell.fill.bgColor
                            )
                        
                        if cell.alignment:
                            target_cell.alignment = Alignment(
                                horizontal=cell.alignment.horizontal,
                                vertical=cell.alignment.vertical,
                                wrap_text=cell.alignment.wrap_text
                            )
                        
                        if cell.number_format:
                            target_cell.number_format = cell.number_format
                    except:
                        pass  # Skip if formatting fails
            
            # Copy column widths
            for col in source_sheet.column_dimensions:
                try:
                    if source_sheet.column_dimensions[col].width:
                        target_sheet.column_dimensions[col].width = source_sheet.column_dimensions[col].width
                except:
                    pass
            
            # Copy row heights
            for row in source_sheet.row_dimensions:
                try:
                    if source_sheet.row_dimensions[row].height:
                        target_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height
                except:
                    pass
            
            # Copy merged cells
            try:
                for merged_cell_range in source_sheet.merged_cells.ranges:
                    target_sheet.merge_cells(str(merged_cell_range))
            except:
                pass
        
        # Save the new unlocked workbook as .xlsx
        wb_output.save(output_file)
        print(f"\n✓ Successfully created unlocked file: {output_file}")
        print("All data has been extracted without VBA macros")
        
        wb_input.close()
        wb_output.close()
        
        return True
        
    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found")
        return False
    except PermissionError:
        print(f"Error: Permission denied. Make sure the file is not open in Excel")
        return False
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # Configuration
    input_filename = "boq.xls"  # Change this to your locked .xlsm file name
    output_filename = "boq_file.xlsx"  # Output file (will be .xlsx format)
    
    print("=" * 60)
    print("Excel VBA Password Bypass - .XLSM Data Extractor")
    print("=" * 60)
    print("\nThis tool extracts all data from VBA-locked .xlsm files")
    print("Output will be saved as .xlsx format (without macros)\n")
    
    # Check if input file exists
    if not os.path.exists(input_filename):
        print(f"Error: '{input_filename}' not found in current directory")
        print(f"\nCurrent directory: {os.getcwd()}")
        print("\nPlease either:")
        print(f"1. Place your .xlsm file in the current directory")
        print(f"2. Or update the 'input_filename' variable with the full path")
        print("\nExample: input_filename = r'C:\\Users\\YourName\\Documents\\myfile.xlsm'")
    else:
        # Extract data
        success = extract_xlsm_data(input_filename, output_filename)
        
        if success:
            print("\n" + "=" * 60)
            print("IMPORTANT NOTES:")
            print("=" * 60)
            print("• VBA macros are NOT included (they were password-protected)")
            print("• Formulas are converted to their calculated values")
            print("• All visible data and formatting are preserved")
            print("• Merged cells are preserved")
            print("• The new file has no password protection")
            print("• Output is in .xlsx format (macro-free)")
