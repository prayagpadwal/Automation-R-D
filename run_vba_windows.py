import win32com.client

# Define the path to your macro-enabled Excel file (.xlsm)
xlsm_path = r"C:\Users\padwa\Downloads\Automation\Vee.xlsm"  # Update with your actual file path

try:
    # Open Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Run in the background

    # Open the workbook (ensure it's .xlsm)
    wb = excel.Workbooks.Open(xlsm_path)

    # Run the VBA Macro (ensure the macro name matches exactly)
    excel.Application.Run("SplitDataByQuantity")

    # Save as .xlsm to retain macros
    wb.Save()

    # Close Excel properly
    wb.Close(SaveChanges=True)
    excel.Quit()

    print("✅ Macro executed successfully.")

except Exception as e:
    print(f"❌ Error: {e}")
