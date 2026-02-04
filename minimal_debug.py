import win32com.client

# Create Visio
visio = win32com.client.Dispatch("Visio.Application")
visio.Visible = True

# Create document
doc = visio.Documents.Add("")
page = doc.Pages.Item(1)

# Open flowchart stencil and drop a Process shape
try:
    stencil = visio.Documents.OpenEx("BASFLO_M.VSSX", 4)
    master = stencil.Masters("Process")
    shape = page.Drop(master, 4, 4)
    shape.Text = "Test"
    print(f"Shape created: {shape.Name}")
except Exception as e:
    print(f"Error creating shape: {e}")
    exit()

# Check connection points section
visSectionConnectionPts = 9
row_count = shape.RowCount(visSectionConnectionPts)
print(f"Existing connection point rows: {row_count}")

# Delete all existing connection points
print(f"Deleting {row_count} existing rows...")
for i in range(row_count - 1, -1, -1):
    shape.DeleteRow(visSectionConnectionPts, i)

print(f"Rows after deletion: {shape.RowCount(visSectionConnectionPts)}")

# Now add ONE connection point
visRowConnectionPts = 0
visTagCnnctPt = 153
visX = 0
visY = 1

print("\nAdding one connection point...")
row_num = shape.AddRow(visSectionConnectionPts, visRowConnectionPts, visTagCnnctPt)
print(f"AddRow returned row index: {row_num}")

# Set formulas using row_num
print("Setting formulas...")
shape.CellsSRC(visSectionConnectionPts, row_num, visX).Formula = "=Width*0.5"
shape.CellsSRC(visSectionConnectionPts, row_num, visY).Formula = "=Height*0.5"

# Verify
row_count = shape.RowCount(visSectionConnectionPts)
print(f"\nFinal row count: {row_count}")

x_val = shape.CellsSRC(visSectionConnectionPts, row_num, visX).Formula
y_val = shape.CellsSRC(visSectionConnectionPts, row_num, visY).Formula
print(f"Row {row_num}: X={x_val}, Y={y_val}")

print("\nCheck ShapeSheet - Developer tab → Show ShapeSheet")
input("Press Enter to close...")

