# CMR Print Manager - 24 Place CMR Document Printing Application

A .NET 9 Windows Forms application for managing and printing CMR (Lettre de Voiture) transport documents with full support for all 24 standardized fields.

## Features

### 📋 24 CMR Numbered Places
The application supports all 24 standard CMR places as per the international road transport convention:

1. **Consignor name/address** - Sender/shipper information
2. **Consignor telephone/reference** - Contact and reference details
3. **Loading instructions** - Special handling or loading requirements
4. **Consignor remarks** - Additional notes from shipper
5. **Documents attached** - List of attached documentation
6. **Seals applied** - Seals and security information
7. **Type of packaging/marks** - Packaging details and marks
8. **Goods description** - Detailed description of goods
9. **Goods gross weight (kg)** - Total weight in kilograms
10. **Loading order number** - Reference to loading order
11. **Transport charges place** - Where charges are payable
12. **Consignee name/address** - Recipient information
13. **Consignee telephone/reference** - Recipient contact details
14. **Unloading instructions** - Special unloading requirements
15. **Remarks by carrier** - Carrier observations
16. **Remarks by consignee** - Recipient notes
17. **Transport authorizations** - Licensing/authorization details
18. **Route information** - Planned route details
19. **Insurance particulars** - Insurance information
20. **Carrier signature** - Carrier authentication
21. **Export/transport date** - Date of transport
22-24. **Signature areas** - Designated signature zones

### 🎨 Visual CMR Layout Preview
- Left panel displays a miniature CMR template with all 24 numbered positions
- Helps visualize where information will be placed on the actual document

### 🔤 Resizable Text Fields
- Adjust font size for each of the 24 places (6-20pt)
- Adjust vertical offset (+/- 100px) to fine-tune text positioning
- Real-time preview of adjustments

### 💾 Template System
- **Save templates** - Store your preferred formatting for future use
- **Load templates** - Quickly apply saved layouts
- **Delete templates** - Remove unused templates
- Templates stored in `%AppData%\CMRPrint\Templates\`

### 👥 Customer Management
- Add/edit customer information (consignor/shipper details)
- Import customers from Excel files
- Quick selection dropdown for frequently used customers

### 📊 Excel Import
- Import multiple customer records from Excel files
- Required columns: Name, Address, City, Country, VatNumber
- Bulk-load shipper data for repeated use

### 🖨️ Print Functions
- **Preview Print** - See exactly how the CMR will print before sending to printer
- **Print CMR** - Send formatted document to any connected printer
- Prints on standard A4 paper with all 24 places positioned correctly

## Installation & Usage

### Requirements
- Windows 10 or later (64-bit)
- .NET 9 (included in the self-contained executable)

### Running the Application
1. Extract `CMRPrint.exe` to any location
2. Double-click to run (no installation required)
3. Application runs completely standalone

### Basic Workflow

#### 1. Set Up Your First Template
- Open the application
- Click "Save" template button to create a default template
- Adjust font sizes and offsets for each place as needed
- Click "Save" again to preserve your settings

#### 2. Add Customers
- **Manual entry:** Click "New", fill in consignor details, click "Save"
- **Import from Excel:** Click "Import Excel" and select your file
  - Excel must have columns: Name, Address, City, Country, VatNumber
- Select customers from the dropdown for quick re-use

#### 3. Fill in Transport Details
- Select a customer from the dropdown (pre-fills consignor info)
- Fill in the 24 CMR places with transport information
- Each field shows its CMR place number for reference

#### 4. Adjust Layout if Needed
- If text doesn't align perfectly on your CMR form:
  - Adjust "Font" size to make text smaller/larger
  - Adjust "Offset" to move text up or down
  - Use "Preview Print" to check alignment
- Save your adjustments as a new template for future use

#### 5. Print
- Click "Preview Print" to see the final result
- Click "Print CMR" to print to your default printer
- Or select a different printer from the print dialog

## Template Management

Templates store your personalized font sizes and vertical offsets for each of the 24 places. This allows you to:

- Create different templates for different CMR printers or forms
- Quickly switch between layouts
- Save time on repeated jobs
- Share templates with other users

**Templates are saved to:** `%AppData%\CMRPrint\Templates\`

### Creating Templates
1. Adjust all font sizes and offsets to match your CMR form
2. Click the "Save" button
3. Enter a descriptive name (default: `CMR_YYYY-MM-DD_HHMM`)
4. Template is saved with all your customizations

### Loading Templates
1. Select a template from the dropdown
2. Click "Load"
3. All previously saved settings are applied

### Deleting Templates
1. Select a template from the dropdown
2. Click "Delete"
3. Confirm deletion

## File Locations

- **Application:** `CMRPrint.exe` (self-contained, standalone)
- **Templates:** `%AppData%\CMRPrint\Templates\`
- **Source Code:** Available in GitHub if modifications needed

## Excel Import Format

Create an Excel file with the following columns:

| Name | Address | City | Country | VatNumber |
|------|---------|------|---------|-----------|
| Acme Logistics | 123 Main St | Berlin | Germany | DE123456789 |
| Global Transport | 456 Oak Ave | Paris | France | FR987654321 |

The application will ignore rows where the Name column is empty.

## Printing Tips

- **Alignment:** Use "Preview Print" and template adjustments to ensure perfect alignment
- **Paper:** Standard A4 works best
- **Test Print:** Do a test print before bulk printing
- **CMR Form:** Align the application output with your pre-printed CMR forms or use blank paper with a background image

## Technical Details

- **Framework:** .NET 9
- **Platform:** Windows only (Win-x64)
- **Architecture:** Self-contained single-file executable
- **Dependencies:** ClosedXML for Excel support
- **Build:** Release optimized with trimming disabled for compatibility

## Troubleshooting

### Text alignment is off
- Adjust the "Offset" values for specific places
- Use Preview Print to see real-time changes
- Save adjustments as a new template

### Customer import fails
- Verify Excel file has required columns: Name, Address, City, Country, VatNumber
- Ensure Name column is not empty for rows to import
- Try re-saving the Excel file in .xlsx format

### Print preview is blank
- Ensure at least the consignor name is filled in
- Check that a printer is configured on your system
- Try installing/updating printer drivers

### Templates not saving
- Verify you have write access to `%AppData%\CMRPrint\Templates\`
- Check available disk space
- Try running as Administrator if permissions issues occur

## Support & Feedback

This is a standalone application for CMR document printing. For bug reports or feature requests, consult the development team.

---

**Version:** 1.0  
**Release Date:** June 2026  
**Executable Location:** `c:\CMRPrint\bin\Release\net9.0-windows\win-x64\publish\CMRPrint.exe`
