# Malca-PDF-to-Excel
# The goal is to transfer the information from a Malca invoice into an excel file.
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import _ from 'lodash';

async function processInvoice() {
  try {
    console.log("Starting invoice processing...");
    
    const response = await window.fs.readFile('Invoice_497655.pdf', { encoding: 'utf8' });
    
    // Initialize array to store all shipment entries
    const shipments = [];
    
    // Parse the content line by line
    const lines = response.split('\n');
    
    // Extract header info
    const invoiceNumber = "497655";
    const invoiceDate = "11-Dec-24";
    const customer = "REALHQ";
    
    // Process each line
    let currentShipment = null;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      // Match shipment pattern
      const shipmentMatch = line.match(/(\d{2}\/\d{2}\/\d{4})\s+([^\/]+)\/([^\s]+)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)/);
      
      if (shipmentMatch) {
        if (currentShipment) {
          shipments.push(currentShipment);
        }
        
        currentShipment = {
          'Send Date': shipmentMatch[1],
          'Origin': shipmentMatch[2].trim(),
          'Destination': shipmentMatch[3].trim(),
          'Shipment No': shipmentMatch[4],
          'Weight (LB)': parseFloat(shipmentMatch[5].replace(',', '')),
          'Customs Value': parseFloat(shipmentMatch[6].replace(',', '')),
          'Total Charges': 0,
          'Charges': []
        };
      }
      
      // Match charge details
      const chargeMatch = line.match(/(FREIGHT & LIABILITY|FUEL CHARGE|SECURITY FEE|WEIGHT CHARGE)\s+([\d,.]+)/i);
      if (chargeMatch && currentShipment) {
        const charge = {
          type: chargeMatch[1].trim(),
          amount: parseFloat(chargeMatch[2].replace(',', ''))
        };
        currentShipment.Charges.push(charge);
        currentShipment['Total Charges'] += charge.amount;
      }
    }
    
    // Add last shipment
    if (currentShipment) {
      shipments.push(currentShipment);
    }
    
    // Create Excel workbook
    const wb = XLSX.utils.book_new();
    
    // Flatten charges into columns
    const flattenedShipments = shipments.map(s => ({
      'Send Date': s['Send Date'],
      'Origin': s['Origin'],
      'Destination': s['Destination'],
      'Shipment No': s['Shipment No'],
      'Weight (LB)': s['Weight (LB)'],
      'Customs Value': s['Customs Value'],
      'Total Charges': s['Total Charges'],
      'Freight & Liability': _.sumBy(s.Charges.filter(c => c.type.toUpperCase().includes('FREIGHT')), 'amount'),
      'Fuel Charge': _.sumBy(s.Charges.filter(c => c.type.toUpperCase().includes('FUEL')), 'amount'),
      'Security Fee': _.sumBy(s.Charges.filter(c => c.type.toUpperCase().includes('SECURITY')), 'amount'),
      'Weight Charge': _.sumBy(s.Charges.filter(c => c.type.toUpperCase().includes('WEIGHT')), 'amount')
    }));
    
    // Create worksheet from flattened data
    const ws = XLSX.utils.json_to_sheet(flattenedShipments);
    
    // Add the worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Shipments');
    
    // Add summary sheet
    const summaryData = [{
      'Invoice Number': invoiceNumber,
      'Invoice Date': invoiceDate,
      'Customer': customer,
      'Total Shipments': shipments.length,
      'Total Weight': _.sumBy(shipments, 'Weight (LB)'),
      'Total Customs Value': _.sumBy(shipments, 'Customs Value'),
      'Total Charges': _.sumBy(shipments, 'Total Charges')
    }];
    
    const summaryWs = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');
    
    console.log("Processing complete.");
    console.log(`Processed ${shipments.length} shipments.`);
    
    // Write file
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    return buffer;
    
  } catch (error) {
    console.error("Error processing invoice:", error);
    throw error;
  }
}

// Execute the function
processInvoice().catch(console.error);
[Invoice_497655 - $52k Dec.pdf](https://github.com/user-attachments/files/18592854/Invoice_497655.-.52k.Dec.pdf)
[Invoice_497655 - $52k Dec.xlsx](https://github.com/user-attachments/files/18592856/Invoice_497655.-.52k.Dec.xlsx)
