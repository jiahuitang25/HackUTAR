function generateInvoice() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const invoices = {}; // Object to hold invoice data
  
    // Loop through the data to collect invoice information
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const invoiceNumber = row[0];
      const clientName = row[1];
      const clientEmail = row[2];
      const itemDescription = row[3];
      const quantity = row[4];
      const unitPrice = row[5];
      const status = row[7]; // Column H (Status)
  
      // Only process pending invoices
      if (status === "Pending") {
        // Initialize the invoice if it doesn't exist
        if (!invoices[invoiceNumber]) {
          invoices[invoiceNumber] = {
            clientName: clientName,
            clientEmail: clientEmail,
            items: [],
            totalAmount: 0
          };
        }
  
        // Calculate total item amount and add the item to the invoice
        const totalItemAmount = quantity * unitPrice;
        invoices[invoiceNumber].items.push({
          description: itemDescription,
          quantity: quantity,
          unitPrice: unitPrice,
          total: totalItemAmount
        });
        invoices[invoiceNumber].totalAmount += totalItemAmount;
  
        // Update the Total Amount in the sheet
        sheet.getRange(i + 1, 7).setValue(totalItemAmount); // Column G
      }
    }
  
    // Generate and send invoices
    for (const invoiceNumber in invoices) {
      const invoiceData = invoices[invoiceNumber];
      const subject = `Invoice #${invoiceNumber} for ${invoiceData.clientName}`;
      const body = `
    <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <title>Invoice</title>
    <style>
      /* CSS Reset */
      body, p, table, th, td {
        margin: 0;
        padding: 0;
        font-family: Arial, sans-serif;
      }
  
      body {
        background-color: #f4f4f4;
        line-height: 1.6;
      }
  
      .invoice-container {
        max-width: 700px;
        margin: 20px auto;
        background-color: white;
        border-radius: 8px;
        box-shadow: 0 0 15px rgba(0, 0, 0, 0.2);
        padding: 40px;
        border: 1px solid #ddd;
      }
  
      .invoice-header {
        text-align: center;
        margin-bottom: 30px;
      }
  
      .invoice-header img {
        max-width: 100px;
        margin-bottom: 10px;
      }
  
      .invoice-header h1 {
        color: #007bff;
        font-size: 32px;
        border-bottom: 2px solid #007bff;
        display: inline-block;
        padding-bottom: 5px;
      }
  
      .company-name {
        font-size: 18px;
        color: #333;
        margin-bottom: 5px;
      }
  
      .invoice-info {
        margin-bottom: 20px;
      }
  
      .invoice-info p {
        margin-bottom: 8px;
        font-size: 16px;
      }
  
      .invoice-info strong {
        color: #333;
      }
  
      .invoice-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 20px;
      }
  
      .invoice-table th, .invoice-table td {
        padding: 12px;
        text-align: left;
        border-bottom: 1px solid #ddd;
      }
  
      .invoice-table th {
        background-color: #f4f4f4;
        font-size: 16px;
      }
  
      .invoice-table tbody tr:hover {
        background-color: #f9f9f9;
      }
  
      .invoice-total {
        text-align: right;
        font-size: 18px;
        margin-top: 20px;
      }
  
      .invoice-total strong {
        color: #333;
      }
  
      .invoice-footer {
        text-align: center;
        color: #777;
        font-size: 14px;
        margin-top: 30px;
      }
    </style>
  </head>
  <body>
    <div class="invoice-container">
      <div class="invoice-header">
        <img src="https://drive.google.com/uc?export=download&id=1HlXDi4noTPaeAg-CY55KuRBuvOS98ZTQ" alt="Company Logo">
        <div class="company-name"><strong>Agacia Furniture Sdn.Bhd</strong></div>
        <h1>Invoice</h1>
      </div>
  
      <div class="invoice-info">
        <p><strong>Invoice Number:</strong> ${invoiceNumber}</p>
        <p><strong>Invoice Date:</strong> ${new Date().toLocaleString()}</p>
      </div>
  
      <div class="invoice-info">
        <p><strong>Client Name:</strong> ${invoiceData.clientName}</p>
        <p><strong>Client Email:</strong> ${invoiceData.clientEmail}</p>
      </div>
  
      <table class="invoice-table">
        <thead>
          <tr>
            <th>Item No</th>
            <th>Item Description</th>
            <th>Quantity</th>
            <th>Unit Price</th>
            <th>Total Amount</th>
          </tr>
        </thead>
        <tbody>
          ${invoiceData.items.map((item, index) => `
            <tr>
              <td>${index + 1}</td>
              <td>${item.description}</td>
              <td>${item.quantity}</td>
              <td>RM${item.unitPrice.toFixed(2)}</td>
              <td>RM${item.total.toFixed(2)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
  
      <p class="invoice-total"><strong>Total: RM${invoiceData.totalAmount.toFixed(2)}</strong></p>
  
      <div class="invoice-footer">
        <p>Thank you for your business!</p>
      </div>
    </div>
  </body>
  </html>`;
  
      // Send the invoice email
      MailApp.sendEmail({
        to: invoiceData.clientEmail,
        subject: subject,
        htmlBody: body // Use htmlBody to send the HTML formatted email
      });
  
      // Update the status to "Sent" for all items in this invoice
      for (let j = 1; j < data.length; j++) {
        if (data[j][0] == invoiceNumber) {
          Logger.log(`Updating status for invoice number: ${invoiceNumber} at row: ${j + 1}`);
          sheet.getRange(j + 1, 8).setValue("Sent"); // Column H
        }
      }
    }
  }