const express = require("express");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const app = express();
const port = process.env.PORT || 3000;

const publicFolder = path.join(__dirname, "public");
if (!fs.existsSync(publicFolder)) {
  fs.mkdirSync(publicFolder);
}


app.use(express.json());
app.use(express.static("public"));

app.post("/", async (req, res) => {
    const { success, content } = req.body;

    if (!success || !Array.isArray(content)) {
      return res.status(400).json({ error: "Invalid data format." });
    }
  
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Orders Report");
  
      // Define columns
      worksheet.columns = [
        { header: "Order No.", key: "order_id", width: 15 },
        { header: "Order Date", key: "created_at", width: 20 },
        { header: "Due Status", key: "is_overdue", width: 15 },
        { header: "Amount", key: "remaining_amount", width: 10 },
        { header: "Name", key: "customer_name", width: 25 },
        { header: "Company Name", key: "customer_company", width: 30 },
        { header: "Email", key: "customer_email", width: 30 },
        { header: "Suburb", key: "customer_suburb", width: 25 },
        { header: "State", key: "customer_state", width: 25 },
      ];
  
      // Add rows
      content.forEach((item) => {
        worksheet.addRow({
          order_id: typeof item.order_id === "string" ? item.order_id.replace(/#/g, "") : item.order_id,
          created_at: item.created_at,
          is_overdue: item.is_overdue ? "DUE" : "",
          remaining_amount: Number(item.remaining_amount),
          customer_name: item.customer_name,
          customer_company: item.customer_company,
          customer_email: item.customer_email,
          customer_suburb: item.customer_suburb,
          customer_state: item.customer_state,
        });

        if(item.is_overdue){
          worksheet.getRow(worksheet._rows.length).getCell(4).fill = { 
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'F08080' },
          } 
        }
      });
  
      // Save file in public folder
      const fileName = new Date().toLocaleDateString('en-AU').replace(/\//g, "-") +".xlsx";
      const filePath = path.join(publicFolder, fileName);
      await workbook.xlsx.writeFile(filePath);
      
      const baseUrl = `${req.protocol}://${req.get("host")}`; 

      res.json({
        success: true,
        message: "Excel file created successfully.",
        downloadUrl: `${baseUrl}/${fileName}`,
      });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: "Failed to generate Excel file." });
    }
});

app.listen(port, () => console.log(`App listening on port ${port}`));