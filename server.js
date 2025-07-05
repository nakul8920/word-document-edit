const express = require('express');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const multer = require('multer');

const app = express();
app.use(express.static('public'));
app.use('/background image', express.static(path.join(__dirname, 'background image')));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Handle form submission and generate docx
app.post('/generate', multer().none(), (req, res) => {
    try {
        const data = req.body;
        const templatePath = path.join(__dirname, 'templates', 'Ats PR 1B Vs 1S.docx');
        
        // Check if template file exists
        if (!fs.existsSync(templatePath)) {
            console.error('Template file not found:', templatePath);
            return res.status(404).send('Template file not found');
        }
        
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        // Map form fields to placeholders
        doc.setData({
            SELLER_1_NAME: data.seller1Name || '',
            SELLER_1_PAN: data.seller1Pan,
            SELLER_1_AADHAR: data.seller1Aadhar,
            SELLER_1_HUSBAND_WIFE_DAUGHTER_NAME: data.seller1Relation,
            SELLER_1_ADDRESS: data.seller1Address,
            UNIT_TYPE: data.unitType,
            UNIT_SIZE: data.unitSize,
            UNIT_NO: data.unitNo,
            FLOOR_NO: data.floorNo,
            TOWER_NO: data.towerNo,
            PROPERTY_NAME: data.propertyName,
            PROPERTY_ADDRESS: data.propertyAddress,
            BUYER_1_NAME: data.buyer1Name,
            BUYER_1_PAN: data.buyer1Pan,
            BUYER_1_AADHAR: data.buyer1Aadhar,
            BUYER_1_HUSBAND_WIFE_DAUGHTER_NAME: data.buyer1Relation,
            BUYER_1_ADDRESS: data.buyer1Address,
            COMPANY_NAME: data.companyName,
            COMPANY_ADDRESS: data.companyAddress
        });

        doc.render();
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename=customized.docx'
        });
        res.send(buf);
    } catch (error) {
        console.error('Error in /generate:', error);
        res.status(500).send('Error generating document: ' + error.message);
    }
});

// Handle 2 Buyer, 2 Seller form submission
app.post('/generate2', multer().none(), (req, res) => {
    try {
        const data = req.body;
        console.log('Received data for /generate2:', data); // Debug log
        console.log('unitSize value:', data.unitSize); // Specific check for unitSize
        const templatePath = path.join(__dirname, 'templates', 'Ats PR 2B Vs 2S.docx');
        
        // Check if template file exists
        if (!fs.existsSync(templatePath)) {
            console.error('Template file not found:', templatePath);
            return res.status(404).send('Template file not found');
        }
        
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        // Map form fields to placeholders for 2 Buyer, 2 Seller
        const templateData = {
            SELLER_1_NAME: data.seller1Name || '',
            SELLER_1_PAN: data.seller1Pan,
            SELLER_1_AADHAR: data.seller1Aadhar,
            SELLER_1_HUSBAND_WIFE_DAUGHTER_NAME: data.seller1Relation,
            SELLER_2_NAME: data.seller2Name || '',
            SELLER_2_PAN: data.seller2Pan,
            SELLER_2_AADHAR: data.seller2Aadhar,
            SELLER_2_HUSBAND_WIFE_DAUGHTER_NAME: data.seller2Relation,
            SELLER_2_ADDRESS: data.seller2Address,
            UNIT_NO: data.unitNo,
            UNIT_SIZE: data.unitSize,
            UNIT_TYPE: data.unitType,
            FLOOR_NO: data.floorNo,
            TOWER_NO: data.towerNo,
            PROPERTY_NAME: data.propertyName,
            PROPERTY_ADDRESS: data.propertyAddress,
            BUYER_1_NAME: data.buyer1Name,
            BUYER_1_PAN: data.buyer1Pan,
            BUYER_1_AADHAR: data.buyer1Aadhar,
            BUYER_1_HUSBAND_WIFE_DAUGHTER_NAME: data.buyer1Relation,
            BUYER_2_NAME: data.buyer2Name,
            BUYER_2_PAN: data.buyer2Pan,
            BUYER_2_AADHAR: data.buyer2Aadhar,
            BUYER_2_HUSBAND_WIFE_DAUGHTER_NAME: data.buyer2Relation,
            BUYER_2_ADDRESS: data.buyer2Address,
            COMPANY_NAME: data.companyName,
            COMPANY_ADDRESS: data.companyAddress
        };
        
        console.log('Template data being set:', templateData); // Log the data being set
        doc.setData(templateData);

        doc.render();
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename=customized.docx'
        });
        res.send(buf);
    } catch (error) {
        console.error('Error in /generate2:', error);
        res.status(500).send('Error generating document: ' + error.message);
    }
});

// Handle 1 Buyer, 2 Seller form submission
app.post('/generate3', multer().none(), (req, res) => {
    try {
        const data = req.body;
        const templatePath = path.join(__dirname, 'templates', 'Ats PR 1B Vs 2S.docx');
        
        // Check if template file exists
        if (!fs.existsSync(templatePath)) {
            console.error('Template file not found:', templatePath);
            return res.status(404).send('Template file not found');
        }
        
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        // Map form fields to placeholders for 1 Buyer, 2 Seller
        doc.setData({
            SELLER_1_NAME: data.seller1Name || '',
            SELLER_1_PAN: data.seller1Pan,
            SELLER_1_AADHAR: data.seller1Aadhar,
            SELLER_1_HUSBAND_WIFE_DAUGHTER_NAME: data.seller1Relation,
            SELLER_2_NAME: data.seller2Name || '',
            SELLER_2_PAN: data.seller2Pan,
            SELLER_2_AADHAR: data.seller2Aadhar,
            SELLER_2_HUSBAND_WIFE_DAUGHTER_NAME: data.seller2Relation,
            SELLER_2_ADDRESS: data.seller2Address,
            UNIT_NO: data.unitNo,
            UNIT_SIZE: data.unitSize,
            UNIT_TYPE: data.unitType,
            FLOOR_NO: data.floorNo,
            TOWER_NO: data.towerNo,
            PROPERTY_NAME: data.propertyName,
            PROPERTY_ADDRESS: data.propertyAddress,
            BUYER_1_NAME: data.buyer1Name,
            BUYER_1_PAN: data.buyer1Pan,
            BUYER_1_AADHAR: data.buyer1Aadhar,
            BUYER_1_HUSBAND_WIFE_DAUGHTER_NAME: data.buyer1Relation,
            BUYER_1_ADDRESS: data.buyer1Address,
            COMPANY_NAME: data.companyName,
            COMPANY_ADDRESS: data.companyAddress
        });

        doc.render();
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename=customized.docx'
        });
        res.send(buf);
    } catch (error) {
        console.error('Error in /generate3:', error);
        res.status(500).send('Error generating document: ' + error.message);
    }
});

// Handle 2 Buyer, 1 Seller form submission
app.post('/generate4', multer().none(), (req, res) => {
    try {
        const data = req.body;
        const templatePath = path.join(__dirname, 'templates', 'Ats PR 2B Vs 1S.docx');
        
        // Check if template file exists
        if (!fs.existsSync(templatePath)) {
            console.error('Template file not found:', templatePath);
            return res.status(404).send('Template file not found');
        }
        
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        // Map form fields to placeholders for 2 Buyers, 1 Seller
        doc.setData({
            SELLER_1_NAME: data.seller1Name || '',
            SELLER_1_PAN: data.seller1Pan,
            SELLER_1_AADHAR: data.seller1Aadhar,
            SELLER_1_HUSBAND_WIFE_DAUGHTER_NAME: data.seller1Relation,
            SELLER_1_ADDRESS: data.seller1Address,
            UNIT_NO: data.unitNo,
            UNIT_SIZE: data.unitSize,
            UNIT_TYPE: data.unitType,
            FLOOR_NO: data.floorNo,
            TOWER_NO: data.towerNo,
            PROPERTY_NAME: data.propertyName,
            PROPERTY_ADDRESS: data.propertyAddress,
            BUYER_1_NAME: data.buyer1Name,
            BUYER_1_PAN: data.buyer1Pan,
            BUYER_1_AADHAR: data.buyer1Aadhar,
            BUYER_1_HUSBAND_WIFE_DAUGHTER_NAME: data.buyer1Relation,
            BUYER_2_NAME: data.buyer2Name,
            BUYER_2_PAN: data.buyer2Pan,
            BUYER_2_AADHAR: data.buyer2Aadhar,
            BUYER_2_HUSBAND_WIFE_DAUGHTER_NAME: data.buyer2Relation,
            BUYER_2_ADDRESS: data.buyer2Address,
            COMPANY_NAME: data.companyName,
            COMPANY_ADDRESS: data.companyAddress
        });

        doc.render();
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename=customized.docx'
        });
        res.send(buf);
    } catch (error) {
        console.error('Error in /generate4:', error);
        res.status(500).send('Error generating document: ' + error.message);
    }
});

// Add a test route to check if server is running
app.get('/test', (req, res) => {
    res.json({ message: 'Server is running successfully!' });
});

app.listen(3000, () => {
    console.log('Server running at http://localhost:3000');
    console.log('Test the server at: http://localhost:3000/test');
}); 