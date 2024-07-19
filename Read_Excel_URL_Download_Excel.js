//express API
//โค้ดสำหรับอ่านข้อมูลจากไฟล์ Excel และ Lookup Excel สำหรรับสร้าง Report ไฟล์ Excel และมีการทำให้สามารถ Download ไฟล์ได้โดยมีการ response URL
const express = require('express');
const router = express.Router();
const fs = require('fs');
const path = require('path');
const PORT = 443;
const ExcelJS = require('exceljs');

router.use(express.json());

// Function to format date
const formatDate = (dateString) => {
    const date = new Date(dateString);
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
};

// กำหนดเส้นทาง GET สำหรับ endpoint '/DATA/Hospital/Health/Rider'
router.get('/DATA/Hospital/Health/Rider', async (req, res) => {
    // ดึงค่าพารามิเตอร์จาก query string
    let hospital = req.query.hospital; // ชื่อโรงพยาบาล
    let date_start = req.query.date_start; // วันที่เริ่มต้นในรูปแบบ YYYY-MM-DD

    // ตัวแปรสำหรับเก็บผลรวมของราคาวางบิลและจำนวน order ของ rider และ iShip
    let totalDisbursementCost = 0; // ผลรวมของราคาวางบิล
    let riderCount = 0; // จำนวน order ของ rider
    let iShipCount = 0; // จำนวน order ของ iShip

    // แยกวันที่เริ่มต้นออกเป็นปี, เดือน, และวัน
    const dateParts = date_start.split(' '); // แยกวันที่และเวลาออกจากกัน
    const [year, month, day] = dateParts[0].split('-'); // แยกปี, เดือน, และวันจากวันที่
    const extractedMonth = month; // เก็บค่าเดือนในตัวแปร extractedMonth

    // สร้างออบเจ็กต์เพื่อแปลงเดือนเป็นภาษาไทย
    const monthNamesInThai = {
        '01': 'มกราคม',
        '02': 'กุมภาพันธ์',
        '03': 'มีนาคม',
        '04': 'เมษายน',
        '05': 'พฤษภาคม',
        '06': 'มิถุนายน',
        '07': 'กรกฎาคม',
        '08': 'สิงหาคม',
        '09': 'กันยายน',
        '10': 'ตุลาคม',
        '11': 'พฤศจิกายน',
        '12': 'ธันวาคม'
    };

    // แปลงชื่อเดือนเป็นภาษาไทย
    const monthNameInThai = monthNamesInThai[extractedMonth];
    // สร้างชื่อไฟล์ที่ประกอบด้วยชื่อโรงพยาบาลและเดือนเป็นภาษาไทย
    const namefile = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567`;

    try {
        // กำหนดเส้นทางของไฟล์ Excel ที่จะอ่าน
        const excelFilePath = `C:\\office\\Hospital\\New_Data\\${hospital}.xlsx`;
        const disbursementFilePath = `C:\\office\\Hospital\\FileLookUP\\disbursement_cost.xlsx`;

        // ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
        if (!fs.existsSync(excelFilePath) || !fs.existsSync(disbursementFilePath)) {
            return res.status(404).send('ไม่พบไฟล์ Excel ที่ระบุ');
        }

        // อ่านไฟล์ Excel เข้า Workbook
        const inputWorkbook = new ExcelJS.Workbook();
        await inputWorkbook.xlsx.readFile(excelFilePath);

        const inputWorksheet = inputWorkbook.getWorksheet('Sheet1'); // ปรับชื่อ Sheet ตามต้องการ
        const outputWorkbook = new ExcelJS.Workbook();
        const outputWorksheet = outputWorkbook.addWorksheet('Data');

        // เพิ่ม Header ให้กับไฟล์ output ที่จะสร้าง
        let row = 4; // กำหนดแถวเริ่มต้นของ Header
        outputWorksheet.getCell(`A${row}`).value = 'ลำดับ';
        outputWorksheet.getCell(`B${row}`).value = 'ชื่อผู้ใช้บริการ';
        outputWorksheet.getCell(`C${row}`).value = 'วันที่สร้างออเดอร์';
        outputWorksheet.getCell(`D${row}`).value = 'เลขที่การจัดส่ง';
        outputWorksheet.getCell(`E${row}`).value = 'เลขออเดอร์';
        outputWorksheet.getCell(`F${row}`).value = 'HN';
        outputWorksheet.getCell(`G${row}`).value = 'บริการขนส่ง';
        outputWorksheet.getCell(`H${row}`).value = 'สิทธิการรักษา';
        outputWorksheet.getCell(`I${row}`).value = 'ราคาวางบิล';
        outputWorksheet.getCell(`J${row}`).value = 'Status';

        // กำหนดชื่อไฟล์ Excel ที่จะสร้าง
        outputWorksheet.getCell('E2').value = `ประวัติการรอรับยา ${hospital} ประจำเดือน${monthNameInThai} 2567`;

        // อ่านไฟล์ disbursement cost เข้า Workbook
        const disbursementWorkbook = new ExcelJS.Workbook();
        await disbursementWorkbook.xlsx.readFile(disbursementFilePath);
        const disbursementWorksheet = disbursementWorkbook.getWorksheet('Sheet1'); // ปรับชื่อ Sheet ตามต้องการ

        // ประมวลผลข้อมูลจาก inputWorksheet ไปยัง outputWorksheet
        let rowCounter = 1; // ตัวนับแถว

        // วนลูปแต่ละแถวใน inputWorksheet
        inputWorksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // ข้ามแถวหัวเรื่อง

            const serviceRight = row.getCell('C').value; // สมมติว่าคอลัมน์ C มีสิทธิการรักษา
            const statusActive = row.getCell('E').value; // สมมติว่าคอลัมน์ E มีสถานะ
            const referenceNumber = row.getCell('F').value; // สมมติว่าคอลัมน์ F มีเลขอ้างอิง
            const orderDate = row.getCell('K').value; // สมมติว่าคอลัมน์ K มีวันที่สร้างออเดอร์
            const name = row.getCell('L').value; // สมมติว่าคอลัมน์ L มีชื่อผู้ใช้บริการ
            const order_no = row.getCell('G').value; // สมมติว่าคอลัมน์ G มีเลขออเดอร์
            const cardHN = row.getCell('N').value; // สมมติว่าคอลัมน์ N มี HN
            const order_type = row.getCell('D').value; // สมมติว่าคอลัมน์ D มีประเภทการจัดส่ง

            // รูปแบบวันที่ของออเดอร์
            const formattedOrderDate = formatDate(orderDate);

            // ค้นหาค่า disbursement cost ที่ตรงกับ hospital
            let disbursementCost = ''; // ตัวแปรสำหรับเก็บราคาวางบิล
            disbursementWorksheet.eachRow((disRow, disRowNumber) => {
                if (disRowNumber === 1) return; // ข้ามแถวหัวเรื่อง

                const disHospital = disRow.getCell('B').value; // สมมติว่าคอลัมน์ B มีชื่อโรงพยาบาล

                if (disHospital === hospital) {
                    disbursementCost = disRow.getCell('C').value; // สมมติว่าคอลัมน์ C มีราคาวางบิล
                    totalDisbursementCost += disbursementCost; // รวมยอดราคาวางบิล
                }
            });

            // ตรวจสอบสิทธิการรักษาและสถานะ
            if ((serviceRight === 'บัตรทอง' || serviceRight === 'UC') &&
                (statusActive === 'success' || statusActive === 'waiting_transport' || statusActive === 'waiting_conference')) {
                outputWorksheet.addRow([rowCounter++, name, formattedOrderDate, order_no, referenceNumber, cardHN, order_type, serviceRight, disbursementCost]); // เพิ่มแถวพร้อมตัวนับ, เลขอ้างอิง, และวันที่ออเดอร์

                // นับประเภทการจัดส่ง rider และ non-rider
                if (order_type.toLowerCase() === 'rider') {
                    riderCount++; // เพิ่มจำนวน rider
                } else {
                    iShipCount++; // เพิ่มจำนวน iShip
                }
            }
        });

        // สร้าง formatter เพื่อรูปแบบการแสดงผลตัวเลข
        const formatter = new Intl.NumberFormat('th-TH', {
            currency: 'THB',
        });

        // เพิ่มแถวรวมยอดราคาวางบิล
        const lastRowNumber = outputWorksheet.rowCount + 1; // กำหนดแถวสุดท้าย
        outputWorksheet.mergeCells(`A${lastRowNumber}:H${lastRowNumber}`); // รวมเซลล์ตั้งแต่คอลัมน์ A ถึง H ในแถวสุดท้าย
        outputWorksheet.getCell(`A${lastRowNumber}`).alignment = { horizontal: 'center' }; // จัดแนวกลางของเซลล์
        outputWorksheet.getCell(`A${lastRowNumber}`).value = 'รวม'; // กำหนดค่า 'รวม' ในเซลล์
        outputWorksheet.getCell(`I${lastRowNumber}`).value = formatter.format(totalDisbursementCost); // กำหนดค่าผลรวมในเซลล์ I แถวสุดท้าย

        // เพิ่มจำนวน rider และ iShip
        outputWorksheet.getCell(`B${lastRowNumber + 2}`).value = `Rider: ${riderCount}`; // กำหนดจำนวน rider ในเซลล์
        outputWorksheet.getCell(`B${lastRowNumber + 3}`).value = `iShip: ${iShipCount}`; // กำหนดจำนวน iShip ในเซลล์
        outputWorksheet.getCell(`C${lastRowNumber + 2}`).value = `${riderCount}`; // กำหนดจำนวน rider ในเซลล์
        outputWorksheet.getCell(`C${lastRowNumber + 3}`).value = `${iShipCount}`; // กำหนดจำนวน iShip ในเซลล์

        // บันทึกไฟล์ outputWorkbook
        const filePath = path.join(__dirname, `${namefile}.xlsx`); // กำหนดเส้นทางไฟล์
        await outputWorkbook.xlsx.writeFile(filePath); // บันทึกไฟล์

        // กำหนด URL สำหรับดาวน์โหลดไฟล์
        const downloadURL = `http://localhost:${PORT}/downloadExcel/${namefile}`;
        res.send({ downloadURL }); // ส่ง URL ให้กับผู้ใช้งาน
    } catch (error) {
        console.error('เกิดข้อผิดพลาดในการดำเนินการ:', error); // แสดงข้อผิดพลาดใน console
        res.status(500).send('ขออภัย เกิดข้อผิดพลาดในการดำเนินการ'); // ส่งข้อความแสดงข้อผิดพลาด
    }
});

// กำหนดเส้นทาง GET สำหรับ endpoint '/downloadExcel/:filename'
router.get('/downloadExcel/:filename', (req, res) => {
    const { filename } = req.params; // ดึงชื่อไฟล์จากพารามิเตอร์
    const file = `${__dirname}/${filename}.xlsx`; // กำหนดเส้นทางไฟล์
    res.download(file, () => { // ส่งไฟล์ให้ผู้ใช้งานดาวน์โหลด
        fs.unlink(file, (err) => { // ลบไฟล์หลังจากดาวน์โหลดเสร็จ
            if (err) {
                console.error('เกิดข้อผิดพลาดในการลบไฟล์:', err); // แสดงข้อผิดพลาดใน console
            } else {
                console.log('ไฟล์ถูกลบแล้ว:', file); // แสดงข้อความว่าไฟล์ถูกลบแล้ว
            }
        });
    });
});


module.exports = router;
