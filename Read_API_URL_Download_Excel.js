//express API
//โค้ดสำหรับอ่านข้อมูลจาก API ต้นทางและนำข้อมูลจาก API มาทำการจัดรูปแบบใหม่ก่อนจะก่อนจะสร้าง Report เป็นไฟล์ Excel และทำการ response URL สำหรับ Dowmload ไฟล์ Report
const express = require('express');
const axios = require('axios');
const router = express();
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

router.get('/DATA/Hospital/Health/Rider', async (req, res) => {
    // กำหนดเส้นทาง HTTP GET สำหรับ URL '/DATA/Hospital/Health/Rider' โดยใช้ฟังก์ชัน async เพื่อรองรับการทำงานแบบอะซิงค์

    let hospital = req.param("hospital");
    // ดึงค่าพารามิเตอร์ 'hospital' จาก URL query และเก็บไว้ในตัวแปร 'hospital'

    let date_start = req.param("date_start");
    // ดึงค่าพารามิเตอร์ 'date_start' จาก URL query และเก็บไว้ในตัวแปร 'date_start'

    let date_end = req.param("date_end");
    // ดึงค่าพารามิเตอร์ 'date_end' จาก URL query และเก็บไว้ในตัวแปร 'date_end'

    let totalDisbursementCost = 0;
    // กำหนดตัวแปร 'totalDisbursementCost' เป็น 0 เพื่อใช้เก็บยอดรวมของ 'disbursement_cost' ที่จะคำนวณในภายหลัง

    const BEARER_TOKEN = ''; // ใส่ Bearer Token ที่นี่
    // กำหนดค่าตัวแปร Bearer Token สำหรับการยืนยันตัวตนในการเรียกใช้ API ที่ต้องการการยืนยันตัวตน

    const dateParts = date_start.split(' ');
    // แยกวันที่ในตัวแปร 'date_start' ออกเป็นสองส่วน (วันและเวลา) โดยใช้ช่องว่าง (' ') เป็นตัวคั่น

    const [year, month, day] = dateParts[0].split('-');
    // แยกปี, เดือน, และวัน จากส่วนที่ 1 ของวันที่ (รูปแบบ YYYY-MM-DD) โดยใช้เครื่องหมาย '-' เป็นตัวคั่น

    const extractedMonth = month; // This is the '05' part you need
    // เก็บค่าเดือนที่แยกออกมาในตัวแปร 'extractedMonth' (เช่น '05')

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
    // สร้างวัตถุ 'monthNamesInThai' ที่เก็บชื่อเดือนในภาษาไทยตามค่าเดือนที่ได้จาก 'extractedMonth'

    const monthNameInThai = monthNamesInThai[extractedMonth];
    // ดึงชื่อเดือนในภาษาไทยจาก 'monthNamesInThai' ตามค่าของ 'extractedMonth'

    const namefile = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567`;
    // สร้างชื่อไฟล์ Excel โดยรวมชื่อโรงพยาบาลและชื่อเดือนในภาษาไทยตามที่กำหนด

    try {
        const medicalRights = ['บัตรทอง', 'UC'];
        // สร้างอาร์เรย์ 'medicalRights' เพื่อเก็บประเภทสิทธิบัตรที่ต้องการดึงข้อมูล

        let combinedData = [];
        // สร้างอาร์เรย์ 'combinedData' เพื่อเก็บข้อมูลที่รวมจากหลายแหล่ง

        for (const rights of medicalRights) {
            // ใช้ลูป for เพื่อดึงข้อมูลสำหรับแต่ละประเภทสิทธิ

            const response = await axios.get('https://dashboard.manageai.co.th/gateway/api/v1/telepharmacy/transactions', {
                params: {
                    order_type: '',
                    hospital: hospital,
                    status_active: 'success,waiting_transport,waiting_conference',
                    medical_rights: rights,
                    start_date: date_start,
                    end_date: date_end
                },
                headers: {
                    Authorization: `Bearer ${BEARER_TOKEN}`
                }
            });
            // ใช้ axios เพื่อส่งคำขอ HTTP GET ไปยัง API ที่กำหนด โดยระบุพารามิเตอร์ที่จำเป็นและ Bearer Token สำหรับการยืนยันตัวตน

            combinedData = combinedData.concat(response.data.data);
            // รวมข้อมูลที่ดึงได้จาก API เข้าไปใน 'combinedData' โดยการใช้เมธอด concat

        }

        // Filter out entries with transport_process set to FALSE
        const filteredData = combinedData.filter(item => item.transport_process !== false);
        // กรองข้อมูลใน 'combinedData' โดยเลือกเฉพาะรายการที่มีค่า 'transport_process' ไม่เท่ากับ false

        const sortedData = filteredData.sort((a, b) => {
            const dateA = new Date(a.date_receive_parcel_order);
            const dateB = new Date(b.date_receive_parcel_order);

            if (dateA.getDate().toString().slice(-2) === '01' && dateB.getDate().toString().slice(-2) !== '01') {
                return -1;
            } else if (dateA.getDate().toString().slice(-2) !== '01' && dateB.getDate().toString().slice(-2) === '01') {
                return 1;
            } else {
                return dateA - dateB;
            }
        });
        // เรียงลำดับข้อมูลตามวันที่ 'date_receive_parcel_order' โดยให้วันที่ 1 ของเดือนอยู่ก่อน จากนั้นเรียงลำดับวันที่ตามปกติ

        const nameAndDateSet = new Set();
        // สร้าง Set ใหม่ชื่อว่า 'nameAndDateSet' สำหรับเก็บข้อมูลที่ไม่ซ้ำกัน

        const duplicates = [];
        // สร้างอาร์เรย์ 'duplicates' เพื่อเก็บรายการที่ซ้ำกัน

        const dataWithNoDuplicates = sortedData.filter((item, index) => {
            const formattedOrderDate = formatDate(item.order_date);
            const key = `${item.name}-${formattedOrderDate}`;

            if (nameAndDateSet.has(key)) {
                duplicates.push(item);
                return false;
            } else {
                nameAndDateSet.add(key);
                return true;
            }
        }).map((item, index) => {
            return { ...item, no: index + 1, formattedOrderDate: formatDate(item.order_date) };
        });
        // กรองข้อมูลที่ซ้ำกันออกจาก 'sortedData' โดยใช้ Set เพื่อตรวจสอบ และจัดรูปแบบวันที่ จากนั้นเพิ่มหมายเลขลำดับลงในข้อมูลที่ไม่ซ้ำกัน

        const workbook = new ExcelJS.Workbook();
        // สร้างวัตถุ 'workbook' ใหม่สำหรับการสร้างไฟล์ Excel

        const worksheet = workbook.addWorksheet('Data');
        // เพิ่ม worksheet ใหม่ชื่อว่า 'Data' ลงใน workbook

        let row = 4;
        // กำหนดหมายเลขแถวเริ่มต้นเป็น 4 สำหรับการเพิ่มข้อมูล

        worksheet.getCell(`A${row}`).value = 'ลำดับ';
        worksheet.getCell(`B${row}`).value = 'ชื่อผู้ใช้บริการ';
        worksheet.getCell(`C${row}`).value = 'วันที่สร้างออเดอร์';
        worksheet.getCell(`D${row}`).value = 'เลขที่การจัดส่ง';
        worksheet.getCell(`E${row}`).value = 'เลขออเดอร์';
        worksheet.getCell(`F${row}`).value = 'HN';
        worksheet.getCell(`G${row}`).value = 'บริการขนส่ง';
        worksheet.getCell(`H${row}`).value = 'สิทธิการรักษา';
        worksheet.getCell(`I${row}`).value = 'ราคาวางบิล';
        // กำหนดชื่อคอลัมน์ในแถวที่ 4 ของ worksheet

        dataWithNoDuplicates.forEach(item => {
            row++;
            // เพิ่มหมายเลขแถวทีละ 1 เพื่อเพิ่มข้อมูลในแต่ละแถว

            worksheet.getCell(`A${row}`).value = item.no;
            worksheet.getCell(`B${row}`).value = item.name;
            worksheet.getCell(`C${row}`).value = item.formattedOrderDate; // ใช้วันที่ที่จัดรูปแบบแล้ว
            worksheet.getCell(`D${row}`).value = item.order_no;
            worksheet.getCell(`E${row}`).value = item.reference_id;
            worksheet.getCell(`F${row}`).value = item.cardHN;
            worksheet.getCell(`G${row}`).value = item.order_type;
            worksheet.getCell(`H${row}`).value = item.medical_rights;
            worksheet.getCell(`I${row}`).value = item.disbursement_cost;
            // เพิ่มข้อมูลแต่ละรายการลงใน worksheet โดยเพิ่มข้อมูลในแต่ละเซลล์

            totalDisbursementCost += item.disbursement_cost;
            // คำนวณยอดรวมของ 'disbursement_cost' โดยบวกค่า 'disbursement_cost' ของแต่ละรายการเข้าไปใน 'totalDisbursementCost'
        });

        const formatter = new Intl.NumberFormat('th-TH', {
            currency: 'THB',
        });
        // สร้างฟอร์แมตเตอร์สำหรับจัดรูปแบบตัวเลขในรูปแบบของสกุลเงินไทย

        worksheet.mergeCells(`A${row + 1}:H${row + 1}`);
        // รวมเซลล์ A ถึง H ในแถวถัดไป (แถวที่ 'row + 1') เพื่อแสดงผลรวม

        worksheet.getCell(`A${row + 1}`).alignment = { horizontal: 'center' };
        // จัดแนวข้อความในเซลล์ A ของแถว 'row + 1' ให้อยู่กลาง

        worksheet.getCell(`A${row + 1}`).value = 'รวม';
        // กำหนดข้อความ 'รวม' ในเซลล์ A ของแถว 'row + 1'

        worksheet.getCell(`I${row + 1}`).value = formatter.format(totalDisbursementCost);
        // กำหนดยอดรวมของ 'totalDisbursementCost' ในเซลล์ I ของแถว 'row + 1' โดยใช้ฟอร์แมตเตอร์สำหรับการจัดรูปแบบ

        worksheet.getCell('E2').value = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567`;
        // กำหนดชื่อไฟล์ในเซลล์ E2 ของ worksheet

        const filePath = path.join(__dirname, `${namefile}.xlsx`);
        // สร้างพาธสำหรับการบันทึกไฟล์ Excel โดยใช้ชื่อไฟล์ที่กำหนด

        await workbook.xlsx.writeFile(filePath);
        // บันทึกไฟล์ Excel ไปยังพาธที่กำหนด

        const downloadURL = `http://localhost:${PORT}/downloadExcel/${namefile}`;
        // สร้าง URL สำหรับการดาวน์โหลดไฟล์ Excel

        res.send({ downloadURL });
        // ส่ง URL สำหรับการดาวน์โหลดไฟล์ Excel กลับไปยังผู้ร้องขอ
    } catch (error) {
        console.error('เกิดข้อผิดพลาดในการเรียก API:', error);
        // แสดงข้อผิดพลาดในกรณีที่เกิดข้อผิดพลาดในการดึงข้อมูลหรือบันทึกไฟล์

        res.status(500).send('ขออภัย เกิดข้อผิดพลาดในการเรียก API');
        // ส่งข้อความข้อผิดพลาดกลับไปยังผู้ร้องขอ
    }
});

router.get('/DATA/Hospital/Health/Rider/Add', async (req, res) => {
    // กำหนดเส้นทาง HTTP GET สำหรับ URL '/DATA/Hospital/Health/Rider/Add' โดยใช้ฟังก์ชัน async เพื่อรองรับการทำงานแบบอะซิงค์

    let hospital = req.param("hospital");
    // ดึงค่าพารามิเตอร์ 'hospital' จาก URL query และเก็บไว้ในตัวแปร 'hospital'

    let date_start = req.param("date_start");
    // ดึงค่าพารามิเตอร์ 'date_start' จาก URL query และเก็บไว้ในตัวแปร 'date_start'

    let date_end = req.param("date_end");
    // ดึงค่าพารามิเตอร์ 'date_end' จาก URL query และเก็บไว้ในตัวแปร 'date_end'

    let totalDisbursementCost = 0;
    // กำหนดตัวแปร 'totalDisbursementCost' เป็น 0 เพื่อใช้เก็บยอดรวมของ 'disbursement_cost' ที่จะคำนวณในภายหลัง

    const BEARER_TOKEN = ''; // ใส่ Bearer Token ที่นี่
    // กำหนดค่าตัวแปร Bearer Token สำหรับการยืนยันตัวตนในการเรียกใช้ API ที่ต้องการการยืนยันตัวตน

    const dateParts = date_start.split(' ');
    // แยกวันที่ในตัวแปร 'date_start' ออกเป็นสองส่วน (วันและเวลา) โดยใช้ช่องว่าง (' ') เป็นตัวคั่น

    const [year, month, day] = dateParts[0].split('-');
    // แยกปี, เดือน, และวัน จากส่วนที่ 1 ของวันที่ (รูปแบบ YYYY-MM-DD) โดยใช้เครื่องหมาย '-' เป็นตัวคั่น

    const extractedMonth = month; // This is the '05' part you need
    // เก็บค่าเดือนที่แยกออกมาในตัวแปร 'extractedMonth' (เช่น '05')

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
    // สร้างวัตถุ 'monthNamesInThai' ที่เก็บชื่อเดือนในภาษาไทยตามค่าเดือนที่ได้จาก 'extractedMonth'

    const monthNameInThai = monthNamesInThai[extractedMonth];
    // ดึงชื่อเดือนในภาษาไทยจาก 'monthNamesInThai' ตามค่าของ 'extractedMonth'

    const namefile = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567`;
    // สร้างชื่อไฟล์ Excel โดยรวมชื่อโรงพยาบาลและชื่อเดือนในภาษาไทยตามที่กำหนด

    try {
        const medicalRights = ['บัตรทอง', 'UC'];
        // สร้างอาร์เรย์ 'medicalRights' เพื่อเก็บประเภทสิทธิบัตรที่ต้องการดึงข้อมูล

        let combinedData = [];
        // สร้างอาร์เรย์ 'combinedData' เพื่อเก็บข้อมูลที่รวมจากหลายแหล่ง

        for (const rights of medicalRights) {
            // ใช้ลูป for เพื่อดึงข้อมูลสำหรับแต่ละประเภทสิทธิ

            const response = await axios.get('https://dashboard.manageai.co.th/gateway/api/v1/telepharmacy/transactions', {
                params: {
                    order_type: '',
                    hospital: hospital,
                    status_active: 'success,waiting_transport,waiting_conference',
                    medical_rights: rights,
                    start_date: date_start,
                    end_date: date_end
                },
                headers: {
                    Authorization: `Bearer ${BEARER_TOKEN}`
                }
            });
            // ใช้ axios เพื่อส่งคำขอ HTTP GET ไปยัง API ที่กำหนด โดยระบุพารามิเตอร์ที่จำเป็นและ Bearer Token สำหรับการยืนยันตัวตน

            combinedData = combinedData.concat(response.data.data);
            // รวมข้อมูลที่ดึงได้จาก API เข้าไปใน 'combinedData' โดยการใช้เมธอด concat

        }

        // Filter out entries with transport_process set to FALSE
        const filteredData = combinedData.filter(item => item.transport_process !== false);
        // กรองข้อมูลใน 'combinedData' โดยเลือกเฉพาะรายการที่มีค่า 'transport_process' ไม่เท่ากับ false

        const sortedData = filteredData.sort((a, b) => {
            const dateA = new Date(a.date_receive_parcel_order);
            const dateB = new Date(b.date_receive_parcel_order);

            if (dateA.getDate().toString().slice(-2) === '01' && dateB.getDate().toString().slice(-2) !== '01') {
                return -1;
            } else if (dateA.getDate().toString().slice(-2) !== '01' && dateB.getDate().toString().slice(-2) === '01') {
                return 1;
            } else {
                return dateA - dateB;
            }
        });
        // เรียงลำดับข้อมูลตามวันที่ 'date_receive_parcel_order' โดยให้วันที่ 1 ของเดือนอยู่ก่อน จากนั้นเรียงลำดับวันที่ตามปกติ

        const nameAndDateSet = new Set();
        // สร้าง Set ใหม่ชื่อว่า 'nameAndDateSet' สำหรับเก็บข้อมูลที่ไม่ซ้ำกัน

        const duplicates = [];
        // สร้างอาร์เรย์ 'duplicates' เพื่อเก็บรายการที่ซ้ำกัน

        const dataWithNoDuplicates = sortedData.filter((item, index) => {
            const formattedOrderDate = formatDate(item.order_date);
            const key = `${item.name}-${formattedOrderDate}`;

            if (nameAndDateSet.has(key)) {
                duplicates.push(item);
                return false;
            } else {
                nameAndDateSet.add(key);
                return true;
            }
        }).map((item, index) => {
            return { ...item, no: index + 1, formattedOrderDate: formatDate(item.order_date) };
        });
        // กรองข้อมูลที่ซ้ำกันออกจาก 'sortedData' โดยใช้ Set เพื่อตรวจสอบ และจัดรูปแบบวันที่ จากนั้นเพิ่มหมายเลขลำดับลงในข้อมูลที่ไม่ซ้ำกัน

        const workbook = new ExcelJS.Workbook();
        // สร้างวัตถุ 'workbook' ใหม่สำหรับการสร้างไฟล์ Excel

        const worksheet = workbook.addWorksheet('Data');
        // เพิ่ม worksheet ใหม่ชื่อว่า 'Data' ลงใน workbook

        let row = 4;
        // กำหนดหมายเลขแถวเริ่มต้นเป็น 4 สำหรับการเพิ่มข้อมูล

        worksheet.getCell(`A${row}`).value = 'ลำดับ';
        worksheet.getCell(`B${row}`).value = 'ชื่อผู้ใช้บริการ';
        worksheet.getCell(`C${row}`).value = 'วันที่สร้างออเดอร์';
        worksheet.getCell(`D${row}`).value = 'เลขที่การจัดส่ง';
        worksheet.getCell(`E${row}`).value = 'เลขออเดอร์';
        worksheet.getCell(`F${row}`).value = 'HN';
        worksheet.getCell(`G${row}`).value = 'บริการขนส่ง';
        worksheet.getCell(`H${row}`).value = 'สิทธิการรักษา';
        worksheet.getCell(`I${row}`).value = 'ราคาวางบิล';
        worksheet.getCell(`J${row}`).value = 'box_size_detail';
        worksheet.getCell(`K${row}`).value = 'box_size_name';
        // กำหนดชื่อคอลัมน์ในแถวที่ 4 ของ worksheet

        dataWithNoDuplicates.forEach(item => {
            row++;
            // เพิ่มหมายเลขแถวทีละ 1 เพื่อเพิ่มข้อมูลในแต่ละแถว

            worksheet.getCell(`A${row}`).value = item.no;
            worksheet.getCell(`B${row}`).value = item.name;
            worksheet.getCell(`C${row}`).value = item.formattedOrderDate; // ใช้วันที่ที่จัดรูปแบบแล้ว
            worksheet.getCell(`D${row}`).value = item.order_no;
            worksheet.getCell(`E${row}`).value = item.reference_id;
            worksheet.getCell(`F${row}`).value = item.cardHN;
            worksheet.getCell(`G${row}`).value = item.order_type;
            worksheet.getCell(`H${row}`).value = item.medical_rights;
            worksheet.getCell(`I${row}`).value = item.disbursement_cost;
            worksheet.getCell(`J${row}`).value = item.box_size_detail;
            worksheet.getCell(`K${row}`).value = item.box_size_name;
            // เพิ่มข้อมูลแต่ละรายการลงใน worksheet โดยเพิ่มข้อมูลในแต่ละเซลล์

            totalDisbursementCost += item.disbursement_cost;
            // คำนวณยอดรวมของ 'disbursement_cost' โดยบวกค่า 'disbursement_cost' ของแต่ละรายการเข้าไปใน 'totalDisbursementCost'
        });

        const formatter = new Intl.NumberFormat('th-TH', {
            currency: 'THB',
        });
        // สร้างตัวจัดรูปแบบสำหรับสกุลเงินไทย (THB) โดยใช้ Intl.NumberFormat

        worksheet.mergeCells(`A${row + 1}:H${row + 1}`);
        worksheet.getCell(`A${row + 1}`).alignment = { horizontal: 'center' };
        worksheet.getCell(`A${row + 1}`).value = 'รวม';
        worksheet.getCell(`I${row + 1}`).value = formatter.format(totalDisbursementCost);
        // รวมเซลล์ A ถึง H ในแถวที่ 'row + 1' เพื่อแสดงยอดรวม และจัดแนวข้อความให้อยู่กลางในเซลล์ A ของแถวดังกล่าว

        worksheet.getCell('E2').value = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567`;
        // กำหนดข้อความที่เซลล์ E2 เพื่อใช้เป็นชื่อไฟล์ใน worksheet

        const filePath = path.join(__dirname, `${namefile}.xlsx`);
        // สร้างพาธสำหรับการบันทึกไฟล์ Excel โดยใช้ชื่อไฟล์ที่กำหนด

        await workbook.xlsx.writeFile(filePath);
        // บันทึกไฟล์ Excel ไปยังพาธที่กำหนด

        const downloadURL = `http://localhost:${PORT}/downloadExcel/${namefile}`;
        // สร้าง URL สำหรับการดาวน์โหลดไฟล์ Excel

        res.send({ downloadURL });
        // ส่ง URL สำหรับการดาวน์โหลดไฟล์ Excel กลับไปยังผู้ร้องขอ
    } catch (error) {
        console.error('เกิดข้อผิดพลาดในการเรียก API:', error);
        // แสดงข้อผิดพลาดในกรณีที่เกิดข้อผิดพลาดในการดึงข้อมูลหรือบันทึกไฟล์

        res.status(500).send('ขออภัย เกิดข้อผิดพลาดในการเรียก API');
        // ส่งข้อความข้อผิดพลาดกลับไปยังผู้ร้องขอ
    }
});

router.get('/DATA/Hospital/Health/Rider/Exclude', async (req, res) => {
    // กำหนดเส้นทาง HTTP GET สำหรับ URL '/DATA/Hospital/Health/Rider/Exclude' โดยใช้ฟังก์ชัน async เพื่อรองรับการทำงานแบบอะซิงค์

    let hospital = req.param("hospital");
    // ดึงค่าพารามิเตอร์ 'hospital' จาก URL query และเก็บไว้ในตัวแปร 'hospital'

    let date_start = req.param("date_start");
    // ดึงค่าพารามิเตอร์ 'date_start' จาก URL query และเก็บไว้ในตัวแปร 'date_start'

    let date_end = req.param("date_end");
    // ดึงค่าพารามิเตอร์ 'date_end' จาก URL query และเก็บไว้ในตัวแปร 'date_end'

    let totalDisbursementCost = 0;
    // กำหนดตัวแปร 'totalDisbursementCost' เป็น 0 เพื่อใช้เก็บยอดรวมของ 'disbursement_cost' ที่จะคำนวณในภายหลัง

    const BEARER_TOKEN = ''; // ใส่ Bearer Token ที่นี่
    // กำหนดค่าตัวแปร Bearer Token สำหรับการยืนยันตัวตนในการเรียกใช้ API ที่ต้องการการยืนยันตัวตน

    const dateParts = date_start.split(' ');
    // แยกวันที่ในตัวแปร 'date_start' ออกเป็นสองส่วน (วันและเวลา) โดยใช้ช่องว่าง (' ') เป็นตัวคั่น

    const [year, month, day] = dateParts[0].split('-');
    // แยกปี, เดือน, และวัน จากส่วนที่ 1 ของวันที่ (รูปแบบ YYYY-MM-DD) โดยใช้เครื่องหมาย '-' เป็นตัวคั่น

    const extractedMonth = month; // This is the '05' part you need
    // เก็บค่าเดือนที่แยกออกมาในตัวแปร 'extractedMonth' (เช่น '05')

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
    // สร้างวัตถุ 'monthNamesInThai' ที่เก็บชื่อเดือนในภาษาไทยตามค่าเดือนที่ได้จาก 'extractedMonth'

    const monthNameInThai = monthNamesInThai[extractedMonth];
    // ดึงชื่อเดือนในภาษาไทยจาก 'monthNamesInThai' ตามค่าของ 'extractedMonth'

    const namefile = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567 (ไม่มีข้อมูล บัตรทอง และ UC)`;
    // สร้างชื่อไฟล์ Excel โดยรวมชื่อโรงพยาบาลและชื่อเดือนในภาษาไทยตามที่กำหนด พร้อมเพิ่มข้อความระบุว่าไม่มีข้อมูลสิทธิบัตรทองและ UC

    try {
        const excludeRights = ['บัตรทอง', 'UC'];
        // สร้างอาร์เรย์ 'excludeRights' เพื่อเก็บประเภทสิทธิบัตรที่ต้องการกรองออก

        let combinedData = [];
        // สร้างอาร์เรย์ 'combinedData' เพื่อเก็บข้อมูลที่รวมจากหลายแหล่ง

        const response = await axios.get('https://dashboard.manageai.co.th/gateway/api/v1/telepharmacy/transactions', {
            params: {
                order_type: '',
                hospital: hospital,
                status_active: 'success,waiting_transport,waiting_conference',
                start_date: date_start,
                end_date: date_end
            },
            headers: {
                Authorization: `Bearer ${BEARER_TOKEN}`
            }
        });
        // ใช้ axios เพื่อส่งคำขอ HTTP GET ไปยัง API ที่กำหนด โดยระบุพารามิเตอร์ที่จำเป็นและ Bearer Token สำหรับการยืนยันตัวตน

        combinedData = response.data.data.filter(item => !excludeRights.includes(item.medical_rights));
        // กรองข้อมูลใน response.data.data โดยเลือกเฉพาะรายการที่ไม่รวมอยู่ใน 'excludeRights'

        // Filter out entries with transport_process set to FALSE
        const filteredData = combinedData.filter(item => item.transport_process !== false);
        // กรองข้อมูลที่เหลือออกจาก 'combinedData' โดยเลือกเฉพาะรายการที่มีค่า 'transport_process' ไม่เท่ากับ false

        const sortedData = filteredData.sort((a, b) => {
            const dateA = new Date(a.date_receive_parcel_order);
            const dateB = new Date(b.date_receive_parcel_order);

            if (dateA.getDate().toString().slice(-2) === '01' && dateB.getDate().toString().slice(-2) !== '01') {
                return -1;
            } else if (dateA.getDate().toString().slice(-2) !== '01' && dateB.getDate().toString().slice(-2) === '01') {
                return 1;
            } else {
                return dateA - dateB;
            }
        });
        // เรียงลำดับข้อมูลตามวันที่ 'date_receive_parcel_order' โดยให้วันที่ 1 ของเดือนอยู่ก่อน จากนั้นเรียงลำดับวันที่ตามปกติ

        const nameAndDateSet = new Set();
        // สร้าง Set ใหม่ชื่อว่า 'nameAndDateSet' สำหรับเก็บข้อมูลที่ไม่ซ้ำกัน

        const duplicates = [];
        // สร้างอาร์เรย์ 'duplicates' เพื่อเก็บรายการที่ซ้ำกัน

        const dataWithNoDuplicates = sortedData.filter((item, index) => {
            const formattedOrderDate = formatDate(item.order_date);
            const key = `${item.name}-${formattedOrderDate}`;

            if (nameAndDateSet.has(key)) {
                duplicates.push(item);
                return false;
            } else {
                nameAndDateSet.add(key);
                return true;
            }
        }).map((item, index) => {
            return { ...item, no: index + 1, formattedOrderDate: formatDate(item.order_date) };
        });
        // กรองข้อมูลที่ซ้ำกันออกจาก 'sortedData' โดยใช้ Set เพื่อตรวจสอบ และจัดรูปแบบวันที่ จากนั้นเพิ่มหมายเลขลำดับลงในข้อมูลที่ไม่ซ้ำกัน

        const workbook = new ExcelJS.Workbook();
        // สร้างวัตถุ 'workbook' ใหม่สำหรับการสร้างไฟล์ Excel

        const worksheet = workbook.addWorksheet('Data');
        // เพิ่ม worksheet ใหม่ชื่อว่า 'Data' ลงใน workbook

        let row = 4;
        // กำหนดหมายเลขแถวเริ่มต้นเป็น 4 สำหรับการเพิ่มข้อมูล

        worksheet.getCell(`A${row}`).value = 'ลำดับ';
        worksheet.getCell(`B${row}`).value = 'ชื่อผู้ใช้บริการ';
        worksheet.getCell(`C${row}`).value = 'วันที่สร้างออเดอร์';
        worksheet.getCell(`D${row}`).value = 'เลขที่การจัดส่ง';
        worksheet.getCell(`E${row}`).value = 'เลขออเดอร์';
        worksheet.getCell(`F${row}`).value = 'HN';
        worksheet.getCell(`G${row}`).value = 'บริการขนส่ง';
        worksheet.getCell(`H${row}`).value = 'สิทธิการรักษา';
        // กำหนดชื่อคอลัมน์ในแถวที่ 4 ของ worksheet

        dataWithNoDuplicates.forEach(item => {
            row++;
            // เพิ่มหมายเลขแถวทีละ 1 เพื่อเพิ่มข้อมูลในแต่ละแถว

            worksheet.getCell(`A${row}`).value = item.no;
            worksheet.getCell(`B${row}`).value = item.name;
            worksheet.getCell(`C${row}`).value = item.formattedOrderDate; // ใช้วันที่ที่จัดรูปแบบแล้ว
            worksheet.getCell(`D${row}`).value = item.order_no;
            worksheet.getCell(`E${row}`).value = item.reference_id;
            worksheet.getCell(`F${row}`).value = item.cardHN;
            worksheet.getCell(`G${row}`).value = item.order_type;
            worksheet.getCell(`H${row}`).value = item.medical_rights;
            // เพิ่มข้อมูลแต่ละรายการลงใน worksheet โดยเพิ่มข้อมูลในแต่ละเซลล์

            totalDisbursementCost += item.disbursement_cost;
            // คำนวณยอดรวมของ 'disbursement_cost' โดยบวกค่า 'disbursement_cost' ของแต่ละรายการเข้าไปใน 'totalDisbursementCost'
        });

        const formatter = new Intl.NumberFormat('th-TH', {
            currency: 'THB',
        });
        // สร้างตัวจัดรูปแบบสำหรับสกุลเงินไทย (THB) โดยใช้ Intl.NumberFormat

        worksheet.getCell('E2').value = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567 (ไม่มีข้อมูล บัตรทอง และ UC)`;
        // กำหนดข้อความที่เซลล์ E2 เพื่อใช้เป็นชื่อไฟล์ใน worksheet

        const filePath = path.join(__dirname, `${namefile}.xlsx`);
        // สร้างพาธสำหรับการบันทึกไฟล์ Excel โดยใช้ชื่อไฟล์ที่กำหนด

        await workbook.xlsx.writeFile(filePath);
        // บันทึกไฟล์ Excel ไปยังพาธที่กำหนด

        const downloadURL = `http://localhost:${PORT}/downloadExcel/${namefile}`;
        // สร้าง URL สำหรับการดาวน์โหลดไฟล์ Excel

        res.send({ downloadURL });
        // ส่ง URL สำหรับการดาวน์โหลดไฟล์ Excel กลับไปยังผู้ร้องขอ
    } catch (error) {
        console.error('เกิดข้อผิดพลาดในการเรียก API:', error);
        // แสดงข้อผิดพลาดในกรณีที่เกิดข้อผิดพลาดในการดึงข้อมูลหรือบันทึกไฟล์

        res.status(500).send('ขออภัย เกิดข้อผิดพลาดในการเรียก API');
        // ส่งข้อความข้อผิดพลาดกลับไปยังผู้ร้องขอ
    }
});

router.get('/DATA/Hospital/Health/Rider/Exclude/Add', async (req, res) => {
    // กำหนดเส้นทาง HTTP GET สำหรับ URL '/DATA/Hospital/Health/Rider/Exclude/Add' โดยใช้ฟังก์ชัน async เพื่อรองรับการทำงานแบบอะซิงค์

    let hospital = req.param("hospital");
    // ดึงค่าพารามิเตอร์ 'hospital' จาก URL query และเก็บไว้ในตัวแปร 'hospital'

    let date_start = req.param("date_start");
    // ดึงค่าพารามิเตอร์ 'date_start' จาก URL query และเก็บไว้ในตัวแปร 'date_start'

    let date_end = req.param("date_end");
    // ดึงค่าพารามิเตอร์ 'date_end' จาก URL query และเก็บไว้ในตัวแปร 'date_end'

    let totalDisbursementCost = 0;
    // กำหนดตัวแปร 'totalDisbursementCost' เป็น 0 เพื่อใช้เก็บยอดรวมของ 'disbursement_cost' ที่จะคำนวณในภายหลัง

    const BEARER_TOKEN = ''; // ใส่ Bearer Token ที่นี่
    // กำหนดค่าตัวแปร Bearer Token สำหรับการยืนยันตัวตนในการเรียกใช้ API ที่ต้องการการยืนยันตัวตน

    const dateParts = date_start.split(' ');
    // แยกวันที่ในตัวแปร 'date_start' ออกเป็นสองส่วน (วันและเวลา) โดยใช้ช่องว่าง (' ') เป็นตัวคั่น

    const [year, month, day] = dateParts[0].split('-');
    // แยกปี, เดือน, และวัน จากส่วนที่ 1 ของวันที่ (รูปแบบ YYYY-MM-DD) โดยใช้เครื่องหมาย '-' เป็นตัวคั่น

    const extractedMonth = month; // This is the '05' part you need
    // เก็บค่าเดือนที่แยกออกมาในตัวแปร 'extractedMonth' (เช่น '05')

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
    // สร้างวัตถุ 'monthNamesInThai' ที่เก็บชื่อเดือนในภาษาไทยตามค่าเดือนที่ได้จาก 'extractedMonth'

    const monthNameInThai = monthNamesInThai[extractedMonth];
    // ดึงชื่อเดือนในภาษาไทยจาก 'monthNamesInThai' ตามค่าของ 'extractedMonth'

    const namefile = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567 (ไม่มีข้อมูล บัตรทอง และ UC)`;
    // สร้างชื่อไฟล์ Excel โดยรวมชื่อโรงพยาบาลและชื่อเดือนในภาษาไทยตามที่กำหนด พร้อมเพิ่มข้อความระบุว่าไม่มีข้อมูลสิทธิบัตรทองและ UC

    try {
        const excludeRights = ['บัตรทอง', 'UC'];
        // สร้างอาร์เรย์ 'excludeRights' เพื่อเก็บประเภทสิทธิบัตรที่ต้องการกรองออก

        let combinedData = [];
        // สร้างอาร์เรย์ 'combinedData' เพื่อเก็บข้อมูลที่รวมจากหลายแหล่ง

        const response = await axios.get('https://dashboard.manageai.co.th/gateway/api/v1/telepharmacy/transactions', {
            params: {
                order_type: '',
                hospital: hospital,
                status_active: 'success,waiting_transport,waiting_conference',
                start_date: date_start,
                end_date: date_end
            },
            headers: {
                Authorization: `Bearer ${BEARER_TOKEN}`
            }
        });
        // ใช้ axios เพื่อส่งคำขอ HTTP GET ไปยัง API ที่กำหนด โดยระบุพารามิเตอร์ที่จำเป็นและ Bearer Token สำหรับการยืนยันตัวตน

        combinedData = response.data.data.filter(item => !excludeRights.includes(item.medical_rights));
        // กรองข้อมูลใน response.data.data โดยเลือกเฉพาะรายการที่ไม่รวมอยู่ใน 'excludeRights'

        // Filter out entries with transport_process set to FALSE
        const filteredData = combinedData.filter(item => item.transport_process !== false);
        // กรองข้อมูลที่เหลือออกจาก 'combinedData' โดยเลือกเฉพาะรายการที่มีค่า 'transport_process' ไม่เท่ากับ false

        const sortedData = filteredData.sort((a, b) => {
            const dateA = new Date(a.date_receive_parcel_order);
            const dateB = new Date(b.date_receive_parcel_order);

            if (dateA.getDate().toString().slice(-2) === '01' && dateB.getDate().toString().slice(-2) !== '01') {
                return -1;
            } else if (dateA.getDate().toString().slice(-2) !== '01' && dateB.getDate().toString().slice(-2) === '01') {
                return 1;
            } else {
                return dateA - dateB;
            }
        });
        // เรียงลำดับข้อมูลตามวันที่ 'date_receive_parcel_order' โดยให้วันที่ 1 ของเดือนอยู่ก่อน จากนั้นเรียงลำดับวันที่ตามปกติ

        const nameAndDateSet = new Set();
        // สร้าง Set ใหม่ชื่อว่า 'nameAndDateSet' สำหรับเก็บข้อมูลที่ไม่ซ้ำกัน

        const duplicates = [];
        // สร้างอาร์เรย์ 'duplicates' เพื่อเก็บรายการที่ซ้ำกัน

        const dataWithNoDuplicates = sortedData.filter((item, index) => {
            const formattedOrderDate = formatDate(item.order_date);
            const key = `${item.name}-${formattedOrderDate}`;

            if (nameAndDateSet.has(key)) {
                duplicates.push(item);
                return false;
            } else {
                nameAndDateSet.add(key);
                return true;
            }
        }).map((item, index) => {
            return { ...item, no: index + 1, formattedOrderDate: formatDate(item.order_date) };
        });
        // กรองข้อมูลที่ซ้ำกันออกจาก 'sortedData' โดยใช้ Set เพื่อตรวจสอบ และจัดรูปแบบวันที่ จากนั้นเพิ่มหมายเลขลำดับลงในข้อมูลที่ไม่ซ้ำกัน

        const workbook = new ExcelJS.Workbook();
        // สร้างวัตถุ 'workbook' ใหม่สำหรับการสร้างไฟล์ Excel

        const worksheet = workbook.addWorksheet('Data');
        // เพิ่ม worksheet ใหม่ชื่อว่า 'Data' ลงใน workbook

        let row = 4;
        // กำหนดหมายเลขแถวเริ่มต้นเป็น 4 สำหรับการเพิ่มข้อมูล

        worksheet.getCell(`A${row}`).value = 'ลำดับ';
        worksheet.getCell(`B${row}`).value = 'ชื่อผู้ใช้บริการ';
        worksheet.getCell(`C${row}`).value = 'วันที่สร้างออเดอร์';
        worksheet.getCell(`D${row}`).value = 'เลขที่การจัดส่ง';
        worksheet.getCell(`E${row}`).value = 'เลขออเดอร์';
        worksheet.getCell(`F${row}`).value = 'HN';
        worksheet.getCell(`G${row}`).value = 'บริการขนส่ง';
        worksheet.getCell(`H${row}`).value = 'สิทธิการรักษา';
        worksheet.getCell(`I${row}`).value = 'box_size_detail';
        worksheet.getCell(`J${row}`).value = 'box_size_name';
        // กำหนดชื่อคอลัมน์ในแถวที่ 4 ของ worksheet โดยเพิ่มฟิลด์ 'box_size_detail' และ 'box_size_name'

        dataWithNoDuplicates.forEach(item => {
            row++;
            // เพิ่มหมายเลขแถวทีละ 1 เพื่อเพิ่มข้อมูลในแต่ละแถว

            worksheet.getCell(`A${row}`).value = item.no;
            worksheet.getCell(`B${row}`).value = item.name;
            worksheet.getCell(`C${row}`).value = item.formattedOrderDate; // ใช้วันที่ที่จัดรูปแบบแล้ว
            worksheet.getCell(`D${row}`).value = item.order_no;
            worksheet.getCell(`E${row}`).value = item.reference_id;
            worksheet.getCell(`F${row}`).value = item.cardHN;
            worksheet.getCell(`G${row}`).value = item.order_type;
            worksheet.getCell(`H${row}`).value = item.medical_rights;
            worksheet.getCell(`I${row}`).value = item.box_size_detail;
            worksheet.getCell(`J${row}`).value = item.box_size_name;
            // เพิ่มข้อมูลแต่ละรายการลงใน worksheet โดยเพิ่มข้อมูลในแต่ละเซลล์

            totalDisbursementCost += item.disbursement_cost;
            // คำนวณยอดรวมของ 'disbursement_cost' โดยบวกค่า 'disbursement_cost' ของแต่ละรายการเข้าไปใน 'totalDisbursementCost'
        });

        const formatter = new Intl.NumberFormat('th-TH', {
            currency: 'THB',
        });
        // สร้างตัวแปร 'formatter' สำหรับการจัดรูปแบบตัวเลขเป็นสกุลเงินบาท

        worksheet.getCell('E2').value = `ประวัติการรอรับยา${hospital} ประจำเดือน${monthNameInThai} 2567 (ไม่มีข้อมูล บัตรทอง และ UC)`;
        // ตั้งค่าเซลล์ 'E2' ของ worksheet เป็นชื่อไฟล์ที่จัดรูปแบบตามที่กำหนด

        const filePath = path.join(__dirname, `${namefile}.xlsx`);
        // สร้างพาธสำหรับการบันทึกไฟล์ Excel โดยใช้ชื่อไฟล์ที่กำหนด

        await workbook.xlsx.writeFile(filePath);
        // บันทึกไฟล์ Excel ไปยังพาธที่กำหนด

        const downloadURL = `http://localhost:${PORT}/downloadExcel/${namefile}`;
        // สร้าง URL สำหรับการดาวน์โหลดไฟล์ Excel

        res.send({ downloadURL });
        // ส่ง URL สำหรับการดาวน์โหลดไฟล์ Excel กลับไปยังผู้ร้องขอ
    } catch (error) {
        console.error('เกิดข้อผิดพลาดในการเรียก API:', error);
        // แสดงข้อผิดพลาดในกรณีที่เกิดข้อผิดพลาดในการดึงข้อมูลหรือบันทึกไฟล์

        res.status(500).send('ขออภัย เกิดข้อผิดพลาดในการเรียก API');
        // ส่งข้อความข้อผิดพลาดกลับไปยังผู้ร้องขอ
    }
});

router.get('/downloadExcel/:filename', (req, res) => {
    // กำหนดเส้นทาง HTTP GET สำหรับ URL '/downloadExcel/:filename' โดยใช้ฟังก์ชันที่รับสองพารามิเตอร์: req และ res

    const { filename } = req.params;
    // ดึงค่าพารามิเตอร์ 'filename' จาก URL query และเก็บไว้ในตัวแปร 'filename'
    // เช่น ถ้า URL เป็น '/downloadExcel/filename123', 'filename' จะเป็น 'filename123'

    const file = `${__dirname}/${filename}.xlsx`;
    // สร้างพาธเต็มของไฟล์โดยใช้ชื่อไฟล์ที่ได้รับจาก URL และกำหนดนามสกุล '.xlsx'
    // __dirname เป็นตัวแปรที่เก็บพาธของไดเรกทอรีที่ปัจจุบันของไฟล์สคริปต์

    res.download(file, () => {
        // ใช้ res.download เพื่อให้ผู้ใช้สามารถดาวน์โหลดไฟล์ที่กำหนด โดยระบุพาธของไฟล์
        // ฟังก์ชัน callback ที่จะถูกเรียกหลังจากที่ไฟล์ถูกดาวน์โหลด

        fs.unlink(file, (err) => {
            // ใช้ fs.unlink เพื่อลบไฟล์หลังจากดาวน์โหลดเสร็จ
            // ตัวแปร 'file' เป็นพาธของไฟล์ที่ต้องการลบ

            if (err) {
                // ถ้ามีข้อผิดพลาดในการลบไฟล์
                console.error('เกิดข้อผิดพลาดในการลบไฟล์:', err);
                // แสดงข้อผิดพลาดในคอนโซล
            } else {
                console.log('ไฟล์ถูกลบแล้ว:', file);
                // ถ้าลบไฟล์สำเร็จ แสดงข้อความในคอนโซลว่าไฟล์ถูกลบแล้ว
            }
        });
    });
});


module.exports = router;
