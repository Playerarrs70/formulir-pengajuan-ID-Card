// assets/js/report-generator.js

/**
 * Membuat laporan harian yang terfilter untuk absensi MAHASISWA.
 * Dipanggil dari history_admin.html
 */
function generateDailyReport(reportData) {
    // Definisikan style untuk digunakan kembali
    const titleStyle = { font: { bold: true, sz: 16 }, alignment: { horizontal: "center", vertical: "center" } };
    const headerStyle = { font: { bold: true }, fill: { fgColor: { rgb: "E9ECEF" } }, border: { bottom: { style: "thin" } } };
    const subHeaderStyle = { font: { bold: true, sz: 12 } };
    const boldStyle = { font: { bold: true } };

    // Siapkan data dalam format array of arrays
    const data = [
        [{v: "Filtered Daily Student Attendance Report", s: titleStyle}, {}, {}, {}, {}, {}],
        [],
        [{v: "Report Generation Date:", s: boldStyle}, new Date().toLocaleString('en-US', { dateStyle: 'full', timeStyle: 'long' })],
        [{v: "Data For Date:", s: boldStyle}, reportData.reportDate || "All Week"],
        [],
        [{v: "Active Filters", s: subHeaderStyle}],
        [{v: "Pleton:", s: boldStyle}, reportData.filters.pleton || "All"],
        [{v: "Major:", s: boldStyle}, reportData.filters.prodi || "All"],
        [{v: "Status:", s: boldStyle}, reportData.filters.status || "All"],
        [],
        [{v: "Filtered Summary", s: subHeaderStyle}],
        [`Total Records: ${reportData.total}`, `Present: ${reportData.present}`, `Absent: ${reportData.absent}`],
        [],
    ];

    if (reportData.presentList.length > 0) {
        data.push([{v: "Present Students", s: { font: { bold: true, color: { rgb: "198754" }}}}]);
        data.push(["No.", "Name", "NIM", "Major", "Pleton", "Time"].map(h => ({v: h, s: headerStyle})));
        reportData.presentList.forEach((student, index) => {
            data.push([index + 1, student.name, student.nim, student.major, student.pleton, student.time]);
        });
        data.push([]);
    }

    if (reportData.absentList.length > 0) {
        data.push([{v: "Absent Students", s: { font: { bold: true, color: { rgb: "DC3545" }}}}]);
        data.push(["No.", "Name", "NIM", "Major", "Pleton"].map(h => ({v: h, s: headerStyle})));
        reportData.absentList.forEach((student, index) => {
            data.push([index + 1, student.name, student.nim, student.major, student.pleton]);
        });
    }
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }]; // Gabungkan sel judul
    ws['!cols'] = [{ wch: 5 }, { wch: 30 }, { wch: 18 }, { wch: 25 }, { wch: 15 }, { wch: 10 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Student Daily Report");
    XLSX.writeFile(wb, `Student_Daily_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
}


/**
 * Membuat laporan mingguan untuk satu MAHASISWA.
 * Dipanggil dari modal di history_admin.html
 */
function generateWeeklyReport(reportData) {
    const data = [];
    const titleStyle = { font: { bold: true, sz: 16 }, alignment: { horizontal: "center" } };
    const headerStyle = { font: { bold: true }, fill: { fgColor: { rgb: "E9ECEF" } }, border: { top: { style: "thin" }, bottom: { style: "thin" } } };
    const boldStyle = {font:{bold:true}};

    data.push([{v: "Student Weekly Attendance Report", s: titleStyle}, {}, {}, {}]);
    data.push([]);
    data.push([{v: "Name:", s:boldStyle}, reportData.studentName]);
    data.push([{v: "NIM:", s:boldStyle}, reportData.nim]);
    data.push([{v: "Major:", s:boldStyle}, reportData.major]);
    data.push([{v: "Pleton:", s:boldStyle}, reportData.pleton]);
    data.push([{v: "Primary Mentor:", s:boldStyle}, reportData.mentorUtama]);
    data.push([{v: "Assistant Mentor:", s:boldStyle}, reportData.mentorAsisten]);
    data.push([]);
    data.push([{v: "Weekly Summary", s:{font:{bold:true, sz:12}}}]);
    const headers = ["Day", "Date", "Status", "Time"];
    data.push(headers.map(h => ({v: h, s: headerStyle})));
    reportData.weeklyData.forEach(day => { data.push([day.day, day.date, day.status, day.time]); });

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    ws['!cols'] = [ { wch: 15 }, { wch: 18 }, { wch: 15 }, { wch: 15 } ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Student Weekly Report");
    XLSX.writeFile(wb, `Weekly_Report_${reportData.studentName}.xlsx`);
}


// --- BARU: Fungsi Laporan Mingguan untuk Mentor ---
/**
 * Membuat laporan mingguan untuk satu MENTOR.
 * Dipanggil dari modal di history_mentors.html
 */
function generateMentorWeeklyReport(reportData) {
    const data = [];
    // Styles
    const titleStyle = { font: { bold: true, sz: 16 }, alignment: { horizontal: "center" } };
    const headerStyle = { font: { bold: true }, fill: { fgColor: { rgb: "E9ECEF" } }, border: { top: { style: "thin" }, bottom: { style: "thin" } } };
    const boldStyle = { font: { bold: true } };

    // Header Laporan
    data.push([{v: "Mentor Weekly Attendance Report", s: titleStyle}, {}, {}, {}]);
    data.push([]);
    data.push([{v: "Name:", s:boldStyle}, reportData.mentorName]);
    data.push([{v: "Pleton:", s:boldStyle}, reportData.pleton]);
    data.push([{v: "Role:", s:boldStyle}, reportData.role]);
    data.push([]);
    data.push([{v: "Weekly Summary", s:{font:{bold:true, sz:12}}}]);
    
    // Tabel Data
    const headers = ["Day", "Date", "Status", "Check-in Time"];
    data.push(headers.map(h => ({v: h, s: headerStyle})));
    reportData.weeklyData.forEach(day => {
        const dateObj = new Date(day.date + 'T00:00:00');
        const dayName = dateObj.toLocaleDateString('en-US', { weekday: 'long' });
        data.push([
            dayName, 
            day.date, 
            day.status, 
            day.checkinTime || '-'
        ]); 
    });

    // Membuat file Excel
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    ws['!cols'] = [ { wch: 15 }, { wch: 18 }, { wch: 15 }, { wch: 15 } ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Mentor Weekly Report");
    XLSX.writeFile(wb, `Mentor_Weekly_Report_${reportData.mentorName}.xlsx`);
}