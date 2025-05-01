document.addEventListener("DOMContentLoaded", function () {
  const classSelect = document.getElementById("classSelect");
  const classTables = document.querySelectorAll(".classTable");
  const violationInputs = document.querySelectorAll(".violation");
  const saveBtn = document.getElementById("saveBtn");
  const resetBtn = document.getElementById("resetBtn");
  const undoBtn = document.getElementById("undoBtn");
  const exportBtn = document.getElementById("exportBtn");
  const cleanOldBtn = document.getElementById("cleanOldAttendanceBtn");
  const exportRecentBtn = document.getElementById("exportRecentAttendanceBtn");

  let currentClass = "";
  let undoStack = [];

  const studentList = [
    { id: "012345678901", name: "Nguyễn Văn A", classId: "10A1" },
    { id: "023456789012", name: "Trần Thị B", classId: "10A2" },
    { id: "034567890123", name: "Lê Văn C", classId: "11A1" },
  ];
  localStorage.setItem("studentList", JSON.stringify(studentList));

  const schoolTotalCells = {};
  for (let grade of [10, 11, 12]) {
    const count = grade === 11 ? 10 : 9;
    for (let i = 1; i <= count; i++) {
      const classId = `${grade}A${i}`;
      const el = document.getElementById(`school-total-${classId}`);
      if (el) schoolTotalCells[classId] = el;
    }
  }

  function showClassTable() {
    classTables.forEach(table => table.classList.remove("active"));
    const selected = classSelect.value;
    const table = document.getElementById(selected);
    if (table) {
      table.classList.add("active");
      currentClass = selected;
    }
  }

  function calculateTotals() {
    const selectedClass = classSelect.value;
    let schoolTotal = 0;
    document.querySelectorAll(`.classTable.active tbody tr`).forEach(row => {
      let rowTotal = 0;
      row.querySelectorAll(".violation").forEach(input => {
        rowTotal += parseInt(input.value) || 0;
      });
      row.querySelector(".total").textContent = 200 - rowTotal;
      schoolTotal += rowTotal;
    });
    if (schoolTotalCells[selectedClass]) {
      schoolTotalCells[selectedClass].textContent = 200 - schoolTotal;
    }
    sortClassesByGrade();
  }

  function sortClassesByGrade() {
    const classIdsByGrade = {
      '10': ['10A1','10A2','10A3','10A4','10A5','10A6','10A7','10A8','10A9'],
      '11': ['11A1','11A2','11A3','11A4','11A5','11A6','11A7','11A8','11A9','11A10'],
      '12': ['12A1','12A2','12A3','12A4','12A5','12A6','12A7','12A8','12A9']
    };
    ["10", "11", "12"].forEach(grade => {
      const data = classIdsByGrade[grade].map(classId => {
        return {
          classId,
          totalPoints: parseInt(schoolTotalCells[classId]?.textContent) || 0
        };
      });
      data.sort((a, b) => b.totalPoints - a.totalPoints);
      const tbody = document.querySelector(`#sortedTable${grade} tbody`);
      if (tbody) {
        tbody.innerHTML = "";
        data.forEach(({ classId, totalPoints }) => {
          const row = document.createElement("tr");
          row.innerHTML = `<td>${classId}</td><td>${totalPoints}</td>`;
          tbody.appendChild(row);
        });
      }
    });
  }

  function resetAll() {
    document.querySelectorAll(".violation").forEach(i => i.value = 0);
    document.querySelectorAll("textarea").forEach(t => t.value = "");
    calculateTotals();
  }

  function collectData() {
    const data = {};
    classTables.forEach(table => {
      const classId = table.id;
      const classData = [];
      table.querySelectorAll("tbody tr").forEach(row => {
        const rowData = {
          day: row.querySelector("td").textContent,
          violations: []
        };
        row.querySelectorAll(".violation").forEach(input => {
          rowData.violations.push(input.value);
        });
        classData.push(rowData);
      });
      data[classId] = classData;
    });
    return data;
  }

  saveBtn?.addEventListener("click", () => {
    const violations = collectData();
    const attendance = JSON.parse(localStorage.getItem("attendance") || "[]");
  
    // ✅ In ra dữ liệu sắp gửi
    console.log("📤 Dữ liệu gửi lên server:");
    console.log("violations:", violations);
    console.log("attendance:", attendance);
  
    // Gọi ping trước để đánh thức server (tránh Render bị "ngủ đông")
    fetch("https://myserverrrr.onrender.com/", { method: "GET" })
      .catch(() => {}) // Bỏ qua lỗi nếu ping thất bại
      .finally(() => {
        // Gửi dữ liệu thật sự
        fetch("https://myserverrrr.onrender.com/save", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(violations)
        })
          .then(async res => {
            if (!res.ok) {
              const errorText = await res.text(); // đọc chi tiết lỗi từ server
              throw new Error(`❌ Server trả lỗi ${res.status}: ${errorText}`);
            }
            return res.json();
          })
          .then(res => {
            console.log("✅ Phản hồi từ server:", res);
            alert(res.message || "✅ Lưu thành công");
          })
          .catch(err => {
            console.error("🚫 Gặp lỗi khi lưu dữ liệu:", err);
            alert("❌ Lỗi kết nối server hoặc dữ liệu không hợp lệ. Xem console để biết chi tiết.");
          });
      });
  });
  

  resetBtn?.addEventListener("click", () => {
    if (confirm("Bạn chắc chắn muốn đặt lại tất cả?")) resetAll();
  });

  exportBtn?.addEventListener("click", () => {
    const diemDanh = JSON.parse(localStorage.getItem("attendance") || "[]");
    const viPham = collectData();
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(diemDanh);
    XLSX.utils.book_append_sheet(wb, ws1, "DiemDanh");
    Object.keys(viPham).forEach(classId => {
      const rows = viPham[classId].map(row => {
        const obj = { Thứ: row.day };
        const violationTitles = [
          "Nghỉ học không phép",
          "Đi muộn",
          "Không đồng phục",
          "Không đeo thẻ",
          "Vệ sinh không tốt",
          "Ăn quà vặt",
          "Vi phạm ATGT",
          "Hút thuốc lá trong trường",
          "Vi phạm khác",
          "Không quản được giờ trống",
          "Xếp loại giờ"
        ];
        
        row.violations.forEach((val, i) => {
          obj[violationTitles[i] || `Vi phạm ${i+1}`] = val;
        });
        
        return obj;
      });
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, classId);
    });
    XLSX.writeFile(wb, "du_lieu_tong_hop.xlsx");
  });

  cleanOldBtn?.addEventListener("click", () => {
    const allLogs = JSON.parse(localStorage.getItem("attendance") || "[]");
    const today = new Date();
    const recentLogs = allLogs.filter(entry => {
      const [day, month, year] = entry.date.split("/").map(Number);
      const entryDate = new Date(year, month - 1, day);
      return (today - entryDate) / (1000 * 60 * 60 * 24) <= 7;
    });
    localStorage.setItem("attendance", JSON.stringify(recentLogs));
    alert("✅ Đã xóa các bản điểm danh cũ hơn 7 ngày.");
  });

  exportRecentBtn?.addEventListener("click", () => {
    const allLogs = JSON.parse(localStorage.getItem("attendance") || "[]");
    const today = new Date();
    const recentLogs = allLogs.filter(entry => {
      const [day, month, year] = entry.date.split("/").map(Number);
      const entryDate = new Date(year, month - 1, day);
      return (today - entryDate) / (1000 * 60 * 60 * 24) <= 7;
    });
    if (recentLogs.length === 0) return alert("⚠️ Không có học sinh nào điểm danh trong 7 ngày gần nhất.");
    const ws = XLSX.utils.json_to_sheet(recentLogs);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DiemDanh_7ngay");
    XLSX.writeFile(wb, "diem_danh_7_ngay_gan_nhat.xlsx");
  });

  let qrScanner;
  document.getElementById("startScanBtn")?.addEventListener("click", () => {
    qrScanner = new Html5Qrcode("reader");
    qrScanner.start(
      { facingMode: "environment" },
      { fps: 10, qrbox: 250 },
      msg => {
        try {
          const data = JSON.parse(msg);
          const list = JSON.parse(localStorage.getItem("studentList") || "[]");
          const found = list.find(s => s.id === data.id);
          if (found) {
            const now = new Date();
            const today = now.toLocaleDateString();
            const logs = JSON.parse(localStorage.getItem("attendance") || "[]");
            const already = logs.find(l => l.id === found.id && l.date === today);
            if (!already) {
              logs.push({ id: found.id, name: found.name, classId: found.classId, time: now.toLocaleTimeString(), date: today });
              localStorage.setItem("attendance", JSON.stringify(logs));
              document.getElementById("status").innerText = `✅ Điểm danh: ${found.name} (${found.classId})`;
            } else {
              document.getElementById("status").innerText = `✅ ${found.name} đã điểm danh hôm nay.`;
            }
          } else {
            document.getElementById("status").innerText = "❌ Không tìm thấy học sinh";
          }
        } catch {
          document.getElementById("status").innerText = "❌ QR không hợp lệ";
        }
      },
      err => console.warn("QR Error:", err)
    );
    document.getElementById("startScanBtn").style.display = "none";
    document.getElementById("stopScanBtn").style.display = "inline-block";
  });

  document.getElementById("stopScanBtn")?.addEventListener("click", () => {
    if (qrScanner) {
      qrScanner.stop();
      document.getElementById("startScanBtn").style.display = "inline-block";
      document.getElementById("stopScanBtn").style.display = "none";
    }
  });

  window.showAttendance = function () {
    const data = JSON.parse(localStorage.getItem("attendance") || "[]");
    const tbody = document.querySelector("#attendanceTable tbody");
    tbody.innerHTML = "";
    data.forEach(d => {
      const row = tbody.insertRow();
      row.innerHTML = `<td>${d.name}</td><td>${d.classId}</td><td>${d.time}</td><td>${d.date}</td>`;
    });
    document.getElementById("attendanceTable").style.display = "table";
  };

  violationInputs.forEach(input => input.addEventListener("input", calculateTotals));
  classSelect.addEventListener("change", () => {
    showClassTable();
    calculateTotals();
  });
  showClassTable();
  calculateTotals();
});
function getUncheckedStudents() {
  const studentList = JSON.parse(localStorage.getItem("studentList") || "[]");
  const attendance = JSON.parse(localStorage.getItem("attendance") || "[]");
  const today = new Date().toLocaleDateString();

  const checkedIdsToday = attendance
    .filter(a => a.date === today)
    .map(a => a.id);

  return studentList.filter(s => !checkedIdsToday.includes(s.id));
}

function showUnchecked() {
  const list = getUncheckedStudents();
  const ul = document.getElementById("uncheckedList");
  ul.innerHTML = "";

  if (list.length === 0) {
    ul.innerHTML = "<li>✅ Tất cả học sinh đã điểm danh hôm nay.</li>";
    return;
  }

  list.forEach(s => {
    const li = document.createElement("li");
    li.textContent = `${s.name} (${s.classId})`;
    ul.appendChild(li);
  });
}
