function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sổ Nề Nếp Lớp 10A5')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== CÁC HÀM TIỆN ÍCH CHUẨN HÓA ====================
function chuanHoaHoTen(str) {
  if (!str) return "";
  return str.toString().trim().toLowerCase().replace(/\s+/g, ' ');
}

function chuanHoaNgaySinh(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var d = ("0" + val.getDate()).slice(-2);
    var m = ("0" + (val.getMonth() + 1)).slice(-2);
    var y = val.getFullYear();
    return d + "/" + m + "/" + y;
  }
  var s = val.toString().trim();
  // Nếu là yyyy-mm-dd (chuẩn date picker HTML5)
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) {
    var parts = s.split("-");
    var y = parts[0];
    var m = ("0" + parts[1]).slice(-2);
    var d = ("0" + parts[2]).slice(-2);
    return d + "/" + m + "/" + y;
  }
  // Nếu là dd/mm/yyyy hoặc d/m/yyyy hoặc dd-mm-yyyy
  var parts2 = s.split(/[\/\-\.]/);
  if (parts2.length === 3) {
    if (parts2[0].length === 4) { // yyyy/mm/dd
      var y = parts2[0];
      var m = ("0" + parts2[1]).slice(-2);
      var d = ("0" + parts2[2]).slice(-2);
      return d + "/" + m + "/" + y;
    } else { // dd/mm/yyyy
      var d = ("0" + parts2[0]).slice(-2);
      var m = ("0" + parts2[1]).slice(-2);
      var y = parts2[2];
      return d + "/" + m + "/" + y;
    }
  }
  return s.replace(/[^0-9]/g, '');
}

function layChuSoNgaySinh(strDate) {
  var ch = chuanHoaNgaySinh(strDate);
  return (ch || "").replace(/[^0-9]/g, '');
}

// ==================== 1. XÁC THỰC MÃ PIN CHO CÁN SỰ & GVCN ====================
function kiemTraDangNhap(vaiTro, maPIN) {
  try {
    var pinNhap = (maPIN || "").toString().trim();
    var v = (vaiTro || "").toString().trim();

    var to = "all";
    var ten = "Lớp trưởng / Lớp phó / GVCN";
    var pinChuan = "9999";

    if (v === "1" || v.indexOf("Tổ 1") !== -1 || v.indexOf("Tổ trưởng Tổ 1") !== -1) {
      to = "1";
      ten = "Tổ trưởng Tổ 1";
      pinChuan = "1001";
    } else if (v === "2" || v.indexOf("Tổ 2") !== -1 || v.indexOf("Tổ trưởng Tổ 2") !== -1) {
      to = "2";
      ten = "Tổ trưởng Tổ 2";
      pinChuan = "1002";
    } else if (v === "3" || v.indexOf("Tổ 3") !== -1 || v.indexOf("Tổ trưởng Tổ 3") !== -1) {
      to = "3";
      ten = "Tổ trưởng Tổ 3";
      pinChuan = "1003";
    } else if (v === "4" || v.indexOf("Tổ 4") !== -1 || v.indexOf("Tổ trưởng Tổ 4") !== -1) {
      to = "4";
      ten = "Tổ trưởng Tổ 4";
      pinChuan = "1004";
    }

    // Khớp mã PIN chuẩn hoặc mã PIN vạn năng của GVCN (9999)
    if (pinNhap === pinChuan || pinNhap === "9999") {
      return {
        success: true,
        thanhCong: true,
        tenVaiTro: ten,
        vaiTro: to,
        message: "Đăng nhập thành công!",
        thongBao: "Đăng nhập thành công!"
      };
    }

    // Kiểm tra tiếp trong Google Sheets nếu người dùng đổi PIN trong Sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetHS = ss.getSheetByName("DanhSachHocSinh");
    if (sheetHS) {
      var dataHS = sheetHS.getDataRange().getValues();
      for (var i = 1; i < dataHS.length; i++) {
        var toSheet = dataHS[i][2] ? dataHS[i][2].toString().replace(/[^0-9]/g, '') : "";
        var chucVu = dataHS[i][3] ? dataHS[i][3].toString() : "";
        var pinSheet = dataHS[i][4] ? dataHS[i][4].toString().trim() : "";

        if (pinNhap === pinSheet && pinSheet !== "") {
          if (to !== "all" && toSheet === to) {
            return {
              success: true,
              thanhCong: true,
              tenVaiTro: dataHS[i][1] || ten,
              vaiTro: to,
              message: "Đăng nhập thành công!",
              thongBao: "Đăng nhập thành công!"
            };
          }
          if (to === "all" && (chucVu.indexOf("Lớp") !== -1 || chucVu.indexOf("GV") !== -1)) {
            return {
              success: true,
              thanhCong: true,
              tenVaiTro: dataHS[i][1] || ten,
              vaiTro: "all",
              message: "Đăng nhập thành công!",
              thongBao: "Đăng nhập thành công!"
            };
          }
        }
      }
    }

    return {
      success: false,
      thanhCong: false,
      message: "❌ Sai Mã PIN! Vui lòng kiểm tra lại.",
      thongBao: "❌ Sai Mã PIN! Vui lòng kiểm tra lại."
    };
  } catch (e) {
    return {
      success: false,
      thanhCong: false,
      message: "Lỗi: " + e.toString(),
      thongBao: "Lỗi: " + e.toString()
    };
  }
}

// ==================== 2. LẤY DỮ LIỆU KHỞI TẠO CHO FORM ====================
function layDuLieuKhoiTao() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetHS = ss.getSheetByName("DanhSachHocSinh");
    var sheetHV = ss.getSheetByName("DanhMucHanhVi");
    
    var dsHocSinh = [];
    if (sheetHS) {
      var dataHS = sheetHS.getDataRange().getValues();
      for (var i = 1; i < dataHS.length; i++) {
        if (dataHS[i][1]) {
          var toSo = dataHS[i][2] ? dataHS[i][2].toString().replace(/[^0-9]/g, '') : "";
          dsHocSinh.push({
            maHS: dataHS[i][0] ? dataHS[i][0].toString() : "",
            hoTen: dataHS[i][1].toString(),
            to: toSo,
            chucVu: dataHS[i][3] ? dataHS[i][3].toString() : ""
          });
        }
      }
    }
    
    var dsHanhVi = [];
    if (sheetHV) {
      var dataHV = sheetHV.getDataRange().getValues();
      for (var j = 1; j < dataHV.length; j++) {
        if (dataHV[j][1]) {
          dsHanhVi.push({
            maHV: dataHV[j][0] ? dataHV[j][0].toString() : "",
            tenHV: dataHV[j][1].toString(),
            loai: dataHV[j][2] ? dataHV[j][2].toString() : "Vi phạm",
            diem: dataHV[j][3] ? Number(dataHV[j][3]) : 0
          });
        }
      }
    }
    
    return { dsHocSinh: dsHocSinh, dsHanhVi: dsHanhVi };
  } catch (e) {
    return { dsHocSinh: [], dsHanhVi: [] };
  }
}

// ==================== 3. LƯU BÁO CÁO NỀ NẾP HÀNG NGÀY ====================
function luuBaoCao(duLieu) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNK = ss.getSheetByName("NhatKyBaoCao");
    if (!sheetNK) {
      sheetNK = ss.insertSheet("NhatKyBaoCao");
      sheetNK.appendRow(["Mã", "Thời gian", "Tuần", "Họ và tên", "Hành vi", "Điểm", "Ghi chú", "Người báo cáo"]);
    }
    
    var id = Utilities.getUuid().substring(0, 8);
    var thoiGian = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "HH:mm:ss dd/MM/yyyy");
    
    sheetNK.appendRow([
      id,
      thoiGian,
      duLieu.tuan,
      duLieu.hocSinh,
      duLieu.hanhVi,
      duLieu.diem,
      duLieu.ghiChu,
      duLieu.nguoiBaoCao
    ]);
    
    return {
      success: true,
      thanhCong: true,
      message: "✅ Đã lưu báo cáo nề nếp thành công!",
      thongBao: "✅ Đã lưu báo cáo nề nếp thành công!"
    };
  } catch (error) {
    return {
      success: false,
      thanhCong: false,
      message: "Lỗi: " + error.toString(),
      thongBao: "Lỗi: " + error.toString()
    };
  }
}

// ==================== 4. LẤY DỮ LIỆU THEO TUẦN ĐỂ XUẤT SỔ TỔNG HỢP PDF ====================
function layDuLieuTheoTuan(tuanChon) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNK = ss.getSheetByName("NhatKyBaoCao");
    var danhSachTuan = [];
    if (!sheetNK) {
      return { success: true, thanhCong: true, danhSach: [], tuan: tuanChon };
    }

    // Tra cứu Tổ của từng học sinh từ DanhSachHocSinh
    var mapTo = {};
    var sheetHS = ss.getSheetByName("DanhSachHocSinh");
    if (sheetHS) {
      var dataHS = sheetHS.getDataRange().getValues();
      for (var k = 1; k < dataHS.length; k++) {
        if (dataHS[k][1]) {
          var toSo = dataHS[k][2] ? dataHS[k][2].toString().replace(/[^0-9]/g, '') : "";
          mapTo[dataHS[k][1].toString().trim()] = toSo ? ("Tổ " + toSo) : "";
        }
      }
    }

    var data = sheetNK.getDataRange().getValues();
    var soTuanChon = (tuanChon || "").toString().replace(/[^0-9]/g, '');

    for (var i = 1; i < data.length; i++) {
      var tuanDong = data[i][2] ? data[i][2].toString().trim() : "";
      var soTuanDong = tuanDong.replace(/[^0-9]/g, '');

      // So khớp linh hoạt: trùng tên chữ hoặc trùng số tuần
      var khop = (tuanDong.toLowerCase() === (tuanChon || "").toString().trim().toLowerCase()) ||
                 (soTuanDong !== "" && soTuanDong === soTuanChon);

      if (khop) {
        var tenHS = data[i][3] ? data[i][3].toString().trim() : "";
        var tgRaw = data[i][1];
        var tgStr = "";
        if (tgRaw instanceof Date) {
          tgStr = Utilities.formatDate(tgRaw, "Asia/Ho_Chi_Minh", "HH:mm dd/MM/yyyy");
        } else {
          tgStr = tgRaw ? tgRaw.toString().trim() : "";
        }

        danhSachTuan.push({
          thoiGian: tgStr,
          hocSinh: tenHS,
          to: mapTo[tenHS] || "",
          hanhVi: data[i][4] ? data[i][4].toString() : "",
          diem: data[i][5] !== "" ? Number(data[i][5]) : 0,
          ghiChu: data[i][6] ? data[i][6].toString() : "",
          nguoiBao: data[i][7] ? data[i][7].toString() : ""
        });
      }
    }
    return {
      success: true,
      thanhCong: true,
      danhSach: danhSachTuan,
      tuan: tuanChon
    };
  } catch (e) {
    return {
      success: false,
      thanhCong: false,
      message: e.toString(),
      thongBao: e.toString()
    };
  }
}

// ==================== 5. TÍNH NĂNG ĐỘT PHÁ SKKN: PHỤ HUYNH TRA CỨU BẢO MẬT ====================
// Sử dụng HỌ VÀ TÊN + NGÀY THÁNG NĂM SINH để bảo vệ quyền riêng tư học sinh (Luật Trẻ em 2016)
function traCuuPhuHuynh(hoTenNhap, ngaySinhNhap, tuanChon) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetHS = ss.getSheetByName("DanhSachHocSinh");
    if (!sheetHS) {
      return {
        success: false,
        thanhCong: false,
        message: "Chưa thiết lập bảng DanhSachHocSinh trong hệ thống!",
        thongBao: "Chưa thiết lập bảng DanhSachHocSinh trong hệ thống!"
      };
    }

    var tenChuanNhap = chuanHoaHoTen(hoTenNhap);
    var chuSoNsNhap = layChuSoNgaySinh(ngaySinhNhap);

    if (!tenChuanNhap || !chuSoNsNhap) {
      return {
        success: false,
        thanhCong: false,
        message: "Vui lòng nhập đầy đủ Họ và tên và Ngày tháng năm sinh của học sinh!",
        thongBao: "Vui lòng nhập đầy đủ Họ và tên và Ngày tháng năm sinh của học sinh!"
      };
    }

    var dataHS = sheetHS.getDataRange().getValues();
    if (dataHS.length < 2) {
      return {
        success: false,
        thanhCong: false,
        message: "Danh sách học sinh hiện đang trống.",
        thongBao: "Danh sách học sinh hiện đang trống."
      };
    }

    // Tự động nhận diện cột Ngày sinh (mặc định Cột F - chỉ số 5, hoặc tìm theo tiêu đề)
    var colNgaySinh = 5;
    var headers = dataHS[0];
    for (var c = 0; c < headers.length; c++) {
      var h = headers[c] ? headers[c].toString().toLowerCase() : "";
      if (h.indexOf("sinh") !== -1 || h.indexOf("ngày sinh") !== -1 || h.indexOf("ngaysinh") !== -1) {
        colNgaySinh = c;
        break;
      }
    }

    var hocSinhTimThay = null;
    for (var i = 1; i < dataHS.length; i++) {
      var tenDong = chuanHoaHoTen(dataHS[i][1]);
      var nsDong = dataHS[i][colNgaySinh];
      var chuSoNsDong = layChuSoNgaySinh(nsDong);

      if (tenDong === tenChuanNhap && chuSoNsDong === chuSoNsNhap) {
        var toSo = dataHS[i][2] ? dataHS[i][2].toString().replace(/[^0-9]/g, '') : "";
        hocSinhTimThay = {
          maHS: dataHS[i][0] ? dataHS[i][0].toString() : "",
          hoTen: dataHS[i][1] ? dataHS[i][1].toString().trim() : "",
          to: toSo ? ("Tổ " + toSo) : "",
          chucVu: dataHS[i][3] ? dataHS[i][3].toString() : "Học sinh",
          ngaySinh: chuanHoaNgaySinh(nsDong)
        };
        break;
      }
    }

    if (!hocSinhTimThay) {
      return {
        success: false,
        thanhCong: false,
        message: "❌ Thông tin không khớp! Không tìm thấy học sinh với Họ tên và Ngày sinh đã nhập. Quý Phụ huynh vui lòng kiểm tra lại.",
        thongBao: "❌ Thông tin không khớp! Không tìm thấy học sinh với Họ tên và Ngày sinh đã nhập. Quý Phụ huynh vui lòng kiểm tra lại."
      };
    }

    // Lấy dữ liệu vi phạm / khen thưởng từ NhatKyBaoCao
    var sheetNK = ss.getSheetByName("NhatKyBaoCao");
    var danhSachNhatKy = [];
    var tongDiemThayDoi = 0;
    var tongCong = 0;
    var tongTru = 0;

    var soTuanChon = (tuanChon || "").toString().replace(/[^0-9]/g, '');
    var locTatCa = (!tuanChon || tuanChon === "all" || tuanChon === "Tất cả" || tuanChon.toLowerCase().indexOf("tất cả") !== -1);

    if (sheetNK) {
      var dataNK = sheetNK.getDataRange().getValues();
      var tenChuanHS = chuanHoaHoTen(hocSinhTimThay.hoTen);

      for (var j = 1; j < dataNK.length; j++) {
        var tenNhatKy = chuanHoaHoTen(dataNK[j][3]);
        if (tenNhatKy === tenChuanHS) {
          var tuanDong = dataNK[j][2] ? dataNK[j][2].toString().trim() : "";
          var soTuanDong = tuanDong.replace(/[^0-9]/g, '');

          var khopTuan = locTatCa || (tuanDong.toLowerCase() === (tuanChon || "").toString().trim().toLowerCase()) || (soTuanDong !== "" && soTuanDong === soTuanChon);

          if (khopTuan) {
            var diemSo = dataNK[j][5] !== "" ? Number(dataNK[j][5]) : 0;
            tongDiemThayDoi += diemSo;
            if (diemSo > 0) tongCong += diemSo;
            if (diemSo < 0) tongTru += Math.abs(diemSo);

            var tgRaw = dataNK[j][1];
            var tgStr = "";
            if (tgRaw instanceof Date) {
              tgStr = Utilities.formatDate(tgRaw, "Asia/Ho_Chi_Minh", "HH:mm dd/MM/yyyy");
            } else {
              tgStr = tgRaw ? tgRaw.toString().trim() : "";
            }

            danhSachNhatKy.push({
              thoiGian: tgStr,
              tuan: tuanDong || ("Tuần " + soTuanDong),
              hanhVi: dataNK[j][4] ? dataNK[j][4].toString() : "",
              diem: diemSo,
              ghiChu: dataNK[j][6] ? dataNK[j][6].toString() : "",
              nguoiBao: dataNK[j][7] ? dataNK[j][7].toString() : ""
            });
          }
        }
      }
    }

    var xepLoai = "Tốt";
    if (tongTru === 0) {
      xepLoai = "Tốt";
    } else if (tongTru <= 2) {
      xepLoai = "Khá";
    } else if (tongTru <= 5) {
      xepLoai = "Trung bình";
    } else {
      xepLoai = "Cần rèn luyện thêm";
    }

    var nhanXet = "";
    if (tongTru === 0) {
      nhanXet = "🌟 Tuyệt vời! Trong tuần qua em chấp hành rất nghiêm túc mọi nề nếp của trường, lớp. Gia đình hãy cùng khen ngợi và động viên em tiếp tục phát huy nhé!";
    } else if (tongTru === 1) {
      nhanXet = "🔔 Nhìn chung em thực hiện tốt nề nếp, chỉ có sơ suất nhỏ bị trừ 1 điểm. Gia đình nhắc con chú ý chuẩn bị chu đáo hơn để tuần tới đạt kết quả trọn vẹn nhé!";
    } else if (tongTru >= 2 && tongTru <= 5) {
      nhanXet = "📝 Tuần này em có vi phạm nhỏ về nề nếp (bị trừ " + tongTru + " điểm). Kính mong gia đình nhắc nhở, động viên con ngủ đủ giấc và tập trung hơn trong giờ học nhé!";
    } else if (tongTru >= 6 && tongTru <= 10) {
      nhanXet = "⚠️ Tuần này nề nếp của em có phần giảm sút (bị trừ " + tongTru + " điểm). Kính đề nghị gia đình sát sao đôn đốc, cùng phối hợp uốn nắn để em chấn chỉnh kịp thời.";
    } else {
      nhanXet = "❗ Tuần này em vi phạm nề nếp khá nhiều (bị trừ " + tongTru + " điểm). Kính mong quý phụ huynh sắp xếp trao đổi trực tiếp với GVCN để cùng có biện pháp hỗ trợ em tốt nhất.";
    }

    return {
      success: true,
      thanhCong: true,
      hocSinh: hocSinhTimThay,
      tuan: locTatCa ? "Toàn bộ học kỳ" : tuanChon,
      tongCong: tongCong,
      tongTru: tongTru,
      tongDiemThayDoi: tongDiemThayDoi,
      xepLoai: xepLoai,
      nhanXet: nhanXet,
      danhSach: danhSachNhatKy
    };

  } catch (err) {
    return {
      success: false,
      thanhCong: false,
      message: "Lỗi hệ thống: " + err.toString(),
      thongBao: "Lỗi hệ thống: " + err.toString()
    };
  }
}
