/* ========================================
   CÔNG ĐOÀN SỐ – Application Logic
   ======================================== */

// ── DATA: DL_CHAM_LO (48 records from Excel) ──
const DL_CHAM_LO = [
  {ma:"LS-001",ngay:46048,nam:2026,thang:1,noiDung:"Chi tiền viếng bố đẻ đc Lò Văn Sơn – CĐ Hạ tấng",loai:"Hiếu/viếng",soTien:800000,nguoiHoTro:null,doanVien:"Lò Văn Sơn",quanHe:null,lyDo:"Mất",nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46048,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-002",ngay:46055,nam:2026,thang:2,noiDung:"Chi tiền tặng quà ĐVCĐ nghỉ việc CĐ TT Hạ tầng",loai:"Quà đoàn viên nghỉ việc",soTien:2500000,nguoiHoTro:null,doanVien:null,quanHe:null,lyDo:null,nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46055,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-003",ngay:46066,nam:2026,thang:2,noiDung:"Chi tien tham vo dc Tuan sinh con CD TT Ha tang",loai:"Hỷ/Sinh con",soTien:300000,nguoiHoTro:null,doanVien:"Đồng chí Tuấn",quanHe:null,lyDo:null,nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46066,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-004",ngay:46063,nam:2026,thang:2,noiDung:"Chứng từ chi tiền sinh nhật quý 4/2025",loai:"Sinh nhật",soTien:14100000,nguoiHoTro:null,doanVien:null,quanHe:null,lyDo:null,nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46063,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-005",ngay:46063,nam:2026,thang:2,noiDung:"Chứng từ chi tiền sinh nhật quý 1/2026 + 02/2026",loai:"Sinh nhật",soTien:9600000,nguoiHoTro:null,doanVien:null,quanHe:null,lyDo:null,nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46063,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-006",ngay:46063,nam:2026,thang:2,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Hiếu/viếng",soTien:1000000,nguoiHoTro:"Ngô Minh Hòa",doanVien:"Nguyễn Minh Cường",quanHe:null,lyDo:"Bác bị mất ngày 14/11/2025",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46063,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-007",ngay:46063,nam:2026,thang:2,noiDung:"Chi thăm hỏi đồng chí Bùi Văn Tình",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Bùi Văn Tình",doanVien:"Bùi Văn Tình",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Bùi Văn Tình",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46063,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-008",ngay:46063,nam:2026,thang:2,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Trần Đức Tân",doanVien:"Trần Tuyên",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46063,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-009",ngay:46063,nam:2026,thang:2,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Nguyễn Đắc Thao",doanVien:"Nguyễn Mạnh Cường",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46063,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-010",ngay:46091,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Hỷ/Sinh con",soTien:300000,nguoiHoTro:"Nguyễn Thị Hồng Vân",doanVien:"Ôn Đức Phương",quanHe:"Vợ",lyDo:"Sinh em bé",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-011",ngay:46091,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Hỷ/Sinh con",soTien:300000,nguoiHoTro:"Nguyễn Thị Thanh Hải",doanVien:"Cao Văn Huy",quanHe:"Vợ",lyDo:"Sinh em bé",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-012",ngay:46091,nam:2026,thang:3,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Hiếu/viếng",soTien:1000000,nguoiHoTro:"Đỗ Thị Niên",doanVien:"Trần Anh Quân",quanHe:null,lyDo:"Bác bị mất ngày 26/1/2026",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-013",ngay:46091,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Nguyễn Thị Phương Anh",doanVien:"Nguyễn Thị Phương Anh",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Ngô Thị Anh",diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-014",ngay:46091,nam:2026,thang:3,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Hiếu/viếng",soTien:500000,nguoiHoTro:"Ma Văn Tuất",doanVien:"Ma Văn Nam",quanHe:"Bố đẻ",lyDo:"Bác bị mất ngày 18/2/2026",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:"TTVT Thanh Thủy",trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-015",ngay:46091,nam:2026,thang:3,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Hiếu/viếng",soTien:1000000,nguoiHoTro:"Đinh Thị Vân",doanVien:"Ngô Anh Đức",quanHe:"Mẹ vợ",lyDo:"Bác bị mất ngày 28/2/2026",nguoiThucHien:"Ngô Thị Anh",diaBan:null,donVi:"VNPT Vĩnh Tường",trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-016",ngay:46091,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Bùi Văn Xuân",doanVien:"Bùi Văn Đông",quanHe:"Bố",lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-017",ngay:46091,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Nguyễn Trường Sơn",doanVien:"Nguyễn Trường Sơn",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-018",ngay:46091,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Bùi Văn Lệch",doanVien:"Bùi Văn Nhất",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-019",ngay:46091,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Hỷ/Sinh con",soTien:300000,nguoiHoTro:"Đường Thị Ngọc Trang",doanVien:"Nguyễn Tuấn Linh",quanHe:"Vợ",lyDo:"Sinh em bé",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46091,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-020",ngay:46104,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Nguyễn Bảo Ngọc",doanVien:"Hoàng Phương Nhung",quanHe:"Con gái",lyDo:"Ốm Đi viện",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46104,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-021",ngay:46104,nam:2026,thang:3,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Hiếu/viếng",soTien:1000000,nguoiHoTro:"Nguyễn Văn Hồng",doanVien:"Nguyễn Văn Quân",quanHe:"Bố đẻ",lyDo:"Bác bị mất ngày 20/03/2026",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:"TTVT Vĩnh Tường",trangThai:"Đã hoàn thành",ngayHT:46104,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-022",ngay:46104,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Lê Quân Nhạc",doanVien:"Bùi Thanh Quang",quanHe:"Bố vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46104,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-023",ngay:46104,nam:2026,thang:3,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Hoàng Trung Quân",doanVien:"Hoàng Trung Dũng",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46104,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-024",ngay:46119,nam:2026,thang:4,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Tăng Bá Trứ",doanVien:"Nguyễn Thị Ánh Tuyết",quanHe:"bố",lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46119,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-025",ngay:46119,nam:2026,thang:4,noiDung:"Thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:500000,nguoiHoTro:"Nguyễn Thị Thanh Tuyết",doanVien:"Hoàng Xuân Sửu",quanHe:"vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46119,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-026",ngay:46126,nam:2026,thang:4,noiDung:"Chi tiền sinh nhật tháng 03/2026",loai:"Sinh nhật",soTien:4500000,nguoiHoTro:null,doanVien:null,quanHe:null,lyDo:null,nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46126,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-027",ngay:46126,nam:2026,thang:4,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Bùi Văn Nhất",doanVien:"Bùi Văn Nhất",quanHe:"đồng chí",lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46126,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-028",ngay:46128,nam:2026,thang:4,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Nguyễn Thị Phương",doanVien:"Nguyễn Minh Chí",quanHe:"vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46128,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-029",ngay:46128,nam:2026,thang:4,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Hạ Thị Ái",doanVien:"Đặng Hồng Quân",quanHe:"Mẹ vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46128,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-030",ngay:46128,nam:2026,thang:4,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Hiếu/viếng",soTien:1000000,nguoiHoTro:"Phạm Thị Liên",doanVien:"Nguyễn Văn Mạnh",quanHe:"Mẹ vợ",lyDo:"Bác bị mất ngày 30/3/2026",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:"TTVT Thanh Sơn",trangThai:"Đã hoàn thành",ngayHT:46128,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-031",ngay:46128,nam:2026,thang:4,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Hiếu/viếng",soTien:1000000,nguoiHoTro:"Đặng Đình Bình",doanVien:"Đặng Văn Đình",quanHe:"Bố đẻ",lyDo:"Bác bị mất ngày 01/4/2026",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:"TTVT Tam Dương",trangThai:"Đã hoàn thành",ngayHT:46128,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-032",ngay:46145,nam:2026,thang:5,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Lê Thị Vân Anh",doanVien:"Nguyễn Đức Hòa",quanHe:"vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46145,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-033",ngay:46145,nam:2026,thang:5,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Đỗ Thị Thanh Giang",doanVien:"Bùi Đình Cương",quanHe:"vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46145,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-034",ngay:46153,nam:2026,thang:5,noiDung:"Chi tiền sinh nhật tháng 04/2026",loai:"Sinh nhật",soTien:4500000,nguoiHoTro:null,doanVien:null,quanHe:null,lyDo:null,nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46153,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-035",ngay:46153,nam:2026,thang:5,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Nguyễn Thị Thu Huyền",doanVien:"Đoàn Quốc Cường",quanHe:"vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46153,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-036",ngay:46153,nam:2026,thang:5,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Phạm Văn Tùy",doanVien:"Phạm Văn Tùy",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46153,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-037",ngay:46153,nam:2026,thang:5,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Bùi Thị Thanh Thúy",doanVien:"Nguyễn Tuấn Linh",quanHe:"vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46153,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-038",ngay:46163,nam:2026,thang:5,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Tai nạn/rủi ro",soTien:3000000,nguoiHoTro:"Nguyễn Văn Cường",doanVien:"Nguyễn Văn Cường",quanHe:null,lyDo:"Tai nạn lao động",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46163,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-039",ngay:46175,nam:2026,thang:6,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Nguyễn Văn Đức",doanVien:"Nguyễn Văn Đức",quanHe:null,lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46175,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-040",ngay:46175,nam:2026,thang:6,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Phạm Thị Duyên",doanVien:"Đinh Quang Vinh",quanHe:"vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Nguyễn Văn Phây",diaBan:"Vĩnh phúc",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46175,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-041",ngay:46175,nam:2026,thang:6,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Hiếu/viếng",soTien:1000000,nguoiHoTro:"Phạm Văn Sinh",doanVien:"Phạm Văn Tùy",quanHe:"Bố đẻ",lyDo:"Bác bị mất ngày 12/5/2026",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:"TTVT Thanh Sơn",trangThai:"Đã hoàn thành",ngayHT:46175,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-042",ngay:46175,nam:2026,thang:6,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Tai nạn/rủi ro",soTien:3000000,nguoiHoTro:"Bùi Tuấn Anh",doanVien:"Bùi Tuấn Anh",quanHe:null,lyDo:"Tai nạn lao động",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46175,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-043",ngay:46183,nam:2026,thang:6,noiDung:"Chi tiền sinh nhật tháng 05/2026",loai:"Sinh nhật",soTien:1200000,nguoiHoTro:null,doanVien:null,quanHe:null,lyDo:null,nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Đã hoàn thành",ngayHT:46183,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-044",ngay:46183,nam:2026,thang:6,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Hỷ/Sinh con",soTien:200000,nguoiHoTro:"Trần Thị Vân Anh",doanVien:"Nguyễn Tiến Cường",quanHe:"vợ",lyDo:"Sinh em bé",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46183,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-045",ngay:46205,nam:2026,thang:7,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Hà Thị Hoa",doanVien:"Bùi Quang Hải",quanHe:"vợ",lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46205,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-046",ngay:46205,nam:2026,thang:7,noiDung:"Chi tiền thăm hỏi đoàn viên công đoàn",loai:"Ốm đau/đi viện",soTien:300000,nguoiHoTro:"Bùi Đình Cương",doanVien:"Bùi Đình Cương",quanHe:"đồng chí",lyDo:"Ốm Đi viện",nguoiThucHien:"Dương Văn Nam",diaBan:"Hòa Bình",donVi:null,trangThai:"Đã hoàn thành",ngayHT:46205,soNgayXL:0,canhBao:"Đúng hạn"},
  {ma:"LS-047",ngay:46205,nam:2026,thang:7,noiDung:"Chi tiền thanh toán lễ thăm viếng",loai:"Tai nạn/rủi ro",soTien:3300000,nguoiHoTro:"Lê Văn Quyền",doanVien:"Lê Văn Quyền",quanHe:null,lyDo:"Tai nạn giao thông",nguoiThucHien:"Lê Anh Tuấn",diaBan:"Phú Thọ",donVi:null,trangThai:"Đang xử lý",ngayHT:null,soNgayXL:null,canhBao:"Đang xử lý"},
  {ma:"LS-048",ngay:46205,nam:2026,thang:7,noiDung:"Chi tiền sinh nhật tháng 06/2026",loai:"Sinh nhật",soTien:0,nguoiHoTro:null,doanVien:null,quanHe:null,lyDo:null,nguoiThucHien:null,diaBan:null,donVi:null,trangThai:"Chờ duyệt",ngayHT:null,soNgayXL:null,canhBao:"Chưa duyệt"}
];

// ── DATA: QUY_CHE ──
const QUY_CHE = [
  {ma:"OM_DAU",ten:"Ốm đau/đi viện",tuKhoa:["ốm","đi viện","điều trị","bệnh viện","phẫu thuật","nằm viện"],muc:300000,hoSo:"Đề nghị/Thông tin xác nhận; danh sách chi",thoiHan:"02 ngày",ghiChu:"Áp dụng theo quy chế thực tế",icon:"🏥",color:"type-om-dau",gradient:"linear-gradient(135deg,#3b82f6,#60a5fa)"},
  {ma:"HIEU",ten:"Hiếu/viếng",tuKhoa:["mất","viếng","bố","mẹ","vợ","chồng","tang","qua đời"],muc:1000000,hoSo:"Thông tin sự việc; danh sách chi; chứng từ",thoiHan:"01 ngày",ghiChu:"Mức có thể khác theo quan hệ",icon:"🕊️",color:"type-hieu",gradient:"linear-gradient(135deg,#8b5cf6,#a78bfa)"},
  {ma:"HY_SINH_CON",ten:"Hỷ/Sinh con",tuKhoa:["sinh con","sinh em bé","sinh","hỷ"],muc:300000,hoSo:"Thông tin đoàn viên; danh sách chi",thoiHan:"03 ngày",ghiChu:"Theo quy chế thực tế",icon:"👶",color:"type-hy-sinh",gradient:"linear-gradient(135deg,#ec4899,#f472b6)"},
  {ma:"TAI_NAN_RUI_RO",ten:"Tai nạn/rủi ro",tuKhoa:["tai nạn","hỏa hoạn","rủi ro","tai nạn lao động","tai nạn giao thông"],muc:2000000,hoSo:"Tờ trình; xác minh; quyết định; chứng từ",thoiHan:"01 ngày",ghiChu:"Có thể trình trợ cấp bổ sung",icon:"⚠️",color:"type-tai-nan",gradient:"linear-gradient(135deg,#ef4444,#f87171)"},
  {ma:"SINH_NHAT",ten:"Sinh nhật",tuKhoa:["sinh nhật"],muc:100000,hoSo:"Danh sách sinh nhật; quyết định/đề nghị chi",thoiHan:"Theo tháng",ghiChu:"Mức bình quân/người",icon:"🎂",color:"type-sinh-nhat",gradient:"linear-gradient(135deg,#f59e0b,#fbbf24)"},
  {ma:"NGHI_VIEC",ten:"Quà đoàn viên nghỉ việc",tuKhoa:["nghỉ việc","chia tay","chuyển công tác"],muc:500000,hoSo:"Danh sách; quyết định/đề nghị chi",thoiHan:"05 ngày",ghiChu:"Theo chủ trương từng thời kỳ",icon:"🎁",color:"type-nghi-viec",gradient:"linear-gradient(135deg,#22c55e,#4ade80)"},
  {ma:"KHAC",ten:"Khác",tuKhoa:["khác"],muc:0,hoSo:"Hồ sơ theo nội dung phát sinh",thoiHan:"05 ngày",ghiChu:"BCH xem xét riêng",icon:"📁",color:"type-khac",gradient:"linear-gradient(135deg,#64748b,#94a3b8)"}
];

// ── SUBMITTED CASES (localStorage) ──
function getSubmittedCases() {
  try {
    return JSON.parse(localStorage.getItem('congdoan_submitted') || '[]');
  } catch { return []; }
}
function saveSubmittedCases(cases) {
  localStorage.setItem('congdoan_submitted', JSON.stringify(cases));
}

// ── HELPERS ──
function formatCurrency(n) {
  if (n == null || isNaN(n)) return '0';
  return new Intl.NumberFormat('vi-VN').format(n) + 'đ';
}

function formatCurrencyShort(n) {
  if (n >= 1e9) return (n / 1e9).toFixed(1) + ' tỷ';
  if (n >= 1e6) return (n / 1e6).toFixed(1) + ' tr';
  if (n >= 1e3) return (n / 1e3).toFixed(0) + 'k';
  return n + 'đ';
}

function excelDateToStr(serial) {
  if (!serial || serial < 1) return '--';
  const d = new Date((serial - 25569) * 86400 * 1000);
  return d.toLocaleDateString('vi-VN');
}

function getTypeClass(loai) {
  const map = {
    'Ốm đau/đi viện': 'type-om-dau',
    'Hiếu/viếng': 'type-hieu',
    'Hỷ/Sinh con': 'type-hy-sinh',
    'Tai nạn/rủi ro': 'type-tai-nan',
    'Sinh nhật': 'type-sinh-nhat',
    'Quà đoàn viên nghỉ việc': 'type-nghi-viec',
    'Khác': 'type-khac'
  };
  return map[loai] || 'type-khac';
}

function getStatusBadge(status) {
  if (status === 'Đã hoàn thành') return '<span class="badge badge-success">✓ Hoàn thành</span>';
  if (status === 'Đang xử lý') return '<span class="badge badge-warning">⏳ Đang xử lý</span>';
  if (status === 'Chờ duyệt') return '<span class="badge badge-info">📋 Chờ duyệt</span>';
  return '<span class="badge badge-neutral">' + (status || '--') + '</span>';
}

function getCanhBaoBadge(cb) {
  if (cb === 'Đúng hạn') return '<span class="badge badge-success">Đúng hạn</span>';
  if (cb === 'Đang xử lý') return '<span class="badge badge-warning">Đang xử lý</span>';
  if (cb === 'Chưa duyệt') return '<span class="badge badge-info">Chưa duyệt</span>';
  if (cb === 'Trễ hạn') return '<span class="badge badge-danger">Trễ hạn</span>';
  return '<span class="badge badge-neutral">' + (cb || '--') + '</span>';
}

// ── NAVIGATION ──
const pages = ['dashboard', 'tiepnhan', 'hoso', 'quyche'];
const pageNames = {
  dashboard: { title: 'Dashboard', breadcrumb: 'Tổng quan' },
  tiepnhan: { title: 'Tiếp nhận', breadcrumb: 'Nhập đề nghị' },
  hoso: { title: 'Hồ sơ chăm lo', breadcrumb: 'DL_CHAM_LO' },
  quyche: { title: 'Quy chế', breadcrumb: 'Danh mục & mức chi' }
};

let currentPage = 'dashboard';

function navigateTo(page) {
  if (!pages.includes(page)) return;
  currentPage = page;

  // Hide all pages
  pages.forEach(p => {
    const el = document.getElementById('page-' + p);
    if (el) el.classList.toggle('page-hidden', p !== page);
  });

  // Update nav
  document.querySelectorAll('.nav-item').forEach(n => {
    n.classList.toggle('active', n.dataset.page === page);
  });
  document.querySelectorAll('.bottom-nav-item').forEach(n => {
    n.classList.toggle('active', n.dataset.page === page);
  });

  // Update topbar
  const info = pageNames[page];
  document.getElementById('topbarTitle').textContent = info.title;
  document.getElementById('topbarPage').textContent = info.breadcrumb;

  // Close mobile sidebar
  closeSidebar();

  // Re-animate page
  const pageEl = document.getElementById('page-' + page);
  pageEl.style.animation = 'none';
  pageEl.offsetHeight; // trigger reflow
  pageEl.style.animation = '';

  // Re-render if needed
  if (page === 'dashboard') renderDashboard();
  if (page === 'hoso') renderHoSoTable();
  if (page === 'quyche') renderQuyche();
  if (page === 'tiepnhan') renderSubmitted();

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── SIDEBAR MOBILE ──
function openSidebar() {
  document.getElementById('sidebar').classList.add('open');
  document.getElementById('sidebarOverlay').classList.add('active');
}

function closeSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebarOverlay').classList.remove('active');
}

// ── TOAST ──
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const icons = { success: '✅', error: '❌', info: 'ℹ️' };
  const toast = document.createElement('div');
  toast.className = 'toast ' + type;
  toast.innerHTML = '<span class="toast-icon">' + icons[type] + '</span><span>' + message + '</span>';
  container.appendChild(toast);
  setTimeout(() => {
    toast.classList.add('fade-out');
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

// ── ANIMATED COUNTER ──
function animateCounters() {
  document.querySelectorAll('.kpi-value[data-count]').forEach(el => {
    const target = parseInt(el.dataset.count);
    const isCurrency = el.dataset.format === 'currency';
    const duration = 1200;
    const start = performance.now();

    function update(now) {
      const elapsed = now - start;
      const progress = Math.min(elapsed / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3); // ease-out cubic
      const current = Math.round(target * eased);

      if (isCurrency) {
        el.textContent = formatCurrencyShort(current);
      } else {
        el.textContent = new Intl.NumberFormat('vi-VN').format(current);
      }

      if (progress < 1) requestAnimationFrame(update);
    }

    requestAnimationFrame(update);
  });
}

// ── CHARTS ──
let chartLoai, chartDiaBan, chartThang;

function destroyCharts() {
  if (chartLoai) { chartLoai.destroy(); chartLoai = null; }
  if (chartDiaBan) { chartDiaBan.destroy(); chartDiaBan = null; }
  if (chartThang) { chartThang.destroy(); chartThang = null; }
}

function buildCharts() {
  destroyCharts();

  // Compute data from DL_CHAM_LO
  const allData = [...DL_CHAM_LO, ...getSubmittedCases().filter(c => c.trangThai === 'Đã hoàn thành')];

  // Chart 1: Loai cham lo (Donut)
  const loaiMap = {};
  allData.forEach(r => {
    const k = r.loai || 'Khác';
    loaiMap[k] = (loaiMap[k] || 0) + 1;
  });

  const loaiLabels = Object.keys(loaiMap);
  const loaiData = Object.values(loaiMap);
  const loaiColors = ['#818cf8', '#a78bfa', '#f472b6', '#f87171', '#fbbf24', '#4ade80', '#94a3b8'];

  const ctx1 = document.getElementById('chartLoai').getContext('2d');
  chartLoai = new Chart(ctx1, {
    type: 'doughnut',
    data: {
      labels: loaiLabels,
      datasets: [{
        data: loaiData,
        backgroundColor: loaiColors.slice(0, loaiLabels.length),
        borderColor: 'rgba(252,250,248,0.95)',
        borderWidth: 3,
        hoverBorderWidth: 0,
        hoverOffset: 8
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '60%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: '#475569',
            font: { family: 'Inter', size: 11 },
            padding: 12,
            usePointStyle: true,
            pointStyleWidth: 8
          }
        },
        tooltip: {
          backgroundColor: 'rgba(255,255,255,0.95)',
          titleColor: '#1e293b',
          bodyColor: '#475569',
          borderColor: 'rgba(0,0,0,0.08)',
          borderWidth: 1,
          cornerRadius: 8,
          padding: 12,
          callbacks: {
            label: function(ctx) {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = ((ctx.parsed / total) * 100).toFixed(1);
              return ctx.label + ': ' + ctx.parsed + ' lượt (' + pct + '%)';
            }
          }
        }
      }
    }
  });

  // Chart 2: Dia ban (Bar)
  const diaBanMap = {};
  const diaBanMoneyMap = {};
  allData.forEach(r => {
    const k = r.diaBan || 'Chưa xác định';
    diaBanMap[k] = (diaBanMap[k] || 0) + 1;
    diaBanMoneyMap[k] = (diaBanMoneyMap[k] || 0) + (r.soTien || 0);
  });

  const dbLabels = Object.keys(diaBanMap);
  const dbCounts = Object.values(diaBanMap);
  const dbMoney = dbLabels.map(k => diaBanMoneyMap[k] / 1e6); // millions

  const ctx2 = document.getElementById('chartDiaBan').getContext('2d');
  chartDiaBan = new Chart(ctx2, {
    type: 'bar',
    data: {
      labels: dbLabels,
      datasets: [
        {
          label: 'Số lượt',
          data: dbCounts,
          backgroundColor: 'rgba(99, 102, 241, 0.7)',
          borderColor: 'rgba(99, 102, 241, 1)',
          borderWidth: 1,
          borderRadius: 6,
          yAxisID: 'y'
        },
        {
          label: 'Kinh phí (triệu)',
          data: dbMoney,
          backgroundColor: 'rgba(14, 165, 233, 0.4)',
          borderColor: 'rgba(14, 165, 233, 1)',
          borderWidth: 1,
          borderRadius: 6,
          yAxisID: 'y1'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          labels: { color: '#475569', font: { family: 'Inter', size: 11 }, usePointStyle: true, pointStyleWidth: 8 }
        },
        tooltip: {
          backgroundColor: 'rgba(255,255,255,0.95)',
          titleColor: '#1e293b',
          bodyColor: '#475569',
          borderColor: 'rgba(0,0,0,0.08)',
          borderWidth: 1,
          cornerRadius: 8,
          padding: 12
        }
      },
      scales: {
        x: {
          ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
          grid: { color: 'rgba(0,0,0,0.06)' }
        },
        y: {
          position: 'left',
          ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
          grid: { color: 'rgba(0,0,0,0.06)' },
          title: { display: true, text: 'Số lượt', color: '#64748b', font: { family: 'Inter', size: 11 } }
        },
        y1: {
          position: 'right',
          ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
          grid: { display: false },
          title: { display: true, text: 'Triệu đồng', color: '#64748b', font: { family: 'Inter', size: 11 } }
        }
      }
    }
  });

  // Chart 3: Monthly trend (Line)
  const monthMap = {};
  const monthMoney = {};
  allData.forEach(r => {
    const m = r.thang || 0;
    if (m >= 1 && m <= 12) {
      monthMap[m] = (monthMap[m] || 0) + 1;
      monthMoney[m] = (monthMoney[m] || 0) + (r.soTien || 0);
    }
  });

  const months = Array.from({length: 12}, (_, i) => i + 1);
  const monthLabels = months.map(m => 'T' + m);
  const monthCounts = months.map(m => monthMap[m] || 0);
  const monthMoneyData = months.map(m => (monthMoney[m] || 0) / 1e6);

  const ctx3 = document.getElementById('chartThang').getContext('2d');
  chartThang = new Chart(ctx3, {
    type: 'line',
    data: {
      labels: monthLabels,
      datasets: [
        {
          label: 'Số lượt',
          data: monthCounts,
          borderColor: '#6366f1',
          backgroundColor: 'rgba(99, 102, 241, 0.1)',
          fill: true,
          tension: 0.4,
          pointBackgroundColor: '#6366f1',
          pointBorderColor: '#ffffff',
          pointBorderWidth: 2,
          pointRadius: 5,
          pointHoverRadius: 8,
          yAxisID: 'y'
        },
        {
          label: 'Kinh phí (triệu)',
          data: monthMoneyData,
          borderColor: '#0ea5e9',
          backgroundColor: 'rgba(14, 165, 233, 0.05)',
          fill: true,
          tension: 0.4,
          pointBackgroundColor: '#0ea5e9',
          pointBorderColor: '#ffffff',
          pointBorderWidth: 2,
          pointRadius: 5,
          pointHoverRadius: 8,
          yAxisID: 'y1'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'index',
        intersect: false
      },
      plugins: {
        legend: {
          labels: { color: '#475569', font: { family: 'Inter', size: 11 }, usePointStyle: true, pointStyleWidth: 8 }
        },
        tooltip: {
          backgroundColor: 'rgba(255,255,255,0.95)',
          titleColor: '#1e293b',
          bodyColor: '#475569',
          borderColor: 'rgba(0,0,0,0.08)',
          borderWidth: 1,
          cornerRadius: 8,
          padding: 12
        }
      },
      scales: {
        x: {
          ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
          grid: { color: 'rgba(0,0,0,0.06)' }
        },
        y: {
          position: 'left',
          ticks: { color: '#64748b', font: { family: 'Inter', size: 11 }, stepSize: 2 },
          grid: { color: 'rgba(0,0,0,0.06)' },
          title: { display: true, text: 'Số lượt', color: '#64748b', font: { family: 'Inter', size: 11 } }
        },
        y1: {
          position: 'right',
          ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
          grid: { display: false },
          title: { display: true, text: 'Triệu đồng', color: '#64748b', font: { family: 'Inter', size: 11 } }
        }
      }
    }
  });
}

// ── DASHBOARD: Recent table ──
function renderRecentTable() {
  const tbody = document.getElementById('recentTableBody');
  const recent = [...DL_CHAM_LO].reverse().slice(0, 8);
  tbody.innerHTML = recent.map(r => `
    <tr onclick="showDetail('${r.ma}')">
      <td>${r.ma}</td>
      <td style="max-width:250px;overflow:hidden;text-overflow:ellipsis">${r.noiDung}</td>
      <td><span class="type-tag ${getTypeClass(r.loai)}">${r.loai}</span></td>
      <td style="font-weight:600;color:var(--text-primary)">${formatCurrency(r.soTien)}</td>
      <td>${getStatusBadge(r.trangThai)}</td>
      <td>${getCanhBaoBadge(r.canhBao)}</td>
    </tr>
  `).join('');
}

function renderDashboard() {
  // Update KPIs with real data
  const allData = [...DL_CHAM_LO, ...getSubmittedCases().filter(c => c.trangThai === 'Đã hoàn thành')];
  const totalLuot = allData.length;
  const totalTien = allData.reduce((s, r) => s + (r.soTien || 0), 0);
  const diaBans = new Set(allData.map(r => r.diaBan).filter(Boolean));
  const avg = totalLuot > 0 ? Math.round(totalTien / totalLuot) : 0;

  const kpiEls = document.querySelectorAll('.kpi-value[data-count]');
  if (kpiEls[0]) kpiEls[0].dataset.count = totalLuot;
  if (kpiEls[1]) kpiEls[1].dataset.count = totalTien;
  if (kpiEls[2]) kpiEls[2].dataset.count = diaBans.size || 4;
  if (kpiEls[3]) kpiEls[3].dataset.count = avg;

  animateCounters();
  buildCharts();
  renderRecentTable();

  // Update badge
  const pending = getSubmittedCases().filter(c => c.trangThai === 'Chờ duyệt').length;
  document.getElementById('badgePending').textContent = pending;
}

// ── HO SO TABLE ──
let hosoPage = 1;
const hosoPerPage = 15;

function getFilteredHoSo() {
  const search = (document.getElementById('searchHoSo')?.value || '').toLowerCase().trim();
  const filterLoai = document.getElementById('filterLoai')?.value || '';
  const filterDiaBan = document.getElementById('filterDiaBan')?.value || '';
  const filterThang = document.getElementById('filterThang')?.value || '';
  const filterTrangThai = document.getElementById('filterTrangThai')?.value || '';

  const all = [...DL_CHAM_LO, ...getSubmittedCases()];

  return all.filter(r => {
    if (search) {
      const haystack = [r.ma, r.noiDung, r.loai, r.doanVien, r.nguoiHoTro, r.diaBan, r.nguoiThucHien]
        .filter(Boolean).join(' ').toLowerCase();
      if (!haystack.includes(search)) return false;
    }
    if (filterLoai && r.loai !== filterLoai) return false;
    if (filterDiaBan && r.diaBan !== filterDiaBan) return false;
    if (filterThang && String(r.thang) !== filterThang) return false;
    if (filterTrangThai && r.trangThai !== filterTrangThai) return false;
    return true;
  });
}

function renderHoSoTable() {
  const filtered = getFilteredHoSo();
  const totalPages = Math.max(1, Math.ceil(filtered.length / hosoPerPage));
  if (hosoPage > totalPages) hosoPage = totalPages;

  const start = (hosoPage - 1) * hosoPerPage;
  const pageData = filtered.slice(start, start + hosoPerPage);

  const tbody = document.getElementById('hosoTableBody');
  if (pageData.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9"><div class="empty-state"><div class="empty-icon">📭</div><h3>Không tìm thấy hồ sơ</h3><p>Thử thay đổi bộ lọc</p></div></td></tr>';
  } else {
    tbody.innerHTML = pageData.map(r => `
      <tr onclick="showDetail('${r.ma}')">
        <td>${r.ma}</td>
        <td>T${r.thang}/${r.nam}</td>
        <td style="max-width:220px;overflow:hidden;text-overflow:ellipsis">${r.noiDung}</td>
        <td><span class="type-tag ${getTypeClass(r.loai)}">${r.loai}</span></td>
        <td style="font-weight:600;color:var(--text-primary)">${formatCurrency(r.soTien)}</td>
        <td>${r.doanVien || '--'}</td>
        <td>${r.diaBan || '--'}</td>
        <td>${getStatusBadge(r.trangThai)}</td>
        <td>${getCanhBaoBadge(r.canhBao)}</td>
      </tr>
    `).join('');
  }

  // Summary
  const totalTien = filtered.reduce((s, r) => s + (r.soTien || 0), 0);
  document.getElementById('hosoSummary').innerHTML = `
    <div class="summary-item">Tổng: <span class="summary-value">${filtered.length} hồ sơ</span></div>
    <div class="summary-item">Kinh phí: <span class="summary-value">${formatCurrency(totalTien)}</span></div>
    <div class="summary-item">Hoàn thành: <span class="summary-value">${filtered.filter(r => r.trangThai === 'Đã hoàn thành').length}</span></div>
    <div class="summary-item">Đang xử lý: <span class="summary-value">${filtered.filter(r => r.trangThai !== 'Đã hoàn thành').length}</span></div>
  `;

  // Pagination
  document.getElementById('paginationInfo').textContent =
    `Hiển thị ${start + 1}–${Math.min(start + hosoPerPage, filtered.length)} / ${filtered.length}`;

  document.getElementById('totalRecords').textContent = filtered.length + ' hồ sơ';

  const controls = document.getElementById('paginationControls');
  let paginationHTML = '';
  paginationHTML += `<button class="pagination-btn" onclick="goPage(${hosoPage - 1})" ${hosoPage <= 1 ? 'disabled' : ''}>‹</button>`;
  for (let i = 1; i <= totalPages; i++) {
    if (totalPages <= 7 || i === 1 || i === totalPages || Math.abs(i - hosoPage) <= 1) {
      paginationHTML += `<button class="pagination-btn ${i === hosoPage ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
    } else if (i === hosoPage - 2 || i === hosoPage + 2) {
      paginationHTML += `<span style="color:var(--text-tertiary);padding:0 4px">…</span>`;
    }
  }
  paginationHTML += `<button class="pagination-btn" onclick="goPage(${hosoPage + 1})" ${hosoPage >= totalPages ? 'disabled' : ''}>›</button>`;
  controls.innerHTML = paginationHTML;
}

function goPage(p) {
  hosoPage = p;
  renderHoSoTable();
}

// ── DETAIL MODAL ──
function showDetail(ma) {
  const all = [...DL_CHAM_LO, ...getSubmittedCases()];
  const record = all.find(r => r.ma === ma);
  if (!record) return;

  document.getElementById('modalTitle').textContent = 'Hồ sơ ' + record.ma;

  const fields = [
    ['Mã hồ sơ', record.ma],
    ['Ngày phát sinh', record.ngay ? excelDateToStr(record.ngay) : (record.ngayStr || '--')],
    ['Nội dung', record.noiDung],
    ['Loại chăm lo', `<span class="type-tag ${getTypeClass(record.loai)}">${record.loai}</span>`],
    ['Số tiền', `<strong style="color:var(--primary-600)">${formatCurrency(record.soTien)}</strong>`],
    ['Người được hỗ trợ', record.nguoiHoTro || '--'],
    ['Đoàn viên liên quan', record.doanVien || '--'],
    ['Quan hệ', record.quanHe || '--'],
    ['Lý do', record.lyDo || '--'],
    ['Người thực hiện', record.nguoiThucHien || '--'],
    ['Địa bàn', record.diaBan || '--'],
    ['Đơn vị', record.donVi || '--'],
    ['Trạng thái', getStatusBadge(record.trangThai)],
    ['Cảnh báo', getCanhBaoBadge(record.canhBao)]
  ];

  document.getElementById('modalBody').innerHTML = fields.map(([label, value]) =>
    `<div class="detail-row"><span class="detail-label">${label}</span><span class="detail-value">${value}</span></div>`
  ).join('');

  // Show approve/reject buttons for pending cases
  const footer = document.getElementById('modalFooter');
  if (record.trangThai === 'Chờ duyệt') {
    footer.innerHTML = `
      <button class="btn btn-sm btn-success" onclick="approveCase('${record.ma}')">✅ Duyệt</button>
      <button class="btn btn-sm btn-secondary" onclick="rejectCase('${record.ma}')" style="background:rgba(239,68,68,0.1);color:#dc2626;border-color:rgba(239,68,68,0.2)">❌ Từ chối</button>
      <button class="btn btn-sm btn-secondary" onclick="closeModal()">Đóng</button>
    `;
  } else {
    footer.innerHTML = `<button class="btn btn-sm btn-secondary" onclick="closeModal()">Đóng</button>`;
  }

  document.getElementById('detailModal').classList.add('active');
}

function approveCase(ma) {
  const cases = getSubmittedCases();
  const idx = cases.findIndex(c => c.ma === ma);
  if (idx === -1) return;

  cases[idx].trangThai = 'Đã hoàn thành';
  cases[idx].canhBao = 'Đúng hạn';
  cases[idx].ngayHT = Date.now();
  saveSubmittedCases(cases);

  closeModal();
  showToast(`Đã duyệt hồ sơ ${ma} thành công!`, 'success');

  // Refresh views
  renderSubmitted();
  updatePendingBadge();
  if (currentPage === 'hoso') renderHoSo();
  if (currentPage === 'dashboard') { buildCharts(); renderRecentTable(); animateCounters(); }
}

function rejectCase(ma) {
  const cases = getSubmittedCases();
  const idx = cases.findIndex(c => c.ma === ma);
  if (idx === -1) return;

  cases[idx].trangThai = 'Từ chối';
  cases[idx].canhBao = 'Từ chối';
  saveSubmittedCases(cases);

  closeModal();
  showToast(`Đã từ chối hồ sơ ${ma}`, 'warning');

  renderSubmitted();
  updatePendingBadge();
  if (currentPage === 'hoso') renderHoSo();
}

function updatePendingBadge() {
  const pending = getSubmittedCases().filter(c => c.trangThai === 'Chờ duyệt').length;
  document.getElementById('badgePending').textContent = pending;
}

function closeModal() {
  document.getElementById('detailModal').classList.remove('active');
}

// ── QUY CHE ──
function renderQuyche() {
  const grid = document.getElementById('quycheGrid');
  grid.innerHTML = QUY_CHE.map(q => `
    <div class="quyche-card">
      <div style="position:absolute;top:0;left:0;right:0;height:3px;background:${q.gradient}"></div>
      <div class="quyche-card-header">
        <div class="quyche-card-icon" style="background:${q.gradient.replace('135deg', '135deg')}22">${q.icon}</div>
        <div>
          <div class="quyche-card-title">${q.ten}</div>
          <div class="quyche-card-code">${q.ma}</div>
        </div>
      </div>
      <div class="quyche-amount">${q.muc > 0 ? formatCurrency(q.muc) : 'Theo BCH'}</div>
      <div class="quyche-info">
        <div class="quyche-info-row">
          <span class="info-label">Từ khóa:</span>
          <span class="info-value">${q.tuKhoa.join(', ')}</span>
        </div>
        <div class="quyche-info-row">
          <span class="info-label">Hồ sơ:</span>
          <span class="info-value">${q.hoSo}</span>
        </div>
        <div class="quyche-info-row">
          <span class="info-label">Thời hạn:</span>
          <span class="info-value">${q.thoiHan}</span>
        </div>
        <div class="quyche-info-row">
          <span class="info-label">Ghi chú:</span>
          <span class="info-value">${q.ghiChu}</span>
        </div>
      </div>
    </div>
  `).join('');
}

// ── TIEP NHAN: Form Logic ──
let nextId = 1;

function generateMaHoSo() {
  const submitted = getSubmittedCases();
  const year = new Date().getFullYear();
  const maxId = submitted.reduce((max, c) => {
    const match = c.ma.match(/TN-\d+-(\d+)/);
    return match ? Math.max(max, parseInt(match[1])) : max;
  }, 0);
  nextId = maxId + 1;
  const ma = `TN-${year}-${String(nextId).padStart(4, '0')}`;
  document.getElementById('inputMaHoSo').value = ma;
  document.getElementById('formMaHoSo').textContent = 'Mã: ' + ma;
  return ma;
}

function classify(text) {
  if (!text || text.trim().length < 2) return null;
  const lower = text.toLowerCase();

  for (const q of QUY_CHE) {
    if (q.ma === 'KHAC') continue; // check last
    for (const kw of q.tuKhoa) {
      if (lower.includes(kw.toLowerCase())) {
        return q;
      }
    }
  }
  return QUY_CHE.find(q => q.ma === 'KHAC');
}

function handleNoiDungChange() {
  const text = document.getElementById('inputNoiDung').value;
  const result = classify(text);

  if (result && text.trim().length > 3) {
    // Set classification
    document.getElementById('inputLoaiChamLo').value = result.ten;
    document.getElementById('aiClassLabel').innerHTML = `📋 Gợi ý: <strong>${result.ten}</strong>`;

    // Set amount
    if (result.muc > 0) {
      document.getElementById('inputMucDeNghi').value = new Intl.NumberFormat('vi-VN').format(result.muc);
      document.getElementById('aiMucLabel').innerHTML = `💰 Mức tham khảo: <strong>${formatCurrency(result.muc)}</strong>`;
    }

    // Show suggestion box
    const suggestionDiv = document.getElementById('aiSuggestion');
    suggestionDiv.style.display = 'block';

    const hoSoList = result.hoSo.split(';').map(s => s.trim());
    const daCo = (document.getElementById('inputHoSoDaCo').value || '').toLowerCase();

    document.getElementById('aiSuggestionContent').innerHTML = `
      <div style="margin-bottom:8px"><strong>Phân loại:</strong> ${result.icon} ${result.ten}</div>
      <div style="margin-bottom:8px"><strong>Mức tham khảo:</strong> ${result.muc > 0 ? formatCurrency(result.muc) : 'Theo BCH xem xét'}</div>
      <div style="margin-bottom:8px"><strong>Thời hạn xử lý:</strong> ${result.thoiHan}</div>
      <div style="margin-bottom:4px"><strong>Hồ sơ cần có:</strong></div>
      <ul>
        ${hoSoList.map(h => {
          const hasThat = daCo && daCo.includes(h.toLowerCase().substring(0, 4));
          return `<li class="${hasThat ? 'ok' : 'missing'}">${hasThat ? '✅' : '⚠️'} ${h}</li>`;
        }).join('')}
      </ul>
      <div style="margin-top:8px;font-size:11px;color:var(--text-tertiary)">💡 ${result.ghiChu}</div>
    `;
  } else {
    document.getElementById('aiSuggestion').style.display = 'none';
    document.getElementById('aiClassLabel').textContent = '';
    document.getElementById('aiMucLabel').textContent = '';
  }
}

function resetForm() {
  document.getElementById('formTiepNhan').reset();
  document.getElementById('aiSuggestion').style.display = 'none';
  document.getElementById('aiClassLabel').textContent = '';
  document.getElementById('aiMucLabel').textContent = '';
  generateMaHoSo();
  // Set today's date
  document.getElementById('inputNgay').valueAsDate = new Date();
}

function handleFormSubmit(e) {
  e.preventDefault();

  const ma = document.getElementById('inputMaHoSo').value;
  const ngayStr = document.getElementById('inputNgay').value;
  const nguoiDeNghi = document.getElementById('inputNguoiDeNghi').value;
  const diaBan = document.getElementById('inputDiaBan').value;
  const doanVien = document.getElementById('inputDoanVien').value;
  const nguoiHoTro = document.getElementById('inputNguoiHoTro').value;
  const quanHe = document.getElementById('inputQuanHe').value;
  const noiDung = document.getElementById('inputNoiDung').value;
  const loai = document.getElementById('inputLoaiChamLo').value;
  const mucStr = document.getElementById('inputMucDeNghi').value;
  const uuTien = document.getElementById('inputUuTien').value;

  if (!nguoiDeNghi || !diaBan || !noiDung) {
    showToast('Vui lòng điền đầy đủ thông tin bắt buộc', 'error');
    return;
  }

  const soTien = parseInt(mucStr.replace(/\D/g, '')) || 0;
  const now = new Date();
  const dateParts = ngayStr.split('-');
  const thang = parseInt(dateParts[1]);
  const nam = parseInt(dateParts[0]);

  const newCase = {
    ma,
    ngay: null,
    ngayStr: ngayStr,
    nam,
    thang,
    noiDung,
    loai: loai || 'Khác',
    soTien,
    nguoiHoTro: nguoiHoTro || null,
    doanVien: doanVien || null,
    quanHe: quanHe || null,
    lyDo: noiDung,
    nguoiThucHien: nguoiDeNghi,
    diaBan: diaBan !== 'Khác' ? diaBan : null,
    donVi: null,
    trangThai: 'Chờ duyệt',
    ngayHT: null,
    soNgayXL: null,
    canhBao: 'Chưa duyệt',
    uuTien,
    nguoiDeNghi,
    createdAt: now.toISOString()
  };

  const cases = getSubmittedCases();
  cases.push(newCase);
  saveSubmittedCases(cases);

  showToast(`Đã gửi đề nghị ${ma} thành công!`, 'success');

  resetForm();
  renderSubmitted();

  // Update badge
  const pending = cases.filter(c => c.trangThai === 'Chờ duyệt').length;
  document.getElementById('badgePending').textContent = pending;
}

function renderSubmitted() {
  const cases = getSubmittedCases();
  const tbody = document.getElementById('submittedTableBody');

  if (cases.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6"><div class="empty-state"><div class="empty-icon">📭</div><h3>Chưa có đề nghị nào</h3><p>Nhập form bên trên để gửi đề nghị mới</p></div></td></tr>';
    return;
  }

  tbody.innerHTML = cases.slice().reverse().map(r => `
    <tr onclick="showDetail('${r.ma}')">
      <td>${r.ma}</td>
      <td>${r.ngayStr || '--'}</td>
      <td>${r.nguoiDeNghi || '--'}</td>
      <td><span class="type-tag ${getTypeClass(r.loai)}">${r.loai}</span></td>
      <td style="font-weight:600;color:var(--text-primary)">${formatCurrency(r.soTien)}</td>
      <td>${getStatusBadge(r.trangThai)}</td>
    </tr>
  `).join('');
}

// ── EVENT LISTENERS ──
function setupEventListeners() {
  // Sidebar nav
  document.querySelectorAll('.nav-item[data-page]').forEach(item => {
    item.addEventListener('click', () => navigateTo(item.dataset.page));
  });

  // Bottom nav
  document.querySelectorAll('.bottom-nav-item[data-page]').forEach(item => {
    item.addEventListener('click', () => navigateTo(item.dataset.page));
  });

  // Hamburger
  document.getElementById('btnMenuToggle').addEventListener('click', openSidebar);
  document.getElementById('sidebarOverlay').addEventListener('click', closeSidebar);

  // Modal close on overlay click
  document.getElementById('detailModal').addEventListener('click', (e) => {
    if (e.target === document.getElementById('detailModal')) closeModal();
  });

  // Escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
  });

  // Form
  document.getElementById('formTiepNhan').addEventListener('submit', handleFormSubmit);

  // Auto-classify on typing (debounced)
  let debounceTimer;
  document.getElementById('inputNoiDung').addEventListener('input', () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(handleNoiDungChange, 300);
  });

  // Also re-trigger when "Hồ sơ đã có" changes
  document.getElementById('inputHoSoDaCo').addEventListener('input', () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(handleNoiDungChange, 300);
  });

  // Ho so filters
  ['searchHoSo', 'filterLoai', 'filterDiaBan', 'filterThang', 'filterTrangThai'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', () => { hosoPage = 1; renderHoSoTable(); });
      el.addEventListener('change', () => { hosoPage = 1; renderHoSoTable(); });
    }
  });
}

// ── INIT ──
function init() {
  // Set current date
  const now = new Date();
  document.getElementById('currentDate').textContent = now.toLocaleDateString('vi-VN', {
    weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric'
  });

  // Set today's date in form
  document.getElementById('inputNgay').valueAsDate = new Date();

  setupEventListeners();

  // Generate first case ID
  generateMaHoSo();

  // Render initial page
  renderDashboard();
  renderQuyche();
  renderSubmitted();

  // Update badge
  const pending = getSubmittedCases().filter(c => c.trangThai === 'Chờ duyệt').length;
  document.getElementById('badgePending').textContent = pending;
}

// Start
document.addEventListener('DOMContentLoaded', init);
