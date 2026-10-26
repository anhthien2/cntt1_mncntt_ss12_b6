### Phần 1. Phân tích: Bảng Excel "Xấu" vs Bảng "Đúng"

**Bảng XẤU (sẽ bị trừ điểm) có đặc điểm:**

1.  Gộp ô - Merge Cells ở hàng tiêu đề
2.  2-3 hàng Header - Ví dụ hàng 1 là tiêu đề chung, hàng 2 mới là tên cột
3.  Có hàng Tổng cộng / Ghi chú nằm giữa dữ liệu
4.  Để hàng trống, cột trống giữa bảng để cho "thoáng"
5.  Mỗi loại dữ liệu để 1 sheet khác nhau rồi cộng tay

Hậu quả: Không dùng được `Ctrl + Shift + L` để lọc, không dùng được VLOOKUP, PivotTable sẽ lỗi.

**Bảng ĐÚNG (đạt chuẩn) phải:**

1.  **Hàng 1 duy nhất là Header Row:** Mỗi cột 1 tên, không gộp ô. VD: `A1: Mã Nhà | B1: Diện tích | C1: Giá bán | D1: Mức hoa hồng`
2.  Dữ liệu liên tục từ hàng 2 trở xuống, không có hàng trống
3.  Không có hàng tổng ở giữa, tổng để cuối cùng hoặc sheet khác
4.  Mỗi cột chỉ chứa 1 loại dữ liệu (cột Giá chỉ chứa số)

### Phần 2. Triển khai - Bộ học liệu Smart-Home

**A. Word - File gửi đối tác:**  
Dùng cấu trúc Heading 2 như hôm trước để bật Navigation Pane:

1.  Tổng quan dự án Smart-Home
2.  Phân tích thị trường
3.  Kiến trúc hệ thống IoT
4.  Yêu cầu phần cứng
5.  Yêu cầu phần mềm
6.  Lộ trình triển khai
7.  Phân tích chi phí - lợi ích
8.  Rủi ro và giải pháp
9.  Đội ngũ nhân sự
10.  Liên hệ hợp tác
    
# Phần3


### Lỗi "trôi vùng dữ liệu" là gì?

Bạn viết công thức ở ô E2:  
`=VLOOKUP(D2, BangTraCuu!A2:B4, 2, FALSE)`

Khi bạn kéo xuống E3, Excel tự hiểu bạn muốn dò xuống 1 hàng, nên nó tự đổi thành:  
`E3 = VLOOKUP(D3, BangTraCuu!A3:B5, 2, FALSE)`

Nhưng bảng tra cứu của bạn chỉ có từ A2 đến B4, A3:B5 đã bị lệch ra ngoài -> nó không tìm thấy "Trung bình" nữa -> báo lỗi `#N/A`. Đó gọi là **trôi vùng**.

### Cách sửa bằng dấu  $ - Khóa vùng

Dấu `$` có tác dụng **KHÓA** không cho Excel tự đổi số khi kéo.

Bạn sửa lại thành:  
`=VLOOKUP(D2, BangTraCuu!$A$2:$B$4, 2, FALSE)`

Ý nghĩa từng loại khóa:

-   `$A$2:$B$4` - **Khóa tuyệt đối:** Khóa cả cột và hàng. Kéo đi đâu nó vẫn đứng yên ở A2:B4. **Đây là cái bạn phải dùng cho bảng tra cứu.**
-   `$A2` - Khóa cột A, nhưng hàng 2 vẫn được trôi khi kéo xuống
-   `A$2` - Khóa hàng 2, nhưng cột A vẫn trôi khi kéo sang phải

**Mẹo nhớ nhanh:**

1.  Đặt con trỏ vào giữa `A2:B4` trong công thức
2.  Bấm phím **F4** trên bàn phím 1 lần -> nó sẽ tự thêm `$` thành `$A$2:$B$4`
3.  Enter rồi kéo xuống -> hết lỗi trôi 100%.


