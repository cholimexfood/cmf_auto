**THÔNG TIN BOT Ver2.6**

_Update Ver2.6:_
- Nếu chạy Bot sau 9h:30 AM hằng ngày thì sẽ vào thẳng chế độ Manual.

_Update Ver2.5.1:_
- Fix lỗi thời gian chờ lỗi (số âm) khi chạy Script Update.

_Update Ver2.5: Ver tối ưu tiến trình - tiết kiệm thời gian (Dự kiến trước 10h hoàn thành)_
- Tool tự chạy vào 7h50 AM
- Giảm số lần check đơn thành 3 lần (mỗi lần cách nhau 20 phút).
- Nếu đơn hàng xử lý xong: Down - Tạo Form - Tạo CSV - Import sẵn dữ liệu - chờ 9h chạy Script
- Nếu đơn hàng chưa xử lý xong: 9h30 kiểm tra lần cuối:
    - Nếu xong: Auto toàn bộ tiến trình.
    - Nếu chưa xong: Kết thúc Auto

_Update Ver2.4.2: Tối ưu tiến trình (giảm thời gian)_
- Tool Download và Create Form Excel chạy ngay sau khi đơn hàng được duyệt xong

_Update Ver2.4.1:_
- Fix lỗi Crash Tools do Chrome cập nhật.
- Fix lỗi tối ưu Log DB không chạy.
- Bổ sung lệnh manual tối ưu DB.

_Update Ver2.4:_
- Tối ưu Log DB: giảm dung lượng Log mỗi ngày trước khi Import Data
- Fixed: lỗi check đơn lần cuối không hoạt động
- Fixed: lỗi không xoá được File kết quả check đơn hàng -> dẫn đến kết quả thông báo bị sai
- Tối ưu hoá đăng nhập vào thoát web

_Update Ver2.3.1:_
- Cho phép ấn N trong 3s để bỏ qua chế độ Auto khi khởi động
- Fix lỗi màu sắc khi Print thông báo trong cmd

_Update Ver2.3:_
- Không cho chạy lệnh Manual khi đang "Auto"
- Fix thời gian chờ qua 10h => Kết thúc "Auto"

_Update Ver2.2:_
- Lỗi luôn trả về giá trị đã duyệt đơn hàng xong (Fixed)

_**Bot có 2 chức năng Auto và Manual:**_

_Auto:_

- Bot sẽ tự khởi động vào khoảng 7h50 AM mỗi ngày trừ ngày Chủ Nhật.
- Bot sẽ kiểm tra Trạng thái đơn hàng tối đa 4 lần (mỗi lần cách nhau 20 phút).
- Nếu đơn hàng được duyệt trước 9h30 AM thì đúng 9h30 AM Bot sẽ tự động Tải báo cáo và import vào SQL, chương trình sẽ kết thúc chế độ Auto khi hoàn thành.
- Nếu đến 9h30 AM đơn hàng vẫn chưa được duyệt, Bot sẽ chờ đến 10h00 AM kiểm tra trạng thái đơn hàng lần cuối:
    + Nếu đã duyệt hết -> BOT sẽ tải báo cáo và import SQL, kết thúc khi hoàn thành.
    + Nếu chưa duyệt hết -> BOT sẽ kết thúc chế độ Auto -> Người dùng sẽ theo dõi và dùng chế độ Manual qua lệnh trên Telegram.

_Manual:_

- start - Lời chào | Lệnh này là lời chào của Bot.
- manualsql - Chạy Manual All | Lệnh này Bot sẽ tiền hành Tải và import SQL không quan tâm đơn hàng duyệt hết hay chưa (người dùng hãy check thủ công đơn hàng và kích hoạt nếu cần)
- ktttdh - Bước 1: Kiểm tra trạng thái đơn hàng | Lệnh này kiểm tra đơn hàng đã duyệt hết chưa
- downrp - Bước 2: Tải Tuyến - Khách hàng - Đơn hàng | Lệnh này chỉ tải báo cáo Tuyến - khách hàng - đơn hàng
- copytoform - Bước 3: Tạo Form Import | Lệnh này tạo form excel để chuẩn bị dữ liệu cho import
- importtosql - Bước 4: Import SQL - Run Script | Lệnh này sẽ chuyển đổi dữ liệu và Import + Chạy sript để xử lý dữ liệu
- runscript - Run Script | Lệnh này chỉ chạy Script

_**Quá trình hoạt động của Bot đều sẽ thông báo qua nhóm telegram:**_
- Nhóm user sẽ nhận được thông tin kết quả cần.
- Nhóm Dev sẽ nhận được thông tin quá trình hoạt động của Bot - Lỗi nếu có.
