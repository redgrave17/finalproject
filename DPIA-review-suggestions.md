# Rà soát & đề xuất bổ sung — Hồ sơ Đánh giá tác động xử lý dữ liệu cá nhân (Mẫu Đ24-DLCN-01)

**Tài liệu được rà soát:** `Mau ho so danh gia tac dong xu ly du lieu ca nhan 5 (2) (1).docx` (Google Drive, cập nhật 10/06/2026)
**Bối cảnh:** Hồ sơ DPIA của Công ty Cổ phần GAPIT, được chuẩn bị trong khuôn khổ dự án **eSIM Du Lịch trên Zalo Mini App** (theo thư mục tài liệu dự án trên Drive).
**Nguồn tham chiếu:** Tài liệu dự án eSIM trên Drive (SRS Mini-app, Quy trình CSKH eSIM v1.0, Chính sách BVDLCN eSIM, Timeline & MoM, checklist training CX); repo `Paparusi/legal-ai-agent` và dataset `duyet/vietnamese-legal-documents-dataset` (công cụ tra cứu/kiểm chứng văn bản pháp luật Việt Nam).

---

## A. Vấn đề lớn nhất: nội dung mô tả sai dịch vụ

Toàn bộ phần thuyết minh (Mục V.1, V.2, VI.1, VI.2) đang mô tả **dịch vụ tin nhắn đa kênh CPaaS** ("GAPIT cung cấp dịch vụ tin nhắn đa kênh (CPaaS)... gửi tin marketing automation..."), trong khi hồ sơ này phục vụ dịch vụ **eSIM du lịch bán qua Zalo Mini App**. Hai phương án:

1. Nếu hồ sơ chỉ dành cho dịch vụ eSIM → viết lại VI.1 và các mục mô tả theo luồng thực tế của dịch vụ (đề xuất nội dung bên dưới).
2. Nếu hồ sơ bao trùm nhiều dịch vụ của GAPIT → bổ sung mục liệt kê từng dịch vụ/hệ thống xử lý DLCN, trong đó eSIM Du Lịch là một hạng mục riêng với mô tả riêng.

**Đề xuất nội dung thay thế cho VI.1 (bối cảnh kinh doanh):**

> GAPIT cung cấp dịch vụ eSIM du lịch quốc tế ("eSIM Du Lịch") cho người dùng Việt Nam thông qua Zalo Mini App (link: zalo.me/s/4552695804193899018). Khách hàng chọn gói eSIM theo quốc gia/khu vực, thanh toán trực tuyến (qua VietinBank/ZaloPay), cung cấp địa chỉ email (bắt buộc) để nhận: (1) email xác nhận đơn hàng, (2) email chứa mã QR kích hoạt eSIM, (3) hóa đơn điện tử (phát hành qua VNPT, nếu khách yêu cầu). Dữ liệu đơn hàng được lưu trữ, quản lý tập trung trên hệ thống CMS do GAPIT xây dựng, đặt tại Việt Nam, phục vụ tra cứu, đối soát và chăm sóc khách hàng (qua Zalo OA "eSIM Du Lịch" và email support.esim@gapit.com.vn).

---

## B. Bổ sung theo từng mục của biểu mẫu

### Mục I — Bên Kiểm soát dữ liệu
- Điền các trường còn trống: **Mã số thuế (1c), điện thoại (7), email (9), website (10), mạng xã hội (11)** — mục 11 nên ghi Zalo OA "eSIM Du Lịch" (ID OA: 1537525979457573204) và các kênh chính thức khác; **ngày thành lập (12)** kèm bản sao GCN ĐKKD.
- Mục 13 (ngành nghề) của phần I đang để trống trong khi bảng mã ngành lại nằm ở phần II — chuyển bảng mã ngành về đúng phần I (GAPIT là Bên Kiểm soát).
- Mục 15: bổ sung ngày sinh, giới tính, chức vụ, trình độ, SĐT, email của người đứng đầu bộ phận BVDLCN (hiện chỉ có tên Nguyễn Văn Long).
- Mục 16: bảng nhân sự BVDLCN đang trống — liệt kê ít nhất các đầu mối thực tế của dự án (ví dụ PIC vận hành CMS, đầu mối CX/Helpdesk).

### Mục II — Bên Xử lý dữ liệu (theo hợp đồng)
Hiện đang điền lại chính GAPIT — chưa đúng bản chất. Theo tài liệu dự án, các bên xử lý/tham gia xử lý theo hợp đồng cần kê khai (kèm bản sao hợp đồng theo yêu cầu của mẫu):

| Đối tác | Vai trò | Dữ liệu liên quan |
|---|---|---|
| **Consortio** (NCC eSIM, nước ngoài) | Cung cấp/kích hoạt gói eSIM qua API CMS | Thông tin đơn hàng cần thiết để phát hành QR eSIM |
| **VNPT** | Phát hành hóa đơn điện tử | Họ tên/tên đơn vị, MST, địa chỉ, email nhận hóa đơn |
| **VietinBank, ZaloPay** | Kênh thanh toán/thu hộ | Thông tin giao dịch thanh toán (lưu ý: có thể là bên kiểm soát độc lập theo pháp luật ngân hàng/TGTT — cần ghi rõ vai trò trong hợp đồng) |
| **Zalo (VNG)** | Nền tảng Mini App, Zalo OA | Định danh tài khoản Zalo, dữ liệu tương tác trên Mini App/OA |
| **HubSpot** (nếu triển khai theo MoM 13/03/2026) | CRM lưu email khiếu nại, trạng thái xử lý | Email, nội dung khiếu nại — *là dịch vụ cloud nước ngoài → kéo theo nghĩa vụ chuyển DLCN ra nước ngoài (xem mục V.12)* |

### Mục III — Bên thứ ba
Đang trống hoàn toàn. Cần rà soát và kê khai: đại lý B2B (theo thư mục "Hợp đồng B2B" — đại lý tiếp nhận danh sách người đặt, nhận code qua email và đẩy mã về Zalo khách hàng → có tiếp xúc DLCN của người dùng cuối), và các bên nhận chia sẻ dữ liệu khác (nếu có).

### Mục IV — Nhà phát triển / tích hợp API
Đang trống, trong khi hệ thống tích hợp nhiều API: **API CMS–Consortio, API thanh toán ZaloPay, API thanh toán VietinBank, API hóa đơn VNPT, Zalo Mini App SDK**. Cần điền:
- Danh sách loại hình nhà phát triển (mục 1).
- Quyền và điều khoản hạn chế khi tích hợp API (mục 2, 3) — trích từ hợp đồng/tài liệu API từng bên.
- DLCN trao đổi qua API (mục 4): email, mã đơn hàng, giá trị giao dịch, thông tin xuất hóa đơn.
- Quy định hành vi sai trái khi sử dụng API (mục 5).

### Mục V — Hoạt động xử lý DLCN
- **V.3.1 / V.3.2:** chưa tích (√) loại dữ liệu nào. Với dịch vụ eSIM, đối chiếu Chính sách BVDLCN eSIM (tài liệu "9. BVDLCN eSIM.docx") thì cần tích tối thiểu: *họ tên; địa chỉ liên hệ (email); số điện thoại; thông tin tài khoản số; dữ liệu phản ánh hoạt động/lịch sử hoạt động trên không gian mạng (hành vi trong app); MST cá nhân (khi xuất hóa đơn)*. Chính sách BVDLCN eSIM còn liệt kê cả *ngày sinh, tài khoản ngân hàng, số hộ chiếu/CMND/CCCD, IP, IMEI, vị trí* — nếu thực tế Mini App **không** thu các trường này thì nên sửa Chính sách cho khớp; nếu có thu thì phải tích tương ứng (số CMND, hộ chiếu...) và cân nhắc mục dữ liệu nhạy cảm (thông tin giao dịch qua tổ chức trung gian thanh toán đã được mẫu xếp vào nhóm nhạy cảm).
- **V.4 + V.5 (hình thức lấy sự đồng ý):** V.5 đang trống. Mô tả cơ chế thực tế: checkbox tại bước hoàn tất đơn hàng trên Mini App — *"Thông qua việc hoàn tất đơn hàng, bạn đồng ý với Điều khoản Dịch vụ và Chính sách Quyền riêng tư"* — kèm link 2 văn bản; đính kèm Điều khoản dịch vụ eSIM và Chính sách BVDLCN eSIM làm phụ lục. **Lưu ý tuân thủ:** sự đồng ý cho mục đích *marketing* phải tách riêng (opt-in riêng), không được gộp vào việc hoàn tất đơn hàng — Điều 11 NĐ 13/2023/NĐ-CP yêu cầu đồng ý theo từng mục đích; nội dung 2.3 (marketing automation) hiện mâu thuẫn với cơ chế đồng ý đang mô tả.
- **V.6, V.7:** trống — ước lượng dung lượng dữ liệu (GB) và số lượng chủ thể dữ liệu dự kiến từ số liệu đơn hàng CMS và dự báo doanh số trong Timeline/P&L của dự án.
- **V.8 (bảng thời gian xử lý):** đang dùng nội dung mẫu từ tài liệu khác ("Thu thập dữ liệu cá nhân từ Bên A" — không có "Bên A" nào trong hồ sơ). Viết lại theo luồng eSIM: thu thập tại checkout → lưu CMS → gửi email QR (~10 phút) → phát hành hóa đơn → lưu trữ phục vụ CSKH/đối soát → xóa/hủy theo chính sách.
- **V.9 (thời gian lưu trữ) — xung đột pháp lý cần xử lý:** quy định "xóa sau 3 năm từ tương tác cuối" mâu thuẫn với nghĩa vụ lưu trữ chứng từ kế toán/hóa đơn **10 năm** (Luật Kế toán 2015, NĐ 123/2020/NĐ-CP). Đề xuất lập bảng lưu trữ theo loại dữ liệu: dữ liệu hóa đơn/giao dịch 10 năm; dữ liệu tài khoản/CSKH 3 năm; dữ liệu marketing đến khi rút đồng ý.
- **V.12 + V.13 (chuyển DLCN ra nước ngoài) — điểm trọng yếu:** đang bỏ trống Có/Không. Vì NCC eSIM **Consortio là tổ chức nước ngoài** nhận dữ liệu đơn hàng qua API (và HubSpot nếu dùng), gần như chắc chắn phải tích **Có**, điền thông tin Bên chuyển/Bên nhận tại mục 13, và **lập thêm Hồ sơ đánh giá tác động chuyển dữ liệu cá nhân ra nước ngoài (Mẫu Đ24-DLCN-02)** theo Điều 25 NĐ 13/2023/NĐ-CP — hồ sơ này chưa thấy trong thư mục dự án.
- **V.14 (biện pháp bảo vệ):**
  - 14.1: chỉ dẫn chiếu "Chính sách xử lý DLCN thu thập từ bên ngoài ngày 21/08/2023". Bổ sung: **Chính sách bảo vệ dữ liệu cá nhân (áp dụng cho dịch vụ eSIM)** — nêu số hiệu/ngày ban hành; **Quy trình CSKH dịch vụ eSIM v1.0**; quy chế phân quyền CMS (nếu có văn bản).
  - 14.2: hiện chỉ có mã hóa khi truyền, firewall, antivirus. Bổ sung các biện pháp thực tế của hệ thống CMS/Mini App: phân quyền truy cập theo vai trò (RBAC) trên CMS, xác thực 2 lớp cho tài khoản quản trị, nhật ký truy cập (audit log), sao lưu định kỳ, mã hóa dữ liệu khi lưu trữ (nếu áp dụng). (Timeline dự án ghi nhận đã có bước "Đánh giá về bảo mật, chống hack" 17/12 — dẫn chiếu kết quả này.)
  - 14.3: trống — nêu tiêu chuẩn áp dụng (ví dụ ISO/IEC 27001) hoặc ghi rõ "chưa áp dụng tiêu chuẩn, kế hoạch áp dụng năm…".
  - 14.6: tích Có/Không về thông báo xử lý dữ liệu nhạy cảm — nếu xác định có xử lý dữ liệu nhạy cảm (giao dịch thanh toán) thì phải "Có" kèm cơ chế thông báo.

### Mục VI — Đánh giá tác động
- **Lỗi cấu trúc 2.5 ↔ 2.6:** mục 2.5 (tác động đối với hệ thống pháp luật) có 5 tiểu mục **trống**, trong khi nội dung đang nằm dưới mục 2.6 lại phân tích... hệ thống pháp luật. Chuyển khối nội dung từ 2.6 lên 2.5, và **viết mới mục 2.6 (tác động đối với lợi ích của chủ thể dữ liệu)**: lợi ích (nhận eSIM tức thời qua email, tra cứu lịch sử mua, hỗ trợ khiếu nại có SLA); rủi ro (lộ email → spam/lừa đảo; lộ mã QR → mất quyền sử dụng gói eSIM đã mua; lộ thông tin hóa đơn); biện pháp giảm thiểu (chỉ gửi QR tới email đã xác nhận thanh toán, giới hạn truy cập CMS, quy trình xử lý khiếu nại theo Quy trình CSKH v1.0).
- **VI.3 (lấy ý kiến):** bổ sung kênh lấy ý kiến từ chính chủ thể dữ liệu/người dùng — dự án đã có vòng "Publish app để toàn bộ Gapiter test và feedback" (10/4) và cơ chế ghi nhận cải tiến app (thư mục 8) — dẫn chiếu làm bằng chứng.
- **VI.4 (giám sát, đánh giá):** thay đoạn mô tả chung chung bằng phân công cụ thể: Bộ phận Bảo vệ DLCN (người đứng đầu: Nguyễn Văn Long) chủ trì; tần suất rà soát theo 14.4 (tháng/quý/năm); cập nhật hồ sơ khi thay đổi NCC, kênh thanh toán, hoặc phạm vi dữ liệu.

### Mục VII — Danh mục văn bản pháp luật
Cập nhật và bổ sung (có thể dùng dataset `duyet/vietnamese-legal-instruct` — nguồn vbpl.vn — và tính năng legal search của `legal-ai-agent` để kiểm tra hiệu lực và số điều chính xác trước khi nộp):
- **Luật Bảo vệ dữ liệu cá nhân 2025** (hiệu lực 01/01/2026) — đã dẫn nhưng cần **kiểm tra lại số điều** (hồ sơ ghi "Điều 8, Điều 14") theo bản chính thức; vì Luật đã có hiệu lực, nên đưa thành căn cứ pháp lý chính và rà soát xem Mẫu Đ24-DLCN-01 (ban hành kèm NĐ 13/2023) có còn là biểu mẫu hiện hành hay đã có nghị định/biểu mẫu mới thay thế trước khi nộp A05 (Bộ Công an).
- Luật An toàn thông tin mạng 2015 (Điều 16–19 về bảo vệ thông tin cá nhân trên mạng).
- Luật Viễn thông 2023 (dịch vụ eSIM/viễn thông).
- Luật Bảo vệ quyền lợi người tiêu dùng 2023 (bán hàng cho người tiêu dùng qua nền tảng số).
- Luật Giao dịch điện tử 2023.
- Nghị định 52/2013/NĐ-CP & 85/2021/NĐ-CP về thương mại điện tử (bán hàng qua Mini App).
- Nghị định 91/2020/NĐ-CP về chống tin nhắn rác, thư điện tử rác (gửi email/tin nhắn marketing).
- Nghị định 123/2020/NĐ-CP về hóa đơn, chứng từ + Luật Kế toán 2015 (thời hạn lưu trữ — liên quan mục V.9).
- Nước ngoài: GDPR chỉ nên giữ nếu thực sự phục vụ khách hàng EU; nếu không, ghi rõ phạm vi khách hàng là người dùng Việt Nam.

### Lỗi hình thức cần sửa
- "Nghị định **53/2023**/NĐ-CP" tại các mục II.3, III.3, V.13(e) → đúng là **53/2022/NĐ-CP** (mục I.3 đang ghi đúng).
- "Webstie" → "Website" (lặp ở 4 chỗ).
- Mục V.13(p) ghi chú sai (đang là ghi chú của mục phân công BVDLCN thay vì giấy tờ thành lập).

---

## C. Lưu ý thủ tục
- Hồ sơ DPIA phải được lập và **luôn có sẵn** để phục vụ kiểm tra của Cục An ninh mạng và phòng, chống tội phạm sử dụng công nghệ cao (A05) — Bộ Công an, và gửi 01 bản chính trong **60 ngày** kể từ ngày bắt đầu xử lý dữ liệu (Điều 24 NĐ 13/2023/NĐ-CP). App đã golive từ tháng 4/2026 → cần đối chiếu mốc thời hạn này.
- Nếu xác nhận có chuyển dữ liệu ra nước ngoài (Consortio/HubSpot): lập thêm **Mẫu Đ24-DLCN-02** trong cùng thời hạn.
- Cập nhật hồ sơ khi có thay đổi: thêm NCC, thêm kênh bán (đại lý B2B, kế hoạch Taobao/Alibaba sẽ phát sinh chuyển dữ liệu xuyên biên giới mới và đối tượng chủ thể dữ liệu là người nước ngoài).

## D. Về hai repo GitHub tham chiếu
- `duyet/vietnamese-legal-documents-dataset`: dataset 467K cặp QA sinh từ 127K văn bản trên vbpl.vn — hữu ích để **kiểm chứng hiệu lực, số điều, cơ quan ban hành** của các văn bản liệt kê tại Mục VII (các loại QA `meta_status`, `legal_basis`, `key_provisions`).
- `Paparusi/legal-ai-agent`: nền tảng AI tra cứu pháp luật + review tài liệu (40K+ văn bản, tính năng compliance check, sửa/soạn điều khoản). Có thể dùng để rà soát chéo bản DPIA và Chính sách BVDLCN eSIM; lưu ý corpus crawl sẵn trong repo hiện thiên về luật lao động, nên với NĐ 13/2023 và Luật BVDLCN 2025 cần dùng tính năng search trực tuyến của nó hoặc đối chiếu trực tiếp trên vbpl.vn/thuvienphapluat.vn trước khi chốt trích dẫn.
