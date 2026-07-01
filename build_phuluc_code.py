# -*- coding: utf-8 -*-
"""Phụ lục: Quy trình NHẬN – XỬ LÝ – GIAO Mã code zVAS (đính kèm HĐ Đại lý zVAS).
Tổng hợp từ slide quy trình nhận code + chính sách đại lý VNG + form HĐ VNG-GAPIT."""
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
OUT='/home/user/finalproject/output/PhuLuc_QuyTrinh_NhanXuLyGiao_MaCode_zVAS.docx'
d=docx.Document(); st=d.styles['Normal']; st.font.name='Times New Roman'; st.font.size=Pt(12)
NAVY=RGBColor(0x1B,0x2F,0x5A); ZALO=RGBColor(0x00,0x68,0xFF); GRAY=RGBColor(0x55,0x55,0x55)
def H(t,sz=13,sb=10,center=False,color=NAVY):
    p=d.add_paragraph();r=p.add_run(t);r.bold=True;r.font.size=Pt(sz);r.font.color.rgb=color
    p.paragraph_format.space_before=Pt(sb);p.paragraph_format.space_after=Pt(4)
    if center:p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    return p
def P(t,just=True,it=False,bold=False,ind=None):
    p=d.add_paragraph();r=p.add_run(t);r.italic=it;r.bold=bold
    if just:p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    if ind:p.paragraph_format.left_indent=Cm(ind)
    p.paragraph_format.space_after=Pt(3);return p
def shade(cell,hexc):
    cell._tc.get_or_add_tcPr().append(docx.oxml.parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="%s"/>'%hexc))

H('PHỤ LỤC 03',sz=15,center=True,sb=0)
H('QUY TRÌNH NHẬN – XỬ LÝ – GIAO MÃ CODE zVAS CHO KHÁCH HÀNG',sz=12.5,center=True,sb=2,color=RGBColor(0,0,0))
P('(Đính kèm và là một phần không tách rời của Hợp đồng Đại lý phân phối Dịch vụ zVAS số ……/2026/HĐĐL-zVAS/GPT '
  'giữa Công ty Cổ phần GAPIT — “Bên A” và Đại lý — “Bên B”.)',just=False,it=True)

H('1. Mục đích và phạm vi')
P('Phụ lục này quy định trình tự Bên B (Đại lý) nhận Mã code zVAS từ Bên A (GAPIT), xử lý/quản lý Mã code và giao '
  'Mã code tới khách hàng/Người Dùng Cuối, bảo đảm thống nhất trong vận hành và tuân thủ chính sách của nhà phát hành (VNG/Zalo).')

H('2. Sơ đồ tổng quát luồng Mã code')
tbl=d.add_table(rows=1,cols=7); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
flow=[('VNG / Zalo\n(Nhà phát hành)','DCE6F7'),('→','FFFFFF'),('BÊN A – GAPIT\n(Phân phối · CMS/zbox.vn)','DCE6F7'),
      ('→','FFFFFF'),('BÊN B – Đại lý\n(Nhận & quản lý code)','DCE6F7'),('→','FFFFFF'),('Khách hàng / NDC\n(Kích hoạt tại zbox.vn)','DCE6F7')]
for i,(t,c) in enumerate(flow):
    cell=tbl.rows[0].cells[i]; cell.text=''
    pr=cell.paragraphs[0]; pr.alignment=WD_ALIGN_PARAGRAPH.CENTER
    run=pr.add_run(t); run.bold=(t!='→'); run.font.size=Pt(9 if t!='→' else 14)
    if c!='FFFFFF': shade(cell,c)
P('')

H('3. Bước 1 — NHẬN Mã code (Bên B nhận từ Bên A)')
for t in [
 '3.1. Đặt hàng: Bên B đặt hàng trên hệ thống CMS của Bên A và/hoặc qua email đặt hàng được chỉ định; hệ thống chốt đơn theo từng giao dịch — KHÔNG tách lẻ Mã code trong một đơn.',
 '3.2. Thanh toán (trả trước): Bên B thanh toán 100% giá trị đơn theo Giá cấp cho Đại lý (đã bao gồm chiết khấu) tại thời điểm chốt đơn. Hệ thống ghi nhận trạng thái “Thanh toán thành công” kèm mã giao dịch.',
 '3.3. Bàn giao Mã code: Sau khi tiền được ghi có vào tài khoản của Bên A, Bên A bàn giao Mã code zVAS theo từng đơn (không tách lẻ) trong vòng 03 (ba) ngày làm việc, qua: (a) hệ thống CMS/zbox.vn (mục “Danh sách code đã mua”); (b) email; hoặc (c) phương thức khác do Hai Bên thỏa thuận bằng văn bản.',
 '3.4. Kiểm tra khi nhận: Bên B kiểm tra số lượng, loại dịch vụ, tình trạng Mã code ngay khi nhận. Mã code lỗi / không sử dụng được / đã bị sử dụng trước khi bàn giao → xử lý theo Mục 7 (Bên A thu hồi và cấp lại trong tối đa 24 Giờ làm việc).',
]: P(t,ind=0.5)

H('4. Bước 2 — XỬ LÝ & QUẢN LÝ Mã code (trên zbox.vn)')
for t in [
 '4.1. Truy cập: Bên B truy cập tài khoản tại zbox.vn → chọn Menu “Quản lý Tài Khoản” → mục “Quản Lý Mã Code”.',
 '4.2. Chức năng: xem “Danh sách code đã mua” gồm Mã code, Dịch vụ, Ngày tạo, Ngày hết hạn, Trạng thái (Khả dụng / Đã kích hoạt / Hết hạn); sao chép Mã code; và tải file Excel để đối chiếu, quản lý tồn kho.',
 '4.3. Theo dõi tồn kho & hạn kích hoạt: Bên B chủ động theo dõi số lượng Mã code khả dụng và thời hạn hiệu lực kích hoạt của từng Mã code; ưu tiên phân phối các Mã code sắp hết hạn.',
 '4.4. Nguyên tắc dữ liệu: trường hợp Bên B sử dụng hệ thống của Bên A để gắn/ghi nhận Mã code cho SĐT/tài khoản người dùng, áp dụng nguyên tắc khử định danh và tối thiểu hóa dữ liệu theo Thỏa thuận Xử lý Dữ liệu Cá nhân (DPA) đính kèm Hợp đồng.',
]: P(t,ind=0.5)

H('5. Bước 3 — GIAO Mã code cho khách hàng (02 mô hình)')
t2=d.add_table(rows=4,cols=3); t2.style='Table Grid'; t2.alignment=WD_TABLE_ALIGNMENT.CENTER
hdr=['Mô hình','Cách giao Mã code','Áp dụng / Lưu ý']
for j,h in enumerate(hdr):
    c=t2.rows[0].cells[j]; c.text=''; r=c.paragraphs[0].add_run(h); r.bold=True; r.font.color.rgb=RGBColor(0xFF,0xFF,0xFF); r.font.size=Pt(10); shade(c,'1B2F5A')
rows=[
 ['Quy trình TIÊU CHUẨN','Bên B (Đại lý) nhận Mã code từ hệ thống, sau đó CẤP LẠI (trao) Mã code cho khách hàng/Người Dùng Cuối qua kênh của Bên B.','Phù hợp khi Đại lý tự quản lý và trực tiếp chăm sóc khách hàng của mình.'],
 ['Quy trình NÂNG CAO','Mã code được GỬI TRỰC TIẾP vào Zalo của khách hàng.','Áp dụng theo TỪNG đơn hàng, không tách lẻ.'],
 ['Nghĩa vụ chung','Bên B giao đúng, đủ Mã code cho khách hàng của đơn hàng; bảo mật Mã code, không để lộ/lọt.','Không niêm yết/bán công khai dưới Giá sàn công khai của Zalo (theo Điều 4 Hợp đồng).'],
]
for i,row in enumerate(rows):
    for j,v in enumerate(row):
        c=t2.rows[i+1].cells[j]; c.text=''; rr=c.paragraphs[0].add_run(v); rr.font.size=Pt(9.5)
        if j==0: rr.bold=True; rr.font.color.rgb=NAVY

H('6. Bước 4 — KÍCH HOẠT (Khách hàng / Người Dùng Cuối)')
for t in [
 '6.1. Người Dùng Cuối kích hoạt Mã code tại zbox.vn/activate-code: đăng nhập ĐÚNG tài khoản Zalo/ZingMP3 → nhập Mã code → kích hoạt. Sau khi kích hoạt thành công, tài khoản được sử dụng các tính năng nâng cao tương ứng theo Thời hạn Gói.',
 '6.2. Hỗ trợ kích hoạt: Bên B là đầu mối chăm sóc khách hàng cấp 1 (hướng dẫn kích hoạt, tiếp nhận vướng mắc); leo thang Bên A/VNG khi vượt phạm vi xử lý.',
]: P(t,ind=0.5)

H('7. Xử lý Mã code lỗi và hỗ trợ (dẫn chiếu Điều 6, Điều 9 Hợp đồng)')
for t in [
 'Mã code lỗi / không kích hoạt được không do lỗi của Bên B hoặc Người Dùng Cuối: Bên A thu hồi và cấp lại Mã code khác trong tối đa 24 (hai mươi bốn) Giờ làm việc kể từ khi nhận yêu cầu hợp lệ.',
 'Sử dụng sai hướng dẫn dẫn đến tài khoản bị khóa/giới hạn: Bên A không có nghĩa vụ hủy, thu hồi hoặc hoàn trả Mã code đó.',
 'Mã code đã thanh toán không được hoàn tiền dưới mọi hình thức (kể cả chưa kích hoạt), trừ Mã code lỗi thuộc trách nhiệm Bên A — theo chính sách của nhà phát hành (VNG).',
 'SLA hỗ trợ của Bên A: phản hồi ban đầu ≤ 24 Giờ làm việc; xử lý ≤ 02 ngày làm việc kể từ khi nhận đủ thông tin hợp lệ, trừ trường hợp phức tạp hoặc bất khả kháng.',
]: P('• '+t,ind=0.5)

H('8. Bảo mật và bảo vệ dữ liệu cá nhân')
P('Bên B bảo mật Mã code và chỉ giao cho đúng khách hàng của đơn hàng. Dữ liệu cá nhân của khách hàng/Người Dùng '
  'Cuối được xử lý theo Thỏa thuận Xử lý Dữ liệu Cá nhân (DPA) đính kèm Hợp đồng; áp dụng nguyên tắc khử định danh/'
  'tối thiểu hóa. Riêng dữ liệu định danh phục vụ eKYC là dữ liệu cá nhân NHẠY CẢM — cần sự đồng ý riêng của chủ thể '
  'dữ liệu và áp dụng biện pháp bảo mật nghiêm ngặt theo Luật Bảo vệ Dữ liệu Cá nhân số 91/2025/QH15 và Nghị định số '
  '356/2025/NĐ-CP.')

# Bảng tóm tắt 5 bước
H('9. Bảng tóm tắt quy trình')
t3=d.add_table(rows=1,cols=4); t3.style='Table Grid'; t3.alignment=WD_TABLE_ALIGNMENT.CENTER
for j,h in enumerate(['Bước','Chủ thể','Hành động chính','Hệ thống / Kênh']):
    c=t3.rows[0].cells[j]; c.text=''; r=c.paragraphs[0].add_run(h); r.bold=True; r.font.color.rgb=RGBColor(0xFF,0xFF,0xFF); r.font.size=Pt(10); shade(c,'0068FF')
steps=[
 ['1. Đặt & Thanh toán','Bên B','Đặt đơn (không tách lẻ) · trả trước 100%','CMS / email'],
 ['2. Nhận code','Bên A → Bên B','Bàn giao Mã code ≤ 03 ngày LV sau khi tiền ghi có','CMS / zbox.vn / email'],
 ['3. Quản lý code','Bên B','Xem/tra cứu/tải Excel · theo dõi tồn kho & hạn','zbox.vn → Quản Lý Mã Code'],
 ['4. Giao code','Bên B','Cấp lại cho KH (tiêu chuẩn) hoặc gửi thẳng Zalo KH (nâng cao)','Kênh Đại lý / Zalo KH'],
 ['5. Kích hoạt','KH / NDC','Đăng nhập Zalo/ZingMP3 → nhập code → kích hoạt','zbox.vn/activate-code'],
]
for row in steps:
    cells=t3.add_row().cells
    for j,v in enumerate(row):
        cells[j].text=''; rr=cells[j].paragraphs[0].add_run(v); rr.font.size=Pt(9.3)
        if j==0: rr.bold=True; rr.font.color.rgb=NAVY

d.add_paragraph()
sg=d.add_table(rows=1,cols=2); sg.alignment=WD_TABLE_ALIGNMENT.CENTER
a=sg.rows[0].cells[0].paragraphs[0]; a.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=a.add_run('ĐẠI DIỆN BÊN A\n(CÔNG TY CỔ PHẦN GAPIT)\n(Ký, ghi rõ họ tên, đóng dấu)'); r.bold=True
b=sg.rows[0].cells[1].paragraphs[0]; b.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=b.add_run('ĐẠI DIỆN BÊN B\n(ĐẠI LÝ)\n(Ký, ghi rõ họ tên, đóng dấu)'); r.bold=True

d.save(OUT); print('SAVED',OUT)
