# -*- coding: utf-8 -*-
"""Revise phụ lục quy trình Mã code: bổ sung 2 luồng nhận code (CMS / Email thủ công theo HĐ final).
Highlight toàn bộ thông tin bổ sung."""
import copy, docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
Y=WD_COLOR_INDEX.YELLOW; NAVY=RGBColor(0x1B,0x2F,0x5A)
SRC='/root/.claude/uploads/0ae05b9e-80f9-5907-8558-033a769d7c2e/55f10006-PhuLuc_QuyTrinh_MaCode_zVAS_VNG_GAPIT.docx'
OUT='/home/user/finalproject/output/PhuLuc_QuyTrinh_MaCode_zVAS_REVISED_2Luong.docx'
d=docx.Document(SRC)
def hl_run(r): r.font.highlight_color=Y
def find_p(needle):
    for p in d.paragraphs:
        if needle in p.text: return p
    raise ValueError('NOT FOUND: '+needle)
def new_p_after(cursor, text, bold=False, size=12, label=None):
    el=OxmlElement('w:p'); cursor.addnext(el); p=Paragraph(el,None)
    if label:
        r=p.add_run(label+' '); r.bold=True; r.font.color.rgb=RGBColor(0xC0,0,0); r.font.size=Pt(9)
    r=p.add_run(text); r.bold=bold; r.font.name='Times New Roman'; r.font.size=Pt(size); hl_run(r)
    return el

# neo: sau đoạn Hỗ trợ của Bước 4 (trước mục "3. Xử lý Mã code lỗi")
anchor=find_p('Hỗ trợ: Bên B hỗ trợ, hướng dẫn khách hàng và người dùng cuối')
cur=anchor._p

cur=new_p_after(cur,'Bước 3B — Hai luồng nhận Mã code của khách hàng (qua CMS và qua email)',bold=True,size=13,label='【BỔ SUNG】')
cur=new_p_after(cur,'Tùy năng lực hệ thống và thỏa thuận giữa các Bên, khách hàng của Bên B (đối tác bán lại, khách hàng '
 'doanh nghiệp) nhận Mã code theo một trong hai luồng sau:')

# Bảng so sánh 2 luồng
tb=d.add_table(rows=1,cols=3); tb.style='Table Grid'; tb.alignment=WD_TABLE_ALIGNMENT.CENTER
def shade(cell,hexc): cell._tc.get_or_add_tcPr().append(docx.oxml.parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="%s"/>'%hexc))
for j,h in enumerate(['Tiêu chí','Luồng 1 — Nhận qua hệ thống CMS','Luồng 2 — Nhận thủ công qua email (theo Hợp đồng)']):
    c=tb.rows[0].cells[j]; c.text=''; r=c.paragraphs[0].add_run(h); r.bold=True; r.font.color.rgb=RGBColor(0xFF,0xFF,0xFF); r.font.size=Pt(9.5); shade(c,'1B2F5A')
rows=[
 ('Cách nhận Mã code','Khách hàng đăng nhập hệ thống CMS/zbox.vn của Bên B và nhận Mã code trực tiếp trên tài khoản của mình.',
  'Bên B gửi Mã code tới đúng địa chỉ email mà khách hàng đã đăng ký tại Hợp đồng/Phụ lục 02.'),
 ('Hệ thống / kênh','CMS / zbox.vn (mục “Danh sách code đã mua”, “Quản lý mã code”).','Thư điện tử (email) — địa chỉ chỉ định tại Phụ lục 02.'),
 ('Thời hạn bàn giao','Hiển thị và khả dụng ngay sau khi đơn được xác nhận và ghi nhận thanh toán.','Trong vòng 03 (ba) ngày làm việc kể từ khi Bên B nhận đủ thanh toán.'),
 ('Xác nhận bàn giao','Hệ thống ghi nhận trạng thái bàn giao; khách hàng đối chiếu danh sách Mã code trên CMS.','Xác nhận bàn giao qua email; lập Biên bản bàn giao – nghiệm thu (Phụ lục 03) nếu có yêu cầu.'),
 ('Thời điểm hoàn tất & rủi ro','Hoàn tất khi Mã code khả dụng trên tài khoản CMS của khách hàng.','Hoàn tất khi Bên B gửi Mã code tới email đã đăng ký; kể từ thời điểm đó, rủi ro chuyển sang khách hàng.'),
 ('Quản lý / tra cứu','Tra cứu, sao chép, tải file Excel và theo dõi trạng thái (khả dụng / đã kích hoạt / hết hạn) trực tiếp trên CMS/zbox.vn.','Tra cứu lịch sử giao dịch tại zbox.vn → Quản lý tài khoản; khách hàng lưu trữ email chứa Mã code.'),
 ('Bảo mật','Phân quyền tài khoản CMS; áp dụng khử định danh và tối thiểu hóa dữ liệu khi phù hợp.','Khách hàng chịu trách nhiệm bảo mật địa chỉ email và hệ thống email nhận Mã code (bao gồm an toàn đường truyền).'),
 ('Phù hợp với','Đối tác/khách hàng có năng lực hệ thống, số lượng lớn, cần chủ động quản lý Mã code.','Khách hàng doanh nghiệp nhận theo từng đơn hàng, không sử dụng hệ thống CMS.'),
]
for tt,l1,l2 in rows:
    cs=tb.add_row().cells
    for j,v in enumerate([tt,l1,l2]):
        cs[j].text=''; r=cs[j].paragraphs[0].add_run(v); r.font.size=Pt(9); hl_run(r)
        if j==0: r.bold=True; r.font.color.rgb=NAVY
# di chuyển bảng ra sau con trỏ
cur.addnext(tb._tbl); cur=tb._tbl

cur=new_p_after(cur,'Các bước của Luồng 1 (qua CMS): (1) Khách hàng đặt hàng và thanh toán trả trước theo từng đơn; '
 '(2) Bên B xác nhận đơn, Mã code được ghi có vào tài khoản CMS/zbox.vn của khách hàng; (3) Khách hàng đăng nhập '
 'CMS/zbox.vn, xem “Danh sách code đã mua”, sao chép hoặc tải file Excel; (4) Khách hàng cấp Mã code cho người dùng '
 'cuối và hướng dẫn kích hoạt.',label='【BỔ SUNG】')
cur=new_p_after(cur,'Các bước của Luồng 2 (thủ công qua email): (1) Khách hàng gửi đơn và thanh toán 100% trước; '
 '(2) Trong 03 ngày làm việc, Bên B gửi Mã code tới đúng địa chỉ email đã đăng ký; (3) Việc bàn giao được coi là '
 'hoàn tất kể từ khi email được gửi tới địa chỉ đã đăng ký (rủi ro chuyển sang khách hàng); (4) Khách hàng lưu trữ '
 'email, cấp Mã code cho người dùng cuối và hướng dẫn kích hoạt.',label='【BỔ SUNG】')
cur=new_p_after(cur,'Điểm chung của cả hai luồng: người dùng cuối kích hoạt tại zbox.vn/activate-code; Mã code đã '
 'thanh toán và đã bàn giao thành công không được hoàn trả, hoàn tiền hoặc quy đổi (trừ Mã code lỗi thuộc trách '
 'nhiệm Bên B/nhà phát hành); khách hàng không niêm yết hoặc bán công khai thấp hơn giá bán lẻ trực tiếp do VNG/'
 'nhà phát hành công bố.',label='【BỔ SUNG】')

# Cập nhật dòng "Giao code" trong bảng tóm tắt (T1) -> nêu 2 luồng + highlight
t_sum=d.tables[1]
for row in t_sum.rows:
    if row.cells[0].text.strip().startswith('4.'):
        row.cells[2].text=''; r=row.cells[2].paragraphs[0].add_run('Giao Mã code cho khách hàng qua CMS (Luồng 1) hoặc thủ công qua email (Luồng 2)'); r.font.size=Pt(9.3); hl_run(r)
        row.cells[3].text=''; r=row.cells[3].paragraphs[0].add_run('CMS/zbox.vn hoặc email'); r.font.size=Pt(9.3); hl_run(r)

# Ghi chú ở phần mục đích (đầu tài liệu) về bản revised
mp=find_p('Phụ lục này quy định trình tự')
r=mp.add_run(' (Bản cập nhật: bổ sung hai luồng nhận Mã code của khách hàng — qua hệ thống CMS và thủ công qua email '
 'theo Hợp đồng dịch vụ giữa GAPIT và khách hàng doanh nghiệp/đối tác bán lại.)'); hl_run(r); r.italic=True

d.save(OUT); print('SAVED',OUT)
