# -*- coding: utf-8 -*-
"""Tạo DPA cho ĐẠI LÝ zVAS từ KHUNG GỐC mẫu PDPA(VNI) (sao chép nguyên văn),
chỉ thay phần 'hình thức dịch vụ' lấy từ bản DPA Đại lý zVAS cũ.
Quy ước PDPA(VNI): Bên A = Bên Kiểm Soát (Đại lý), Bên B = GAPIT (Bên Xử Lý)."""
import docx, os
from docx.shared import Pt

SRC = '/root/.claude/uploads/0ae05b9e-80f9-5907-8558-033a769d7c2e/74d6ff3c-PDPAVNI.docx'

def set_para_text(p, text):
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ''
    else:
        p.add_run(text)

def set_cell(cell, lines):
    if isinstance(lines, str): lines = [lines]
    cell._tc.clear_content()
    for ln in lines:
        p = cell.add_paragraph(); r = p.add_run(ln)
        r.font.name = 'Times New Roman'; r.font.size = Pt(12)

d = docx.Document(SRC)
plist = list(d.paragraphs)
def fp(needle):
    for p in plist:
        if needle in p.text: return p
    raise ValueError('not found: '+needle)

# Tiêu đề
set_para_text(fp('m theo H'),
    '(Kèm theo Hợp đồng Đại lý phân phối Dịch vụ zVAS số ……/2026/HĐĐL-zVAS/GPT ký ngày …… tháng …… năm 2026)')
# Bên A = Đại lý (Bên Kiểm Soát)
set_para_text(fp('Party A'),
    'Bên A/ Party A: ………………………………………… (ĐẠI LÝ – Bên Kiểm Soát Dữ Liệu Cá Nhân)')
# Đại diện GAPIT (Bên B)
gap_idx = next(i for i,p in enumerate(plist) if 'GAPIT' in p.text and 'Nh' in p.text and 'cung c' in p.text)
for p in plist[gap_idx:gap_idx+8]:
    t = p.text.strip()
    if t.startswith('Đại diện'): set_para_text(p, 'Đại diện\t\t: Ông Nguyễn Văn Long')
    elif t.startswith('Chức vụ'): set_para_text(p, 'Chức vụ\t\t: Giám đốc')
# bỏ dòng tiếng Anh lạc
for p in plist:
    if 'ARTICLE 2' in p.text: set_para_text(p, '')

# Điều 2 — số hợp đồng
p_md = fp('phù hợp với mục đích, phạm vi dịch vụ')
set_para_text(p_md,
    'Mục đích Xử Lý Dữ Liệu Cá Nhân: phù hợp với mục đích, phạm vi dịch vụ, công việc, quyền và nghĩa vụ '
    'của Các Bên theo Hợp đồng Đại lý phân phối Dịch vụ zVAS số ……/2026/HĐĐL-zVAS/GPT '
    '(sau đây gọi chung là “Hợp đồng”), cụ thể như sau:')
# mô tả dịch vụ
mi = plist.index(p_md); target=None
for p in plist[mi+1:mi+4]:
    if not p.text.strip(): target=p; break
if target is None: target=plist[mi+1]
set_para_text(target,
    'Theo đó, Bên A làm đại lý phân phối các dịch vụ giá trị gia tăng trên nền tảng Zalo và ZingMP3 '
    '(“Dịch vụ zVAS”) dưới dạng Mã code zVAS do Bên B cung cấp tới khách hàng/Người Dùng Cuối của Bên A; '
    'Bên B thay mặt Bên A thực hiện Xử Lý Dữ Liệu Cá Nhân của khách hàng/Người Dùng Cuối do Bên A chuyển giao '
    'nhằm: tiếp nhận và xử lý đơn hàng của Bên A; bàn giao Mã code zVAS; hỗ trợ kích hoạt và xử lý sự cố kỹ thuật '
    'theo yêu cầu hỗ trợ do Bên A chuyển tiếp; lập và xuất hóa đơn theo Hợp đồng.')

# Bảng dữ liệu
t0 = d.tables[0]
set_cell(t0.rows[1].cells[0],
    'Họ, tên đệm và tên; Số điện thoại; Địa chỉ email; Địa chỉ liên hệ; Thông tin về tài khoản số của cá nhân '
    '(ví dụ: định danh tài khoản Zalo/ZingMP3 phục vụ kích hoạt); Thông tin đặt hàng, kích hoạt và hỗ trợ '
    '(mã đơn hàng, loại dịch vụ, lịch sử yêu cầu hỗ trợ); Dữ liệu khác giúp xác định Chủ Thể Dữ Liệu Cá Nhân.')
set_cell(t0.rows[1].cells[1],
    'Phục vụ Hợp đồng Đại lý phân phối Dịch vụ zVAS số ……/2026/HĐĐL-zVAS/GPT: hỗ trợ Bên A phân phối Dịch vụ zVAS '
    '(xử lý đơn hàng của Bên A, bàn giao Mã code, hỗ trợ kích hoạt và xử lý sự cố kỹ thuật do Bên A chuyển tiếp, '
    'lập và xuất hóa đơn).')
set_cell(t0.rows[1].cells[2],
    'Khách hàng/Người Dùng Cuối của Bên A (Đại lý) được Bên A phân phối và cấp Mã code zVAS; đầu mối liên hệ của '
    'Bên A phục vụ đặt hàng và xử lý hỗ trợ.')
set_cell(t0.rows[3].cells[0],
    'Không áp dụng theo mặc định. Theo bản chất Dịch vụ zVAS, Bên B chủ yếu xử lý dữ liệu cá nhân cơ bản; chỉ xử lý '
    'dữ liệu cá nhân nhạy cảm khi Bên A chỉ dẫn cụ thể bằng văn bản và bổ sung danh mục tại Phụ lục, phù hợp với '
    'Pháp Luật Bảo Vệ Dữ Liệu.')
set_cell(t0.rows[3].cells[1], 'Không áp dụng theo mặc định.')
set_cell(t0.rows[3].cells[2], 'Không áp dụng theo mặc định.')

# Bảng chữ ký
t1 = d.tables[-1]
set_cell(t1.rows[0].cells[-1],
         ['ĐẠI DIỆN BÊN B', 'CÔNG TY CỔ PHẦN GAPIT', 'Ông Nguyễn Văn Long – Giám đốc', '(Ký, ghi rõ họ tên, đóng dấu)'])
set_cell(t1.rows[0].cells[0],
         ['ĐẠI DIỆN BÊN A', '(ĐẠI LÝ)', '(Ký, ghi rõ họ tên, đóng dấu)'])

os.makedirs('/home/user/finalproject/output', exist_ok=True)
out='/home/user/finalproject/output/Thoa_thuan_XLDLCN_zVAS_GAPIT_DaiLy_PDPAVNI.docx'
d.save(out); print('SAVED', out)
