# -*- coding: utf-8 -*-
"""Tạo 2 DPA (eSIM, zVAS) từ KHUNG GỐC mẫu PDPA(VNI): sao chép nguyên văn mẫu,
chỉ thay phần 'hình thức dịch vụ' (số hợp đồng, mục đích/phạm vi, bảng dữ liệu).
Quy ước theo PDPA(VNI): Bên A = Bên Kiểm Soát (khách hàng), Bên B = GAPIT (Bên Xử Lý)."""
import docx
from docx.shared import Pt

SRC = '/root/.claude/uploads/0ae05b9e-80f9-5907-8558-033a769d7c2e/74d6ff3c-PDPAVNI.docx'

def set_para_text(p, text):
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ''
    else:
        p.add_run(text)

def find_para(doc, needle):
    for p in doc.paragraphs:
        if needle in p.text:
            return p
    raise ValueError('not found: '+needle)

def set_cell(cell, lines):
    if isinstance(lines, str): lines = [lines]
    cell._tc.clear_content()
    for ln in lines:
        p = cell.add_paragraph()
        r = p.add_run(ln)
        r.font.name = 'Times New Roman'; r.font.size = Pt(12)

def build(params, out):
    d = docx.Document(SRC)
    plist = list(d.paragraphs)

    def fp(needle):
        for p in plist:
            if needle in p.text:
                return p
        raise ValueError('not found: '+needle)

    set_para_text(fp('K...m theo H...p ...'.replace('...','')) if False else fp('m theo H'), params['title_ref'])
    set_para_text(fp('Party A'), params['party_a'])

    gap_idx = next(i for i,p in enumerate(plist) if 'GAPIT' in p.text and 'Nh' in p.text and 'cung c' in p.text)
    for p in plist[gap_idx:gap_idx+8]:
        t = p.text.strip()
        if t.startswith('\u0110\u1ea1i di\u1ec7n'):
            set_para_text(p, '\u0110\u1ea1i di\u1ec7n\t\t: \u00d4ng Nguy\u1ec5n V\u0103n Long')
        elif t.startswith('Ch\u1ee9c v\u1ee5'):
            set_para_text(p, 'Ch\u1ee9c v\u1ee5\t\t: Gi\u00e1m \u0111\u1ed1c')

    for p in plist:
        if 'ARTICLE 2' in p.text:
            set_para_text(p, '')

    p_mucdich = fp('ph\u00f9 h\u1ee3p v\u1edbi m\u1ee5c \u0111\u00edch, ph\u1ea1m vi d\u1ecbch v\u1ee5')
    set_para_text(p_mucdich,
        'M\u1ee5c \u0111\u00edch X\u1eed L\u00fd D\u1eef Li\u1ec7u C\u00e1 Nh\u00e2n: ph\u00f9 h\u1ee3p v\u1edbi m\u1ee5c \u0111\u00edch, ph\u1ea1m vi d\u1ecbch v\u1ee5, c\u00f4ng vi\u1ec7c, '
        'quy\u1ec1n v\u00e0 ngh\u0129a v\u1ee5 c\u1ee7a C\u00e1c B\u00ean theo ' + params['contract_name'] +
        ' (sau \u0111\u00e2y g\u1ecdi chung l\u00e0 \u201cH\u1ee3p \u0111\u1ed3ng\u201d), c\u1ee5 th\u1ec3 nh\u01b0 sau:')

    mi = plist.index(p_mucdich)
    target = None
    for p in plist[mi+1:mi+4]:
        if not p.text.strip():
            target = p; break
    if target is None:
        target = plist[mi+1]
    set_para_text(target, params['purpose'])

    # 7) Bảng dữ liệu (Table 0): hàng 1 = cơ bản, hàng 3 = nhạy cảm
    t0 = d.tables[0]
    set_cell(t0.rows[1].cells[0], params['basic_data'])
    set_cell(t0.rows[1].cells[1], params['basic_purpose'])
    set_cell(t0.rows[1].cells[2], params['basic_subject'])
    set_cell(t0.rows[3].cells[0], params['sensitive_data'])
    set_cell(t0.rows[3].cells[1], params['sensitive_purpose'])
    set_cell(t0.rows[3].cells[2], params['sensitive_subject'])

    # 8) Bảng chữ ký (Table 1): điền đại diện Bên B = GAPIT
    t1 = d.tables[-1]
    # cột phải = Bên B
    set_cell(t1.rows[0].cells[-1],
             ['ĐẠI DIỆN BÊN B', 'CÔNG TY CỔ PHẦN GAPIT', 'Ông Nguyễn Văn Long – Giám đốc', '(Ký, ghi rõ họ tên, đóng dấu)'])
    set_cell(t1.rows[0].cells[0],
             ['ĐẠI DIỆN BÊN A', params['party_a_sign'], '(Ký, ghi rõ họ tên, đóng dấu)'])

    d.save(out)
    print('SAVED', out)

# ===================== eSIM =====================
esim = dict(
    title_ref='(Kèm theo Hợp đồng Đại lý phân phối SIM/eSIM du lịch số ……/2026/HĐĐL/GPT ký ngày …… tháng …… năm 2026)',
    party_a='Bên A/ Party A: ………………………………………… (ĐẠI LÝ – Bên Kiểm Soát Dữ Liệu Cá Nhân)',
    party_a_sign='(ĐẠI LÝ)',
    contract_name='Hợp đồng Đại lý phân phối SIM/eSIM du lịch số ……/2026/HĐĐL/GPT',
    purpose=('Theo đó, Bên A làm đại lý phân phối Sản phẩm SIM/eSIM du lịch do Bên B cung cấp tới khách hàng cuối; '
             'Bên B thay mặt Bên A thực hiện Xử Lý Dữ Liệu Cá Nhân của khách hàng cuối nhằm: tiếp nhận và xử lý đơn hàng '
             'trên hệ thống CMS; cấp và gửi mã QR/thông tin kích hoạt eSIM cho khách hàng cuối qua email; xuất hóa đơn '
             'giá trị gia tăng (qua đối tác phát hành hóa đơn điện tử – VNPT) tương ứng với đơn hàng đã thanh toán; '
             'xử lý thanh toán qua kênh thanh toán tích hợp (VietinBank); hỗ trợ kỹ thuật, hướng dẫn kích hoạt và chăm sóc '
             'khách hàng; lưu trữ, tra cứu và đối soát đơn hàng theo Hợp đồng.'),
    basic_data=('Họ và tên; Số điện thoại; Địa chỉ liên hệ, địa chỉ email (để nhận mã QR/eSIM và hóa đơn); '
                'Số định danh cá nhân/số hộ chiếu (nếu khách hàng cung cấp khi đăng ký theo yêu cầu của quốc gia đích); '
                'Thông tin về tài khoản số của cá nhân (ví dụ: tài khoản Zalo); Mã số thuế cá nhân và thông tin xuất hóa đơn '
                '(khi có yêu cầu); Thông tin đơn hàng (mã đơn, loại SIM/eSIM, gói cước, số lượng, giá, trạng thái xử lý); '
                'Dữ liệu khác giúp xác định Chủ Thể Dữ Liệu Cá Nhân.'),
    basic_purpose=('Phục vụ Hợp đồng Đại lý phân phối SIM/eSIM du lịch số ……/2026/HĐĐL/GPT: xử lý đơn hàng, cấp và gửi '
                   'mã QR/eSIM, thanh toán, xuất hóa đơn, hỗ trợ và chăm sóc khách hàng, lưu trữ – tra cứu – đối soát.'),
    basic_subject=('Khách hàng cuối mua Sản phẩm SIM/eSIM du lịch thông qua Bên A (Đại lý); đầu mối liên hệ của Bên A '
                   'phục vụ đặt hàng và xuất hóa đơn.'),
    sensitive_data=('Theo nguyên tắc không thu thập. Trường hợp phát sinh (ví dụ: thông tin giao dịch tài chính qua tổ chức '
                    'trung gian thanh toán, dữ liệu vị trí, hình ảnh giấy tờ định danh), Các Bên bổ sung danh mục tại Phụ lục '
                    'và áp dụng biện pháp bảo mật theo Pháp Luật Bảo Vệ Dữ Liệu.'),
    sensitive_purpose='Không áp dụng theo mặc định.',
    sensitive_subject='Không áp dụng theo mặc định.',
)

# ===================== zVAS =====================
zvas = dict(
    title_ref='(Kèm theo Hợp đồng Cung cấp Dịch vụ zVAS số ……/2026/HĐDV-zVAS/GPT ký ngày …… tháng …… năm 2026)',
    party_a='Bên A/ Party A: ………………………………………… (KHÁCH HÀNG DOANH NGHIỆP – Bên Kiểm Soát Dữ Liệu Cá Nhân)',
    party_a_sign='(KHÁCH HÀNG DOANH NGHIỆP)',
    contract_name='Hợp đồng Cung cấp Dịch vụ zVAS số ……/2026/HĐDV-zVAS/GPT',
    purpose=('Theo đó, Bên B cung cấp các dịch vụ giá trị gia tăng trên nền tảng Zalo và ZingMP3 (“Dịch vụ zVAS”) cho Bên A; '
             'Bên B thay mặt Bên A thực hiện Xử Lý Dữ Liệu Cá Nhân của Người Dùng Cuối do Bên A chỉ định nhằm: tiếp nhận và '
             'xử lý đơn hàng; bàn giao và hỗ trợ kích hoạt Mã code zVAS; onboarding, hỗ trợ kỹ thuật và chăm sóc khách hàng; '
             'lập và xuất hóa đơn theo Hợp đồng.'),
    basic_data=('Họ, tên đệm và tên; Số điện thoại; Địa chỉ email; Địa chỉ liên hệ; Thông tin về tài khoản số của cá nhân '
                '(ví dụ: định danh tài khoản Zalo/ZingMP3 phục vụ kích hoạt); Thông tin đặt hàng, kích hoạt và hỗ trợ '
                '(mã đơn hàng, loại dịch vụ, lịch sử yêu cầu hỗ trợ); Dữ liệu khác giúp xác định Chủ Thể Dữ Liệu Cá Nhân.'),
    basic_purpose=('Phục vụ Hợp đồng Cung cấp Dịch vụ zVAS số ……/2026/HĐDV-zVAS/GPT: cung cấp Dịch vụ zVAS (xử lý đơn hàng, '
                   'bàn giao và kích hoạt Mã code, onboarding, hỗ trợ kỹ thuật và chăm sóc khách hàng, lập và xuất hóa đơn).'),
    basic_subject=('Người Dùng Cuối do Bên A chỉ định (cán bộ, nhân viên, khách hàng hoặc cá nhân khác của Bên A) được cấp và '
                   'sử dụng Mã code zVAS; đầu mối liên hệ của Bên A phục vụ đặt hàng và xuất hóa đơn.'),
    sensitive_data=('Không áp dụng theo mặc định. Theo bản chất Dịch vụ zVAS, Bên B chủ yếu xử lý dữ liệu cá nhân cơ bản; chỉ '
                    'xử lý dữ liệu cá nhân nhạy cảm khi Bên A chỉ dẫn cụ thể bằng văn bản và bổ sung danh mục tại Phụ lục, phù '
                    'hợp với Pháp Luật Bảo Vệ Dữ Liệu.'),
    sensitive_purpose='Không áp dụng theo mặc định.',
    sensitive_subject='Không áp dụng theo mặc định.',
)

import os
os.makedirs('/home/user/finalproject/output', exist_ok=True)
build(esim, '/home/user/finalproject/output/Thoa_thuan_XLDLCN_eSIM_GAPIT_DaiLy_PDPAVNI.docx')
build(zvas, '/home/user/finalproject/output/Thoa_thuan_XLDLCN_zVAS_GAPIT_KHDN_PDPAVNI.docx')
