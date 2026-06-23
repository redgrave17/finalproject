# -*- coding: utf-8 -*-
"""Thêm 2 điều khoản (bảo vệ GAPIT=Bên B) vào HĐ VNG–GAPIT, highlight phần bổ sung.
ĐK1: thời hạn hỗ trợ đổi/hoàn Mã code lỗi từ bàn giao đến kích hoạt + gia hạn khi gần hết hạn.
ĐK2: hỗ trợ Bên B khi Bên A (VNG) tạm ngừng/chấm dứt dịch vụ với Mã code đã mua chưa bán/chưa kích hoạt."""
import copy, docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor, Pt

SRC = '/root/.claude/uploads/0ae05b9e-80f9-5907-8558-033a769d7c2e/6a26830c-ZALG_R1__Form_HD_Cung_C_p_DV_zVas_VNG_fb1_2.docx'
OUT = '/home/user/finalproject/output/HD_zVAS_VNG_GAPIT_BO_SUNG_DIEU_KHOAN.docx'
Y = WD_COLOR_INDEX.YELLOW

d = docx.Document(SRC)

def para_text(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))

def find_p(needle):
    for el in d.element.iter(qn('w:p')):
        if needle in para_text(el):
            return Paragraph(el, None)
    raise ValueError('NOT FOUND: '+needle)

# font tham chiếu (tài liệu pháp lý VN)
def ref_font(anchor):
    return 'Times New Roman', Pt(13)

def make_para_after(anchor, segments, ref_pPr, fname, fsize):
    """segments: list of (text, bold, highlight). Trả về Paragraph mới."""
    new_p = OxmlElement('w:p')
    anchor._p.addnext(new_p)
    if ref_pPr is not None:
        pPr = copy.deepcopy(ref_pPr)
        for nm in pPr.findall(qn('w:numPr')):  # bỏ auto-number để không lệch số
            pPr.remove(nm)
        new_p.append(pPr)
    np = Paragraph(new_p, None)
    for text, bold, hl in segments:
        r = np.add_run(text)
        r.bold = bold
        if fname: r.font.name = fname
        if fsize: r.font.size = fsize
        if hl: r.font.highlight_color = Y
        if bold and hl is False:
            r.font.color.rgb = RGBColor(0xC0, 0, 0)
    return np

# Điểm chèn: sau đoạn "cung cấp lại Mã code zVas khác trong vòng tối đa 24 giờ làm việc"
anchor = find_p('tối đa 24')
fname, fsize = ref_font(anchor)
ref_pPr = anchor._p.find(qn('w:pPr'))

# Nội dung 2 điều khoản (mỗi gạch đầu dòng = 1 đoạn)
blocks = [
 # ĐIỀU KHOẢN 1
 [('(BỔ SUNG) Thời hạn hỗ trợ đối với Mã code bị lỗi (từ thời điểm bàn giao đến khi kích hoạt):', True, True)],
 [('(i) Mã code zVas phải được kích hoạt trong thời hạn 12 (mười hai) tháng kể từ ngày Bên A phát hành/bàn giao '
   'Mã code cho Bên B (“Thời hạn kích hoạt”).', False, True)],
 [('(ii) Trong suốt Thời hạn kích hoạt, nếu Mã code bị lỗi, không sử dụng được, đã bị sử dụng trước khi bàn giao '
   'hoặc không kích hoạt được mà không do lỗi của Bên B/Người Dùng Cuối, Bên A có trách nhiệm hủy/thu hồi và cấp '
   'lại Mã code khác (hoặc hoàn tiền tương ứng nếu không thể cấp lại) trong vòng tối đa 24 (hai mươi bốn) Giờ làm '
   'việc kể từ khi tiếp nhận yêu cầu hợp lệ của Bên B.', False, True)],
 [('(iii) Trường hợp Bên B phát hiện lỗi và gửi yêu cầu hợp lệ tới Bên A trước thời điểm hết Thời hạn kích hoạt '
   '(kể cả khi gần hết hạn), nghĩa vụ hỗ trợ đổi/cấp lại của Bên A vẫn được thực hiện đầy đủ ngay cả khi việc xử lý '
   'hoàn tất sau thời điểm hết Thời hạn kích hoạt. Mã code được cấp lại có Thời hạn kích hoạt và Thời hạn Gói mới '
   'tương đương Mã code ban đầu, tính từ ngày cấp lại.', False, True)],
 [('(iv) Trường hợp do đặc thù kỹ thuật cần thêm thời gian, Bên A thông báo cho Bên B và gia hạn thời hạn xử lý một '
   'cách hợp lý; khoảng thời gian Bên A xử lý lỗi không được tính trừ vào Thời hạn kích hoạt/Thời hạn Gói của Bên B.', False, True)],
 # ĐIỀU KHOẢN 2
 [('(BỔ SUNG) Hỗ trợ Bên B khi Bên A tạm ngừng/chấm dứt cung cấp Dịch vụ zVas đối với Mã code Bên B đã mua nhưng '
   'chưa phân phối/chưa kích hoạt:', True, True)],
 [('(i) Trường hợp Bên A tạm ngừng, thay đổi hoặc chấm dứt cung cấp một hoặc nhiều Dịch vụ zVas trong khi Bên B còn '
   'nắm giữ Mã code đã thanh toán nhưng chưa bàn giao cho Người Dùng Cuối, hoặc đã bàn giao nhưng chưa kích hoạt, '
   'Bên A phải thông báo bằng văn bản cho Bên B tối thiểu 30 (ba mươi) ngày trước thời điểm tạm ngừng/chấm dứt.', False, True)],
 [('(ii) Đối với các Mã code đã thanh toán nhưng chưa kích hoạt của dịch vụ bị tạm ngừng/chấm dứt, Bên A hỗ trợ Bên B '
   'theo một hoặc kết hợp các phương án sau (theo thỏa thuận của Hai Bên): (a) gia hạn Thời hạn kích hoạt để Bên B '
   'tiếp tục phân phối; (b) chuyển đổi sang Mã code của dịch vụ tương đương về giá trị; hoặc (c) hoàn tiền cho Bên B '
   'phần giá trị Mã code chưa kích hoạt theo đơn giá Bên B đã thanh toán, trong vòng 30 (ba mươi) ngày kể từ ngày '
   'Hai Bên thống nhất phương án.', False, True)],
 [('(iii) Trong thời gian tạm ngừng, Bên A bảo lưu hiệu lực các Mã code đã thanh toán của Bên B và khôi phục khả năng '
   'kích hoạt khi dịch vụ được cung cấp trở lại; thời gian tạm ngừng không được tính trừ vào Thời hạn kích hoạt của '
   'Bên B.', False, True)],
 [('(iv) Việc tạm ngừng/chấm dứt do Bên A không làm miễn trừ nghĩa vụ hỗ trợ và quyết toán nêu trên của Bên A đối với '
   'Bên B.', False, True)],
]

cur = anchor
for seg in blocks:
    cur = make_para_after(cur, seg, ref_pPr, fname, fsize)

import os
os.makedirs('/home/user/finalproject/output', exist_ok=True)
d.save(OUT)
print('SAVED', OUT)
