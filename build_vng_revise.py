# -*- coding: utf-8 -*-
"""Sửa 2 điều khoản bổ sung trong HĐ VNG–GAPIT theo góp ý pháp chế:
 - Bỏ khái niệm 'GAPIT phải kích hoạt trong 12 tháng' (rào cản); làm rõ thời hạn kích hoạt là cho NDC sau khi bán.
 - Bỏ 'mua lại/hoàn tiền/gia hạn bắt buộc'; thay bằng hỗ trợ thúc đẩy bán (refer KHDN) nỗ lực thiện chí; chuyển đổi chỉ khi 2 bên đồng ý.
Highlight phần sửa."""
import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt
Y=WD_COLOR_INDEX.YELLOW
SRC='/root/.claude/uploads/0ae05b9e-80f9-5907-8558-033a769d7c2e/c17dec0b-HD_zVAS_VNG_GAPIT_BO_SUNG_DIEU_KHOAN.docx'
OUT='/home/user/finalproject/output/HD_zVAS_VNG_GAPIT_SUA_THEO_GOPY.docx'
d=docx.Document(SRC)

def ptext(el): return ''.join(t.text or '' for t in el.iter(qn('w:t')))
def find(needle):
    for el in d.element.iter(qn('w:p')):
        if needle in ptext(el): return Paragraph(el,None)
    raise ValueError('NOT FOUND: '+needle)
def setp(needle,new,bold=False):
    p=find(needle)
    if p.runs:
        p.runs[0].text=new
        for r in p.runs[1:]: r.text=''
        r0=p.runs[0]
    else:
        r0=p.add_run(new)
    r0.bold=bold; r0.font.highlight_color=Y; r0.font.name='Times New Roman'; r0.font.size=Pt(13)

# ---------- ĐIỀU KHOẢN 1: Mã code lỗi (sửa lại, bỏ rào cản 12 tháng) ----------
setp('(BỔ SUNG) Thời hạn hỗ trợ đối với Mã code bị lỗi',
 '(BỔ SUNG – ĐÃ CHỈNH THEO GÓP Ý) Bảo đảm và hỗ trợ xử lý đối với Mã code bị lỗi:',bold=True)
setp('Mã code zVas phải được kích hoạt trong thời hạn 12 (mười hai) tháng',
 '(i) Bên A bảo đảm Mã code bàn giao cho Bên B là hợp lệ, chưa sử dụng và có thể kích hoạt trong thời hạn hiệu lực '
 'kích hoạt do Bên A công bố cho từng đợt phát hành. Thời hạn hiệu lực kích hoạt áp dụng cho việc kích hoạt của '
 'Người Dùng Cuối SAU KHI mua; không phải là nghĩa vụ buộc Bên B phải kích hoạt Mã code khi chưa bán cho Người '
 'Dùng Cuối.')
setp('Trong suốt Thời hạn kích hoạt, nếu Mã code bị lỗi',
 '(ii) Trường hợp Mã code bị lỗi, không sử dụng được, đã bị sử dụng trước khi bàn giao, hoặc không kích hoạt được '
 'mà không do lỗi của Bên B/Người Dùng Cuối, Bên A hủy/thu hồi và cấp lại Mã code khác cùng loại với thời hạn hiệu '
 'lực kích hoạt tương đương, trong vòng tối đa 24 (hai mươi bốn) Giờ làm việc kể từ khi nhận yêu cầu hợp lệ; trường '
 'hợp không thể cấp lại, Hai Bên thống nhất phương án xử lý phù hợp.')
setp('Trường hợp Bên B phát hiện lỗi và gửi yêu cầu hợp lệ',
 '(iii) Quyền yêu cầu đổi/cấp lại Mã code lỗi của Bên B được bảo lưu nếu Bên B gửi yêu cầu hợp lệ tới Bên A trước '
 'thời điểm Mã code hết thời hạn hiệu lực kích hoạt (kể cả khi gần hết hạn), ngay cả khi việc xử lý của Bên A hoàn '
 'tất sau thời điểm đó.')
setp('Trường hợp do đặc thù kỹ thuật cần thêm thời gian',
 '(iv) Khoảng thời gian Bên A xử lý lỗi (kiểm tra, thu hồi, cấp lại) không bị tính trừ vào thời hạn hiệu lực kích '
 'hoạt của Mã code thay thế.')

# ---------- ĐIỀU KHOẢN 2: Hỗ trợ tiêu thụ / khi ngừng dịch vụ (sửa lại theo thực tế) ----------
setp('(BỔ SUNG) Hỗ trợ Bên B khi Bên A tạm ngừng',
 '(BỔ SUNG – ĐÃ CHỈNH THEO GÓP Ý) Hỗ trợ tiêu thụ Mã code còn tồn và khi Bên A thay đổi/ngừng cung cấp Dịch vụ zVas:',bold=True)
setp('Trường hợp Bên A tạm ngừng, thay đổi hoặc chấm dứt cung cấp một hoặc nhiều Dịch vụ zVas trong khi Bên B còn',
 '(i) Các Bên thống nhất: theo chính sách chung của Bên A, Mã code zVas đã bàn giao cho Bên B sẽ không được nhận '
 'lại, không mua lại và không hoàn tiền.')
setp('Đối với các Mã code đã thanh toán nhưng chưa kích hoạt của dịch vụ bị tạm ngừng',
 '(ii) Hỗ trợ thúc đẩy tiêu thụ (nỗ lực thiện chí): trường hợp Bên B còn Mã code đã mua chưa phân phối hết, Bên A '
 'nỗ lực hợp lý hỗ trợ Bên B tiêu thụ thông qua các chương trình thương mại theo từng thời kỳ, bao gồm giới thiệu/'
 'kết nối khách hàng doanh nghiệp tiềm năng cho Bên B và phối hợp thúc đẩy bán. Đây là hỗ trợ trên cơ sở nỗ lực '
 'thiện chí, không cấu thành nghĩa vụ bảo đảm doanh số, mua lại hay hoàn tiền.')
setp('Trong thời gian tạm ngừng, Bên A bảo lưu hiệu lực',
 '(iii) Khi Bên A dự kiến thay đổi, tạm ngừng hoặc chấm dứt cung cấp một Dịch vụ zVas, Bên A thông báo cho Bên B '
 'tối thiểu 30 (ba mươi) ngày trước; trong thời gian đó và một khoảng thời gian chuyển tiếp hợp lý, Bên A bảo lưu '
 'thời hạn hiệu lực kích hoạt của các Mã code Bên B đã mua (chưa kích hoạt) thuộc dịch vụ đó để Người Dùng Cuối '
 'tiếp tục kích hoạt, và phối hợp hỗ trợ Bên B tiêu thụ phần còn lại.')
setp('Việc tạm ngừng/chấm dứt do Bên A không làm miễn trừ',
 '(iv) Trường hợp Hai Bên cùng có nhu cầu, Hai Bên có thể thỏa thuận chuyển đổi phần Mã code chưa kích hoạt sang '
 'Mã code của dịch vụ khác tương đương về giá trị; việc chuyển đổi chỉ thực hiện khi cả Hai Bên đồng ý, không mang '
 'tính bắt buộc.')

d.save(OUT); print('SAVED',OUT)
