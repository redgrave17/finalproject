# -*- coding: utf-8 -*-
"""Tách Business Roadmap zVAS thành 3 slide riêng:
 1) Roadmap Now/Next/Later  2) Biên lợi nhuận GAPIT & ĐL cấp II  3) Order Flow zVAS CMS (swimlane).
Branding GAPIT/Zalo."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

C=lambda h:RGBColor.from_string(h)
ZALO=C('0068FF'); NAVY=C('001A40'); BLUE2=C('003F88'); ORANGE=C('F5862C')
TEAL=C('0F6E56'); GREEN=C('22C55E'); PURPLE=C('534AB7'); AMBER=C('BA7517')
CLBLUE=C('185FA5'); GRAY=C('718096'); GRAY2=C('4A5568'); LIGHT=C('F7F8F8')
WHITE=C('FFFFFF'); LN=C('E2E8F0'); SLATE=C('94A3B8')
HEAD='Arial'; BODY='Calibri'
LOGO='/home/user/finalproject/assets/image4.png'

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)

def rect(sl,x,y,w,h,fill,line=None,rounded=False,lw=0.75):
    sh=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                           Inches(x),Inches(y),Inches(w),Inches(h))
    sh.shadow.inherit=False
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb=fill
    if line is None: sh.line.fill.background()
    else: sh.line.color.rgb=line; sh.line.width=Pt(lw)
    return sh

def txt(sl,x,y,w,h,paras,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,wrap=True):
    tb=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=tb.text_frame
    tf.word_wrap=wrap; tf.vertical_anchor=anchor
    tf.margin_left=Inches(0.05); tf.margin_right=Inches(0.05); tf.margin_top=Inches(0.02); tf.margin_bottom=Inches(0.01)
    for i,(runs,sz,sb) in enumerate(paras):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align; p.space_after=Pt(0); p.space_before=Pt(sb); p.line_spacing=1.03
        for t,c,b,fn in runs:
            r=p.add_run(); r.text=t; r.font.size=Pt(sz); r.font.bold=b; r.font.color.rgb=c; r.font.name=fn
    return tb

def bullets(sl,x,y,w,h,items,clr,sz=11,sa=5):
    tb=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=tb.text_frame
    tf.word_wrap=True; tf.margin_left=Inches(0.05); tf.margin_right=Inches(0.05); tf.margin_top=Inches(0.02)
    for i,it in enumerate(items):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.space_after=Pt(sa); p.space_before=Pt(0); p.line_spacing=1.02
        r=p.add_run(); r.text='▪ '; r.font.size=Pt(sz); r.font.color.rgb=clr; r.font.bold=True; r.font.name=BODY
        r2=p.add_run(); r2.text=it; r2.font.size=Pt(sz); r2.font.color.rgb=GRAY2; r2.font.name=BODY

def frame(sl,chip,title,subtitle_paras):
    rect(sl,0,0,13.333,0.62,LIGHT)
    sl.shapes.add_picture(LOGO,Inches(0.45),Inches(0.12),height=Inches(0.40))
    txt(sl,1.78,0.12,3.0,0.40,[([('×  Zalo VAS',ZALO,True,HEAD)],14,0)],anchor=MSO_ANCHOR.MIDDLE)
    rect(sl,10.95,0.10,2.10,0.40,ORANGE,rounded=True)
    txt(sl,10.95,0.10,2.10,0.40,[([(chip,WHITE,True,BODY)],8.5,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    txt(sl,0.45,0.70,12.6,0.46,[([(title,NAVY,True,HEAD)],25,0)])
    txt(sl,0.45,1.22,12.6,0.30,subtitle_paras)
    rect(sl,0.45,7.20,12.45,0.22,BLUE2,rounded=True)
    txt(sl,0.6,7.20,12.1,0.22,[([('GAPIT JSC  |  gapit.com.vn',WHITE,True,BODY),
       ('     —  Business Roadmap sản phẩm zVAS',WHITE,False,BODY)],8,0)],anchor=MSO_ANCHOR.MIDDLE)

blank=prs.slide_layouts[6]

# ============================================================ SLIDE 1: ROADMAP
s1=prs.slides.add_slide(blank)
frame(s1,'ROADMAP · 1/3','Business Roadmap zVAS — Now · Next · Later',
 [([('Roadmap theo kết quả/outcome & horizon — không cố định ngày, ưu tiên theo chủ đề (theme) và giá trị.',GRAY,False,BODY)],11,0),
  ([('Khung tham khảo: Adrenalin — “Roadmapping Your Digital Product: A Definitive Guide”.',GRAY,False,BODY)],9,1)])
cols=[
 (ZALO,'NOW','Q2–Q3 / 2026','Go-live & Validate','Outcome: vận hành B2B ổn định · dòng tiền trả trước · khởi động mạng đại lý',
  ['Hoàn thiện zVAS CMS: catalog, báo giá (prorate + VAT), hóa đơn, kích hoạt',
   'Code pool & cảnh báo hạn kích hoạt 6/3/2/1 tháng (trong 12 tháng từ nhập kho)',
   'Go-live KHDN + ký HĐ Đại lý của GAPIT; DPA theo mẫu PDPA(VNI)',
   'Chuẩn hóa cơ chế trả trước + bảng giá & chiết khấu theo bậc',
   'Đào tạo CSKH cấp 1; quy trình đổi Mã code lỗi (≤ 24h)']),
 (TEAL,'NEXT','Q4/2026 – Q1/2027','Scale & Enable','Outcome: nhân rộng kênh đại lý · tăng MRR · giảm thao tác thủ công',
  ['Mở rộng mạng Đại lý của GAPIT: RACI, onboarding, KPI doanh số tối thiểu',
   'Portal Đại lý / Affiliate; báo cáo MRR & tồn kho Mã code',
   'Tự động đối soát CSV zbox.vn; auto-select code sắp hết hạn',
   'Thêm bundle (zBusiness + zCloud + zStyle)',
   'Khuyến mại có kiểm soát giá sàn công khai (VNG)']),
 (PURPLE,'LATER','2027 +','Optimize & Expand','Outcome: tối ưu biên lợi nhuận · mở rộng hệ sinh thái · tích hợp sâu',
  ['API hóa khi VNG mở API → loại bỏ thao tác CSV thủ công',
   'Cross-sell eSIM / SIM doanh nghiệp; combo đa dịch vụ',
   'Analytics hành vi & dự báo nhu cầu',
   'Tối ưu pricing & biên theo bậc số lượng',
   'Loyalty cho Đại lý & KHDN; mở rộng vùng/ngành']),
]
xs=[0.45,4.64,8.83]; cw=4.05
for (clr,hz,tf_,tag,oc,items),x in zip(cols,xs):
    rect(s1,x,1.70,cw,0.92,clr,rounded=True)
    txt(s1,x+0.18,1.76,cw-0.34,0.34,[([(hz,WHITE,True,HEAD),('    '+tf_,WHITE,False,BODY)],16,0)])
    txt(s1,x+0.18,2.18,cw-0.34,0.34,[([(tag,WHITE,True,BODY)],12,0)])
    rect(s1,x,2.70,cw,3.95,WHITE,line=LN)
    rect(s1,x,2.70,cw,0.62,C('F4F7FB'))
    txt(s1,x+0.14,2.74,cw-0.28,0.56,[([(oc,clr,True,BODY)],9.5,0)])
    bullets(s1,x+0.10,3.42,cw-0.2,3.15,items,clr,sz=11,sa=7)

# ============================================================ SLIDE 2: MARGINS
s2=prs.slides.add_slide(blank)
frame(s2,'BIÊN LỢI NHUẬN · 2/3','Biên Lợi Nhuận — GAPIT & Đại lý của GAPIT',
 [([('Biên GAPIT = tuần tự trừ:  Chiết khấu GAPIT NHẬN từ VNG  −  Chiết khấu GAPIT CẤP cho Đại lý/KHDN.  (Chỉ có GAPIT và Đại lý của GAPIT — không có đại lý cấp II.)',GRAY,False,BODY)],10.5,0)])
# Hai hộp tham chiếu chiết khấu
def refbox(x,title,clr,line2):
    rect(s2,x,1.60,6.10,0.92,WHITE,line=clr,rounded=True,lw=1.25); rect(s2,x,1.60,0.10,0.92,clr)
    txt(s2,x+0.24,1.65,5.7,0.34,[([(title,clr,True,BODY)],10,0)])
    txt(s2,x+0.24,2.02,5.7,0.44,[([(line2,NAVY,True,BODY)],12.5,0)])
refbox(0.45,'①  GAPIT NHẬN từ VNG  (theo SL GAPIT đặt)',BLUE2,'100–299 code: 15%          ·          ≥ 300 code: 20%')
refbox(6.78,'②  GAPIT CẤP cho Đại lý / KHDN  (theo SL của ĐL/KHDN)',TEAL,'50–99: 5%        ·        100–299: 8%        ·        ≥ 300: 10%')
# Ma trận biên
txt(s2,0.45,2.62,12.4,0.24,[([('Biên GAPIT giữ lại  =  ①  −  ②        (mỗi ô:  ',GRAY,True,BODY),('Biên GAPIT',ORANGE,True,BODY),('  /  ',GRAY,True,BODY),('Biên Đại lý/KHDN',GREEN,True,BODY),(' )',GRAY,True,BODY)],10,0)])
mx=0.45; my=2.90; c0=3.30; cc=2.40; rh=0.60
rect(s2,mx,my,c0,rh,NAVY); txt(s2,mx,my,c0,rh,[([('GAPIT mua ↓   \\   ĐL/KHDN →',WHITE,True,BODY)],9,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
heads=['ĐL/KHDN 50–99  (CK 5%)','100–299  (CK 8%)','≥ 300  (CK 10%)']
for j,h in enumerate(heads):
    rect(s2,mx+c0+j*cc,my,cc,rh,BLUE2); txt(s2,mx+c0+j*cc,my,cc,rh,[([(h,WHITE,True,BODY)],8.6,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
matrows=[('≥ 300  (GAPIT −20%)',[('15%','5%'),('12%','8%'),('10%','10%')]),
         ('100–299  (GAPIT −15%)',[('10%','5%'),('7%','8%'),('5%','10%')])]
for i,(rl,vals) in enumerate(matrows):
    yy=my+rh*(i+1)
    rect(s2,mx,yy,c0,rh,C('EAF0F8'),line=LN,lw=0.5); txt(s2,mx,yy,c0,rh,[([(rl,NAVY,True,BODY)],9.5,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    for j,(g,a) in enumerate(vals):
        rect(s2,mx+c0+j*cc,yy,cc,rh,C('FFF6E9') if i==0 else WHITE,line=LN,lw=0.5)
        txt(s2,mx+c0+j*cc,yy,cc,rh,[([(g,ORANGE,True,BODY),('  /  ',GRAY,False,BODY),(a,GREEN,True,BODY)],11.5,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
# Ví dụ minh họa
txt(s2,0.45,4.84,12.0,0.22,[([('Ví dụ minh họa — zBusiness 12 tháng · Giá bán lẻ Zalo 1.990.000đ',GRAY,True,BODY)],10,0)])
def excard(x,clr,head,cost,buy,gp,gpp,ag,agp):
    rect(s2,x,5.10,4.05,1.32,WHITE,line=clr,rounded=True,lw=1.25); rect(s2,x,5.10,0.10,1.32,clr)
    txt(s2,x+0.22,5.15,3.75,0.46,[([(head,clr,True,BODY)],9.2,0)])
    txt(s2,x+0.22,5.62,3.75,0.76,[
       ([('Giá vốn GAPIT: ',GRAY2,False,BODY),(cost,NAVY,True,BODY),('   ·   ĐL/KHDN mua: ',GRAY2,False,BODY),(buy,NAVY,True,BODY)],8.6,0),
       ([('Biên GAPIT: ',GRAY2,False,BODY),(gp+' ('+gpp+')',ORANGE,True,BODY)],9.2,2),
       ([('Biên Đại lý/KHDN: ',GRAY2,False,BODY),(ag+' ('+agp+')',GREEN,True,BODY)],9.2,1)])
excard(0.45,GREEN,'GAPIT mua ≥300 (−20%)  →  bán ĐL ≥300 (CK 10%)','1.592.000đ','1.791.000đ','199.000đ','10%','199.000đ','10%')
excard(4.64,ORANGE,'GAPIT mua ≥300 (−20%)  →  bán ĐL 50–99 (CK 5%)','1.592.000đ','1.890.500đ','298.500đ','15%','99.500đ','5%')
excard(8.83,BLUE2,'GAPIT mua 100–299 (−15%)  →  bán ĐL 100–299 (CK 8%)','1.691.500đ','1.830.800đ','139.300đ','7%','159.200đ','8%')
# Ghi chú
rect(s2,0.45,6.55,12.45,0.58,C('FFF6E9'),line=ORANGE,rounded=True)
txt(s2,0.65,6.59,12.1,0.5,[
 ([('Lưu ý: ',ORANGE,True,BODY),('Chiết khấu tính trên Giá bán lẻ trực tiếp của Zalo (giá sàn công khai). Tổng biên kênh = chiết khấu GAPIT nhận từ VNG (15–20% theo SL GAPIT đặt); GAPIT giữ = phần còn lại sau khi cấp chiết khấu cho Đại lý/KHDN. GAPIT bán trực tiếp KHDN ở giá sàn → giữ trọn 15–20%. Mã code đã mua: không hoàn/đổi (chính sách VNG); không bán công khai dưới giá sàn. Số ví dụ theo zBusiness 12T — gói khác theo Phụ lục.',GRAY2,False,BODY)],8.7,0)])

# ============================================================ SLIDE 3: ORDER FLOW (swimlane)
s3=prs.slides.add_slide(blank)
frame(s3,'ORDER FLOW · 3/3','Order Flow trên zVAS CMS (6 bước · không GapOne)',
 [([('Phân luồng theo các bên tham gia trên hệ thống zVAS CMS (Service Platform) — quản lý code pool, đơn hàng, kích hoạt, hóa đơn.',GRAY,False,BODY)],11,0)])
LX=0.45; LW=1.45; gap=0.06
n=6; CW=(13.05-LX-LW-(n)*gap)/n
def colx(i): return LX+LW+gap+i*(CW+gap)
hy=1.66; hh=0.5
steps=['View Catalog','Chọn Bundle / Seats','Báo giá (Prorate + VAT)','Thanh toán + Hóa đơn','Subscription Active','Pool & Phân bổ seats']
# step headers
for i,st in enumerate(steps):
    rect(s3,colx(i),hy,CW,hh,BLUE2,rounded=True)
    txt(s3,colx(i),hy+0.02,CW,0.20,[([('0'+str(i+1),C('93C5FD'),True,HEAD)],11,0)],align=PP_ALIGN.CENTER)
    txt(s3,colx(i),hy+0.22,CW,0.26,[([(st,WHITE,True,BODY)],7.8,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
# actor rows
rowy=2.22; rh=0.98
actors=[('Client','Owner / Admin',CLBLUE,C('E6F1FB')),
        ('zVAS CMS','System · auto',PURPLE,C('EEEDFE')),
        ('GAPIT Sales/Admin','Internal',TEAL,C('E1F5EE')),
        ('Payment / Ngoài','ZaloPay·Bank·VNPT',AMBER,C('FAEEDA'))]
cells=[
 # Client
 [['Duyệt catalog zBusiness/zCloud/zStyle/SIM','So sánh bundle & bảng giá'],
  ['Chọn bundle + số seat','Chu kỳ tháng/năm · xem giá tạm tính'],
  ['Review quote PDF','Xác nhận prorate · approve đơn'],
  ['Chuyển khoản / ZaloPay','Upload payment proof'],
  ['Nhận email active','Truy cập Client Portal'],
  ['Vào Pool · phân bổ seat','Theo phòng ban / org tree']],
 # CMS
 [['Render catalog theo phân quyền','Pricing rules & feature matrix'],
  ['Validate seat (min/max)','Cảnh báo code hạn 6/3/2/1 thg · tạo order (draft)'],
  ['Tính prorate · VAT 10%','Generate quote PDF'],
  ['Xuất invoice PDF (VAT·MST)','Set pending_payment'],
  ['Tạo subscription · set dates','Auto-select code sắp hết hạn · draft→active'],
  ['Cộng seat vào Pool','available=purchased−allocated · log']],
 # Sales/Admin
 [['Demo catalog cho KH','Khởi tạo order hộ (sales-assist)'],
  ['Cấu hình custom bundle','Enterprise pricing · note nội bộ'],
  ['Approve quote custom','Gửi KH qua email / Zalo'],
  ['⚠ Verify thanh toán thủ công','Mark paid → CMS kích hoạt'],
  ['Cập nhật KPI pipeline','Ghi nhận MRR mới'],
  ['Theo dõi tồn kho code pool','Chuẩn bị activation']],
 # External
 [['—'],['—'],['—'],
  ['ZaloPay/Bank xử lý · transaction ID','e-Contract bởi VNPT'],
  ['—'],['—']],
]
for r,(an,asub,clr,bg) in enumerate(actors):
    yy=rowy+r*rh
    rect(s3,LX,yy,LW,rh,clr,rounded=True)
    txt(s3,LX+0.06,yy+0.10,LW-0.12,0.5,[([(an,WHITE,True,BODY)],9.5,0),([(asub,WHITE,False,BODY)],7.5,1)],anchor=MSO_ANCHOR.MIDDLE)
    for i in range(6):
        items=cells[r][i]
        empty=(items==['—'])
        rect(s3,colx(i),yy,CW,rh,C('F1F5F9') if empty else bg,line=LN,lw=0.5,rounded=True)
        if empty:
            txt(s3,colx(i),yy,CW,rh,[([('—',SLATE,True,BODY)],14,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
        else:
            bullets(s3,colx(i)+0.02,yy+0.04,CW-0.05,rh-0.06,items,clr,sz=7.3,sa=2)
# status row
sy=rowy+4*rh+0.06
txt(s3,LX,sy,LW,0.3,[([('Trạng thái',GRAY,True,BODY)],8.5,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
stat=[('—',C('F1F5F9'),GRAY),('draft',C('EEF2FF'),C('4338CA')),('draft (quoted)',C('EEF2FF'),C('4338CA')),
      ('pending_payment',C('FEF3C7'),C('92400E')),('active ✓',C('DCFCE7'),C('14532D')),('active + pool',C('DCFCE7'),C('14532D'))]
for i,(t,bg,fc) in enumerate(stat):
    rect(s3,colx(i),sy,CW,0.3,bg,rounded=True)
    txt(s3,colx(i),sy,CW,0.3,[([(t,fc,True,BODY)],8,0)],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)

import os
os.makedirs('/home/user/finalproject/output',exist_ok=True)
out='/home/user/finalproject/output/GAPIT_zVAS_Business_Roadmap_3slides.pptx'
prs.save(out); print('SAVED',out)
