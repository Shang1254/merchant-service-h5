# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

doc = Document()
style = doc.styles['Normal']
style.font.name = '微软雅黑'
style.font.size = Pt(10.5)
style._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

def add_heading(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = '微软雅黑'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
    return h

def add_para(text, bold=False, size=10.5):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = '微软雅黑'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.size = Pt(size)
    run.bold = bold
    return p

def add_table(headers, rows, col_widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ''
        p = hdr[i].paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.name = '微软雅黑'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        run.font.size = Pt(9.5)
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ''
            p = cells[i].paragraphs[0]
            run = p.add_run(str(val))
            run.font.name = '微软雅黑'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
            run.font.size = Pt(9)
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)
    return table

# 标题
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('程序部署费用明细 & 定价审查报告')
run.bold = True
run.font.size = Pt(18)
run.font.name = '微软雅黑'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
run.font.color.rgb = RGBColor(0x2a, 0x1f, 0x5e)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub.add_run('基于阿里云/腾讯云/移动云实际价格校准')
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
run.font.name = '微软雅黑'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

doc.add_paragraph()

# ===== 一、服务器/域名/认证费用明细 =====
add_heading('一、第三方费用明细（客户自己承担部分）', level=1)

add_para('以下费用均为客户自行购买，我们不代收、不加价，只提供指导：', size=10)

doc.add_paragraph()
add_table(
    ['费用项', '阿里云', '腾讯云', '移动云', '备注'],
    [
        ['轻量服务器（2核2G）', '99-199元/年\n（活动价）', '70-150元/年\n（活动价）', '约300-500元/年\n（企业级）', '入门级够用，\n活动价很便宜'],
        ['域名（.com）', '约60元/年', '约60元/年', '约60元/年', '.cn域名更便宜，\n约30-50元/年'],
        ['微信小程序认证费', '—', '—', '—', '企业300元/年\n个体工商户30元/年\n个人30元/年'],
        ['SSL证书', '免费版够用', '免费版够用', '免费版够用', 'Let\'s Encrypt 免费，\n不需要买企业版'],
        ['ICP备案', '免费', '免费', '免费', '需要客户自己配合，\n约7-15个工作日'],
        ['合计（企业主体）', '约459-559元/年', '约420-510元/年', '约690-890元/年', '首年约500元左右'],
        ['合计（个人/个体户）', '约189-289元/年', '约150-240元/年', '约380-580元/年', '首年约200-300元'],
    ],
    col_widths=[3, 2.8, 2.8, 2.8, 3.5]
)

doc.add_paragraph()
add_para('重要提醒：', bold=True)
add_para('• 以上是云厂商官方活动价，平时价格会贵一些，但首年活动价确实很便宜。')
add_para('• 服务器第二年续费会涨价，大概是首年价格的2-3倍，要跟客户说清楚。')
add_para('• 微信认证费是每年都要交的，不是一次性的。')
add_para('• 我们不代收这些费用，客户自己在云厂商官网购买，我们只提供操作指导。')

doc.add_page_break()

# ===== 二、自己部署 vs 找人部署对比 =====
add_heading('二、自己部署 vs 找人部署对比', level=1)

add_table(
    ['对比项', '自己折腾部署', '找我们做', '差价/时间'],
    [
        ['开发费用', '0元（自己写代码）\n或买模板99-599元', '展示型3999元\n预约型5999元', '模板便宜但功能有限，\n定制贵但完全匹配需求'],
        ['服务器+域名+认证', '约500元/年', '约500元/年（客户自付）', '一样，都是客户自己出'],
        ['部署时间', '1-3天（新手）', '3-7个工作日', '我们快，因为有经验，\n不会踩坑'],
        ['技术门槛', '需要懂Linux、\nNginx、SSL配置、\n域名解析、备案', '不用懂，\n我们全包', '客户只需要提供资料，\n剩下的我们搞定'],
        ['踩坑风险', '可能踩坑：\n备案被打回、\nSSL配置错误、\n域名解析不对、\n支付接口调不通', '我们有经验，\n避开常见坑', '自己踩坑可能多花\n几天甚至几周时间'],
        ['后续维护', '出问题自己排查，\n网上搜解决方案', '首月免费陪跑，\n后续399元/月', '我们负责，\n出问题找我们'],
        ['总成本（首年）', '模板版：约600-1100元\n（服务器+模板+时间）', '展示型：约4500元\n（开发+服务器）', '定制版贵一些，\n但功能和服务值这个价'],
        ['适合谁', '有技术基础、\n想省钱、\n有时间折腾的人', '没技术基础、\n想省心、\n有正经生意的商家', '—'],
    ],
    col_widths=[2.5, 4, 4, 4.5]
)

doc.add_paragraph()
add_para('一句话总结：', bold=True)
add_para('自己折腾 = 省几千块开发费，但要花1-3天时间学技术，还可能踩坑；')
add_para('找我们做 = 多花几千块，但你只需要管生意，技术的事我们搞定。')

doc.add_page_break()

# ===== 三、定价表重新审查 =====
add_heading('三、定价表重新审查', level=1)

add_heading('1. 程序定制类定价审查', level=2)

add_table(
    ['类型', '原定价', '市场参考价', '审查结论', '建议调整'],
    [
        ['展示型小程序', '3999元', '3000-5000元', '✅ 合理，在市场区间内', '维持3999元'],
        ['预约/表单型', '5999元', '5000-10000元', '✅ 合理，取市场低位', '维持5999元'],
        ['带支付/商城型', '12999元', '1-3万元', '✅ 合理，取市场低位', '维持12999元'],
        ['部署陪跑（首月）', '免费', '—', '✅ 合理，首月免费是行业惯例', '维持免费'],
        ['后续维护', '399元/月\n或999元/次', '300-500元/月\n或800-1500元/次', '✅ 合理，在市场区间内', '维持399元/月'],
    ],
    col_widths=[2.5, 2, 2.5, 3.5, 3]
)

doc.add_paragraph()
add_heading('2. 海报/视频类定价审查', level=2)

add_table(
    ['类型', '原定价', '市场参考价', '审查结论', '建议调整'],
    [
        ['AI快出海报', '69元/张', '39-99元', '✅ 合理，取中间偏下', '维持69元'],
        ['人工精修海报', '199元/张', '新手50-150\n熟手200-500', '✅ 合理，介于新手和熟手之间', '维持199元'],
        ['月度代运营（30张）', '1599元/月', '—', '✅ 合理，30×69=2070，\n套餐价打77折', '维持1599元'],
        ['30秒短视频', '399元/条', '88-168元（低端）\n专业团队更高', '⚠️ 略偏高，\n但我们是AI+人工剪辑', '可降到299元\n或保持399元（质量更好）'],
        ['60-90秒宣传片', '2999元/条', '专业团队8800-16800', '✅ 合理，我们是小团队AI辅助，\n取低位', '维持2999元'],
    ],
    col_widths=[2.5, 2, 3, 3.5, 3]
)

doc.add_paragraph()
add_heading('3. 引流服务类定价审查', level=2)

add_table(
    ['类型', '原定价', '市场参考价', '审查结论', '建议调整'],
    [
        ['单次获客诊断', '499元/次', '千元以内', '✅ 合理，取中间偏低', '维持499元'],
        ['月度引流方案', '1999元/月', '2500-3500元/月', '✅ 合理，取市场低位', '维持1999元'],
        ['季度陪跑', '5499元/季', '—', '✅ 合理，相当于1833元/月', '维持5499元'],
    ],
    col_widths=[2.5, 2, 3, 3.5, 3]
)

doc.add_paragraph()
add_heading('4. 套餐定价审查', level=2)

add_table(
    ['套餐', '原定价', '成本估算', '审查结论', '建议调整'],
    [
        ['轻量试水包', '499元/月', 'AI工具成本约50-100元', '✅ 毛利约65%-75%，\n引流款没问题', '维持499元'],
        ['标准增长包', '2499元/月', '海报+视频+运营约800-1000元', '✅ 毛利约50%-60%，\n主力利润款', '维持2499元'],
        ['年度陪跑包', '29999元/年', '约12000-15000元', '✅ 毛利约50%-60%，\n长期稳定收入', '维持29999元'],
    ],
    col_widths=[2.5, 2, 3, 3.5, 3]
)

doc.add_page_break()

# ===== 四、定价风险提醒 =====
add_heading('四、定价风险提醒（哪些容易亏）', level=1)

add_table(
    ['风险点', '为什么容易亏', '怎么避免'],
    [
        ['无限改稿', '客户改一次又一次，\n时间成本翻倍', '明确写死修改次数：\n海报2次，视频2次，\n超出按50元/次收费'],
        ['程序改需求', '客户做到一半加需求，\n工作量翻倍', '需求确认书签字后，\n改需求按200-300元/页面/次收费'],
        ['部署环境不明确', '客户服务器买错了，\n或者备案没做好，\n卡在部署环节', '提前跟客户说清楚：\n服务器/域名/认证费客户出，\n我们只负责技术指导和部署'],
        ['服务器续费涨价', '客户首年觉得便宜，\n第二年续费涨2-3倍，\n觉得我们坑他', '提前跟客户说清楚：\n首年活动价便宜，\n续费是云厂商收的，\n不是我们定的'],
        ['免费维护期', '客户以为终身免费，\n实际上我们只包首月', '明确写进合同：\n首月免费陪跑，\n后续维护399元/月或999元/次'],
    ],
    col_widths=[2.5, 5, 6.5]
)

doc.add_paragraph()
add_para('最容易亏的3个坑：', bold=True)
add_para('1. 程序改需求——做到一半客户说"能不能再加个XX功能"，不加客户不高兴，加了工作量翻倍。')
add_para('2. 无限改稿——客户改了五六次还不满意，时间全耗在这一个客户身上了。')
add_para('3. 部署环节——客户自己买了服务器但配置不对，或者备案没下来，卡在那干等，你还得陪着折腾。')

# 保存
output_path = r'C:\Users\商子轩\Desktop\海报\定价方案-含部署费用审查版.docx'
doc.save(output_path)
print(f'文档已保存: {output_path}')
