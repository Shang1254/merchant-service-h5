# -*- coding: utf-8 -*-
import re

with open('C:\\Users\\商子轩\\Desktop\\海报\\智能客服H5.html', 'r', encoding='utf-8') as f:
    html = f.read()

with open('C:\\Users\\商子轩\\Desktop\\海报\\qrcode_base64.txt', 'r') as f:
    base64 = f.read().replace('\n', '').replace('\r', '')

# 替换占位框为图片
html = re.sub(
    r'<div class="qr-placeholder">.*?</div>',
    f'<img src="data:image/png;base64,{base64}" style="width: 160px; height: 160px; border-radius: 8px;">',
    html,
    flags=re.DOTALL
)

with open('C:\\Users\\商子轩\\Desktop\\海报\\智能客服H5.html', 'w', encoding='utf-8') as f:
    f.write(html)

print('QR code embedded successfully')
