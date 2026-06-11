import sys, time
sys.stdout.reconfigure(encoding='utf-8')
from playwright.sync_api import sync_playwright

results = []

def ok(label): results.append(('OK', label))
def fail(label, reason=''):
    results.append(('FAIL', label + (' — ' + reason if reason else '')))

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={'width':1400,'height':900})
    pg.goto('http://localhost:8000')
    pg.wait_for_timeout(1500)

    # ── 1. CHAT TAB ──
    pg.screenshot(path='test_01_chat.png')
    welcome = pg.locator('.msg.assistant').first
    if welcome.is_visible():
        ok('Chat: welcome message hiển thị')
    else:
        fail('Chat: welcome message')

    txt = welcome.inner_text()
    if 'Độc Media' in txt:
        ok('Chat: welcome có dấu tiếng Việt')
    else:
        fail('Chat: welcome thiếu dấu - ' + txt[:60])

    sidebar = pg.locator('#sidebar')
    if sidebar.is_visible():
        ok('Chat: sidebar lịch sử hiển thị')
    else:
        fail('Chat: sidebar')

    inp = pg.locator('#user-input')
    inp.fill('2+2 bằng bao nhiêu? Trả lời 1 dòng.')
    pg.locator('#send-btn').click()
    pg.wait_for_timeout(6000)
    msgs = pg.locator('.msg.assistant').all()
    if len(msgs) >= 2:
        ok('Chat: nhận phản hồi AI (' + str(len(msgs)) + ' msgs)')
    else:
        fail('Chat: không có phản hồi AI - chỉ có ' + str(len(msgs)) + ' msgs')
    pg.screenshot(path='test_02_chat_reply.png')

    pg.locator('.btn-new').click()
    pg.wait_for_timeout(800)
    new_welcome = pg.locator('.msg.assistant').first
    if new_welcome.is_visible():
        ok('Chat: tạo chat mới hoạt động')
    else:
        fail('Chat: new chat không hoạt động')

    # ── 2. CONTENT STUDIO ──
    pg.locator('.tab-btn').nth(1).click()
    pg.wait_for_timeout(600)
    pg.screenshot(path='test_03_content.png')
    if pg.locator('#tab-content').is_visible():
        ok('Content Studio: tab mở được')
    else:
        fail('Content Studio: tab không mở')

    labels_html = pg.locator('#tab-content').inner_html()
    if ('Sản phẩm' in labels_html or 'Mô tả' in labels_html or 'client' in labels_html.lower()):
        ok('Content Studio: labels có dấu tiếng Việt')
    else:
        fail('Content Studio: labels thiếu dấu')

    pg.locator('#cs-product').fill('Sofa da cao cấp')
    pg.locator('#cs-gen-btn').click()
    pg.wait_for_timeout(9000)
    output = pg.locator('#cs-output')
    txt_out = output.inner_text().strip()
    if len(txt_out) > 50:
        ok('Content Studio: generate OK (' + str(len(txt_out)) + ' chars)')
    else:
        fail('Content Studio: generate không ra nội dung - ' + txt_out[:80])
    pg.screenshot(path='test_04_content_gen.png')

    # ── 3. VIDEO SCRIPT ──
    pg.locator('.tab-btn').nth(2).click()
    pg.wait_for_timeout(600)
    pg.screenshot(path='test_05_script.png')
    if pg.locator('#tab-script').is_visible():
        ok('Video Script: tab mở được')
    else:
        fail('Video Script: tab không mở')

    script_html = pg.locator('#tab-script').inner_html()
    if 'Sản phẩm' in script_html or 'Thời lượng' in script_html:
        ok('Video Script: labels có dấu')
    else:
        fail('Video Script: labels thiếu dấu')

    pg.locator('#sc-product').fill('Căn hộ mẫu 3 phòng ngủ')
    pg.locator('#sc-gen-btn').click()
    pg.wait_for_timeout(9000)
    out2 = pg.locator('#sc-output')
    txt2 = out2.inner_text().strip()
    if len(txt2) > 50:
        ok('Video Script: generate kịch bản OK (' + str(len(txt2)) + ' chars)')
    else:
        fail('Video Script: generate không ra kịch bản - ' + txt2[:80])
    pg.screenshot(path='test_06_script_gen.png')

    # ── 4. TẠO HÌNH ──
    pg.locator('.tab-btn').nth(3).click()
    pg.wait_for_timeout(600)
    pg.screenshot(path='test_07_image.png')
    if pg.locator('#tab-image').is_visible():
        ok('Tạo Hình: tab mở được')
    else:
        fail('Tạo Hình: tab không mở')

    img_html = pg.locator('#tab-image').inner_html()
    if 'Mô tả' in img_html or 'Phong cách' in img_html or 'Tạo ảnh' in img_html:
        ok('Tạo Hình: labels có dấu')
    else:
        fail('Tạo Hình: labels thiếu dấu - ' + img_html[:120])

    gen_btn = pg.locator('#gen-btn')
    if gen_btn.is_visible():
        ok('Tạo Hình: nút Tạo ảnh hiển thị')
    else:
        fail('Tạo Hình: nút Tạo ảnh không hiển thị')

    # ── 5. FACEBOOK ADS ──
    pg.locator('.tab-btn').nth(4).click()
    pg.wait_for_timeout(1000)
    pg.screenshot(path='test_08_ads.png')
    if pg.locator('#tab-ads').is_visible():
        ok('Facebook Ads: tab mở được')
    else:
        fail('Facebook Ads: tab không mở')

    sel = pg.locator('#ads-account-sel')
    pg.wait_for_timeout(3000)
    opts = sel.locator('option').all()
    if len(opts) > 1:
        ok('Facebook Ads: load tài khoản OK (' + str(len(opts)) + ' options)')
    else:
        fail('Facebook Ads: không load được tài khoản - ' + str(len(opts)) + ' options')
    pg.screenshot(path='test_09_ads_accounts.png')

    btn = pg.locator('#refresh-ads-btn')
    btn.click()
    pg.wait_for_timeout(12000)
    pg.screenshot(path='test_10_ads_data.png')
    body_html = pg.locator('#ads-body').inner_html()
    kpi_count = body_html.count('kpi-card')
    if kpi_count == 8:
        ok('Facebook Ads: đủ 8 KPI cards (grid 4x2)')
    elif kpi_count > 0:
        fail('Facebook Ads: KPI count = ' + str(kpi_count) + ' (cần 8)')
    else:
        fail('Facebook Ads: không có KPI cards — ' + body_html[:200])

    if 'Tổng chi tiêu' in body_html:
        ok('Facebook Ads: KPI "Tổng chi tiêu" có dấu')
    else:
        fail('Facebook Ads: KPI thiếu dấu')

    if 'Số tin nhắn' in body_html or 'Chi phí/Tin nhắn' in body_html:
        ok('Facebook Ads: metric "Tin nhắn" mới có mặt')
    else:
        fail('Facebook Ads: thiếu metric Tin nhắn')

    if 'Số KHTN' in body_html or 'Chi phí/KHTN' in body_html:
        ok('Facebook Ads: metric "KHTN" mới có mặt')
    else:
        fail('Facebook Ads: thiếu metric KHTN')

    table = pg.locator('.camp-table table')
    if table.is_visible():
        ok('Facebook Ads: bảng chiến dịch hiển thị')
        header_html = pg.locator('.camp-table thead').inner_html()
        if 'Chiến dịch' in header_html:
            ok('Facebook Ads: header bảng có dấu')
        else:
            fail('Facebook Ads: header bảng thiếu dấu')
        if 'Tin nhắn' in header_html and 'KHTN' in header_html:
            ok('Facebook Ads: bảng có cột Tin nhắn + KHTN')
        else:
            fail('Facebook Ads: bảng thiếu cột mới - ' + header_html)
    else:
        fail('Facebook Ads: không có bảng chiến dịch')

    # AI analysis — wait longer for streaming
    pg.wait_for_timeout(40000)
    ai_div = pg.locator('#ai-analysis')
    if ai_div.count() > 0:
        ai_txt = ai_div.inner_text().strip()
        if len(ai_txt) > 50:
            ok('Facebook Ads: AI phân tích sinh ra nội dung (' + str(len(ai_txt)) + ' chars)')
        else:
            fail('Facebook Ads: AI phân tích chưa ra nội dung - ' + ai_txt[:80])
    pg.screenshot(path='test_11_ads_ai.png')

    # ── 6. INTAKE CLIENT ──
    pg.locator('.tab-btn').nth(5).click()
    pg.wait_for_timeout(600)
    pg.screenshot(path='test_12_intake.png')
    if pg.locator('#tab-intake').is_visible():
        ok('Intake Client: tab mở được')
    else:
        fail('Intake Client: tab không mở')

    intake_html = pg.locator('#tab-intake').inner_html()
    if 'Tên doanh nghiệp' in intake_html or 'Lĩnh vực' in intake_html:
        ok('Intake Client: labels có dấu')
    else:
        fail('Intake Client: labels thiếu dấu')

    pg.locator('#in-ten').fill('RECO Design')
    pg.locator('#in-sp').fill('Thiết kế nội thất cao cấp')
    pg.locator('#in-usp').fill('Phong cách Á Đông hiện đại')
    pg.locator('#intake-submit-btn').click()
    pg.wait_for_timeout(80000)
    intake_out = pg.locator('#intake-output-body').inner_text().strip()
    if len(intake_out) > 100:
        ok('Intake Client: generate brief OK (' + str(len(intake_out)) + ' chars)')
    else:
        fail('Intake Client: generate không ra brief - ' + intake_out[:80])
    pg.screenshot(path='test_13_intake_gen.png')

    # Intake to Quote button
    to_quote = pg.locator('#intake-to-quote-btn')
    if not to_quote.is_disabled():
        ok('Intake Client: nút "Chuyển sang Báo Giá" enabled')
    else:
        fail('Intake Client: nút Chuyển sang Báo Giá vẫn disabled')

    # ── 7. BÁO GIÁ ──
    pg.locator('.tab-btn').nth(6).click()
    pg.wait_for_timeout(600)
    pg.screenshot(path='test_14_quote.png')
    if pg.locator('#tab-quote').is_visible():
        ok('Báo Giá: tab mở được')
    else:
        fail('Báo Giá: tab không mở')

    quote_html = pg.locator('#tab-quote').inner_html()
    if 'Tên doanh nghiệp' in quote_html and 'Lĩnh vực' in quote_html:
        ok('Báo Giá: labels có dấu')
    else:
        fail('Báo Giá: labels thiếu dấu')

    if 'Nội thất / Thiết kế' in quote_html:
        ok('Báo Giá: select options có dấu')
    else:
        fail('Báo Giá: select options thiếu dấu')

    pg.locator('#q-client').fill('RECO Design & Build')
    pg.locator('#q-budget').fill('15000000')
    pg.locator('#q-submit-btn').click()
    pg.wait_for_timeout(70000)
    quote_out = pg.locator('#quote-output-body').inner_text().strip()
    if len(quote_out) > 100:
        ok('Báo Giá: generate báo giá OK (' + str(len(quote_out)) + ' chars)')
    else:
        fail('Báo Giá: generate không ra báo giá - ' + quote_out[:80])
    pg.screenshot(path='test_15_quote_gen.png')

    exp_btn = pg.locator('#q-export-btn')
    if not exp_btn.is_disabled():
        ok('Báo Giá: nút Xuất Excel enabled sau generate')
    else:
        fail('Báo Giá: nút Xuất Excel vẫn disabled')

    b.close()

print()
print('=' * 60)
passed = sum(1 for r in results if r[0] == 'OK')
failed = sum(1 for r in results if r[0] == 'FAIL')
for status, label in results:
    marker = 'OK ' if status == 'OK' else 'FAIL'
    print(marker + ': ' + label)
print('=' * 60)
print('TỔNG KẾT: ' + str(passed) + '/' + str(passed + failed) + ' PASS | ' + str(failed) + ' FAIL')
