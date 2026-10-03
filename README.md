# HAK_FSC
Created from gas-tools extension

## Chạy test

Test chạy bằng Node.js trên Google Sheets/Drive giả lập, không đụng vào dữ liệu thật.
Thư mục `tests/` chỉ dùng để test — không copy vào Apps Script.

```bash
node tests/server.test.cjs   # logic Code.gs (ghi/đọc dữ liệu, validate, cache, khóa sinh mã)
node tests/ui.test.cjs       # giao diện Index.html trong Chromium headless (cần: npm i -g playwright)
```

Nên chạy cả hai sau mỗi lần sửa `Code.gs` hoặc `Index.html`, trước khi Deploy phiên bản mới.
