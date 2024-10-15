# Kích hoạt Office 2024 Visio Vĩnh viễn
- ![image](https://github.com/user-attachments/assets/892ab962-1334-4126-9b74-42be48da0f04)
- ![image](https://github.com/BsNgChiThanh/Lich-phong-kham/assets/82578024/d575f08f-29b1-4848-83b0-fb5e88dcb50c)
- [Source cài đặt:](https://gravesoft.dev/office_c2r_links)
  - [VisioPro2024Retail](https://raw.githubusercontent.com/BsNgChiThanh/Office2024Visio/IMP/VisioPro2024Retail.exe)
  - [VisioStd2024Retail](https://raw.githubusercontent.com/BsNgChiThanh/Office2024Visio/IMP/VisioStd2024Retail.exe)

Có nhiều cách kích hoạt, xong tôi chỉ ra 2 cách kích hoạt điển hình nhất:

## CÁCH 1:
- [Dùng MAS Tools](https://github.com/BsNgChiThanh/MAS-TOOL)
- Các bạn chạy **PowerShell** bằng quyền **Run as Administrator**, rồi dán câu lệnh sau đâu vào:
  ```php
  irm https://raw.githubusercontent.com/BsNgChiThanh/MAS-TOOL/IMP/MAS.ps1 | iex
  ```
- Làm theo chỉ dẫn của cửa sổ hiện lên.
- Done!

## CÁCH 2
- ĐẦU TIÊN KÍCH HOẠT BẰNG KEY KMS 180 NGÀY:
  - Chạy **cmd** bằng quyền **Run as Administrator** rồi dán đoạn lệnh này vào, nhấn enter:
  ```php
  for %a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%a")
  if exist "%ProgramFiles% (x86)\Microsoft Office\Office1%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%a"))
  (cscript //nologo ospp.vbs /setprt:1688 >nul || goto wshdisabled)&cscript //nologo ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms" >nul
  (for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
  (for /f %%x in ('dir /b ..\root\Licenses16\visioprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
  (for /f %%x in ('dir /b ..\root\Licenses16\visiopro2024vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
  cscript ospp.vbs /setprt:1688
  cscript ospp.vbs /inpkey:YW66X-NH62M-G6YFP-B7KCT-WXGKQ
  cscript ospp.vbs /inpkey:HGRBX-N68QF-6DY8J-CGX4W-XW7KP
  cscript ospp.vbs /sethst:107.175.77.7
  cscript ospp.vbs /sethst:172.16.0.2
  cscript ospp.vbs /act
  ```
  - ![image](https://github.com/user-attachments/assets/387c35e4-30f8-43ac-9141-fc633f4ede9b)
- SAU ĐÓ TẠO TÁC VỤ GIA HẠN OFFICE HÀNG TUẦN:
  - Chạy Windows PowerShell bằng quyền Run as Administrator rồi dán đoạn lệnh này vào, nhấn enter:
    
  ```PHP
  irm https://raw.githubusercontent.com/BsNgChiThanh/Kich-hoat-Office/KichHoatOffice/GiaHanKichHoat.ps1 | iex
  ```
  - ![image](https://github.com/user-attachments/assets/c61d847b-f874-4549-92af-f49985044f7e)
- DONE!
