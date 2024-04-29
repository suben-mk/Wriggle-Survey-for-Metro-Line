# Wriggle Survey
**Wriggle Survey** คือการหาตำแหน่งจุดศูนย์กลางอุโมงค์ (As-built Tunnel Center) ของแต่ละ Ring Tunnel Segment หลังจากการเจาะอุโมงค์เสร็จเรียบร้อยแล้ว โดยการเก็บข้อมูลจะเก็บเป็นข้อมูลพิกัด 3 มิติ (3d-coordinate) รอบๆวงกลมของอุโมงค์ตามตัวอย่างรูปด้านล่าง 
การคำนวณหาตำแหน่งจุดศูนย์กลางอุโมงค์จะคำนวณด้วยวิธี Line of Best-Fit และวิธี Circle of Best-Fit ซึ่งผลลัพธ์ที่ได้จะได้ค่าพิกัด 3 มิติและรัศมีเฉลี่ย (Average Radius) ของตำแหน่งจุดศูนย์กลางอุโมงค์ หลังจากนั้นก็จะนำผลลัพธ์พิกัด 3 มิติไปเทียบกับค่าออกแบบแนวอุโมงค์ (Tunnel Alignment) เพื่อคำนวณหาค่าเยื้องศูนย์จากแนวอุโมงค์ (Tunnel Deviation) จะได้ค่าเยื้องศูนย์ทางราบ (Horizontal Deviation) และค่าเยื้องศูนย์ทางดิ่ง (Verical Deviation)

ผู้เขียนได้เขียนโค้ดสำหรับการคำนวณ Wriggle Survey ไว้ 2 ภาษา คือภาษา Python และภาษา VBA Excel

### Wriggle Survey Points Scheme
![Cover Wriggle 8pt MWA-Model](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/assets/89971741/5bbe4814-a8e9-4ab3-9e8f-6aa5bb5ffdd0)
_(Credit : นายช่างอานนท์ วิศวกรสำรวจอุโมงค์)_

## Workflow
### Python
  **_Python libraries :_** Numpy, Pandas
  1. เตรียมข้อมูล Wriggle Survey และข้อมูลแนวอุโมงค์ (Tunnel Axis) ตาม Format [Import Wriggle Survey&Tunnel Axis Data.xlsx](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/Python/Import%20Wriggle%20Survey%26Tunnel%20Axis%20Data.xlsx)
  2. ตั้งไฟล์ Path, ทิศทางการเจาะอุโมงค์ (Excavation direction) และเส้นผ่านศูนย์กลางอุโมงค์ตามค่าออกแบบ (Tunnel diameter design)
     
     [*Wriggle_Survey_(Best-Fit_Circle_3D)_Rev06.py*](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/Python/Wriggle_Survey_(Best-Fit_Circle_3D)_Rev06.py)
      ```py
      # Path files
      Import_data_path = "Import Wriggle Survey&Tunnel Axis Data.xlsx"
      Export_data_path = "Export Wriggle Survey.xlsx"

      # Tunnel Diameter Design (m.)
      DiaDesign = 3.396

      # Excavation Direction : Select : DIRECT / REVERSE
      Direction = "DIRECT"
      ```
      
  3. รันไฟล์ Python
### VBA
  1. เปิดไฟล์ [VBA - Wriggle Survey Program (Best-Fit Circle 3D) Rev.07.xlsm](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/VBA/VBA%20-%20Wriggle%20Survey%20Program%20(Best-Fit%20Circle%203D)%20Rev.07.xlsm)
  2. เตรียมข้อมูล Wriggle Survey, Tunnel diameter design ที่ Import Wriggle Data Sheet และข้อมูล Tunnel Axis, Excavation direction ที่ Import Tunnel Axis (DTA) sheet
     
     ![2024-04-26_091625](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/assets/89971741/9ed4a691-eb48-4b68-b54c-1e34a2da08d7)

     ![2024-04-26_091656](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/assets/89971741/5bbdde88-954a-45cc-a8b8-b7a608bafdd0)
     
  3. รันโค้ดโดยการ _คลิ๊กปุ่มสีน้ำเงิน Compute! Wriggle Survey_ ที่ Import Wriggle Data Sheet

## Output
### Python
  [Export Wriggle Survey.xlsx](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/Python/Export%20Wriggle%20Survey.xlsx)
### VBA
  Wriggle Comp. sheet และ Wriggle Backup sheet [VBA - Wriggle Survey Program (Best-Fit Circle 3D) Rev.07.xlsm](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/VBA/VBA%20-%20Wriggle%20Survey%20Program%20(Best-Fit%20Circle%203D)%20Rev.07.xlsm)
### Report
  * สามารถออกแบบ Wriggle Survey Report ตามที่ต้องการ โดยทำการดึงข้อมูลการคำนวณจาก Wriggle Backup sheet ไปยังหน้า Report ที่ออกแบบไว้
  * ตัวอย่าง Wriggle Report 1 sheet [VBA - Wriggle Survey Program (Best-Fit Circle 3D) Rev.07.xlsm](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/VBA/VBA%20-%20Wriggle%20Survey%20Program%20(Best-Fit%20Circle%203D)%20Rev.07.xlsm)

  ![VBA - Wriggle Survey Program (Best-Fit Circle 3D) Rev 07](https://github.com/suben-mk/Wriggle-Survey-for-Tunnel-Project/assets/89971741/ad262c01-e154-4578-91ae-4fc17479c412)

