# Wriggle Survey
Wriggle Survey is as-built tunnel center by taking points (3d-coordinate) around circle of tunnel. The code was computed by Line of Best Fit Method and Circle of Best Fit Method.
I created the code 2 languages which're python and vba excel.

![Cover Wriggle 8pt MWA-Model](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/assets/89971741/5bbe4814-a8e9-4ab3-9e8f-6aa5bb5ffdd0)

## Workflow
### Python
  1. Prepare Wriggle Survey data and Tunnel Axis as [Import Wriggle Survey&Tunnel Axis Data.xlsx](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/Python/Import%20Wriggle%20Survey%26Tunnel%20Axis%20Data.xlsx)
  2. Set path file, Excavation direction and Tunnel diameter design
     
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
      
  3. Run python file
### VBA
  1. Open file [VBA - Wriggle Survey Program (Best-Fit Circle 3D) Rev.07.xlsm](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/VBA/VBA%20-%20Wriggle%20Survey%20Program%20(Best-Fit%20Circle%203D)%20Rev.07.xlsm)
  2. Prepare Wriggle Survey data at Import Wriggle Data Sheet and Tunnel Axis at Import Tunnel Axis (DTA) sheet
     
     ![2024-04-26_091625](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/assets/89971741/9ed4a691-eb48-4b68-b54c-1e34a2da08d7)

     ![2024-04-26_091656](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/assets/89971741/5bbdde88-954a-45cc-a8b8-b7a608bafdd0)
     
  3. Run the code by hit the buttom of BLUE COLOR (Import Wriggle Data Sheet)

## Output
### Python
  [Export Wriggle Survey.xlsx](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/Python/Export%20Wriggle%20Survey.xlsx)
### VBA
  Wriggle Comp. sheet and Wriggle Backup sheet [VBA - Wriggle Survey Program (Best-Fit Circle 3D) Rev.07.xlsm](https://github.com/suben-mk/Wriggle-Survey-for-Metro-Line/blob/main/VBA/VBA%20-%20Wriggle%20Survey%20Program%20(Best-Fit%20Circle%203D)%20Rev.07.xlsm)
