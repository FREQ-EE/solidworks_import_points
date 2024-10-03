# SolidWorks Import Points to Sketch

## Overview
This script imports points from a CSV file and creates corresponding points in a SolidWorks 3D sketch.

## Requirements
- SolidWorks with an open part document.
- CSV file containing coordinates in the format: `"Bone Name", "X", "Y", "Z"`.

## Usage
1. Open the Visual Basic for Applications (VBA) editor in SolidWorks (`Tools > Macro > New`).
2. Copy and paste the code provided in `SW_import_points_to_sketch.vb` into the new macro.
3. Update the `filePath` variable in the script to point to your CSV file containing the coordinates.
4. Save the macro as an `.swp` file.
5. Run the macro (`Tools > Macro > Run`) in SolidWorks while the part document is active.
6. The script will create a 3D sketch and add points based on the CSV coordinates.

## Notes
- Ensure that the CSV file is correctly formatted and that the part document is active.
- You must update the `filePath` in the code to specify the correct location of your CSV file.
