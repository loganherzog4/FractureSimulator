# FractureSimulator
Fracture Simulator project built for my 2017 internship as a Petroleum Engineer.


This project was developed without much regard to code maintainability - it was meant to be used entirely internally for our small team of Reservoir Engineers. Therefore, if you notice that I declared variables at the top of subs in massive walls, and that the majority of algorithm mathematics is not commented over...sorry!


I have developed much better coding practices since I made this project; however, it still does exactly what it was designed to do, and was a project that got me hired on for a second internship.


In order to demo the functionality, follow these steps:


1) Download Fracture Simulator.xlsm, Wood.las, and Directional Survey.txt from this repository
2) Open Fracture Simulator.xlsm and enable macros
3) Click the top button, "Run Simulation"
4) A form should pop up. Click the three dots button next to the text field to open a File Explorer window
5) Navigate to where you saved WOOD.las and double click that file to select it
6) Click the green arrow to designate column headers for the .LAS file
7) Set the following column headers:
8) Measured Depth: MD
9) Permeability: Ki
10) Porosity: PhiT (check box for percentage)
11) Water Saturation: Sw (check box for percentage)
12) And leave Pay and Reservoir Flag headers blank.
13) Write in 4400 ft for Top Depth and 5050 ft for Base Depth and click the green arrow to continue.
14) Click the three dots and choose Directional Survey.txt in the File Explorer as your directional survey file.
15) Click the green arrow to designate column headers for the survey file.
16) Set the following column headers:
17) Measured Depth: MD
18) Vertical Depth: TVD
19) Click the green arrow to continue.
20) Use the default 600 acre grid area option.
21) Click the green arrow to continue.
22) Enter the following fracture parameters:
23) Fracture half-length: 135 ft
24) Average fracture width: 0.22 in
25) Fracture height: 85 ft
26) Fracture top depth: 4400 ft TVD
27) Dimensionless fracture conductivity: 0.23
28) Click the green arrow to continue.
29) Choose the option to generate a fine-scale grid.
30) Choose file path and names for the grid files by clicking the three dot buttons next to the text fields. Save them wherever and as whatever you choose.
31) Designate a wellbore radius of 0.25 ft.
32) Click the green arrow to run the macro. The creation of the grid files will take around 2-3 minutes. Enjoy a cup of coffee!
