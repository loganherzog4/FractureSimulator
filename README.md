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
    Measured Depth: MD
    Permeability: Ki
    Porosity: PhiT (check box for percentage)
    Water Saturation: Sw (check box for percentage)
    
   And leave Pay and Reservoir Flag headers blank.
8) Write in 4400 ft for Top Depth and 5050 ft for Base Depth and click the green arrow to continue.
9) Click the three dots and choose Directional Survey.txt in the File Explorer as your directional survey file.
10) Click the green arrow to designate column headers for the survey file.
11) Set the following column headers:
     Measured Depth: MD
     Vertical Depth: TVD
12) Click the green arrow to continue.
13) Use the default 600 acre grid area option.
14) Click the green arrow to continue.
15) Enter the following fracture parameters:
     Fracture half-length: 135 ft
     Average fracture width: 0.22 in
     Fracture height: 85 ft
     Fracture top depth: 4400 ft TVD
     Dimensionless fracture conductivity: 0.23
16) Click the green arrow to continue.
17) Choose the option to generate a fine-scale grid.
18) Choose file path and names for the grid files by clicking the three dot buttons next to the text fields. Save them wherever and as whatever you choose.
19) Designate a wellbore radius of 0.25 ft.
20) Click the green arrow to run the macro. The creation of the grid files will take around 2-3 minutes. Enjoy a cup of coffee!
