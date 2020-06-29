Monocle Version 1.0 By Bradly Shoen
Created on: 2/28/2020

Synopsis: Creates a full screen slideshow of images from a specified location. Used for digital signage purposes.

Detailed Description: Prompts the user either through the GUI or through command line interface for image location, and slide transition delay. The slideshow will automatically update if a new file is added or a file is deleted.

HOW TO USE:
Graphical User Inteface:
1. Launch Monocle.exe
2. Specify slide location. Use the full name of the location. For example: "\\domain\temp$\Bradly\Slides" or "C:\Users\bradly.shoen\Desktop\Slides" etc
3. Specify slide duration (the time between transitions)
4. Press the Start button
Note: Any slides added/removed from the folder will be updated on the display

Command Line Interface for automation purposes:
Run .\Monocle.exe <slide location> <slide duration> <top slide location> <bottom slide location>. For example: .\Monocle.exe \\domain\temp$\Bradly\Slides 5 \\domain\temp$\Bradly\Signs\top.jpg \\domain\temp$\Bradly\Signs\bottom.jpg 


Report any bugs to Bradly Shoen at bradly.shoen@mso.umt.edu

Thank you!