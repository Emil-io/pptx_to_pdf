# What does this do?

This script converts pptx files to pdf files. This happens through comtypes python library, which uses the "Component Object Model" in Windows to open a PowerPoint application and perform the conversion within the PowerPoint application.
For each pptx file, powerPoint is used to create a copy of the file in the pdf format. The a copy is stored in the same folder with the same name, except the .pdf ending. On top, a meta field is added to the pdf that gives it a watermark for potential recreation.

# How to use?

This script only works on Windows systems, that have PowerPoint installed.

To run the conversion, speficy the root directory where the pptx files should be converted. The path of the root directory can be set above the main method. The script then converts all pptx files in any level of subdirectories.
