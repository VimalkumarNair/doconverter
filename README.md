# Doconverter

With the help of this project, you can esaily convert your PDF and Office files to images by using command lines.

## Warning
Doesn't work very stable, yet.

# Requirements

Firstly requirse app [ghostscript](http://http://www.ghostscript.com/) for PDF files.

Again you must have Office programs installed on your system. * You might get an error if you try to convert a file which was created in a newer version of Office than the one you have on your system.

# How to use

Just type `doconverter.exe file_toconvert.doc|xls|ppt|pdf` in your command line. System will create a file called **export** where your file is located and produce the images there.

# Configs

You can make configurations with the help of **doconverter.exe.config** file.

1. **Ghostscript** file location settings
2. Image resolution settings for Word and PowerPoint
3. Resolution settings for PDF files ( Ex : 100,200,...)
4. Excel works in limits which are already set. This will be editable in settings in coming versions.
