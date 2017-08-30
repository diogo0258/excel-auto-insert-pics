
Sometimes I have to write reports in Excel that require some pictures. These pictures should be resized to fit specific cells.
There's usually a template .xlsx that I need to follow, and the pictures I get usually have the original filenames that the camera gave them (like 20170304-131617.jpg).

This repo has some tools that make this task a little faster. Here's the workflow:

- First I make a list.txt file with the names of the pictures that should go in each specific cell. For example, if a picture of a certain appliance XYZ should go into cell A14 in Sheet 1, list.txt would contain the line
    > 1 A14; XYZ

- The Ahk script loads that list, and sits waiting for a F3 press on either a Windows Image Viewer or Explorer (with a file selected) window. When that happens, it shows a searchable listview with the entries in list.txt, with both the names and the destination cells.
- You can show only entries that have not been fulfilled by searching for ':'.

- When a listview entry is selected, the file is copied to the respective "renamed-pics\%SheetNum% %Cell%"
- Optionally, if you press Shift+Enter to select it, the destination file will be opened in Irfanview, where you can easily crop / rotate it.

- After all the pictures I want have been identified, I call the Excel macro AutoAddPicsToCellsBasedOnFileNames on the template. The pics in "renamed-pics\" will be inserted and fit to their respective cells.

- If I need to move things around in Excel, Ctrl+E will:
    - with a cell selected, show a dialog where you can select an image file. The file will then be inserted and fit to the selected cell.
    - with an image selected, it will be fit to the cell behind its topleft corner.

The xlsm file that contains the macros has an example template that you can test these on. There are also a sample list.txt and some pictures in "original-pics\" for testing the whole thing. And some images in demo-imgs to show what it looks like.

Cheers.
