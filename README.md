# cTaskDialog64
cTaskDialog (TaskDialogIndirect implementation) universally compatible with VB6/VBA7/twinBASIC x86/x64

![Screenshot1](https://i.imgur.com/AQEvO9W.gif) ![Screenshot2](https://i.imgur.com/8VvddRR.gif)

![Screensot3](https://i.imgur.com/npGDQVe.jpg)


This is a version of my [cTaskDialog project](https://www.vbforums.com/showthread.php?777021-VB6-TaskDialogIndirect-Complete-class-implementation-of-Vista-Task-Dialogs) that uses conditional compilation to support both VB6/VBA6 and twinBASIC/VBA7 in either x86 or x64. See that page for complete project description and numerous more pictures and examples.

Since people have asked about using this in VBA, it goes back to the earlier method of using a module to help with subclassing, as the self-subclass code in the last VB6 version only works in VB6, and while twinBASIC supports AddressOf on class members, VBA7 does not. Note that there's a bug in the self-sub version that changes the way multiple pages are handled, sending all events through the first page class. So if you use multiple paged Task Dialogs, you'll now need to relocate events for the other pages to their own event Subs (the Demo does this with it's multi-page Demos).

### LongPtr in VB6/VBA?
You'll need to add LongPtr support to use this codebase in VB6/VBA6. [This thread](https://www.vbforums.com/showthread.php?898078-Typelib-to-add-LongPtr-type-to-VB6-for-universal-codebases) provides two methods: via a typelib with an alias, or via an enum. For simplicity this project currently uses the Enum method (defined in modHelper.bas).

### Requirements
This project will work with VB6, VBA6, VBA7 x86/x64, and twinBASIC x86/x64,. Regardless of the project type, you'll need Common Controls 6.0 enabled via manifest.

### Source Code
The class itself can be found in the Export\Sources folder, along with the exported twinBASIC Demo form. The Export\Resources folder has a manifest for comtl6 if you need it.

To use this outside of twinBASIC, you'll need cTaskDialog.cls and modTDHelper.bas from the Export\Sources folder. Both must be added to a project.
