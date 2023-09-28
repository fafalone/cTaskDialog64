# cTaskDialog64
cTaskDialog (TaskDialogIndirect implementation) universally compatible with VB6/VBA7/twinBASIC x86/x64\
(UPDATE: NOW FULLY WORKING IN VBA64!)

![Screenshot1](https://i.imgur.com/AQEvO9W.gif) ![Screenshot2](https://i.imgur.com/8VvddRR.gif)

![Screensot3](https://i.imgur.com/npGDQVe.jpg)


This is a version of my [cTaskDialog project](https://www.vbforums.com/showthread.php?777021-VB6-TaskDialogIndirect-Complete-class-implementation-of-Vista-Task-Dialogs) that uses conditional compilation to support both VB6/VBA6 and twinBASIC/VBA7 in either x86 or x64. See that page for complete project description and numerous more pictures and examples. The demo is provided as a twinBASIC project, but you can get just the cTaskDialog.cls and modTDHelper.bas for VB6/VBA in Export\Sources. The demos are in Form1.frm.twin there too, but you can use the demos from the main project thread too.

Since people have asked about using this in VBA, it goes back to the earlier method of using a module to help with subclassing, as the self-subclass code in the last VB6 version only works in VB6, and while twinBASIC supports AddressOf on class members, VBA7 does not. Note that there's a bug in the self-sub version that changes the way multiple pages are handled, sending all events through the first page class. So if you use multiple paged Task Dialogs, you'll now need to relocate events for the other pages to their own event Subs (the Demo does this with it's multi-page Demos).

### Updates
(28 Sep 2023) Courtesy of brilliant programmer The trick, a fix has finally been identified for use of the callbacks in VBA64.

(23 Nov 2022) Updated to version 1.2.4. Fixed improper VarPtr calls in VBA7x64 routines.

(26 Oct 2022) Updated to version 1.2.3. Fixed positioning bug on some systems. This occured when system visual effects were disabled, which changed the size immediately when the class expected to be able to compare against the old size. Thanks to Wayne Phillips for figuring this out!

(24 Oct 2022) Updated to version 1.2.2. Fixed the issues with the logo, height after expando closed, and font sizes. Positioning issue is proving difficult so might take a little longer; wanted to fix what I could now. The Logo Demo in the twinBASIC project now shows loading a larger logo image based on current DPI (queried from the control, you don't need to implement it), and the Init routine now sets a default date/time that's returned if the datetime is unchecked (it would previously return a date in 1999... seemed wrong. But you shouldn't consider it valid if not checked, when checkboxes are enabled).

### LongPtr in VB6/VBA?
You'll need to add LongPtr support to use this codebase in VB6/VBA6. [This thread](https://www.vbforums.com/showthread.php?898078-Typelib-to-add-LongPtr-type-to-VB6-for-universal-codebases) provides two methods: via a typelib with an alias, or via an enum. For simplicity this project currently uses the Enum method (defined in modHelper.bas).

### Requirements
This project will work with VB6, VBA6, VBA7 x86/x64, and twinBASIC x86/x64,. Regardless of the project type, you'll need Common Controls 6.0 enabled via manifest.

For twinBASIC, you'll need at least Beta 108 (when the PackingAlignment option was added), but at least 154 is recommended due to earlier versions sometimes producing an erroneous error message that GetSystemImageList is ambiguous. If you do use it with an earlier version, restarting the compiler will get rid of that error. [twinBASIC Releases](https://github.com/twinbasic/twinbasic/releases)

### Source Code
The class itself can be found in the Export\Sources folder, along with the exported twinBASIC Demo form. The Export\Resources folder has a manifest for comtl6 if you need it.

To use this outside of twinBASIC, you'll need cTaskDialog.cls and modTDHelper.bas from the Export\Sources folder. Both must be added to a project.

### Customizations
This class is more than just a straight implementation of the native features (though it supports all of those and can be used with just a few lines for very simply dialogs), it also features custom flags that add additional control types: TextBox, ComboBox (with images), Date/Time, and Slider, all of which can be positioned in either the top region, by the buttons, or in the footer, and can be mixed and matched with eachother and all the built in features. There's also an option to add a logo image in the top right and a few other places. Follow the link to the VBForums thread up top for more pictures and demos of how these work (all the demos are in the Demo Project in the source).

![Screenshot4](https://i.imgur.com/1ApJRg1.jpg) ![Screenshot5](https://i.imgur.com/RW6XlJh.jpg)

![Screenshot6](https://i.imgur.com/FGIPojS.jpg) ![Screenshot6](https://i.imgur.com/xcbkWSB.jpg)

