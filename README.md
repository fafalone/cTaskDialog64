# cTaskDialog
### Current Version: v1.6 Universal Compatibility Version

**Quick Start:** Add cTaskDialog.cls and mTDHelper.bas to your project-- these are the only two required files for your code.


cTaskDialog :: A complete class wrapper for `TaskDialogIndirect`, with additional custom features, universally compatible with VB6/VBA7/twinBASIC x86/x64

**Update (v1.6.2, 02 Dec 2025):** Common button icon for Help should not have been excluded.\
**Update (v1.6.0, 02 Dec 2025):** 
- Add Property InputPasswordChar to Set a custom password char.
  Set as a Unicode AscW value (Integer). The black dot is the default.
- Common button icons now supported for Abort, Ignore, Try Again, and Continue.
- Bug fix: Custom button icons off by one if radio buttons present
- Bug fix: Shell32 Icon IDs could fail without explicitly loading shell32.dll. 
- Demos: twinBASIC Password demo updated to show use of the VerifyText checkbox to 
   toggle whether the password is visible by using .InputPasswordChar

**Update (v1.5.3 (1.5 R3), 03 Jun 2025):**
- Bug fix: Public const in class.

**Update (v1.5.2 (1.5 R2), 27 Mar 2025):**
- Changed missed Debug.Print statements to DebugAppend and set useropt_dbg_PrintToImmediate to False by default, so the class will no longer print debug messages unless changed.\
- Bug fix: zzGetCommonButtonIcon and ResultComboData Long instead of LongPtr.\
- Corrected misc spelling mistakes highlighted by the AccessUI version :)\
**Update (v1.5, 19 Mar 2025):** mTDHelper.bas has been restored to its earlier compact form; change was during troubleshooting and unnecessary. No change to class.\
**Update (v1.5, 15 Jun 2024):**
- Class will now attempt to use comctl32.dll 6.0 in the absence of a manifest, since it's impactical to add one to 32bit VBA hosts without one, like Excel. This is activated only immediately prior to the API call and deactivated immediately after, so it won't impact things like Visual Styles outside this class.

- Added lParam options for AddComboItem; obtain from result with ResultComboData.

 - Custom icons were broken in the main demo project (no issue in this class)

 - ComboNewIndex property to provide the last added combo item index.


**Update (v1.4, 19 Jan 2024):** Incorrect versions of mTDSample.bas were being used that did have conditonal PtrSafe declares. This has been fixed in the root dir for the VBP, in the Export dir, in the twinproj, and on VBForums.\
**Update (v1.4, 17 Jan 2024):**\
After review, I've included the undocumented additional common buttons that were used in the AccessUI version (thanks!). The following .CommonButtons are now available, with their return value given in parentheses:

```vba
TDCBF_ABORT_BUTTON     (TD_ABORT)
TDCBF_IGNORE_BUTTON    (TD_IGNORE)
TDCBF_TRYAGAIN_BUTTON  (TD_TRYAGAIN)
TDCBF_CONTINUE_BUTTON  (TD_CONTINUE)

TDCBF_HELP_BUTTON      '**This will raise the Help Event, and will not close the dialog.**
```

The Help button works everywhere, *including MS Access*. Unfortunately, the AccessUI version had a typo; the release had 16384 which isn't anything-- but it looks like they just had a typo originally, there's a comment '104857 which of course makes no sense... but if you convert these values to hex, you find `&H10000, &H20000, &H40000`, and `&H80000` for the other new buttons... `&H100000` is **1048576** in decimal-- so they just cut off a digit when copying it down. `&H100000` works in Access, I checked. 


**Update (v1.3.8, 30 Sep 2023):** Fix for custom buttons in VBA64.\
**Update (v1.3.7, 28 Sep 2023):** NOW FULLY WORKING IN VBA64! Note: You must update mTDHelper.bas too.)

![Screenshot1](https://i.imgur.com/AQEvO9W.gif) ![Screenshot2](https://i.imgur.com/8VvddRR.gif)

![Screensot3](https://i.imgur.com/npGDQVe.jpg)


This is a version of my [cTaskDialog project](https://www.vbforums.com/showthread.php?777021-VB6-TaskDialogIndirect-Complete-class-implementation-of-Vista-Task-Dialogs) that uses conditional compilation to support both VB6/VBA6 and twinBASIC/VBA7 in either x86 or x64. See that page for complete project description and numerous more pictures and examples. The demo is provided as a twinBASIC project, but you can get just the cTaskDialog.cls and modTDHelper.bas for VB6/VBA in Export\Sources. The demos are in Form1.frm.twin there too, but you can use the demos from the main project thread too.

Since people have asked about using this in VBA, it goes back to the earlier method of using a module to help with subclassing, as the self-subclass code in the last VB6 version only works in VB6, and while twinBASIC supports AddressOf on class members, VBA7 does not. Note that there's a bug in the self-sub version that changes the way multiple pages are handled, sending all events through the first page class. So if you use multiple paged Task Dialogs, you'll now need to relocate events for the other pages to their own event Subs (the Demo does this with it's multi-page Demos).

> [!NOTE]
> You can find a number of tutorials for the examples on the [original VB6 project page](https://www.vbforums.com/showthread.php?777021-VB6-TaskDialogIndirect-Complete-class-implementation-of-Vista-Task-Dialogs).

### Updates
(30 Sep 2023) In my excitement over callbacks finally working, I forgot that I had not implemented the `TASKDIALOG_BUTTON_VBA7` alternates for custom buttons. This has now been implemented and basic functionality verified. Please notify of any issues.

(28 Sep 2023) Courtesy of brilliant programmer The trick, a fix has finally been identified for use of the callbacks in VBA 64bit. Note: You must update mTDHelper.bas too.

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

