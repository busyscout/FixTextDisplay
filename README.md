# FixTextDisplay
Visual Studio is an WPF application. 
WPF text rendering engine has a lot of settings, one of which is TextFormattingMode.

Most parts of Visual Studio use TextFormattingMode=Display, which means, they render text according to OS settings.
But there are still some parts (e.g. Test Explorer, Git commit message placeholder, Exception dialog) that use TextFormattingMode=Ideal, 
so they render smoothed (cleartyped) text even if font smoothing is turned off in your OS.

This hurts eyes. 

This extension customizes WPF styles, so all the controls use TextFormattingMode=Display.
