BAS Module: modTVpropbag.bas

Ver. 1.0.1 Written by: CptnVic  29 Feb. 2008
Ver. 1.0.0 Written by: CptnVic   7 Feb. 2008

Purpose: Save and Load Treeview contents & properties
----------------------------------------------------------------------------
New In Version 1.0.1:
   * Added support for node.tags
   * Cleaned up and reduced needed code

----------------------------------------------------------------------------
Known Bugs: None

----------------------------------------------------------------------------
WHY?

Recently, I wrote some code for a company that required saving and restoring
the contents of a treeview control.  My first code submission (an XML routine)
was rejected due their concerns over XML version clashes.  My second submission
(a database version) was rejected over similar concerns.  They finally bought
my third submission which utilized a random access database.

That project is completed, and works fine, however I have given the task further
thought - the result being the BAS module included in this project.

During my third attempt (and under a considerable time constraint!), I researched
various methods of saving and loading treeview contents.  There are numerous
examples of XML, tab delimited, character delimited, etc. methods... I could 
only find one web page that discussed the method used in my module.  Sadly,
the author of that page did not understand the treeview very well and when I 
finally got his code to work it did almost nothing that I needed.

SO... this is what you get... a module that saves and restores treeview contents
and supports multiple treeviews.  This is as close to "plug-n-play" as I can
get you.

----------------------------------------------------------------------------
WHAT?

The demo project contains a form (frmDemo) and the BAS module: modTVpropbag.bas
that does all the work.

The code in frmDemo (and the form, for that matter) is simply there so you
can experiment with the treeview saving, loading, etc.  It's only relationship
to the bas module is that the LoadTree and SaveTree subs are called from there.

In other words, that's the demo part of the project.  Use your own methods to
create nodes/child nodes and modify their properties, then use the SaveTree
sub to save the treeview... and LoadTree to restore the saved treeview contents.

----------------------------------------------------------------------------
Use:

Place the modTVpropbag.bas in your project and save.

To save a treeview:

  SaveTree tView, [sFile], [DoBackUp]
  Where:
        tview = the name of the treeview to save
        sFile = Optional: The name of the file (with path) to load the treeview contents from.
	DoBackUp = Optional: True will backup the treeview.bag file, False skips the backup.

To Load/Re-load a treeview:
  
  LoadTree tView, [lFile]
  
  Where: 
	tview = the name of the treeview to save
	lFile = Optional: The name of the file (with path) to save the treeview contents to.

----------------------------------------------------------------------------
UNDONE:

I have tried to include the most commonly used properties of the treeview and
to support the use of multiple treeviews.

The only thing (that I can think of!) that you may want to implement that is
not currently implemented is:
  * Restoring the selected node after LoadTree (not my preference)
    -- Easily done by writing YourTREEVIEW.SelectedItem.Index to the bag file.

----------------------------------------------------------------------------
TESTED ON:

Uses propertybag class introduced with vb6, and ms treeview control (MS common 
controls 6.0).  Tested on a P3 Win2000 Pro computer.
Has not been tested on subclassed treeview or custom treeview controls - but see 
few problems there unless properties are vastly different.

The propertybag class is a member of VBRun... so this code should run on virtually
any Windows OS.

----------------------------------------------------------------------------
THANKS TO:

Roger Gilchrist:
For taking the time to send me an email that shamed me into cleaning up the code 
from the ver. 1.0.0 submission.  I shouldn't have done it - but when I ported this 
code from the Rnd access version, I failed to clean up a bunch of loose ends before 
submitting the project.  It all worked, but I think you will find this version much 
smaller, cleaner and considerably optimized.  I apoligize to those to those who were
disgusted with my previous effort... I had some knee surgery scheduled and I wanted
to get that project off of my desk since I knew I wouldn't be in the mood to play
with it again for a while.

To all other emailers:
This version implements the node.tag property that you wanted.  This is probably a
better place to store filenames (etc.) than the node.key since you may want to
re-build an index of node.keys to facilitate the on-demand functions I think you
are shooting for.

I have played with this on-demand idea briefly, and if I pursue the idea - you'll
see the code on PSC.

----------------------------------------------------------------------------
FIND A BUG... Questions OR SUGGESTIONS?

Please let me know!

As always, your most ardent admirer!
CptnVic