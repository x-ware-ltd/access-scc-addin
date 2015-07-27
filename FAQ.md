

This page is here to answer common questions that come from people trying to use the Add-In for the first time.

If you don't see the answer to your question here, join the project mailing list and ask the developers directly.

# Issues with first time use #
## Microsoft Access cannot open this file. ##
The full error that may occur states:

"This file is located outside your intranet or on an untrusted site. Microsoft Access will not open the file due to potential security problems"

This is probably because your version of windows has blocked the file after downloading it from the Internet. To resolve this, locate the Add-in %APPDATA%\Roaming\Microsoft\AddIns, right click and select properties. If you see the option to "Unblock", do so. This will solve your issue.

## The Add-In appears to hang on opening or initial import ##
This is more than likely an issue with your setup of SubVersion and TortoiseSVN. You **MUST** have the TortoiseSVN client as well as the command line SVN utilities installed (TortoiseSVN installer offers the option of installing the command line utilities at setup time)

It could also be an authentication issue. The Add-In currently has no method of requesting authentication from the SVN server. When you performed a checkout from the repository you would have been prompted for a username, when you did this, you had the option for Tortoise to remember those details. Tick that box. This will cache your credentials across Tortoise and the SVN cmd line utility, and will ensure the add in works correctly at that point as well.

## Could not find the object 'MSysAccessStorage' ##
The Add-In uses an Access internal table called MSysAccessStorage to derive some status about local objects. This internal table only came about in Access 2002 version databases, and will not be present in databases prior to 2002.

To encounter this error, you are either:
  1. Using Access 2000 or earlier, the addin is only supported from Access XP.
  1. You are using Access XP or 2003, but the default setting for the file format is set to create Access 2000 type databases. (This is the default out of the box)

To remedy the second scenario:
  * Open any database using Access
  * Click on Tools -> Options.
  * Go to the Advanced tab and change the default file format to "Access 2002".

This will set the default for all new databases.

You will then need to recreate your blank database again and start the import using the addin.

## No mapped URL in options window when selecting the SVN folder ##
If the folder is correctly checked out in SVN, you have cached your user authentication details (as detailed above) and you are getting this error, you may have a dependency issue. The Add-In makes use of the windows scripting host libraries for some integration with the Subversion command line tool. If you do not have windows scripting host installed, or it is disabled, you will not be able to use the Add-In.

If not then it may be disabled, or not installed.

[Enable Script Host](http://www.ehow.com/how_8392805_enable-script-host-windows-7.html)

[Install Script Host](http://www.ehow.com/how_7432762_install-windows-script-host.html)

**Note:** The dependency on windows scripting host was removed from version 0.45. If you are using a version prior to this, upgrade to the latest version.