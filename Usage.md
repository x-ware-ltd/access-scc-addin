# Using Microsoft Access SCC Addin #
This will guide you through the steps of managing an Access project using Subversion.

Make sure to follow the installation guide before performing the following steps.

## Exporting database into Subversion ##

  1. Create a folder to use for exporting Access files to text to be imported into Subversion.
  1. Run the Add-In (Tools->Add-Ins->Subversion for Access XP).
  1. Select the folder you created earlier as the SVN Checkout Folder.
  1. Uncheck the "Commit to SVN after export? option.
  1. Make sure all objects are selected on the right.
  1. Click Export Objects, then click OK.
  1. Once this is complete, import the folder into Subversion.

The folder contains your Access database converted into text, which is now imported into Subversion.

## Recreating database from Subversion ##

  1. Checkout a working copy to add the extra folders and files Subversion needs to track a repository.
  1. Create a new Access database.
  1. Run the Add-In (Tools->Add-Ins->Subversion for Access XP).
  1. Select the folder that contains the working copy as the SVN Checkout Folder.
  1. Click on Import Objects, then click OK.