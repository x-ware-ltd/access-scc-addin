# Development #

To build from source, there must be a working copy already installed.
Make sure to follow the installation guide before performing the following steps.

  1. Check out a copy of the source from svn trunk.
  1. Create a new database, with file name **Subversion4AccessXP.mda**
  1. Use the Add-In to import objects from directory containing source
  1. Remove '-svntmp' suffix from forms created incorrectly - a list of which is shown in the results.
  1. This database can then be modified to change the Add-In
  1. Uninstall current version of the Add-In from Tools->Add-Ins->Add-In Manager
  1. Close database, change extension from .mdb to .mda so it is recognised as an Add-In.
  1. Close any instances of Access.
  1. Erase previous version from `%USERPROFILE%\AppData\Roaming\Microsoft\AddIns`.
  1. Follow the installation guide with the new mda.

Note: Be sure to uninstall the Add-In from the Add-In manager before editing it directly.