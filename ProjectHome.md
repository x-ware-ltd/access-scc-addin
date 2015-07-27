# Summary #
This project is an Add In to Microsoft Access to allow users to import and export the entire database to a Version Control System such as Subversion.

It was demonstrated at the UK Access Users Group in 2011 and received enough attention to warrant making it an open source project.

Unlike the Visual Sourcesafe Addin
  * It does not use the MSSCCI interface
  * It allows multiple users to edit the same object without needing to lock it first.
  * It creates a full separate export of all tables (rather than a single binary file)
  * It allows separate export of References
  * It allows separate export of Toolbars

**Stop the press**: As of Access 2013, the SCC layer of Access has been "deprecated". This means if you are already using systems that rely on the MSSCCI layer, such as Visual Sourcesafe, [PushOK SCC plugin](http://www.pushok.com/software/svn.html) or [Unified SCC](http://aigenta.com/products/UnifiedScc.aspx), this will no longer work from Access 2013 onwards.

## Recent Release History ##
0.50
  * Updated to support export of null data in numeric fields. This issue was introduced when support for international number formats was added.

0.49
  * Updated to support version 1.8 of the subversion libraries, and to stop nagging users that are using Access 2007 and above. The previous caveats on Access 2007 and above still exist though.

0.47
  * Update to handle regional formatting of numbers. Prior to this release some regions may have caused errors, and would have stored numeric table data incorrectly in the repository. Thanks to Florian Reitmeit for spotting this one.

## Version Control Systems ##
The only VCS that is supported at the moment is Subversion but future plans could allow support for Git, Mercurial, possibly TFS (Team Foundation Server) and many others if there is enough demand.

## Access Versions ##
This tool was originally developed for Access 97, but it now built around Access XP.
There is no reason why it cannot be built and used in Access 2003, 2007 and 2010, however
any of the new features present in those systems will not work and may cause errors during
export.

Known issues are:
  * Relationships not being imported/exported (we have no need for them, so have not implemented them)
  * Attachment fields in tables do not work
  * Multi valued records
  * Ribbon not being imported/exported
  * Navigation bars not supported
  * 64 bit Access has never been tested at all (doubt I really want to!)

## Requirements ##
You will need the following to use this:
  * A supported version of Access (XP or greater)
  * A recent version of Tortoise SVN, with the command line tools (version 1.7 or greater)
  * Access to a Subversion repository (private or cloud based)

## Instructions ##
[How to install the addin](http://code.google.com/p/access-scc-addin/wiki/Installation)

[How to use the addin](http://code.google.com/p/access-scc-addin/wiki/Usage)

[How to contribute to the source](http://code.google.com/p/access-scc-addin/wiki/Development)

[Frequently Asked Questions](http://code.google.com/p/access-scc-addin/wiki/FAQ)


## How Can I Contribute? ##
The team is comprised of volunteer developers. We are always looking for new contributors that have an interest in Access and Source Code Control, and even people that can just help improve the documentation.

The current developers have a very rough roadmap of where to go forward, but if you have ideas then just exercise your initiative and commit the changes. The use of SCC on this project allows for full oversight, so if things need to be changed after review, then we can easily do that.

If you are interested in joining the team, send an email to steve.falla@talkjcs.com with your contact info. We look forward to working with you!