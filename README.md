This project is an Add In to Microsoft Access to allow users to import and export the entire database to a Version Control System such as Subversion.

It was demonstrated at the UK Access Users Group in 2011 and received enough attention to warrant making it an open source project.

Unlike the Visual Sourcesafe Addin
- It does not use the MSSCCI interface
- It allows multiple users to edit the same object without needing to lock it first.
- It creates a full separate export of all tables (rather than a single binary file)
- It allows separate export of References
- It allows separate export of Toolbars

*Stop the press*: As of Access 2013, the SCC layer of Access has been "deprecated". This means if you are already using systems that rely on the MSSCCI layer, such as Visual Sourcesafe, [PushOK SCC plugin](http://www.pushok.com/software/svn.html) or [Unified SCC](http://aigenta.com/products/UnifiedScc.aspx), this will no longer work from Access 2013 onwards.

# Recent Release History
See [the releases page](https://github.com/x-ware-ltd/access-scc-addin/releases) for a list of releases and descriptions of what has changed.

# Version Control Systems
The only VCS that has enhanced support at the moment is Subversion/TortoiseSVN using `SVN` mode. This includes:
- doing an SVN import before importing into Access
- doing opening the TortoiseSVN commit window after an export from Access

In addition the add-in can be used in `TEXT` mode. This is a source code control agnostic mode that works directly with the file system. Using this mode the add-in can be used with any VCS/SCC system.

Future plans could allow enhanced support for Git, Mercurial, possibly TFS (Team Foundation Server) and many others if there is enough demand.

# Access Versions
This tool was originally developed for Access 97, but it now built around Access XP (2002-2003 format database).
The addin is actively used in Access 2016, but there may be features present in newer versions of Access that will not work and could cause errors on import or export. 

Known issues are:
- Relationships not being imported/exported (we have no need for them, so have not implemented them)
- Attachment fields in tables do not work
- Multi valued records
- Ribbon not actively supported (As they are storred internally in a table they can be imported/exported, but can not be effectively diffed)
- Navigation bars not supported

# Requirements
You will need the following to use this:
- A supported version of Access (XP or greater)
- In `SVN` mode:
    - A recent version of Tortoise SVN, with the command line tools (version 1.7 or greater)
    - Access to a Subversion repository (private or cloud based)

# Instructions
Instructions for the following can be found in the wiki:
- How to install the addin
- How to use the addin
- How to contribute to the source
- Frequently Asked Questions

# Encoding

Access 2002 and earlier encoded all objects as CP1252. In Access 2016 (possibly earlier) the following objects, depending on content, are encoded as either CP1252 or UCS2:
- Queries
- Reports
- Macros
- Forms

Using the `UTF-8 Encoding` option these UCS2 files will be reencoded as utf-8 with a byte order mark (BOM). Regardless of this option, the above objects with files that are UTF-8 BOM encoded will be reencoded as UCS2 on import.

If in addition to the UCS2 encoded files you wish to reencode all CP1252 files you can use the `UTF-8 Encode All` option described below.

# Configuration

A file called `MSAccessProject.config` is looked for in the root of the repository. It is intended that this file is committed into SCC. This is structured as an ini file, so the config for section `MSAccess SCC Add-in`, option `UTF-8 Encoding`, is specified like
```ini
[MSAccess SCC Add-in]
UTF-8 Encoding=True
```

## Configuration options

All MSAccess SCC Add-in configuration has a section of `MSAccess SCC Add-in`, these are the available options:


### `UTF-8 Encoding`
When `UTF-8 Encoding=True`, on export, UCS2 encoded files are converted to UTF-8. This affects the following objects (and the files within these folders):

- Queries
- Reports
- Macros
- Forms

### `UTF-8 Encode All`
When `UTF-8 Encode All=True`, on export, all Access files are converted to UTF-8. This is a superset of the files affected by `UTF-8 Encoding` and as such can replace it. This affects the following objects (and the files within these folders):

- Tables
- Queries
- Forms
- Reports
- Macros
- Modules
- Toolbars
- Extras
- References
- IMEXSpecs

### `Remove GUIDs`
When `Remove GUIDs=True`, on export, GUID blocks are removed from Forms and Reports.

# How Can I Contribute?
The team is comprised of volunteer developers. We are always looking for new contributors that have an interest in Access and Source Code Control, and even people that can just help improve the documentation.

The current developers have a very rough roadmap of where to go forward, but if you have ideas then just exercise your initiative and commit the changes. The use of SCC on this project allows for full oversight, so if things need to be changed after review, then we can easily do that.

If you are interested in have a look at the Issues, pick an Easy or Good First Task and try to solve it. If you need any help getting started just add a comment. We look forward to working with you!
