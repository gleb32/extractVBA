# extract-vba

This is a project designed to help use Source Code Control (SCC) with
VBA projects that are embedded within MS Office files.

The goal of the project is to have a script that can be called from a
git hook or from a CI server.


## Instructions
TODO (after chosing a method and finishing the code...)


## Concept
Microsoft VBA scripts are embedded within their host Axcel, Access, or Word
file. For purposes of SCC, these files are very difficult to use. Not only
are they very hard for a human to diff because they are not plaintext, but
they also don't play nicely with diff-based SCC such as git.

There are two ways to approach this problem, both of which require extracting
the VBA from within the Office file.

### Git Hooks
The first is to create a script that is run before a user commits changes to
his local repository.

Because we use TortoiseGit here, we'd acually use a `start-commit` hook so
that the files are generated before being staged.

With this approach, the user would add their primary office files to a
`.gitignore`. When commiting changes, the script would run on all Office
files within his working directory and would create the VBA files. Only the
VBA files would be committed to the repository.

The problem with this method is that the extraction script *might* take
a nontrivial amount of time to run because it has to open each program
and each file to parse.


### Continuous Integration
The second option is to use continuous integration to extract the VBA.

A user would work on the Office document file and commit those changes to
his local repository. Upon pushing, a CI service would run, extract the VBA,
and push it to the remote repository.

There are a few main disadvantages to this approach:
1.  The user does not get to see his changes when committing, meaning they
    can't use the diff to create a commit message.
2.  The user only gets to view diffs when pushing - depending on the user's
    coding style, this could mean long time deltas between each VBA diff
    (example: if they only push every 20 commits).
3.  The CI server would have to have a MSOffice licence.


## Useful Links
+ How to enable programmatic access to the VBA within a file
  + http://www.cpearson.com/excel/vbe.aspx
+ Extracting the VBA from the file using Python and COM objects
  + http://stackoverflow.com/a/12288193/1354930
  + This was probably the most important research. Thanks @steven-rumbalski!
+ Import and Export VBA code
  + http://www.rondebruin.nl/win/s9/win002.htm
+ Extract VBA code to text files
  + http://www.pretentiousname.com/excel_extractvba/
