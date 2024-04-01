# Export MS Excel Macros

Since [MS Excel](https://en.wikipedia.org/wiki/Microsoft_Excel) workbooks are binary file formats, they are not suitable for revision control.

An MS Excel workbook containing macros, if added to [Git](https://en.wikipedia.org/wiki/Git), should be able to export all the source codes with each commit.

&nbsp;

## Requirement

The Python tool [oletools](https://pypi.org/project/oletools/) is required to export the contained VBA forms & modules:

```
>pip install -U oletools
```

&nbsp;

## Git Hooks

The [Pro Git book](https://git-scm.com/book/en/v2) explains [Git Hooks](https://git-scm.com/book/en/v2/Customizing-Git-Git-Hooks) in great detail:

> Like many other Version Control Systems, Git has a way to fire off custom scripts when certain important actions occur.

A [Bash script](./pre-commit/pre-commit) and a [Python script](./pre-commit/pre-commit.py) together export macros contained in MS Excel workbooks.

These scripts are tested only on MS Windows.

&nbsp;

## Pre-Commit Hook

A **pre-commit** [hook](https://git-scm.com/book/en/v2/Customizing-Git-Git-Hooks) checks first for presence of unwanted worksheets. If such a worksheet is not present, then the pre-commit hook exports VBA modules to **src.vba** directory, so that all source files can be put under revision control.

The pre-commit hook script is taken from the blog post [How to use Git hooks to version-control your Excel VBA code](https://www.xltrail.com/blog/auto-export-vba-commit-hook).
The scripts are modified slightly.

The same company's open source [git-xl](https://github.com/xlwings/git-xl) extension for Git handles direct diff ops between workbook revisions, but it is not much of help inside an IDE.

These two files must first be put into the **.git/hooks** directory of the repository where workbooks are located.

&nbsp;

## Filter Worksheets

The shell script can give an *unwanted* worksheet name.
If it is empty filtering is ignored.

```
WS_UNWANTED='Confidential'
```

A worksheet with exactly this name or containing this name causes the commit to fail.
The Python script may check for exact matches, by default a partial match filters unwanted worksheets.

```
EXACT_MATCH = False
```

&nbsp;

## Example

After first copying the pre-commit hook scripts, an MS Excel file containing two VBA modules is added to the test repository:

```
>TREE /F
Folder PATH listing for volume Windows
Volume serial number is EX47-DA1E
C:.
    LICENSE
    README.md

No subfolders exist

>COPY ..\export-excel-macros\pre-commit\* .git\hooks
..\export-excel-macros\pre-commit\pre-commit
..\export-excel-macros\pre-commit\pre-commit.py
        2 file(s) copied.

>MKDIR workbook

>MOVE ..\test.xlsm workbook\.
        1 file(s) moved.

>git status
On branch main
Untracked files:
  (use "git add <file>..." to include in what will be committed)
        workbook/

nothing added to commit but untracked files present (use "git add" to track)

>git add workbook

>git commit -m "Add a macro-containing MS Excel workbook"
[main 93d90d4] Add a macro-containing MS Excel workbook
 3 files changed, 2174 insertions(+)
 create mode 100644 src.vba/Calculations.bas
 create mode 100644 src.vba/Init.bas
 create mode 100644 workbook/test.xlsm

>git status
On branch main
nothing to commit, working tree clean

>TREE /F
Folder PATH listing for volume Windows
Volume serial number is EX47-DA1E
C:.
│   LICENSE
│   README.md
│
├───src.vba
│       Calculations.bas
│       Init.bas
│
└───workbook
        test.xlsm
```