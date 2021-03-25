# RSuite_Word-template

This repo is for developing & maintaining assets related to the Macmillan RSuite Word template.

Instructions for manually installing &/or packaging the RSuite Word-template are below: for Word for PC (2010/2013/2019) or Word for Mac (2016/2019).

After that are notes for development, testing, and maintenance.

# Installation

## Assets required for installation

 **Files for Installation should be pulled from the 'files_for_install.zip' attachment to the [latest release](https://github.com/macmillanpublishers/RSuite_Word-template/releases/latest).**

* File:  template_switcher.dotm
* Folder:  RSuiteStyleTemplate
* Folder:  MacmillanStyleTemplate

Please read sections: _PC install_ and _Mac install_, for installation target directories and other details.

## PC Install

#### PC Installation Targets:
* folders: MacmillanStyleTemplate and RSuiteStyleTemplate
Both folders (with all of their contents) should be installed here: `C:\Users\username\AppData\Roaming`
* file: template_switcher.dotm
This file should be installed here:
`C:\Users\username\AppData\Roaming\Microsoft\Word\STARTUP`

#### PC Installation Requirements
* Word will need to be quit
* The Word Startup folder will need to be wiped clean prior to installation
* The rest of the files may pre-exist at time of installation, as long as the newer (installing) versions will overwrite existing.
* Package name / Portal Display name (example): **RSuiteStyleTemplatev6.0**
(The version number should match the version number from the [latest release](https://github.com/macmillanpublishers/RSuite_Word-template/releases/latest) that you downloaded assets from.

###### Notes for Macmillan packaging team (PC)
1. Ideally pushed via deployment, but also available in Windows SelfService portal.
2. If it's straightforward, hide the bluescreen (powershell?) window that pops up during installation via portal.
3. It would be nice if we had the 'standalone' installer, i.e. the .bat file to move files into location for our outside composition vendor, but this is also not critical.

## Mac Install

#### Manual install step-by-step
1. Download 'files_for_install.zip' attached to [latest release](https://github.com/macmillanpublishers/RSuite_Word-template/releases/latest).
2. Unzip files_for_install.zip, and open Terminal. _cd_ into the newly unzipped folder, to perform the commands in step 3 & 4 (eg: cd _/Users/username/Downloads/files_for_install_)
3. Run this command to strip apple quarantine from downloaded files:
    ```xattr -dr com.apple.quarantine ./```
4. Run this command to re-set Word doctype for .dotm files:
    ```find ./ -type f -name "*.dotm" | xargs xattr -wx com.apple.FinderInfo "57 58 54 4D 4D 53 57 44 00 10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"```
5. Copy or move the file _template_switcher.dotm_ to this location on your Mac:
_/Users/username/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/_
6. Copy or move both folders, _MacmillanStyleTemplate_ and _RSuiteStyleTemplate_ to this location on your Mac:
_/Users/username/Library/Containers/com.microsoft.Word/Data/Documents_
7. Launch Word. Make sure you see the 'RSuite' & 'Inspect' buttons, and or the 'RSuite Tools' tab in the ribbon, and try some RSuite Tools items out to make sure they work.


#### Mac Installation requirements
* Word should be quit for the installation.
* Pre-existing contents of the following folders (files & folders) should be removed as a pre-install step:
/Users/username/Library/Containers/com.microsoft.Word/Data/Documents
/Users/username/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/
* Package name / Self-Service Display name (example): **RSuiteStyleTemplatev6.0**
(The version number should match the version number from the [latest release](https://github.com/macmillanpublishers/RSuite_Word-template/releases/latest) that you downloaded assets from.

###### Notes for Macmillan packaging team (Mac):

1. This should be available to run via Self-Service, as well as via policy / timed deployment.
2. It would be nice to have a standalone installer pkg as well, for freelancers etc.

# VBA Development

#### Dependencies
* if using gradle install/pkg tools, gradle requires installation of jdk 8 or higher, available [here](https://jdk.java.net/) (The first 'Ready to Use' version should be fine).
* To use the 'devTools' macros below, you may need to enable the following libraries in your VBA editor (Tools > References):
  * ``Microsoft Visual Basic For Applications Extensibility 5.3``
  * ``Microsoft Scripting Runtime``
  * ``Microsoft Forms 2.0 Object Library``
* For *Unit testing*, download and install [RubberDuck](https://rubberduckvba.com/).

#### Overview
Development documentation below is broken into 5 main topics:
* Project maintenance, releases & versioning
* Working with the custom Ribbon
* gradle tools: to install files to local env for testing, to collect and version-tag files for release.
* devSetup tools: macros to facilitate working in the VBA editor
* Unit testing

Additionally, detailed documentation regarding auto-creation of Style Templates using "WordTemplateStyles.xlsm" and "StyleTemplateCreator.docm" can be found in Macmillan Confluence, [here](https://confluence.macmillan.com/display/PWG/Maintaining+Word+Style+Templates).

## Project Maintenance & Releases
Currently work is done on feature branches per feature/bug, and merged into master once tested/verified on Word 2013(PC), Word 2019(PC) and Word 2019(Mac).

Once ready for UAT, Pre-Releases are created in git, named based on a version number (numbering details below). We run gradle 'build' (see gradle section for details) to create asset, 'files_for_install.zip', which is attached to the release.

Following UAT and approval by the business, a 'pre-release' is transitioned to regular release in git. They are deployed to production and staging servers via manual checkout by release-tag, in coordination with Desktop Support (to match their deployment to user workstations).

##### Version Numbering
Versions for this product are named like 'x.y.z', where x y and z are whole numbers (ex: '6.3.1'). The 'x' indicates a new major version, the 'y' indicates a feature release indicating changes in Style-templates, and the 'z' indicates a release with changes to macros/back-end only

##### Version Maintenance
The version is set via file _./version.txt_.

It is manually added as 'Version' custom document property to 'RSuite_Word-template.dotm', 'RSuite.dotx' and 'RSuite_NoColor.dotx' via gradle 'build'.
From there, every time a user attaches the RSuite template to a file, the same 'Version' custom document property is set on that file, making it easy to track attached styles.

##### Version checking by other products
Bookmaker and egalleymaker tools check the major ('x') Version of a styled file to verify style compatibility.
RSuite_Validate tool and 'Document Styles Version' check in the RSuite Tools ribbon (in Word) verify that 'x.y' matches, or surface a warning to user about mismatched templates.


## Custom Ribbon development
The custom ribbon is implemented via custom ribbon xml, stored as part of the binary itself. It is not accessible via the MS VBA IDE, we are using 'Custom UI Editor For Microsoft Office' to directly edit the ribbon.  
No straightforward way to auto-export/import this xml presents itself, so a copy of the xml is separately maintained in the "*custom_ui*" directory, for versioning.

## gradle Tools
### Install via gradle
A quick installation method for development & testing.

Note: Gradle install uses assets from the repo, not 'built' assets from the 'files_for_install' folder (more on that below):
As a result, if installing from source-code between releases, version numbers for the templates may not match each other, or the release.

##### Steps for install:
1. clone repo to your Mac/PC
2. via commandline/Terminal, cd to directory: *_gradle-install*
3. type command for gradle task (varies by OS)
  * on a Mac (or PC bash emulator):
        * type `./gradlew install`
   * on a PC:
        * type `.\gradle.bat install`
4. If Word is running, the install task will fail and suggest that you quit Word.
5. For some reason, on Windows sometimes this installation fails the first time; if you get a Java.io error re: deleting, run installer again.

### Build via gradle
The 'gradlew build' command does the following:

* _If run on a PC_ (with cmd _build_)
  * adds/updates 'Version' document properties for _./RSuiteWord-template_ and style template files in the repo (with value from _./version.txt_),
  * then copies all files required for installation into a 'files_for_install' folder in the root of the repo.
* _If run on a Mac_ (using cmd _force_build_ instead of _build_)
  * only copies all files required for installation into a 'files_for_install' folder in the root of the repo.

This 'files_for_install' folder should then be zipped and uploaded to the corresponding release page in git.

(The _force_build_ command bypasses the version doc-prop step).

##### Steps for build:
Same as gradle '_Steps for install_', above, except use cmd '_build_' or '_force_build_' instead of '_install_'.

## devSetup tools
In an effort to facilitate simpler vba code versioning in GitHub, there are some macros in "devSetup.docm" file to enable easy export/import of modules. There are also macros to apply version numbers to templates and copy installed, 'working' template files back to the repo. Macros detailed below:

##### * Open all Projects
To open all ‘RSuite_Word_Template’ dotm/docm files in the VBA editor for code access, run macro:  *Open_All_Defined_VBA_Projects*
This is a very useful way to access code from installed templates quickly in the VBA editor.

##### * Export Modules and Binaries
Once you're ready to commit some code, there's a tool to export a .dotm/.docm binary and all of its vba-components to the local git repo repository (*all components except custom ribbon).

  1. run macro: *z_Export_or_Import_VBA_Components*
  2. in the pop-up window, select any/all docs with updated code, and click either _'Export'_ option:

    * "_Export file(s) and modules to git repo_"

      Modules are exported to dir: _'src/(file_basename)'_. The .dotm/.docm binary file is copied from its 'installed path' its home in the local git repo. An alert will notify if there is no defined path for the binary (*see 'Setting Import/Export locations' below for more).

    * "_Export modules ONLY_"

      This does the same as above re: modules, but does not write installed .docm .dotm files back to their default locations in the local repo.

##### * Import modules
You may wish at some point to start fresh with a clean set of modules from the repo.  To do this:
1. run macro: **z_Export_or_Import_VBA_Components**
2. in the pop-up window, select any/all docs with updated code, and click _'Import'_.
3. First, for each selected document/project, backups  will be exported for all current vba components, to a default folder appended with suffix '\_BACKUP\_'.
Then table the Import feature imports modules from the same paths as detailed in 'Setting Import/Export locations', below.

##### * Export RSuite_Word-template binary to repo
To just send the working (installed) version of the RSuite_Word-template.dotm to the repo, run macro: "z_copyInstalledRSWTtoRepo".

##### * Set Versions
To update 'Version' custom document property for all 3 key binaries, you can run macro: "updateVersionsForRepoTemplates" (this is the same macro that gradle 'build' uses).

##### Setting Import/Export locations for files
Any *new* .dotm/docm that you export via this macro will export to the same location as the file by default, in a dir called 'src_*(file_basename)*'. To pre-configure a different export path for a given file, add it to devSetup procedure: 'config.defineVBAProjectParams'.

## Unit testing
Unit tests are done using the integrated Rubberduck tool. The testing modules are housed the *RSuite_WordTemplate.dotm* file. (They may also require that devSetup.docm be open for access to config modules.)
Currently unit tests are stored in separate modules per macro/tool, ex:
  * `TestModule1_Cleanup.bas`
A separate `TestHelpers.bas` module is used for shared 'utility' functions and subs.

Each module has corresponding .dotx/.docx test file(s) in ./test_files

#### Running the tests
1. goto the 'Rubberduck' menu in the vba IDE,
2. select 'Unit Tests > Test Explorer'
3. In the Test Explorer click the 'refresh' icon (top left) to detect tests in all open projects.
4. Run tests!

#### Creating tests
Many of the testfiles are Word template (.dotx) files, so to edit the file itself (instead of spawning a new file based on template) you must open via 'File>Open' in Word.

In testfiles for the Cleanup macro & Charstyle macro, initial content for each test is denoted by a heading matching the test's name, plus preceding and trailing double-underscores.  This format is important, b/c this is how the test finds result-strings for assertions.
(Example: If the test sub name is: *TestPCSpecialCharacters_symbol*, then heading for content for that test in Word will be: \_\_TestPCSpecialCharacters_symbol__)

To create new tests, just follow the format of existing tests in detail. Result strings are defined at the top of the module, for reuse with 'multiple runs' test-scenarios.


---
###### * Notes for future development
- Integration tests
