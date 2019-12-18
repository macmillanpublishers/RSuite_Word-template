# RSuite_Word-template

This repo is for storing and maintaining assets related to the Macmillan RSuite Word template.

Instructions for manually installing &/or packaging the RSuite Word-template are below: for Word for PC (2010/2013) or Word for Mac (2016).

## Assets required for installation

* File:  template_switcher.dotm   (from root of RSuite_Word-template)
* Folder:  RSuiteStyleTemplate   (containing:)
  * File:  RSuite_Word-template.dotm   (from root of RSuite_Word-template)
  * Folder:  StyleTemplate_auto-generate   (containing:)
    * File: RSuite_styles.txt
    * File: RSuite_NoColor.dotx   
    * File: RSuite.dotx   
    * File: sections.txt
    * File: breaks.txt
    * File: containers.txt
* Folder:  MacmillanStyleTemplate, + all contents   (from RSuite_Word-template/oldStyleTemplate/MacmillanStyleTemplate)

## PC Installation

#### PC Installation Targets:
* folders: MacmillanStyleTemplate and RSuiteStyleTemplate
Both folders and their contents should be installed here:  *C:\Users\username\AppData\Roaming*
* file: template_switcher.dotm
This file should be installed here:
_C:\Users\username\AppData\Roaming\Microsoft\Word\STARTUP_

#### PC Installation Requirements
-Word will need to be quit

-The Word Startup folder will need to be wiped clean prior to installation

-The rest of the files may pre-exist at time of installation, as long as the newer (installing) versions will overwrite existing.

#### Notes for Macmillan packaging team (PC)
1. Package name / Portal Display name (example): **RSuiteStyleTemplatev6.0**
(The version number should match the one in file: *RSuite_Word-template/StyleTemplate_auto-generate/RSuite.txt*)

2. Ideally pushed via deployment, but also available in Windows SelfService portal.

3. If it's straightforward, hide the bluescreen (powershell?) window that pops up during installation via portal.

4. It would be nice if we had the 'standalone' installer, i.e. the .bat file to move files into location for our outside composition vendor, but this is also not critical.

## Mac Installation
(Word 2016 specific)

#### Manual install step-by-step
1. Download contents of this repo.

2. Unzip the repo, and open Terminal. _cd_ into the root of the unzipped repo, to perform the commands in step 3 & 4 (eg: cd _/Users/username/Downloads/RSuite_Word-template-master_)

3. Run this command to strip apple quarantine from downloaded files:

    ```xattr -dr com.apple.quarantine ./```

4. Run this command to re-set Word doctype for .dotm files:

    ```find ./ -type f -name "*.dotm" | xargs xattr -wx com.apple.FinderInfo "57 58 54 4D 4D 53 57 44 00 10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"```

5. Copy the file template_switcher.dotm on your Mac, here:
_/Users/username/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/_

6. In the folder oldStyleTemplate look for folder MacmillanStyleTemplate.  Copy the whole MacmillanStyleTemplate and all of its contents to this location on your Mac:
_/Users/username/Library/Containers/com.microsoft.Word/Data/Documents_

7. Create a new folder on your Mac:
_/Users/username/Library/Containers/com.microsoft.Word/Data/Documents/RSuiteStyleTemplate_

8. Into this new folder, drop in these two items from the repo:
file: _RSuite_Word-template.dotm_
folder: _StyleTemplate_auto-generate_, with all of its contents

9. Launch Word 2016. Make sure you see the 'RSuite' & 'Inspect' buttons, and or the 'RSuite Tools' tab in the ribbon, and try some RSuite Tools items out to make sure they work.


#### Mac Installation requirements
-Word should be quit for the installation.

-Pre-existing contents of the following folders (files & folders) should be removed as a pre-install step:
/Users/username/Library/Containers/com.microsoft.Word/Data/Documents
/Users/username/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/

#### Notes for Macmillan packaging team (Mac):
1. Package name / Self-Service Display name (example): **RSuiteStyleTemplatev6.0**
(The version number should match the one in file: *RSuite_Word-template/StyleTemplate_auto-generate/RSuite.txt*)

2. This should be available to run via Self-Service, as well as via policy / timed deployment.

3. It would be nice to have a standalone installer pkg as well, for freelancers etc.
