Attribute VB_Name = "RibbonControl"
Option Explicit

Sub LaunchMacros(Optional control As IRibbonControl, Optional buttonID As String)
    ' Calls macro named as "id" attribute in customUI.xml code
    ' Could also be "tag" attribute if necessary
    
    Dim strMacroName As Variant
    strMacroName = control.ID
    Application.Run strMacroName

End Sub


























'                    <button id="AttachTemplateMacro.zz_AttachStyleTemplate" label="Add Manuscript Styles" image="addfile" size="large" onAction="RibbonControl.LaunchMacros" />
'                    <button id="AttachTemplateMacro.zz_AttachBoundMSTemplate" label="Remove Color Guides" image="removefile" size="large" onAction="RibbonControl.LaunchMacros" />
'                    <button id="AttachTemplateMacro.zz_AttachCoverTemplate" label="Add Cover Copy Styles" image="book" size="normal" onAction="RibbonControl.LaunchMacros" />
'                    <button id="PrintStyles.PrintStyles" label="Print Styles in Margin" image="print" size="normal" onAction="RibbonControl.LaunchMacros" />
'                    <button id="ViewStyles.StylesViewLaunch" size="normal" image="eye" onAction="RibbonControl.LaunchMacros" label="View Styles"/>
'                </group>
'                <group id="CustGrp2" label="Manuscript Tools" >
'                    <!--<button id="CastoffMacro.UniversalCastoff" label="Castoff" image="calculator" size="large" onAction="RibbonControl.LaunchMacros" />-->
'                    <button id="Endnotes.EndnoteDeEmbed" label="Unlink Endnotes" image="unlink" size="large" onAction="RibbonControl.LaunchMacros" />
'                    <button id="CleanupMacro.MacmillanManuscriptCleanup" label="Manuscript Cleanup" image="wineglass" size="large" onAction="RibbonControl.LaunchMacros" />
'                </group>
'                <group id="CustGrp3" label="Style Tools" >
'                    <button id="CharacterStyles.MacmillanCharStyles" label="Tag Character Styles" image="tag" size="large" onAction="RibbonControl.LaunchMacros" />
'                    <button id="TagUnstyledParas.TagUnstyledText" label="Style Body Text" image="target" size="large" onAction="RibbonControl.LaunchMacros" />
'                    <button id="Reports.MacmillanStyleReport" label="Style Report" image="file" size="normal" onAction="RibbonControl.LaunchMacros" />
'                    <button id="Reports.BookmakerReqs" label="Bookmaker Check" image="checkfile" size="normal" onAction="RibbonControl.LaunchMacros" />
'                    <button id="LOCtagsMacro.LibraryOfCongressTags" label="CIP Application" image="index" size="normal" onAction="RibbonControl.LaunchMacros" />
'                </group>
'                <group id="CustGrp4" label="Help" >
'                    <button id="VersionCheck.CheckMacmillanGT" label="Macmillan Tools Version" image="info" size="normal" onAction="RibbonControl.LaunchMacros" />
'                    <button id="VersionCheck.CheckMacmillan" label="Macmillan Styles Version" image="info" size="normal" onAction="RibbonControl.LaunchMacros" />
'                    <button id="EasterEggs.Triceratops" image="pizza" size="normal" onAction="RibbonControl.LaunchMacros" />
