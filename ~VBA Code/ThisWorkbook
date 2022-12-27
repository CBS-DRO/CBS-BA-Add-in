private withevents app as application

private sub app_workbookactivate(byval wb as workbook)
    logr.registerall
    classification.registerall
    knn.registerall
    montecarlo.registerall
end sub

private sub app_workbookaddininstall(byval wb as workbook)
    ' check that the old-in is open
    if util.checkoldaddinopen() = true then
        msgbox "the xll version of the cbs ba add-in is opened. please disable the other version in 'file > options > add-ins > manage (excel add-ins) go' and try again.", vbcritical
        thisworkbook.close
    end if
end sub

private sub app_workbookopen(byval wb as workbook)
    ' this event is triggered when a workbook is open
    if wb is nothing then exit sub
    if wb is thisworkbook then exit sub
    if wb.isinplace then exit sub
    ' fix the links
    util.checkandfixlinks wb
end sub

private sub workbook_addininstall()

end sub

private sub workbook_open()
    ' check that the old-in is open
    if util.checkoldaddinopen() = true then
        msgbox "the xll version of the cbs ba add-in is opened. please disable the other version in 'file > options > add-ins > manage (excel add-ins) go' and try again.", vbcritical
        thisworkbook.close
    end if
    
    set app = application
    
    'addmenuoptions
    addshortcuts
    
    ' fix the links on all open workbooks
    dim wb as workbook
    for each wb in application.workbooks
        app_workbookopen wb
    next
end sub

private sub addshortcuts()
    application.onkey "+^m", "montecarlo.showmontecarlodialog"
    application.onkey "+^f", "util.autoformulabox"
    application.onkey "+^v", "util.versionnum"
end sub

private sub addmenuoptions()
    #if mac then
        exit sub
    #end if

    dim cmbbar as commandbar
    dim cmbcontrol as commandbarcontrol
     
    set cmbbar = application.commandbars("worksheet menu bar")
    set cmbcontrol = cmbbar.controls.add(type:=msocontrolpopup, temporary:=true) 'adds a menu item to the menu bar
    
    on error resume next
    application.commandbars("worksheet menu bar").controls("vba new add-in").delete 'delete the menu item if it exists
    on error goto 0
    
    with cmbcontrol
        .caption = "&vba new add-in" 'names the menu item
        with .controls.add(type:=msocontrolbutton) 'adds a dropdown button to the menu item
            .caption = "run montecarlo simulation (ctrl+shift+n)" 'adds a description to the menu item
            .onaction = "montecarlo.showmontecarlodialog" 'runs the specified macro
            .faceid = 1098 'assigns an icon to the dropdown
        end with
        with .controls.add(type:=msocontrolbutton)
            .caption = "clear histogram sheets"
            .onaction = "util.clearhistograms"
            .faceid = 108
        end with
        with .controls.add(type:=msocontrolbutton)
            .caption = "add formula box (ctrl+shift+d)"
            .onaction = "util.autoformulabox"
            .faceid = 385
        end with
        with .controls.add(type:=msocontrolbutton)
            .caption = "version info (ctrl+shift+v)"
            .onaction = "util.versionnum"
            .faceid = 487
        end with
    end with
end sub
