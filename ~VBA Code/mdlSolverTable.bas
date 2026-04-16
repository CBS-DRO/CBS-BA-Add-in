option explicit

public inputcell(1 to 2) as range
public inputname(1 to 2) as string
public outputrange as range
public noutputcells as integer
public outputname() as string
public formcancel as boolean

public ninputvals(1 to 2) as integer
public inputval1() as single
public inputval2() as single
    
public minmaxopt(1 to 2) as boolean
public minval(1 to 2) as single
public maxval(1 to 2) as single
public increment(1 to 2) as single

public originputval(1 to 2) as single
public origchangingcellval() as single
public changingcells as range

public modelsht as worksheet
public modelshtname as string
public shortmodelshtname as string

public stssht as worksheet
public choicecell(1 to 2) as range

public prevonewaysettings as range
public prevtwowaysettings as range

public solvermessage(0 to 20) as string
public solvererrormessage(0 to 20) as string

public const oneway = 1
public const twoway = 2

public tabletype as integer
public engine as integer
public precision as single
'public ms_setting as integer ' multistart (1 for on, 2 for off)
'public usems as boolean

public const epsilon = 0.000001

public sub runsolvertable(control as iribboncontrol)
#if mac then
    msgbox "solvertable is currently not supported on excel for mac", vbexclamation
    exit sub
#end if
    if activeworkbook is nothing then
        msgbox "cbs_ba add-in: there is no workbook currently active. please open a sheet and run this command again", vbexclamation
        exit sub
    end if
    ' save the workbook
    dim modelwbk  as workbook
    set modelwbk = activeworkbook
    ' check if solver is active (this might inactivate the current workbook)
    if not util.checksolverintl() then
        exit sub
    end if
    ' reactivate the workbook
    modelwbk.activate
    set modelsht = activesheet
    modelshtname = activesheet.name
    
    if left(modelshtname, 4) = "sts_" then
        msgbox "the active sheet's name starts with sts_, which indicates that it is " _
            & "a sheet with solvertable results, not the sheet with your model. select a sheet " _
            & "with a solver model, and run solvertable again.", vbinformation, "wrong active sheet"
        goto proc_exit
    end if
            
    on error goto proc_error
    application.screenupdating = true
    
    ' the following check on whether solver is loaded appears to be unnecessary. if solver is not loaded, it
    ' is loaded automatically when solvertable is loaded, and solver can't be unloaded when solvertable is loaded.
'    dim wb as workbook
'    on error resume next
'    set wb = workbooks("solver.xlam")
'    if err.number <> 0 then
'        msgbox "solver isn't loaded. please load it and then try solvertable again.", vbcritical
'        end
'    end if
    
    if not util.rangeexists("solver_adj") then
        msgbox "no valid decision variables found. please set up decision variables in solver and try again.", vbcritical
        exit sub
    end if
    
    call setupsettingsranges
    
    
    tabletype = oneway
    frmoneway.show
    if formcancel then
        exit sub
    end if
    
    application.screenupdating = false
    
    call getoriginalvalues
    call nameoutputcells
    call setupstssheet
    call definesolvermessages
    call runsolver
    call restoreoriginalvalues
    
    stssht.activate
    range("a1").select
      
proc_exit:
    application.screenupdating = true
    exit sub
    
proc_error:
    call showerror("mdlsolvertable", "main", err.number, err.description)
    goto proc_exit
end sub

sub getoriginalvalues()
    dim count as integer
    dim cell as range
    dim nchangingcells as integer
      
    on error goto proc_error
    
    ' find changing cell range and count the number of changing cells.
    set changingcells = range("solver_adj")
    nchangingcells = changingcells.cells.count
    
    ' store the current changing cell values.  these are presumably the optimal
    ' values for the current model with the current inputs.
    redim origchangingcellval(1 to nchangingcells)
    count = 0
    for each cell in changingcells
        count = count + 1
        origchangingcellval(count) = cell.value
    next
    
    ' store the current value(s) of the input cell(s).
    if tabletype = oneway then
        originputval(1) = inputcell(1).value
    else
        originputval(1) = inputcell(1).value
        originputval(2) = inputcell(2).value
    end if
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlsolvertable", "getoriginalvalues", err.number, err.description)
    goto proc_exit
end sub

sub nameoutputcells()
    dim i as integer
    dim nm as name
    dim outputcell as range
    dim namedrange as range, cell as range
    dim rngname as string, excpos as integer
    dim index as integer
    dim isinnamedrange as boolean
    
    noutputcells = outputrange.cells.count
    redim outputname(1 to noutputcells)
    i = 0
    for each outputcell in outputrange
        i = i + 1
        ' use the cell address by default.
        outputname(i) = outputcell.address
        
        ' check if the output cell is part of a named range.
        isinnamedrange = false
        for each nm in activeworkbook.names
            rngname = nm.name
            ' get rid of sheet name, if any.
            if instr(rngname, "!") > 0 then
                excpos = instr(rngname, "!")
                rngname = right(rngname, len(rngname) - excpos)
            end if
            ' want to skip over hidden solver range names.
            if instr(rngname, "solver_") = 0 and instr(rngname, "print_area") = 0 then
                ' the following error check is in case the range name no longer references anything.
                on error resume next
                set namedrange = range(rngname)
                if err.number = 0 then
                    if namedrange.worksheet.name = modelsht.name then
                        if namedrange.cells.count = 1 then
                            if namedrange.address = outputcell.address then
                                outputname(i) = rngname
                                isinnamedrange = true
                            end if
                        else
                            index = 0
                            for each cell in namedrange.cells
                                index = index + 1
                                if cell.address = outputcell.address then
                                    outputname(i) = rngname & "_" & index
                                    isinnamedrange = true
                                    exit for
                                end if
                            next
                        end if
                    end if
                end if
                on error goto 0
            end if
            if isinnamedrange then exit for
        next
    next
end sub

function stsindex() as integer
    ' changed 12/5/2011
    ' now it finds the first "sts_index" sheet that doesn't exist.
    dim ws as worksheet, index as integer
    index = 1
    on error resume next
    do
        ' error is 0 if sheet with this index exists.
        set ws = worksheets("sts_" & index)
        if err = 0 then index = index + 1
    loop until err <> 0
    on error goto 0
    stsindex = index
end function

sub setupstssheet()
    dim nextindex as integer
    dim i as integer ' counter of one-way input value
    dim i1 as integer, i2 as integer ' counters of two-way input values
    dim j as integer ' counter of output cells
    dim coloffset as integer, rowoffset as integer
    dim chttitlecell(1 to 2) as range
    
    set stssht = worksheets.add(after:=worksheets(worksheets.count))
    nextindex = stsindex
    activesheet.name = "sts_" & nextindex
    
    if tabletype = oneway then
        with range("a1")
            .value = "oneway analysis for solver model in " & modelshtname & " worksheet"
            .font.bold = true
        end with
        
        range("a3").value = inputname(1) & " (cell " & inputcell(1).address & _
            ") values along side, output cell(s) along top"
        
        with range("a4")
            ' headings for input values
            for i = 1 to ninputvals(1)
                .offset(i, 0).value = inputval1(i)
                .offset(i, 0).numberformat = inputcell(1).numberformat
            next
            activesheet.names.add name:="inputvalues", _
                refersto:=range(.offset(1, 0), .offset(ninputvals(1), 0))
                
            ' headings for outputs
            for j = 1 to noutputcells
                with .offset(0, j)
                    .value = outputname(j)
                    .orientation = 90
                    .horizontalalignment = xlright
                end with
            next
            
            activesheet.names.add name:="outputaddresses", _
                refersto:=range(.offset(0, 1), .offset(0, outputrange.cells.count))
            activesheet.names.add name:="outputvalues", _
                refersto:=range(.offset(1, 1), .offset(ninputvals(1), outputrange.cells.count))
            range(.offset(1, 1), .offset(ninputvals(1), outputrange.cells.count)) _
                .borderaround xlcontinuous, xlthin, 0
            
            ' place choicecell at least as far to right as column k.
            if outputrange.cells.count <= 8 then
                coloffset = 10
            else
                coloffset = outputrange.cells.count + 2
            end if
            set choicecell(1) = .offset(0, coloffset)
        end with
        
        with choicecell(1)
            set chttitlecell(1) = .offset(-3, 0)
            .orientation = 90
            .interior.themecolor = xlthemecoloraccent6
            .interior.tintandshade = 0.599993896298105
            .offset(-1, 0).value = "data for chart"
            .validation.add type:=xlvalidatelist, formula1:="=outputaddresses"
            .value = range("b4").value
            .horizontalalignment = xlright
            .offset(0, -1).formula = "=match(" & choicecell(1).address & ",outputaddresses,0)"
            .offset(0, -1).font.color = vbwhite
            for i = 1 to ninputvals(1)
                .offset(i, 0).formula = "=index(outputvalues," & i & "," & .offset(0, -1).address & ")"
            next
            activesheet.names.add name:="chartdata", _
                refersto:=range(.offset(1, 0), .offset(ninputvals(1), 0))
        end with
        
        with chttitlecell(1)
            .formula = "=concatenate(""sensitivity of ""," _
                & choicecell(1).address & ","" to "",""" & inputname(1) & """)"
            .font.color = vbwhite
        end with
        
        createchart activesheet.name & "_chart", chttitlecell(1), _
            inputname(1) & " (" & inputcell(1).address & ")", _
            iif(minmaxopt(1), xllinemarkers, xlxyscatterlines), _
            range("inputvalues"), range("chartdata"), _
            coloffset * 48, (4 + ninputvals(1)) * 15 + choicecell(1).height, 384, 225
                   
        createtextbox "when you select an output from the dropdown " _
            & "list in cell " & choicecell(1).address & ", the chart will adapt to that output.", _
            choicecell(1).left + 2 * 48, 3 * 15, 4 * 48, 4 * 15
                
    else ' twoway
        ' store output names in range way to the right for use in hidden cells
        ' to the left of the choice cells.
        with range("az1")
            for j = 1 to noutputcells
                .offset(j, 0).value = outputname(j)
            next
            activesheet.names.add name:="outputaddresses", _
                refersto:=range(.offset(1, 0), .offset(outputrange.cells.count, 0))
        end with
        
        ' headings for tables of outputs
        for j = 1 to noutputcells
            with range("a4").offset((j - 1) * (ninputvals(1) + 2), 0)
                for i1 = 1 to ninputvals(1)
                    with .offset(i1, 0)
                        .value = inputval1(i1)
                        .numberformat = inputcell(1).numberformat
                    end with
                next
                for i2 = 1 to ninputvals(2)
                    with .offset(0, i2)
                        .value = inputval2(i2)
                        .numberformat = inputcell(2).numberformat
                    end with
                next
                if j = 1 then
                    activesheet.names.add name:="inputvalues1", _
                        refersto:=range(.offset(1, 0), .offset(ninputvals(1), 0))
                    activesheet.names.add name:="inputvalues2", _
                        refersto:=range(.offset(0, 1), .offset(0, ninputvals(2)))
                end if
                .value = outputname(j)
                .horizontalalignment = xlright
                activesheet.names.add name:="outputvalues_" & j, _
                    refersto:=range(.offset(1, 1), .offset(ninputvals(1), ninputvals(2)))
                range(.offset(1, 1), .offset(ninputvals(1), ninputvals(2))) _
                    .borderaround xlcontinuous, xlthin, 0
            end with
        next
        
        with range("a1")
            .entirecolumn.autofit
            .value = "twoway analysis for solver model in " & modelshtname & " worksheet"
            .font.bold = true
        end with
        
        range("a3").value = inputname(1) & " (cell " & inputcell(1).address & ") values along side, " _
            & inputname(2) & " (cell " & inputcell(2).address & ") values along top, output cell in corner"
            
        ' place first choicecell at least as far to right as column k.
        with range("a4")
            if ninputvals(2) <= 8 then
                coloffset = 10
            else
                coloffset = ninputvals(2) + 2
            end if
            set choicecell(1) = .offset(0, coloffset)
            set choicecell(2) = .offset(0, coloffset + 4)
        end with

        with choicecell(1)
            set chttitlecell(1) = .offset(-3, 0)
            .orientation = 90
            .interior.themecolor = xlthemecoloraccent6
            .interior.tintandshade = 0.599993896298105
            .offset(-2, 0).value = "output and " & inputname(1) & " value for chart"
            .offset(-1, 0).value = "output"
            .offset(-1, 1).value = inputname(1) & " value"
            .validation.add type:=xlvalidatelist, formula1:="=outputaddresses"
            .offset(0, 1).validation.add type:=xlvalidatelist, formula1:="=inputvalues1"
            .value = outputname(1)
            .offset(0, 1).value = inputval1(1)
            .offset(0, 1).interior.themecolor = xlthemecoloraccent3
            .offset(0, 1).interior.tintandshade = 0.599993896298105
            .horizontalalignment = xlright
            .offset(0, -1).formula = "=match(" & .address & ",outputaddresses,0)"
            .offset(0, -1).font.color = vbwhite
            .offset(1, -1).formula = "=""outputvalues_""&" & .offset(0, -1).address & ""
            .offset(1, -1).font.color = vbwhite
            .offset(0, 2).formula = "=match(" & .offset(0, 1).address & ",inputvalues1,0)"
            .offset(0, 2).font.color = vbwhite
            for i2 = 1 to ninputvals(2)
                .offset(i2, 0).formula = "=index(indirect(" & .offset(1, -1).address & ")," _
                    & .offset(0, 2).address & "," & i2 & ")"
            next
            activesheet.names.add name:="chartdata1", _
                refersto:=range(.offset(1, 0), .offset(ninputvals(2), 0))
        end with

        with choicecell(2)
            set chttitlecell(2) = .offset(-3, 0)
            .orientation = 90
            .interior.themecolor = xlthemecoloraccent6
            .interior.tintandshade = 0.599993896298105
            .offset(-2, 0).value = "output and " & inputname(2) & " value for chart"
            .offset(-1, 0).value = "output"
            .offset(-1, 1).value = inputname(2) & " value"
            .validation.add type:=xlvalidatelist, formula1:="=outputaddresses"
            .offset(0, 1).validation.add type:=xlvalidatelist, formula1:="=inputvalues2"
            .value = outputname(1)
            .offset(0, 1).value = inputval2(1)
            .offset(0, 1).interior.themecolor = xlthemecoloraccent3
            .offset(0, 1).interior.tintandshade = 0.599993896298105
            .horizontalalignment = xlright
            .offset(0, -1).formula = "=match(" & .address & ",outputaddresses,0)"
            .offset(0, -1).font.color = vbwhite
            .offset(1, -1).formula = "=""outputvalues_""&" & .offset(0, -1).address & ""
            .offset(1, -1).font.color = vbwhite
            .offset(0, 2).formula = "=match(" & .offset(0, 1).address & ",inputvalues2,0)"
            .offset(0, 2).font.color = vbwhite
            for i1 = 1 to ninputvals(1)
                .offset(i1, 0).formula = "=index(indirect(" & .offset(1, -1).address & ")," _
                    & i1 & "," & .offset(0, 2).address & ")"
            next
            activesheet.names.add name:="chartdata2", _
                refersto:=range(.offset(1, 0), .offset(ninputvals(1), 0))
        end with

        if ninputvals(1) > ninputvals(2) then
            rowoffset = ninputvals(1) + 4
        else
            rowoffset = ninputvals(2) + 4
        end if
        coloffset = ninputvals(2) + 2
        
        with chttitlecell(1)
            .formula = "=concatenate(""sensitivity of ""," _
                & choicecell(1).address & ","" to "",""" & inputname(2) & """)"
            .font.color = vbwhite
        end with
        with chttitlecell(2)
            .formula = "=concatenate(""sensitivity of ""," _
                & choicecell(2).address & ","" to "",""" & inputname(1) & """)"
            .font.color = vbwhite
        end with
        
        createchart activesheet.name & "_chart1", chttitlecell(1), _
            inputname(2) & " (" & inputcell(2).address & ")", _
            iif(minmaxopt(2), xllinemarkers, xlxyscatterlines), _
            range("inputvalues2"), range("chartdata1"), _
            coloffset * 48 + range("a1").width, rowoffset * 15 + choicecell(1).height, 384, 225
        createchart activesheet.name & "_chart2", chttitlecell(2), _
            inputname(1) & " (" & inputcell(1).address & ")", _
            iif(minmaxopt(1), xllinemarkers, xlxyscatterlines), _
            range("inputvalues1"), range("chartdata2"), _
            48 + 384 + coloffset * 48 + range("a1").width, rowoffset * 15 + choicecell(2).height, 384, 225
                   
        createtextbox "by making appropriate selections in cells " _
            & choicecell(1).address & ", " & choicecell(1).offset(0, 1).address & ", " _
            & choicecell(2).address & ", and " & choicecell(2).offset(0, 1).address _
            & ", you can chart any row (in left chart) or column (in right chart) " _
            & "of any table to the left.", _
            choicecell(2).left + 4 * 48, 3 * 15, 6 * 48, 7 * 15
    end if
end sub

sub createchart(chtname as string, chttitlecell as range, haxistitle as string, _
        chttype as xlcharttype, haxislabels as range, datavals as range, _
        left as double, top as double, width as double, height as double)
    dim chtobj as chartobject
    dim shp as shape
    dim cht as chart
    dim ser as series
    dim haxis as axis
    
    set chtobj = activesheet.chartobjects.add(left, top, width, height)
    chtobj.name = chtname
    set shp = activesheet.shapes(chtname)
    with shp
        .placement = xlfreefloating
        with .line
            .visible = msotrue
            .forecolor.objectthemecolor = msothemecoloraccent1
            .forecolor.tintandshade = 0
            .forecolor.brightness = 0
            .transparency = 0
            .weight = 1.25
        end with
    end with
    
    set cht = chtobj.chart
    with cht
        .charttype = chttype
        .setsourcedata datavals
        .haslegend = false
        .hastitle = true
        .charttitle.characters.font.size = 12
        .charttitle.formula = "=" & stssht.name & "!" & chttitlecell.address
    end with
    
    set haxis = cht.axes(xlcategory)
    with haxis
        .hastitle = true
        .axistitle.text = haxistitle
    end with
    
    set ser = cht.seriescollection(1)
    ser.xvalues = haxislabels
end sub

sub createtextbox(msg as string, left as double, top as double, width as double, height as double)
    dim shp as shape
    set shp = activesheet.shapes.addtextbox(msotextorientationhorizontal, left, top, width, height)
    shp.textframe.characters.text = msg
    with shp.line
        .visible = msotrue
        .forecolor.objectthemecolor = msothemecoloraccent1
        .forecolor.tintandshade = 0
        .forecolor.brightness = 0
        .transparency = 0
        .weight = 1.25
    end with
    shp.placement = xlfreefloating
end sub

sub definesolvermessages()
    on error goto proc_error
    
    ' these are message solver returns, indexed by the integer returned by the
    ' solversolve function. previous messages indexed by -1 and 12 are gone in solver for excel 2010/2013.
    solvermessage(0) = "solver found a solution. all constraints and optimality conditions are satisfied."
    solvermessage(1) = "solver has converged to the current solution. all constraints are satisfied."
    solvermessage(2) = "solver cannot improve the current solution. all constraints are satisfied."
    solvermessage(3) = "stop chosen when the maximum iteration limit was reached."
    solvermessage(4) = "the objective cell values do not converge."
    solvermessage(5) = "solver could not find a feasible solution."
    solvermessage(6) = "solver stopped at user's request."
    solvermessage(7) = "the linearity conditions required by this lp solver are not satisfied."
    solvermessage(8) = "the problem is too large for solver to handle."
    solvermessage(9) = "solver encountered an error value in a target or constraint cell."
    solvermessage(10) = "stop chosen when maximum time limit was reached."
    solvermessage(11) = "there is not enough memory available to solve the problem."
    solvermessage(13) = "error in model. please verify that all cells and constraints are valid."
    solvermessage(14) = "solver found an integer solution within tolerance. all constraints are satisfied."
    solvermessage(15) = "stop chosen when the maximum number of feasible [integer] solutions was reached."
    solvermessage(16) = "stop chosen when the maximum number of feasible [integer] subproblems was reached."
    solvermessage(17) = "solver converged in probability to a global solution."
    solvermessage(18) = "all variables must have both upper and lower bounds."
    solvermessage(19) = "variable bounds conflict in binary or alldifferent constraint."
    solvermessage(20) = "lower and upper bounds on variables allow no feasible solution."
    
    ' these are error messages for the possible errors that can occur.  these are
    ' used as cell notes.
    'solvererrormessage(3) = "stop"
    solvererrormessage(4) = "no convergence"
    solvererrormessage(5) = "not feasible"
    'solvererrormessage(6) = "stopped"
    solvererrormessage(7) = "not linear"
    solvererrormessage(8) = "too large"
    solvererrormessage(9) = "error value"
    'solvererrormessage(10) = "stopped"
    solvererrormessage(11) = "memory"
    solvererrormessage(13) = "error"
    'solvererrormessage(15) = "stopped"
    'solvererrormessage(16) = "stopped"
    solvererrormessage(18) = "no bounds"
    solvererrormessage(19) = "bounds conflict"
    solvererrormessage(20) = "not feasible"
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlsolvertable", "definesolvermessages", err.number, err.description)
    goto proc_exit
end sub

sub runsolver()
    dim i as integer ' counter of one-way input value
    dim i1 as integer, i2 as integer ' counters of two-way input values
    dim cellcount as integer
    dim solvercode as integer
    dim cell as range
    dim message as string
    dim resultcell as range
    
    on error goto proc_error
    
    application.screenupdating = false
    modelsht.activate
    set resultcell = stssht.range("a4")
    
    ' find current solver settings (stripping off beginning equals sign)
    engine = right(activesheet.names("solver_eng").value, len(activesheet.names("solver_eng").value) - 1)
    precision = right(activesheet.names("solver_pre").value, len(activesheet.names("solver_pre").value) - 1)
    
    ' get multistart setting (1 for on, 2 for off) if engine is grg nonlinear
    ' if engine = 1 then
    '    ms_setting = val(right(activesheet.names("solver_msl").value, 1))
    '    if msgbox("do you want to use the multistart option with the grg nonlinear solver? " _
    '        & "this is useful in some models to avoid stopping at local (as opposed to global) " _
    '        & "optimal solutions.", vbyesno, "use multistart") = vbyes then usems = true
    ' end if
    
    if tabletype = oneway then
            
        ' solve the problem for each input value.
        for i = 1 to ninputvals(1)
            inputcell(1).value = inputval1(i)
            optimize solvercode
            
            ' if a "good" message occurs (including stopping), record the results in the table.
            with resultcell
                if solvercode = 0 or solvercode = 1 or solvercode = 2 _
                        or solvercode = 14 or solvercode = 17 _
                        or solvercode = 3 or solvercode = 6 _
                        or solvercode = 10 or solvercode = 15 _
                        or solvercode = 16 then
                    cellcount = 0
                    for each cell in outputrange
                        cellcount = cellcount + 1
                        with .offset(i, cellcount)
                            .value = cell.value
                            .numberformat = cell.numberformat
                        end with
                    next
                
                ' otherwise, for a "bad" result, display the solver error message.
                else
                    with .offset(i, 1)
                        .value = solvererrormessage(solvercode)
                        .interior.colorindex = 40
                    end with
                end if
                
                ' in any case, display the solver message as a note.
                .offset(i, 1).notetext solvermessage(solvercode)
            end with
        next
    
    ' repeat the same operations for a twoway table.
    else
        
        ' solve the problem and record results for each pair of input values.
        for i2 = 1 to ninputvals(2)
            inputcell(2).value = inputval2(i2)
            for i1 = 1 to ninputvals(1)
                inputcell(1).value = inputval1(i1)
                optimize solvercode
                
                cellcount = 0
                for each cell in outputrange
                    cellcount = cellcount + 1
                    with resultcell.offset((cellcount - 1) * (ninputvals(1) + 2) + i1, i2)
                        if solvercode = 0 or solvercode = 1 or solvercode = 2 _
                                or solvercode = 14 or solvercode = 17 _
                                or solvercode = 3 or solvercode = 6 _
                                or solvercode = 10 or solvercode = 15 _
                                or solvercode = 16 then
                            .value = cell.value
                            .numberformat = cell.numberformat
                        else
                            .value = solvererrormessage(solvercode)
                            .interior.colorindex = 40
                        end if
                        .notetext solvermessage(solvercode)
                    end with
                    with resultcell.offset((cellcount - 1) * (ninputvals(1) + 2), 0)
                        range(.offset(1, 1), .offset(ninputvals(1), ninputvals(2))) _
                            .borderaround xlcontinuous, xlthin, 0
                    end with
                next
            next
        next
    end if
    
    ' restore original multistart setting if grg nonlinear is used
    'if engine = 1 then
    '    if ms_setting = 2 then
    '        solveroptions multistart:=false, requirebounds:=false
    '    else
    '        solveroptions multistart:=true, requirebounds:=false
    '    end if
    'end if
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlsolvertable", "runsolver", err.number, err.description)
    goto proc_exit
end sub

sub optimize(solvercode as integer)
#if mac then
    call optimizemac(solvercode)
#else
    call optimizewin(solvercode)
#end if
    
end sub

sub optimizemac(solvercode as integer)
msgbox "does not work for mac"
end sub

sub optimizewin(solvercode as integer)
    dim cellcount as integer
    dim cell as range
    
    on error goto proc_error
    
    ' start with original values in changing cells
    restoreoriginalchangingcellvalues
    
    if engine = 2 then ' simplex lp
        'solvercode = solversolve(true)
        solvercode = application.run("solver.xlam!solversolve", true)
        if solvercode = 7 then ' not linear
            ' try solving as linear model again with larger precision
            activesheet.names("solver_pre").value = 100 * precision
            restoreoriginalchangingcellvalues
            'solvercode = solversolve(true)
            solvercode = application.run("solver.xlam!solversolve", true)
            ' restore precision
            activesheet.names("solver_pre").value = precision
            if solvercode = 7 then ' still not linear
                ' switch to grg with multistart
                activesheet.names("solver_eng").value = 1
                ' solveroptions multistart:=true, requirebounds:=false
                restoreoriginalchangingcellvalues
                'solvercode = solversolve(true)
                solvercode = application.run("solver.xlam!solversolve", true)
                'if solvercode = 9 then ' error in model
                '    ' turn off multistart and try again
                '    solveroptions multistart:=false, requirebounds:=false
                '    restoreoriginalchangingcellvalues
                '    'solvercode = solversolve(true)
                '    solvercode = application.run("solver.xlam!solversolve", true)
                'end if
                ' restore engine
                activesheet.names("solver_eng").value = 2
            end if
        end if
    elseif engine = 1 then ' grg nonlinear
        ' either use multistart or don't, but don't require bounds on variables
        'if usems then
        '    solveroptions multistart:=true, requirebounds:=false
        'else
        '    solveroptions multistart:=false, requirebounds:=false
        'end if
        'solvercode = solversolve(true)
        solvercode = application.run("solver.xlam!solversolve", true)
        'if usems and solvercode = 9 then ' error in model
        '    ' turn off multistart and try again
        '    solveroptions multistart:=false, requirebounds:=false
        '    restoreoriginalchangingcellvalues
        '    'solvercode = solversolve(true)
        '    solvercode = application.run("solver.xlam!solversolve", true)
        'end if
    else ' evolutionary
        'solvercode = solversolve(true)
        solvercode = application.run("solver.xlam!solversolve", true)
    end if
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlsolvertable", "optimize", err.number, err.description)
    goto proc_exit
end sub

sub restoreoriginalchangingcellvalues()
    dim cellcount as integer
    dim cell as range
    cellcount = 0
    for each cell in changingcells
        cellcount = cellcount + 1
        cell.value = origchangingcellval(cellcount)
    next
end sub

public sub restoreoriginalvalues()
    dim i as integer
    dim cellcount as integer
    dim cell as range
    
    on error goto proc_error
    
    ' restore the changing cells to their original values.
    restoreoriginalchangingcellvalues
    
    ' restore the original input value(s).
    if tabletype = oneway then
        inputcell(1).value = originputval(1)
    else
        for i = 1 to 2
            inputcell(i).value = originputval(i)
        next
    end if
    
    application.screenupdating = true
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlsolvertable", "restoreoriginalvalues", err.number, err.description)
    goto proc_exit
end sub

public sub setupsettingsranges()
    dim x as variant
    
    ' check whether the solvertable sheet name would be too long.
    if len(modelshtname) > 27 then
        shortmodelshtname = left(modelshtname, 27)
    else
        shortmodelshtname = modelshtname
    end if
    
    ' check whether a worksheet with the settings for this model already exists.
    on error resume next
    x = worksheets(shortmodelshtname & "_sts").name
    
    ' if the sheet doesn't exist, add it, hide it, and activate the model sheet.
    if err <> 0 then
        worksheets.add.move after:=sheets(activeworkbook.sheets.count)
        with activesheet
            .name = shortmodelshtname & "_sts"
            .visible = xlveryhidden
        end with
        worksheets(modelshtname).activate
    end if
        
    ' set the settings ranges, and format some cells as text.
    on error goto proc_error
    with worksheets(shortmodelshtname & "_sts").range("a1")
        set prevonewaysettings = range(.offset(0, 0), .offset(9, 0))
        set prevtwowaysettings = range(.offset(0, 1), .offset(17, 1))
    end with
    prevonewaysettings.cells(8).numberformat = "@"
    with prevtwowaysettings
        .cells(8).numberformat = "@"
        .cells(15).numberformat = "@"
    end with
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlsolvertable", "setupsettingsranges", err.number, err.description)
    goto proc_exit
end sub




