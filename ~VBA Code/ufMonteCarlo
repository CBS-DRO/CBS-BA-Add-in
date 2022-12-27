
private sub chkhisto_click()

end sub

private sub commandbutton1_click()
    dim rformula as range, rsimtrials as range, rresult as range, rhistogram as range
    dim nsimtrials as long
    dim bprinthistogram as boolean, bprinthistogramandvalues as boolean, blabel as boolean
    dim macroname as string
    dim bhistory as boolean
    dim bhist as boolean
    dim btranspose, browoutput as boolean
    dim browinput as boolean
    dim nvars as integer
    dim r, c as integer
    
    ' read and validate the dialog
    if not getrefeditrange(refformula, rformula, "formula") then exit sub
    
    nsimtrials = getsimtrials(rsimtrials)
    if nsimtrials < 0 then exit sub
        
    if not getrefeditrange(refresult, rresult, "result") then exit sub
    bhist = getrefeditrange2(refhistogram, rhistogram, "histogram")
    bprinthistogram = chkhisto.value
    'btranspose = chktranspose.value
    btranspose = false
    ' check input data orientation
    browinput = checkifrowinput(rformula)
    if browinput = true then
        nvars = rformula.columns.count
        browoutput = not btranspose
    else
        'nvars = rformula.rows.count
        'browoutput = btranspose
        msgbox "when simulating multiple variables, they must be contained in a single row of an excel spreadshet. please re-structure your spreadsheet to ensure all your variables are contained in one single row.", vbexclamation
        exit sub
    end if
    
    if not verifyformula(rformula, browinput) then exit sub
        
    ' check histogram input
    if bhist = true then
        if browinput = true and rhistogram.columns.count <> nvars then
            msgbox "when specifying the variables for which you want to output a histogram, the number of columns should be equal to the number of variables (in this case, " & nvars & ").", vbexclamation
            exit sub
        elseif browinput = false and rhistogram.rows.count <> nvars then
            msgbox "when specifying the variables for which you want to output a histogram, the number of rows should be equal to the number of variables (in this case, " & nvars & ").", vbexclamation
            exit sub
        end if
        
        for r = 1 to rhistogram.rows.count
            for c = 1 to rhistogram.columns.count
                if not isnumeric(rhistogram.cells(r, c)) or isempty(rhistogram.cells(r, c)) then
                    msgbox "when specifying the variables for which you want to output a histogram, every entry must be true/false or 1/0.", vbexclamation
                    exit sub
                end if
            next c
        next r
    end if
    
    'check open workbooks
    if checkifotherworkbookopen() then
        unload me
        exit sub
    end if
    ' check output range
    if browoutput = true then
        if not checkifrangeempty(refresult, rresult, "result", 15, nvars + 1) then exit sub
    else
        if not checkifrangeempty(refresult, rresult, "result", nvars + 1, 15) then exit sub
    end if
    
    
    ' close the dialog
    unload me
    
    ' check if calculation mode is auto
    if application.calculation = xlcalculationmanual then
        msgbox "the current calculation mode is manual." & vbnewline & vbnewline & "to run montecarlo simulation, the calculation mode will be changed to automatic.", vbinformation
    end if
    
    ' save input fields
    application.run "saveinputfields", rformula, rsimtrials, rresult, bprinthistogram, activesheet.name, rhistogram, btranspose, nsimtrials
    
    ' invoke xll
    if bhist = true then
        montecarlo.mcsim rformula, rsimtrials, rresult, bprinthistogram, bprinthistogramandvalues, rhistogram, browinput, browoutput, nsimtrials
    else
        montecarlo.mcsim rformula, rsimtrials, rresult, bprinthistogram, bprinthistogramandvalues, nothing, browinput, browoutput, nsimtrials
    end if
end sub

private function getrefeditrange(byval re, byref r as range, byval fieldname) as boolean
    getrefeditrange = false
    on error resume next
    dim addr as string
    addr = re.value
    if err.number = 0 then
        set r = range(addr)
        if err.number = 0 then
            getrefeditrange = true
            exit function
        end if
    end if
    msgbox "please select a range for " & fieldname, vbexclamation
end function

private function getsimtrials(byref r as range) as long
    'start with an error value
    getsimtrials = -1
    
    on error resume next
    
    ' begin by trying to read whatever is in the sim trials ref box
    dim addr as string
    addr = refsimtrials.value
    
    if (err.number <> 0) or (trim(addr) = "") then
        ' something's gone wrong
        msgbox "please input a number or select a range for the number of trials you want to use in this simulation.", vbexclamation
        exit function
    end if
    
    ' treat addr as a formula and see what's in it
    dim addrcontents
    addrcontents = range(addr).value
    
    if err.number <> 0 then
        ' looks like it's not a valid address - so let's treat the
        ' value in the ref box as a digit
        addrcontents = addr
        err.clear
    else
        ' we a valid range - update r
        set r = range(addr)
    end if
    
    ' check whether we have a valid number; note that we cannot combine the
    ' two statements below with "and", because if we do
    '    isnumeric(addrcontents) and clng(addrcontents) > 0
    ' and addrcontents is not numeric, the second statement will error out
    ' and - unbelievably - it will evaluate as true
    if isnumeric(addrcontents) then
        if (clng(addrcontents) > 0) then
            getsimtrials = clng(addrcontents)
            
            if getsimtrials < 40 then
                msgbox "please use at least 40 simulation trials.", vbexclamation
                getsimtrials = -1
            end if
            
            exit function
        end if
    end if
    
    ' something's gone wrong
    msgbox "the number of simulation trials must be a positive, whole number. please double check the range you selected.", vbexclamation
end function

private function getrefeditrange2(byval re, byref r as range, byval fieldname) as boolean
    getrefeditrange2 = false
    on error resume next
    dim addr as string
    addr = re.value
    if err.number = 0 then
        set r = range(addr)
        if err.number = 0 then
            getrefeditrange2 = true
            exit function
        end if
    end if
end function

private function checkifrangeempty(byval re, byref r as range, byval fieldname, byval nrows, byval ncols) as boolean
    checkifrangeempty = false
    dim cellnum as integer, nonemptycellnum as integer
    r.resize(nrows, ncols).select
    nonemptycellnum = worksheetfunction.counta(selection)
    if nonemptycellnum = 0 then
        checkifrangeempty = true
        r.resize(1, 1).select
        exit function
    end if
    
    if msgbox("data exists in the range " & selection.address & "." & vbnewline & vbnewline & "press ""ok"" to continue and press ""cancel"" to reselect the result area", vbokcancel + vbinformation) = 1 then
        checkifrangeempty = true
        r.resize(1, 1).select
        exit function
    else
        r.resize(1, 1).select
        exit function
    end if
end function

private function checkifotherworkbookopen() as boolean
    
    dim otherworkbooksopen as boolean
    otherworkbooksopen = (application.workbooks.count > 1)
    
    if otherworkbooksopen then
        if msgbox("we have detected that another workbook is open. keeping it open while running montecarlo simulation may slow down the simulation." & vbnewline & vbnewline & _
        "press 'ok' to continue anyway and press 'cancel' to close open workbooks and try running montecarlo simulation again.", vbokcancel + vbinformation) = 1 then
            checkifotherworkbookopen = false
            exit function
        else
            checkifotherworkbookopen = true
            exit function
        end if
    else
        checkifotherworkbookopen = false
    end if
end function

private function checkifrowinput(byref rformula as range) as boolean
    dim nrows, ncols as integer
    nrows = rformula.rows.count
    ncols = rformula.columns.count
    if nrows = 1 and ncols = 1 then
        checkifrowinput = true
    elseif nrows = 1 and ncols = 2 then
        if not isnumeric(rformula.cells(1, 1).value) then
            checkifrowinput = false
        else
            checkifrowinput = true
        end if
    elseif nrows = 2 and ncols = 1 then
        if not isnumeric(rformula.cells(1, 1).value) then
            checkifrowinput = true
        else
            checkifrowinput = false
        end if
    elseif nrows = 2 and ncols = 2 then
        if not isnumeric(rformula.cells(2, 1).value) then
            checkifrowinput = false
        else
            checkifrowinput = true
        end if
    elseif ncols = 1 or ncols = 2 then
        checkifrowinput = false
    else
        checkifrowinput = true
    end if
end function

private function verifyformula(byref rformula as range, byval browinput as boolean) as boolean
    if browinput = true and rformula.rows.count > 2 then
        msgbox "number of rows in formula cannot exceed two", vbexclamation
        verifyformula = false
    elseif browinput = false and rformula.columns.count > 2 then
        msgbox "number of columns in formula cannot exceed two", vbexclamation
        verifyformula = false
    else
        verifyformula = true
    end if
end function

private sub label10_click()

end sub

private sub label12_click()

end sub