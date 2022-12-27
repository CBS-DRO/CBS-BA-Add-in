


option explicit

private sub cmdcancel_click()
    on error goto proc_error
    unload me
    formcancel = true
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("frmoneway", "cmdcancel_click", err.number, err.description)
    goto proc_exit
end sub

private sub cmdok_click()
    on error goto proc_error
    dim isvalid as boolean
    dim tempmin as single
    dim tempmax as single
    dim count as integer
    dim message as string
    dim cell as range
    dim inputvaluerange as range
    
    dim inputstring as string
    dim nvals as integer
    dim startposition as integer
    dim commaposition as integer
    
    formcancel = false
    
    ' check validity of refinputcell.
    call validaterange(refinputcell, isvalid, c_inputcell)
    if not isvalid then
        goto proc_exit
    else
        set inputcell(1) = range(trim(refinputcell.text))
    end if
    
    ' name of input
    if txtinputname.text = "" then
        inputname(1) = "input"
    else
        inputname(1) = txtinputname.text
    end if
    
    ' check the validity of the input values.
    if optminmax.value then
        minmaxopt(1) = true
        
        ' check validity of txtminval and txtmaxval.
        call validatenumber(txtminval, isvalid)
        if not isvalid then
            goto proc_exit
        else
            tempmin = trim(txtminval.text)
        end if
        
        call validatenumber(txtmaxval, isvalid)
        if not isvalid then
            goto proc_exit
        else
            tempmax = trim(txtmaxval.text)
        end if
        
        if tempmin > tempmax then
            msgbox "the minimum input value should not be greater than the " _
                & "maximum input value. try again.", vbexclamation, "invalid values"
            with txtminval
                .selstart = 0
                .sellength = len(.text)
                .setfocus
            end with
            goto proc_exit
        else
            minval(1) = trim(txtminval.text)
            maxval(1) = trim(txtmaxval.text)
        end if
        
        ' check the validity of txtincr
        call validatenumber(txtincr, isvalid, true)
        if not isvalid then
            goto proc_exit
        else
            increment(1) = trim(txtincr.text)
        end if
        
        ' calculate and store the input values.
        count = 1
        do until minval(1) + count * increment(1) > maxval(1) + epsilon
            count = count + 1
        loop
        ninputvals(1) = count
        redim inputval1(1 to ninputvals(1))
        
        for count = 1 to ninputvals(1)
            inputval1(count) = minval(1) + (count - 1) * increment(1)
        next
        
    elseif optlistrange.value then
        minmaxopt(1) = false
        
        ' the user chose to get a list of input values from a range.
        ' first, check that the validity of the selected range.
        call validaterange(refinputrange, isvalid, c_inputrange)
        if not isvalid then
            goto proc_exit
        else
            ' get the input values
            set inputvaluerange = range(trim(refinputrange.text))
            ninputvals(1) = inputvaluerange.cells.count
            redim inputval1(1 to ninputvals(1))
            count = 1
            for each cell in inputvaluerange
                inputval1(count) = cell.value
                count = count + 1
            next
        end if
        
    elseif optlistvals.value then
        minmaxopt(1) = false
        
        ' the user chose to enter a list of input values, separated by commas.
        ' first, check that this is a valid list of numerical values.
        call validatelist(txtinputvals, isvalid)
        if not isvalid then
            goto proc_exit
        else
            ' parse the string into a set of numeric values and store them.
            inputstring = trim(txtinputvals.text)
            nvals = 0
            startposition = 1
            do
                nvals = nvals + 1
                redim preserve inputval1(1 to nvals)
                commaposition = instr(startposition, inputstring, ",")
                if commaposition > 0 then
                    inputval1(nvals) = _
                        trim(mid(inputstring, startposition, commaposition - startposition))
                    startposition = commaposition + 1
                else
                    inputval1(nvals) = trim(mid(inputstring, startposition, len(inputstring)))
                end if
            loop until commaposition = 0
            ninputvals(1) = nvals
        end if
    end if
    
    ' check the validity of the output cells.
    call validaterange(refoutputrange, isvalid, c_outputrange)
    if not isvalid then
        goto proc_exit
    else
        set outputrange = range(trim(refoutputrange.text))
        ' check whether the output cell range includes the input cell.
        if union(outputrange, inputcell(1)).cells.count = _
                outputrange.cells.count then
            message = "the input cell shouldn't be part of the output range."
            msgbox message, vbexclamation, "invalid entry"
            with refinputcell
                .selstart = 0
                .sellength = len(.text)
                .setfocus
            end with
            goto proc_exit
        end if
    end if
    
    ' store the settings.
    with prevonewaysettings
        .cells(1).value = 1
        .cells(2).value = inputcell(1).address
        if optminmax.value then
            .cells(3).value = 1
            .cells(4).value = trim(txtminval.text)
            .cells(5).value = trim(txtmaxval.text)
            .cells(6).value = trim(txtincr.text)
        elseif optlistrange.value then
            .cells(3).value = 2
            .cells(7).value = trim(refinputrange.text)
        elseif optlistvals.value then
            .cells(3).value = 3
            .cells(8).value = trim(txtinputvals.text)
        end if
        .cells(9).value = outputrange.address
        .cells(10).value = inputname(1)
    end with
    
    unload me
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("frmoneway", "cmdok_click", err.number, err.description)
    goto proc_exit
end sub

private sub optminmax_click()
    with txtminval
        .enabled = true
        .setfocus
    end with
    txtmaxval.enabled = true
    txtincr.enabled = true
    refinputrange.enabled = false
    txtinputvals.enabled = false
end sub

private sub optlistrange_click()
    with refinputrange
        .enabled = true
        .setfocus
    end with
    txtminval.enabled = false
    txtmaxval.enabled = false
    txtincr.enabled = false
    txtinputvals.enabled = false
end sub

private sub optlistvals_click()
    with txtinputvals
        .enabled = true
        .setfocus
    end with
    txtminval.enabled = false
    txtmaxval.enabled = false
    txtincr.enabled = false
    refinputrange.enabled = false
end sub

private sub userform_initialize()
    on error goto proc_error
    with prevonewaysettings
        if .cells(1).value = 1 then
            refinputcell.text = .cells(2).value
            if .cells(3).value = 1 then
                optminmax.value = true
                txtminval.enabled = true
                txtmaxval.enabled = true
                txtincr.enabled = true
                txtminval.text = .cells(4).value
                txtmaxval.text = .cells(5).value
                txtincr.text = .cells(6).value
                refinputrange.text = ""
                refinputrange.enabled = false
                txtinputvals.text = ""
                txtinputvals.enabled = false
            elseif .cells(3).value = 2 then
                optlistrange.value = true
                txtminval.text = ""
                txtmaxval.text = ""
                txtincr.text = ""
                txtminval.enabled = false
                txtmaxval.enabled = false
                txtincr.enabled = false
                refinputrange.enabled = true
                refinputrange.text = .cells(7).value
                txtinputvals.enabled = false
                txtinputvals.text = ""
            elseif .cells(3).value = 3 then
                optlistvals.value = true
                txtminval.text = ""
                txtmaxval.text = ""
                txtincr.text = ""
                txtminval.enabled = false
                txtmaxval.enabled = false
                txtincr.enabled = false
                refinputrange.text = ""
                refinputrange.enabled = false
                txtinputvals.enabled = true
                txtinputvals.text = .cells(8).value
            end if
            refoutputrange.text = .cells(9).value
            txtinputname.text = .cells(10).value
        else
            refinputcell.text = ""
            optminmax.value = true
            with txtminval
                .enabled = true
                .text = ""
            end with
            with txtmaxval
                .enabled = true
                .text = ""
            end with
            with txtincr
                .enabled = true
                .text = ""
            end with
            with refinputrange
                .text = ""
                .enabled = false
            end with
            with txtinputvals
                .text = ""
                .enabled = false
            end with
            refoutputrange.text = ""
            txtinputname.text = ""
        end if
    end with
    
    refinputcell.setfocus
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("frmoneway", "userform_initialize", err.number, err.description)
    goto proc_exit
end sub

private sub userform_queryclose(cancel as integer, closemode as integer)
    if closemode = vbformcontrolmenu then
        formcancel = true
    end if
end sub