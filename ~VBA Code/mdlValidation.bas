public const c_inputcell = 1
public const c_inputrange = 2
public const c_outputrange = 3

public sub validatenumber(txtbox as msforms.textbox, isvalid as boolean, _
        optional ispositive as boolean = false)
    ' purpose:  takes a textbox and checks whether the value is numeric (and positive
    '           if the optional parameter is true)
    ' returns:  isvalid, true only for a valid entry
    
    on error goto proc_error
    dim message as string
    isvalid = true
    with txtbox
        if not isnumeric(.text) then
            isvalid = false
            message = "enter a numeric value in this box."
        elseif ispositive and .text <= 0 then
            isvalid = false
            message = "enter a positive value in this box."
        end if
    end with
    if not isvalid then
        msgbox message, vbexclamation, "invalid entry"
        with txtbox
            .selstart = 0
            .sellength = len(.text)
            .setfocus
        end with
    end if
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlvalidation", "validatenumber", err.number, err.description)
    goto proc_exit
end sub

public sub validaterange(refbox as control, isvalid as boolean, _
        rangetype as integer)
    on error goto proc_error
    dim message as string
    dim selectedrange as range
    
    isvalid = true
    on error resume next
    set selectedrange = range(trim(refbox.text))
    if err <> 0 then
        message = "enter a cell address (or range name) in this box."
        isvalid = false
    else
        select case rangetype
            case c_inputcell
                if selectedrange.cells.count > 1 then
                    message = "enter a single input cell."
                    isvalid = false
                elseif selectedrange.hasformula then
                    message = "the input cell should not contain a formula."
                    isvalid = false
                elseif not isnumeric(selectedrange.value) then
                    message = "the input cell should be numeric, not a label."
                    isvalid = false
                end if
                
            case c_inputrange
                for each cell in selectedrange
                    if cell.value = "" then
                        message = "don't include any empty " _
                            & "cells in the input range."
                        isvalid = false
                        exit for
                    end if
                    if cell.hasformula then
                        message = "the input value cell should not contain a formula."
                        isvalid = false
                        exit for
                    end if
                    if not isnumeric(cell.value) then
                        message = "the input value cells should all be numeric."
                        isvalid = false
                        exit for
                    end if
                next
            
            case c_outputrange
                for each cell in selectedrange
                    if cell.value = "" then
                        message = "don't include any empty " _
                            & "cells in the output range."
                        isvalid = false
                        exit for
                    end if
                next
        end select
    end if
    
    if not isvalid then
        msgbox message, vbexclamation, "invalid entry"
        with refbox
            .selstart = 0
            .sellength = len(.text)
            .setfocus
        end with
    end if
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlvalidation", "validaterange", err.number, err.description)
    goto proc_exit
end sub

public sub validatelist(txtbox as msforms.textbox, isvalid as boolean)
    on error goto proc_error
    dim message as string
    dim startposition as integer
    dim commaposition as integer
    dim txtboxstring as string
    dim txtboxnumber as string
    
    isvalid = true
    txtboxstring = txtbox.text
    startposition = 1
    do
        commaposition = instr(startposition, txtboxstring, ",")
        if commaposition = 0 then
            txtboxnumber = trim(mid(txtboxstring, startposition, len(txtboxstring)))
        else
            txtboxnumber = trim(mid(txtboxstring, startposition, commaposition - startposition))
            startposition = commaposition + 1
        end if
        
        if not isnumeric(txtboxnumber) then
            isvalid = false
            message = "enter a list of numbers, separated by commas."
            exit do
        end if
    loop until commaposition = 0
    
    if not isvalid then
        msgbox message, vbexclamation, "invalid entry"
        with txtbox
            .selstart = 0
            .sellength = len(.text)
            .setfocus
        end with
    end if
    
proc_exit:
    exit sub
    
proc_error:
    call showerror("mdlvalidation", "validatetxtboxstring", err.txtboxnumber, err.description)
    goto proc_exit
end sub


