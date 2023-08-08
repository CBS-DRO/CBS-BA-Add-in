
option explicit

sub registermcsim()
application.macrooptions macro:="mcsim", description:= _
    "runs a simple monte carlo simulation in a spreadheet. ", shortcutkey:="t"
end sub

public sub mcsim(rformula as range, rsimtrials as range, rresult as range, bprinthistogram as boolean, _
bprinthistogramandvalues as boolean, rhistogram as range, browinput as boolean, browoutput as boolean, nsimtrials as long)

' mcsim runs a simple simulation in a spreadsheet
'
' to run a simulation
' 1. put any number of formulas (nformula) to simulate in a row
' 2. put the number of simulation trials in a column underneath the first formula
' 3. if you wish to compute percentiles, enter the percentiles in the first column starting in the 9th row
' 3. select the region with nformula+1 columns and 8 or more rows.
'    the region should start one cell to the left of the first formula
' 4. run the simulation by pressing ctrl-shift-m or use tools | macro | run mcsim
'
' programmer: mark broadie
'
' this vba code is adapted from montecarlito created by martin auer.
' see www.montecarlito.com

dim ntrial as long                ' number of simulation trials
dim blnhideapplication as boolean ' true means hide excel for faster performance
dim nformula as integer           ' number of formulas to simulate
dim intbatchsize as integer       ' number of trials before updating display
dim dblstarttime as double        ' timing variables
dim dblfinishtime as double
dim dblelapsetime as double
dim results() as variant          ' variables to store results
dim avgvec() as variant
dim stddevvec() as variant
dim stderrvec() as variant
dim minvec() as variant
dim maxvec() as variant
dim i as long, j as long       ' looping / temporary variables
dim tempvec() as variant
dim percentiles() as variant      'for computing percentiles
dim percentilevalues() as variant
dim npercentile as integer
dim rvariablenames as range

dim onlyzeroone() as boolean     ' for each variable, contains true if the simulation
                                ' output only contains 0s and 1s, for a nicer hisogram

dblstarttime = timer

nformula = rformula.columns.count
if rsimtrials is nothing then
    ntrial = nsimtrials
else
    ntrial = rsimtrials(1, 1)
end if
if (ntrial <= 2) then
   ntrial = 2
end if
    
redim avgvec(1 to nformula)
redim stddevvec(1 to nformula)
redim stderrvec(1 to nformula)
redim minvec(1 to nformula)
redim maxvec(1 to nformula)
redim onlyzeroone(1 to nformula)
redim results(1 to ntrial, 1 to nformula)
redim tempvec(1 to ntrial)
    
' set up for computing percentiles
npercentile = 7
redim percentiles(1 to npercentile)
redim percentilevalues(1 to npercentile, 1 to nformula)
percentiles(1) = 0.01
percentiles(2) = 0.05
percentiles(3) = 0.1
percentiles(4) = 0.5
percentiles(5) = 0.9
percentiles(6) = 0.95
percentiles(7) = 0.99
    
' set intbatchsize
intbatchsize = 100
if ntrial <= 100 then
   intbatchsize = 20
end if
if ntrial > 10000 then
   intbatchsize = int(ntrial / 100)
end if

'get variable names
set rvariablenames = util.getheaderrowsandresizerange(rformula)

'status bar and screen updating
dim bdisplaystatusbarorig as boolean
bdisplaystatusbarorig = application.displaystatusbar
util.turnoffscreenupdate
application.displaystatusbar = true

on error goto errormcsim
' main simulation loop
for i = 1 to ntrial
    
    ' don't display current trial number unless it is a
    ' multiple of intbatchsize or ntrial trials is reached
    doevents
    if (i mod intbatchsize = 0 or i = ntrial) then
        util.updatestatusbar ("simulation trials " & i & "/" & ntrial & "...")
    end if
   
    ' recalculate the spreadsheet and record the results
    util.calculateactiveworksheet
    ' activesheet.calculate
    for j = 1 to nformula
            if iserror(rformula.cells(1, j)) then
                msgbox "formula returned an error. aborting"
                goto cleanexit
            end if
        results(i, j) = rformula.cells(1, j)
    next j
next i
    
' calculate statistics
for j = 1 to nformula
    ' temporary variables
    dim sum as double
    dim min as double
    dim max as double
    dim avg as double
    dim var as double
    dim eps as double
    dim x as double
    dim dx as double
    
    dim binaryonly as boolean
    binaryonly = true
    
    ' first pass over the data to compute the mean, min, and max
    sum = results(1, j)
    min = results(1, j)
    max = results(1, j)
    for i = 2 to ntrial
        x = results(i, j)
        sum = sum + x
        if x > max then
            max = x
        elseif results(i, j) < min then
            min = x
        end if
        
        if (x <> 0) and (x <> 1) then
            binaryonly = false
        end if
    next i
    avg = sum / ntrial
    
    ' second pass over the data to compute the standard deviation
    ' use numerical recipes trick to reduce roundoff error
    eps = 0
    var = 0
    for i = 1 to ntrial
        x = results(i, j)
        dx = x - avg
        eps = eps + dx
        var = var + dx * dx
    next i
    var = (var + eps * eps / ntrial) / (ntrial - 1)
    
    avgvec(j) = avg
    stddevvec(j) = math.sqr(var)
    stderrvec(j) = stddevvec(j) / math.sqr(ntrial)
    minvec(j) = min
    maxvec(j) = max
    onlyzeroone(j) = binaryonly
    
    'compute percentiles, if necessary
    if npercentile > 0 then
        'copy the data to a vector
        matrix2vec tempvec, results, j, ntrial
        for i = 1 to npercentile
            ' compute each percentile
            percentilevalues(i, j) = computepercentile(tempvec, percentiles(i), ntrial)
        next i
    end if
next j
    
dblfinishtime = timer
dblelapsetime = dblfinishtime - dblstarttime
    
' print the result to the spreadsheet
'rngmcout.cells(2, 1) = "number of trials"
'rngmcout.cells(2, 2) = ntrial

rresult.cells(1, 1) = "cpu seconds"
rresult.cells(1, 2) = dblelapsetime

dim svariablenames() as variant
redim svariablenames(1 to nformula)
' svariablenames = rvariablenames.value2

for j = 1 to nformula
    if rvariablenames is nothing then
        svariablenames(j) = "var" & j
    else
        svariablenames(j) = rvariablenames.cells(1, j)
    end if
next j
rresult.cells(2, 1) = "" '"variable name"
call filloutput(2, 2, svariablenames, rresult)

rresult.cells(3, 1) = "average"
call filloutput(3, 2, avgvec(), rresult)

rresult.cells(4, 1) = "standard deviation"
call filloutput(4, 2, stddevvec(), rresult)

rresult.cells(5, 1) = "standard error"
call filloutput(5, 2, stderrvec(), rresult)
    
rresult.cells(6, 1) = "minimum"
call filloutput(6, 2, minvec(), rresult)
    
rresult.cells(7, 1) = "maximum"
call filloutput(7, 2, maxvec(), rresult)
              

rresult.cells(8, 1) = "percentiles"
for i = 1 to npercentile
    rresult.cells(8 + i, 1) = percentiles(i)
    for j = 1 to nformula
        rresult.cells(8 + i, 1 + j) = percentilevalues(i, j)
    next j
next i

' clear all existing formatting
rresult.resize(npercentile + 8, nformula + 1).clearformats

' format the cpu seconds
with rresult.cells(1, 1).resize(1, 2)
    .font.italic = true
    .font.color = rgb(112, 173, 71)
end with

with rresult.cells(1, 1)
    .horizontalalignment = xlright
end with

' format the variable name headers
with rresult.cells(2, 2).resize(1, nformula)
    .font.bold = true
    .horizontalalignment = xlcenter
end with

with rresult.cells(2, 2).resize(1, nformula)
    .borders(xledgetop).linestyle = xlcontinuous
    .borders(xledgetop).weight = xlthin
    
    .borders(xledgebottom).linestyle = xlcontinuous
    .borders(xledgebottom).weight = xlmedium
end with

' format the row headers
with rresult.cells(3, 1).resize(npercentile + 6)
    .horizontalalignment = xlright
    .font.italic = true
end with

' format the percentiles header
with rresult.cells(8, 1)
    .font.bold = true
    .font.color = rgb(256, 0, 0)
end with

' format the cell values
with rresult.cells(3, 2).resize(npercentile + 6, nformula)
    .horizontalalignment = xlcenter
end with

' autofit everything so the columns are wide enough
rresult.resize(npercentile + 8, nformula + 1).columns.autofit

' any cell with non-integer values to be displayed to 3dp
for i = 3 to npercentile + 8
    for j = 2 to nformula + 1
        if isnumeric(rresult.cells(i, j)) then
            if int(rresult.cells(i, j)) <> rresult.cells(i, j) then
                rresult.cells(i, j).numberformat = "0.000"
            end if
        end if
    next j
next i

dim origworksheet as worksheet, ahistogram() as variant
set origworksheet = activesheet
if bprinthistogram then
    dim bpickhistogram as boolean
    if not rhistogram is nothing then
        ahistogram = rhistogram.value2
    end if
    for j = 1 to nformula
        doevents
        util.updatestatusbar "creating histogram " & j & "/" & nformula
        
        if rhistogram is nothing then
            bpickhistogram = true
        else
            bpickhistogram = ahistogram(1, j)
        end if
        if bpickhistogram then
            dim frequency() as long, bin_mid() as double, new_ws_name as string
            dim nbins as long, b as long, statsstr as string, label as string, median as double, skew as double
            
            min = minvec(j)
            max = maxvec(j)
            binaryonly = onlyzeroone(j)
            median = percentilevalues(4, j) ' 50th percentile
            if median - min = 0 then
                skew = 50
            else
                skew = (max - median) / (median - min)
            end if
            if skew > 100 then
                max = percentilevalues(6, j)
            elseif skew < 0.1 then
                min = percentilevalues(2, j)
            end if
            
            calchiststats results, min, max, binaryonly, ntrial, j, frequency, bin_mid, nbins
            new_ws_name = "ba_hist_" & svariablenames(j)
            dim sheet as worksheet, outputrange as range
            if not sheetexists(new_ws_name) then
                set sheet = activeworkbook.sheets.add(after:=activeworkbook.worksheets(activeworkbook.worksheets.count))
                sheet.name = new_ws_name
            else
                set sheet = sheets(new_ws_name)
                sheet.cells.clear
                if sheet.chartobjects.count > 0 then
                    sheet.chartobjects.delete
                end if
            end if
            
            set outputrange = sheet.range("m6:o" & nbins + 6)
            outputrange.cells(1, 1) = "label"
            outputrange.cells(1, 2) = "bin"
            outputrange.cells(1, 3) = "frequency"
            sheet.range("m6:o6").font.bold = true
            
            for b = 1 to nbins
                outputrange.cells(1 + b, 1) = bin_mid(b)
                outputrange.cells(1 + b, 3) = frequency(b)
                if b = 1 then
                    outputrange.cells(1 + b, 2) = "(-inf, " & (bin_mid(1) + bin_mid(2)) / 2 & ")"
                elseif b = nbins then
                    outputrange.cells(1 + b, 2) = "[" & (bin_mid(nbins - 1) + bin_mid(nbins)) / 2 & ", inf)"
                else
                    outputrange.cells(1 + b, 2) = "[" & (bin_mid(b - 1) + bin_mid(b)) / 2 & ", " & (bin_mid(b) + bin_mid(b + 1)) / 2 & ")"
                end if
            next b
            
            statsstr = "average: " & round(avgvec(j), 2) & vbnewline & "std dev: " & round(stddevvec(j), 2) & _
                        vbnewline & "std err: " & round(stderrvec(j), 4)
            sheet.activate
            label = svariablenames(j)
            createhistogram "m7:o" & nbins + 6, statsstr, label
        end if
    next j
end if

origworksheet.activate
cleanexit:
util.turnonscreenupdate
application.displaystatusbar = bdisplaystatusbarorig
util.recoveroldstatusbar
exit sub

errormcsim:
   msgbox "fatal error: " & err.description, vbexclamation
   on error resume next
   goto cleanexit
end sub

private sub calchiststats(results() as variant, min as double, max as double, binaryonly as boolean, n as long, col as long, _
byref frequency() as long, byref bin_mid() as double, byref nbins as long)
    dim binlen as double, min_ as double, bin_end() as double, b as double, i as long
    dim sorted_indices() as long, sorted_results() as variant, last_count as long, curr_count as long
    
    redim sorted_indices(1 to n)
    redim sorted_results(1 to n)
    for i = 1 to n
        sorted_results(i) = results(i, col)
        sorted_indices(i) = i
    next i
    util.mergesort sorted_results, sorted_indices
    
    if binaryonly then
        min_ = -1.5
        binlen = 1
        nbins = 2
    else
        min_ = min
        binlen = computebinlen(max, min_, n)
        nbins = worksheetfunction.floor((max - min_) / binlen, 1) - 1
        if nbins <= 0 then
            nbins = 2
        end if
    end if
    
    redim frequency(1 to nbins)
    redim bin_mid(1 to nbins)
    redim bin_end(1 to nbins)
    
    for b = 1 to nbins
        bin_mid(b) = min_ + binlen / 2 + b * binlen
        bin_end(b) = bin_mid(b) + binlen / 2
    next b
    
    i = 1
    for b = 1 to nbins
        do while i <= n
            if sorted_results(sorted_indices(i)) >= bin_end(b) and b < nbins then exit do
            i = i + 1
        loop
        curr_count = i - 1
        frequency(b) = curr_count - last_count
        last_count = curr_count
    next b
end sub

function sheetexists(sheettofind as string, optional inworkbook as workbook) as boolean
    if inworkbook is nothing then set inworkbook = activeworkbook

    dim sheet as object
    for each sheet in inworkbook.sheets
        if sheettofind = sheet.name then
            sheetexists = true
            exit function
        end if
    next sheet
    sheetexists = false
end function


private function computebinlen(max as double, byref min as double, n as long)

dim rounded_binlen as double, rounded_nbins as long
dim nbins as long, precision as long, i as long, powprecision as double, intervallen as double
    
nbins = worksheetfunction.ceiling(log(n) / log(1.8), 1)
if nbins < 0 then
    nbins = 1
end if
intervallen = (max - min) / nbins

for i = 30 to -30 step -1
    if intervallen > 10 ^ i then exit for
next i
if i < 0 then
    i = i + 1
end if
precision = i
powprecision = 10 ^ precision

dim nearest_round as double, cand1 as double, cand2 as double, cand3 as double
nearest_round = worksheetfunction.floor(intervallen / (powprecision), 1) * powprecision

if intervallen > 1 then
    cand1 = nearest_round
    cand2 = nearest_round
    cand3 = nearest_round + powprecision
else
    cand1 = nearest_round + 0.2 * powprecision
    cand2 = nearest_round + 0.5 * powprecision
    cand3 = nearest_round + powprecision
end if

' choose correct candidate based on the minimum absolute difference with intervallen
if abs(intervallen - cand1) < abs(intervallen - cand2) then
    if abs(intervallen - cand1) < abs(intervallen - cand3) then
        rounded_binlen = cand1
    else
        rounded_binlen = cand3
    end if
elseif abs(intervallen - cand2) < abs(intervallen - cand3) then
    rounded_binlen = cand2
else
    rounded_binlen = cand3
end if
if rounded_binlen <= 0 or intervallen <= 0 then
    rounded_binlen = 1
end if
min = worksheetfunction.floor(min / powprecision, 1) * powprecision
computebinlen = rounded_binlen

end function


' compute a percentile of a data array
' note that the order of elements in the array will change
function computepercentile(data() as variant, percentile as variant, n as long)
    dim position as long
    position = int(percentile * (n - 1)) + 1
    computepercentile = partitionselect(data, 1, n, position)
end function

' partition-based order statistic calculation
private function partitionselect(x() as variant, offset as long, length as long, index as long) as variant
    dim i as long
    dim j as long
    dim m as long
    dim v as variant
    dim a as long
    dim b as long
    dim c as long
    dim d as long
    dim l as long
    dim n as long
    dim s as long
    
    'for small arrays, just use an insertion sort
    if length < 7 then
        for i = offset to length + offset - 1
            for j = i to offset + 1 step -1
                if x(j - 1) > x(j) then
                    swap x, j, j - 1
                end if
            next j
        next i
        partitionselect = x(index)
        exit function
    end if
        
    ' choose a partition element, v
    m = offset + length / 2     ' small arrays, middle element
    if length > 7 then
        l = offset
        n = offset + length - 1
        if length > 40 then ' big arrays, pseudomedian of 9
            s = length / 8
            l = med3(x, l, l + s, l + 2 * s)
            m = med3(x, m - s, m, m + s)
            n = med3(x, n - 2 * s, n - s, n)
        end if
        m = med3(x, l, m, n) ' mid-size, med of 3
    end if
    v = x(m)
    
    ' establish invariant: v* (<v)* (>v)* v*
    a = offset
    b = a
    c = offset + length - 1
    d = c
    do
        do
            if b > c then
                exit do
            end if
            if x(b) > v then
                exit do
            end if
            if x(b) = v then
                swap x, a, b
                a = a + 1
            end if
            b = b + 1
        loop
        
        do
            if c < b then
                exit do
            end if
            if x(c) < v then
                exit do
            end if
            if x(c) = v then
                swap x, c, d
                d = d - 1
            end if
            c = c - 1
        loop
        
        if b > c then
            exit do
        end if
            
        swap x, b, c
        b = b + 1
        c = c - 1
    loop
    
    ' swap partition elements back to middle
    n = offset + length
    s = a - offset
    if s > b - a then
        s = b - a
    end if
    vecswap x, offset, b - s, s
    s = d - c
    if s > n - d - 1 then
        s = n - d - 1
    end if
    vecswap x, b, n - s, s

    ' recursively select from proper partition
    ' first partition
    s = b - a
    if index < offset + s then
        partitionselect = partitionselect(x, offset, s, index)
        exit function
    end if
    ' last partition
    s = d - c
    if index >= n - s then
        partitionselect = partitionselect(x, n - s, s, index)
        exit function
    end if
    ' it must be the middle partition
    partitionselect = v
end function
            
' utiltity function, swaps x(a) with x(b)
private sub swap(x() as variant, a as long, b as long)
    dim t as variant
    t = x(a)
    x(a) = x(b)
    x(b) = t
end sub
    

' utility function, swaps x(a .. (a+n-1)) with x(b .. (b+n-1))
private sub vecswap(x() as variant, a as long, b as long, n as long)
    dim t as variant
    dim i, ap, bp as long
    ap = a
    bp = b
    for i = 1 to n
        t = x(ap)
        x(ap) = x(bp)
        x(bp) = t
        ap = ap + 1
        bp = bp + 1
    next i
end sub

' utility function, returns the index of the median of the three indexed variables
private function med3(x() as variant, a as long, b as long, c as long) as long
    if x(a) < x(b) then
        if x(b) < x(c) then
            med3 = b
        elseif x(a) < x(c) then
            med3 = c
        else
            med3 = a
        end if
    else
        if x(b) > x(c) then
            med3 = b
        elseif x(a) > x(c) then
            med3 = c
        else
            med3 = a
        end if
    end if
end function


sub matrix2vec(bvec() as variant, amatrix() as variant, j as long, nrow as long)

' matrix2vec copies column j of the matrix amatrix to the vector bvec

dim i as long
for i = 1 to nrow
    bvec(i) = amatrix(i, j)
next i
    
end sub

sub filloutput(irow as integer, jcolstart as integer, avec() as variant, rng as range)
    
' filloutput copies the vector avec into row irow of rng starting at column jcolstart
    
dim i as integer
for i = 1 to ubound(avec)
    rng.cells(irow, jcolstart + i - 1) = avec(i)
next i

end sub

public sub showmontecarlodialog()
' keyboard shortcut: ctrl+shift+m

    dim worksheetname as string, hiddenworksheetname as string
    dim bhistory as boolean
    dim rformula as range, rsimtrials as range, rresult as range, rhistogram as range
    dim nsimtrials as long
    dim rformulastr as string, rsimtrialsstr as string, rresultstr as string, rhistogramstr as string
    dim bprinthistogram as boolean
    dim btransposestr as string, btranspose as boolean
    
    ' check the length of the worksheet name
    if activesheet is nothing then
        msgbox "cbs_ba add-in: there is no sheet currently active. please open a sheet and run this command again", vbexclamation
        exit sub
    end if
    if len(activesheet.name) > 20 then
        msgbox "cbs_ba add-in: to run montecarlo simulation, the worksheet name must not exceed 20 characters." & vbnewline & vbnewline & "(the current worksheet name consists of " & len(activesheet.name) & " characters.)", vbexclamation
        exit sub
    end if
    
    ' check the simulation history for this worksheet
    worksheetname = activesheet.name
    hiddenworksheetname = "ba_hidden_" & worksheetname
    bhistory = false
    bhistory = chksimulhistory(hiddenworksheetname)
    
    ' if there is a simulation history, get field values
    if bhistory = true then
        rformulastr = worksheets(hiddenworksheetname).range("a1").value
        set rformula = worksheets(worksheetname).range(rformulastr)
        rsimtrialsstr = worksheets(hiddenworksheetname).range("a2").value
        if isnumeric(rsimtrialsstr) then
            nsimtrials = clng(rsimtrialsstr)
        else
            set rsimtrials = worksheets(worksheetname).range(rsimtrialsstr)
        end if
        rresultstr = worksheets(hiddenworksheetname).range("a3").value
        set rresult = worksheets(worksheetname).range(rresultstr)
        bprinthistogram = worksheets(hiddenworksheetname).range("a4").value
        rhistogramstr = worksheets(hiddenworksheetname).range("a7").value
        if rhistogramstr <> "" then
            set rhistogram = worksheets(worksheetname).range(rhistogramstr)
        end if
        btransposestr = worksheets(hiddenworksheetname).range("a8").value
        btranspose = false
        if btransposestr <> "" then
            btranspose = worksheets(hiddenworksheetname).range("a8").value
        end if
        
        ' set as default values
        ufmontecarlo.refformula.value = rformulastr
        ufmontecarlo.refsimtrials.value = rsimtrialsstr
        ufmontecarlo.refresult.value = rresultstr
        ufmontecarlo.chkhisto.value = bprinthistogram
        ufmontecarlo.commandbutton1.setfocus
        ufmontecarlo.refhistogram.value = rhistogramstr
        'ufmontecarlo.chktranspose.value = btranspose
    end if
    
    ' show the dialog
    ufmontecarlo.show
end sub

private function chksimulhistory(byval worksheetname) as boolean
    chksimulhistory = false
    dim sh as worksheet, flg as boolean
    
    for each sh in worksheets
        if (sh.name = worksheetname) then chksimulhistory = true: exit for
    next
end function

public sub saveinputfields(byref rformula as range, byref rsimtrials as range, byref rresult as range, byval bprinthistogram, byval worksheetname, byref rhistogram as range, byval btranspose, byref nsimtrials as long)
    dim wstest as worksheet
    dim strsheetname as string
    
    strsheetname = "ba_hidden_" & worksheetname
     
    set wstest = nothing
    on error resume next
    set wstest = activeworkbook.worksheets(strsheetname)
    on error goto 0
     
    if wstest is nothing then
        worksheets.add().name = strsheetname
    else
        application.displayalerts = false
        wstest.delete
        application.displayalerts = true
        worksheets.add().name = strsheetname
    end if
    
    ' hide the worksheet
    activeworkbook.sheets(strsheetname).visible = false
    
    with activeworkbook.sheets(strsheetname)
        .range("a1").value = rformula.address
        if rsimtrials is nothing then
            .range("a2").value = nsimtrials
        else
            .range("a2").value = rsimtrials.address
        end if
        .range("a3").value = rresult.address
        .range("a4").value = bprinthistogram
        .range("a6").value = worksheetname
        if not rhistogram is nothing then
            .range("a7").value = rhistogram.address
        end if
        .range("a8").value = btranspose
    end with
    
    activeworkbook.worksheets(worksheetname).activate
end sub


private function get_random()
    application.volatile
    get_random = rnd()
end function



' histogram macro used along with the cbs_ba.xll add-in.
' this is consistent with cbs_ba.xll version 3.01m.

private sub ba_histogram_create()

    dim nbins as integer
    dim rtable as range, rtable2 as range
    dim rchart as range
    dim rbinlabel as range, rfrequency as range
    dim binlabel as variant, frequency as variant
    dim labelspacing as integer
    dim clleft, cltop, clwidth, clheight as double
    dim numdigit as integer
    dim i as integer
    
    set rtable = application.selection
    set rchart = rtable.cells(0, 0).offset(1, 5)
    
    if rtable.columns.count <> 3 then
        msgbox "the input table must contain 3 columns", vbexclamation
        exit sub
    end if
    
    set rbinlabel = rtable.columns(1)
    set rfrequency = rtable.columns(3)
    
    if rtable.rows.count <= 1 then
        nbins = 1
    else
        nbins = rtable.rows.count
    end if
    
    numdigit = len(cstr(rbinlabel.cells(1, 1).value))
    
    activesheet.shapes.addchart(xlcolumnclustered, width:=450, height:=270).select
    with activeworkbook.activechart
    
        'add data series
        .seriescollection.newseries
        .seriescollection(1).values = rfrequency
        .seriescollection(1).xvalues = rbinlabel
                
        '.axes(xlcategory).ticklabels.numberformat = numformat
        
        for i = .seriescollection.count to 2 step -1
            .seriescollection(i).delete
        next i
        
        'axis label
        .axes(xlcategory).ticklabelspacing = 1
        .axes(xlcategory).majortickmark = xlcross
        .axes(xlvalue).majortickmark = xlcross
        .chartgroups(1).gapwidth = 50
        .chartgroups(1).overlap = 100
        
        'axes titles
        .axes(xlcategory, xlprimary).hastitle = true
        .axes(xlcategory, xlprimary).axistitle.characters.text = "bin"
        .axes(xlvalue, xlprimary).hastitle = true
        .axes(xlvalue, xlprimary).axistitle.characters.text = "frequency"
            
        'formatting
        .axes(xlcategory).hasminorgridlines = false
        .axes(xlvalue).hasmajorgridlines = false
        .axes(xlvalue).hasminorgridlines = false
        .haslegend = false
        .seriescollection(1).format.fill.forecolor.rgb = rgb(48, 103, 192)
        .axes(xlcategory).axistitle.format.textframe2.textrange.font.size = 20
        .axes(xlcategory).axistitle.format.textframe2.textrange.font.bold = msofalse
        .axes(xlvalue).axistitle.format.textframe2.textrange.font.size = 20
        .axes(xlvalue).axistitle.format.textframe2.textrange.font.bold = msofalse
        
        'color
        '.seriescollection(1).interior.color = rgb(48, 103, 192)
        
        'axes label
        .axes(xlcategory).ticklabels.font.size = 16
        .axes(xlcategory).ticklabelspacing = application.roundup(numdigit * nbins / 45, 0)
        .axes(xlvalue).ticklabels.font.size = 16
        
        .hastitle = false
        
        '.shapes.addtextbox(msotextorientationhorizontal, 419.25, 42, 147.75, 51.75).formula = "=$a$1"
    end with

end sub

private sub createhistogram(freqtable as string, statsstr as string, label as string)

    activesheet.range(freqtable).select
    
    call ba_histogram_create
        
    activesheet.move after:=sheets(activeworkbook.sheets.count)
    
    with activeworkbook.activechart
        .parent.left = range("b2").left
        .parent.top = range("b2").top
        .shapes.addtextbox(msotextorientationhorizontal, 280, 6.75, _
            180, 100).select
        'activechart.shapes("text box 1").select
        selection.text = statsstr
        selection.font.size = 16
        .axes(xlcategory, xlprimary).hastitle = true
        .axes(xlcategory, xlprimary).axistitle.characters.text = label
    end with
        
    activesheet.range("a1").select
end sub

public function binomsim(n as long, p as double)
    on error goto binomsim_err
    if p < 0 or p > 1 then
        binomsim = "p must be between 0 and 1"
        exit function
    end if
    if n <= 0 then
        binomsim = "n must be positive"
        exit function
    end if
        
    dim r as double
    
    r = get_random()
    binomsim = application.worksheetfunction.binom_inv(n, p, r)
    exit function
binomsim_err:
    binomsim = "fatal error: " & err.description
end function


public function normalsim(mu as double, sigma as double)
    on error goto normalsim_err
    
    if sigma < 0 then
        normalsim = "sigma cannot be negative"
        exit function
    end if
        
    dim r as double
    
    r = get_random()
    normalsim = application.worksheetfunction.norm_inv(r, mu, sigma)
    exit function
normalsim_err:
    normalsim = "fatal error: " & err.description
end function


public function poissonsim(mean as double)
    on error goto poissonsim_err
    
    if mean < 0 then
        poissonsim = "mean must be non-negative"
        exit function
    end if
        
    dim r as double, prod as double, target as double, count as long
    if mean < 10 then
        target = exp(-mean)
        prod = 1
        count = 0
        do while prod > target
            doevents
            r = get_random()
            prod = prod * r
            count = count + 1
        loop
        poissonsim = count
    else
        'algorithm is algorithm ptrs from the transformed rejection method for generating poisson random variables by hormann (1992)
        'cannot be used for mean > 1e7
        if mean > 10000000# then
            poissonsim = "given mean is too high. needs to be less than 10000000 for this algorithm"
            exit function
        end if
            
        dim b as double, a as double, nur as double, u as double, v as double, us as double, invalpha as double, lnmu as double
        dim k as long
        ' step 0
        b = 0.931 + 2.53 * sqr(mean)
        a = -0.059 + 0.02483 * b
        nur = 0.9277 - 3.6224 / (b - 2)
        do while true
            doevents
            ' step 1
            u = get_random() - 0.5
            v = get_random()
            us = 0.5 - abs(u)
            k = int((2 * a / us + b) * u + mean + 0.43)
            
            'step 2
            if us >= 0.07 and v <= nur then
                poissonsim = k
                exit function
            end if
            if (k <= 0) or (us < 0.013 and v > us) then
                goto nextiteration
            end if
            
            'step 3.0
            invalpha = 1.1239 + 1.1328 / (b - 3.4)
            lnmu = log(mean)
            
            'step 3.1
            if log(v * invalpha / (a / (us * us) + b)) <= -mean + k * lnmu - logfactorial(k) then
                poissonsim = k
                exit function
            end if
                
nextiteration:
        loop
    end if
    
    exit function
poissonsim_err:
    poissonsim = "fatal error: " & err.description
end function

public function triangularsim(min as double, mode as double, max as double)

    ' min = minimum value

    ' mode = most likely value

    ' max =  maximum value

    on error goto triangular_err

    dim u as double, f as double

    if mode < min or max < min or max < mode then
        triangularsim = "parameters should be ordered as min < mode < max"
        exit function
    end if

    u = get_random() ' random number between 0 and 1
    f = (mode - min) / (max - min) ' fraction of the distribution's range where the mode occurs

    if u <= f then
        triangularsim = min + sqr(u * (max - min) * (mode - min))
    else
        triangularsim = max - sqr((1 - u) * (max - min) * (max - mode))
    end if

    exit function

triangular_err:
    triangularsim = "fatal error: " & err.description
end function

public function uniformsim(min as double, max as double)
    ' min = minimum value

    ' max =  maximum value

    on error goto uniformsim_err

    dim u as double

    if max < min then
        uniformsim = "min should be smaller than max"
        exit function
    end if

    u = get_random() 'generate a random number between 0 and 1

    uniformsim = (max - min) * u + min ' generate a random number between min and max

    exit function

uniformsim_err:
    uniformsim = "fatal error: " & err.description
end function

private function logfactorial(k as long)
    dim logtable() as variant
    logtable = array(0, 0.693147181, 1.791759469, 3.17805383, 4.787491743, _
                                6.579251212, 8.525161361, 10.6046029, 12.80182748, 15.10441257)
    if k <= 10 then
        logfactorial = logtable(k - 1)
        exit function
    end if
    
    dim log2pi as double, term2 as double, term31 as double, term32 as double, term3 as double
    log2pi = 1.837877066
    term2 = (k + 0.5) * log(k) - k
    term31 = 0.08333 / k
    term32 = -0.002777778 / (k ^ 3)
    term3 = term31 + term32
    logfactorial = 0.5 * log2pi + term2 + term3
    
end function

public function lognormalsim(mu as double, sigma as double)
on error goto lognormalsim_err
    
    if sigma < 0 then
        lognormalsim = "sigma cannot be negative"
        exit function
    end if
    
    dim normalrv as double
    normalrv = normalsim(mu, sigma)
    lognormalsim = exp(normalrv)
    
    exit function
lognormalsim_err:
    lognormalsim = "fatal error: " & err.description
end function

public function expsim(rate as double)
    on error goto expsim_err
    
    if rate <= 0 then
        expsim = "rate must be positive"
        exit function
    end if
        
    dim r as double
    
    r = get_random()
    expsim = -log(r) / rate
    exit function
expsim_err:
    expsim = "fatal error: " & err.description
end function

public sub registerall()
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    dim argdesc() as string
    
    funccat = "cbs ba add-in functions"
    
    'binomsim(n as long, p as double)
    redim argdesc(1 to 2)
    funcname = "binomsim"
    funcdesc = "generate a random number according to the binomial distribution"
    argdesc(1) = "number of bernoulli trials for this binomial random variable"
    argdesc(2) = "probability of each bernoulli trial being 1"
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    'normalsim(mu as long, sigma as double)
    redim argdesc(1 to 2)
    funcname = "normalsim"
    funcdesc = "generate a random number according to the normal distribution"
    argdesc(1) = "mean of the distribution"
    argdesc(2) = "standard deviation of the distribution"
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    'expsim(rate as double)
    redim argdesc(1 to 1)
    funcname = "expsim"
    funcdesc = "generate a random number according to the exponential distribution"
    argdesc(1) = "rate of the exponential distribution"
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    'triangularsim(min as double, mode as double, max as double)
    redim argdesc(1 to 3)
    funcname = "triangularsim"
    funcdesc = "generate a random number according to the triangular distribution"
    argdesc(1) = "minimum value of the triangular distribution"
    argdesc(2) = "mode value of the triangular distribution"
    argdesc(3) = "maximum value of the triangular distribution"
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    'uniformsim(min as double, max as double)
    redim argdesc(1 to 2)
    funcname = "uniformsim"
    funcdesc = "generate a random number according to the uniform distribution"
    argdesc(1) = "minimum value of the uniform distribution"
    argdesc(2) = "maximum value of the uniform distribution"
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    'poissonsim(mean as double)
    redim argdesc(1 to 1)
    funcname = "poissonsim"
    funcdesc = "generate a random number according to the poisson distribution"
    argdesc(1) = "mean of the poisson distribution"
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    registermcsim
end sub

