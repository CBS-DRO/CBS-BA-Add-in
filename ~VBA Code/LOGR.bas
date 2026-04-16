option explicit
 

private function logisticfunction(w as double) as double
    ' internal function to compute the logistic function with optimization for very small or very large values of w
    if w > 100 then
        logisticfunction = 1 - exp(-w)
    elseif w < -100 then
        logisticfunction = exp(w)
    else
        logisticfunction = 1 / (1 + exp(-w))
    end if
end function

private function internallinearregpredict(model() as variant, known_x() as variant, constant as long, _
col_select() as variant, selectedvars as long, n as long, x_cols as long)
    ' internal function to compute prediction according to a linear regression model.
    
    dim k as long, i as long, j as long, jj as long
    
    k = selectedvars + constant 'k is the total number of coefficients we want from the model
    
    
    dim y_estimate() as variant, prod as double
    redim y_estimate(1 to n, 1 to 1)

    for i = 1 to n
        ' if constant is to be included or not
        if constant = 1 then
            prod = model(1, 1)
        else
            prod = 0
        end if
        j = 1 + constant
        for jj = 1 to x_cols
            ' add to product only if col_select is not 0
            if col_select(jj) <> 0 then
                prod = prod + model(1, j) * known_x(i, jj)
                j = j + 1
            end if
        next jj
        y_estimate(i, 1) = prod
    next i
    
    internallinearregpredict = y_estimate
end function


private function internallogisticregpredict(model() as variant, known_x() as variant, constant as long, col_select() as variant, _
selectedvars as long, n as long, x_cols as long)
    ' internal function to compute predictions of probability of outcome being 1 according to a logistic regression model.
    dim k as long, i as long, j as long, jj as long
    
    k = selectedvars + constant 'k is the total number of coefficients we want from the model
    
    
    dim logitprob() as variant, prod as double
    redim logitprob(1 to n, 1 to 1)

    for i = 1 to n
        ' if constant is to be included or not
        if constant = 1 then
            prod = model(1, 1)
        else
            prod = 0
        end if
        j = 1 + constant
        for jj = 1 to x_cols
            ' add to product only if col_select is not 0
            if col_select(jj) <> 0 then
                prod = prod + model(1, j) * known_x(i, jj)
                j = j + 1
            end if
        next jj
        logitprob(i, 1) = logisticfunction(prod)
    next i
    
    internallogisticregpredict = logitprob
end function

public function linearregpredict(model as range, known_x as range, optional constant as boolean = true, _
optional col_select as range = nothing)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: linearregpredict(model as range, known_x as range, optional constant as boolean = true, _
optional col_select as range = nothing, optional log_odds as boolean = false)
'computes probabilities of y being 1 based on the model computed by logisticregtrain on known x values.
'input:
'   model (required): model coefficients. must be a single row.
'   known_x (required): known x values. number of included columns (see col_select below) must equal that in model if constant
'                    is not included (see constant below). if constant is included, number of entries must be one less than
'                    that in model. can be multiple rows, in which case output will have one entry per row.
'                    each entry must be numeric.
'   constant: if true, interpret the first entry in model to be the constant in regression. if omitted, assumed to be true.
'   col_select: a vector of 0/1. must be a single row. number of columns must equal that in known_x. if jth entry is 1, jth
'            column in known_x matrix is included when making a prediction. if omitted, each entry is assumed to be 1.
'            note that each entry in model variable is always included. the parameters constant and col_select should have
'            the same values as they did when running logisticregtrain for correct estimates of the probabilities.
'output:
'   output consists of one row for each row in known_x in the input: the estimate of y when x takes the corresponding value,
'   based on the model.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

on error goto unexpected_error_linr_predict
    
    dim constant_flag as long, col_select_flags() as variant, j as long, selectedvars as long
    
    'check if we have a constant
    if constant then
        constant_flag = 1
    else
        constant_flag = 0
    end if
    
    'handle col_select
    redim col_select_flags(1 to known_x.columns.count)
    if col_select is nothing then
        for j = 1 to known_x.columns.count
            col_select_flags(j) = 1
        next j
        selectedvars = known_x.columns.count
    else
        if col_select.count <> known_x.columns.count then
            linearregpredict = "number of entries in col_select_in must equal the total number columns in known_x"
            exit function
        end if
        selectedvars = application.worksheetfunction.countif(col_select, "=1")
        if selectedvars = 0 then
            linearregpredict = "at least one column must be selected"
            exit function
        end if
        if application.worksheetfunction.countif(col_select, "=0") + selectedvars <> known_x.columns.count then
            linearregpredict = "all values in col_select must be zero or one"
            exit function
        end if
        for j = 1 to col_select.count
            col_select_flags(j) = col_select(j)
        next j
    end if
    
    
    'check if range dimensions are as expected
    if model.rows.count <> 1 then
        linearregpredict = "model coefficients must be in a single row"
        exit function
    end if
    
    if application.worksheetfunction.count(model) <> selectedvars + constant_flag then
        linearregpredict = "incorrect number of model coefficients for prediction with given x"
        exit function
    end if
    
    dim x_arr() as variant, model_arr() as variant, n as long, x_cols as long
    n = known_x.rows.count
    x_cols = known_x.columns.count
    redim model_arr(1 to model.columns.count)
    redim x_arr(1 to n, 1 to x_cols)
    model_arr = model.value2
    if known_x.count = 1 then
        x_arr(1, 1) = known_x.value2
    else
        x_arr = known_x.value2
    end if
    
    linearregpredict = internallinearregpredict(model_arr, x_arr, constant_flag, col_select_flags, selectedvars, n, x_cols)
    exit function
    
    'error handler
unexpected_error_linr_predict:
    linearregpredict = "fatal error: " & err.description
end function

public function logisticregpredict(model as range, known_x as range, optional constant as boolean = true, _
optional col_select as range = nothing, optional not_implemented as boolean = false)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: logisticregpredict(model as range, known_x as range, optional constant as boolean = true, _
optional col_select as range = nothing, optional log_odds as boolean = false)
'computes probabilities of y being 1 based on the model computed by logisticregtrain on known x values.
'input
'   model (required): model coefficients. must be a single row.
'   known_x (required): known x values. number of included columns (see col_select below) must equal that in model if constant
'                    is not included (see constant below). if constant is included, number of entries must be one less than
'                    that in model. can be multiple rows, in which case output will have one entry per row.
'                    each entry must be numeric.
'   constant: if true, interpret the first entry in model to be the constant in regression. if omitted, assumed to be true.
'   col_select: a vector of 0/1. must be a single row. number of columns must equal that in known_x. if jth entry is 1, jth
'            column in known_x matrix is included when making a prediction. if omitted, each entry is assumed to be 1.
'            note that each entry in model variable is always included. the parameters constant and col_select should have
'            the same values as they did when running logisticregtrain for correct estimates of the probabilities.
'   log_odds: included for backward compatibility. value ignored.
'output
'   output consists of one row for each row in known_x in the input: the estimate of the probability of y being 1 when x takes
'   the corresponding value, based on the model.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    on error goto unexpected_error2
    
    dim constant_flag as long, col_select_flags() as variant, j as long, selectedvars as long
    
    'check if we have a constant
    if constant then
        constant_flag = 1
    else
        constant_flag = 0
    end if
    
    'handle col_select
    redim col_select_flags(1 to known_x.columns.count)
    if col_select is nothing then
        for j = 1 to known_x.columns.count
            col_select_flags(j) = 1
        next j
        selectedvars = known_x.columns.count
    else
        if col_select.count <> known_x.columns.count then
            logisticregpredict = "number of entries in col_select_in must equal the total number columns in known_x"
            exit function
        end if
        selectedvars = application.worksheetfunction.countif(col_select, "=1")
        if selectedvars = 0 then
            logisticregpredict = "at least one column must be selected"
            exit function
        end if
        if application.worksheetfunction.countif(col_select, "=0") + selectedvars <> known_x.columns.count then
            logisticregpredict = "all values in col_select must be zero or one"
            exit function
        end if
        for j = 1 to col_select.count
            col_select_flags(j) = col_select(j)
        next j
    end if
    
    
    'check if range dimensions are as expected
    if model.rows.count <> 1 then
        logisticregpredict = "model coefficients must be in a single row"
        exit function
    end if
    
    if application.worksheetfunction.count(model) <> selectedvars + constant_flag then
        logisticregpredict = "incorrect number of model coefficients for prediction with given x"
        exit function
    end if
    
    dim x_arr() as variant, model_arr() as variant, n as long, x_cols as long
    n = known_x.rows.count
    x_cols = known_x.columns.count
    redim model_arr(1 to model.columns.count)
    redim x_arr(1 to n, 1 to x_cols)
    model_arr = model.value2
    if known_x.count = 1 then
        x_arr(1, 1) = known_x.value2
    else
        x_arr = known_x.value2
    end if
    
    logisticregpredict = internallogisticregpredict(model_arr, x_arr, constant_flag, col_select_flags, selectedvars, n, x_cols)
    exit function
    
    'error handler
unexpected_error2:
    logisticregpredict = "fatal error: " & err.description
end function
 
public function linearregtrain(known_y as range, known_x as range, optional constant as boolean = true, _
optional not_implemented1 as long = 0, optional col_select as range = nothing, optional not_implemented2 as range = nothing, _
optional not_implemented3 as boolean, optional not_implemented4 as long)
' original signature: linearregtrain(known_y as range, known_x as range, optional constant as boolean = true, _
optional na_action as long = 0, optional col_select as range = nothing, optional row_weight as range = nothing, _
optional f_test as boolean, optional cluster_id as long)
'runs linear regression on data using linest function.
'
'input:
'   known_y (required): a range containing y. must be a single column.
'                    top row(s) may include header(s). only the first header row will be used when displaying output.
'                    headers must be non-numeric or formatted differently from the data (see listheaderrows property for
'                    more on how excel identifies header rows.)
'   known_x (required): a range containing x. all entries must be numeric. top row(s) may include header(s).
'                    only the first header row will be used when displaying output. number of rows not counting header(s)
'                    must match those of known_y.
'   constant: boolean indicating whether to include constant or not. assumed true if omitted.
'   na_action: included for backward compatibility. value not used.
'   col_select: a range containing 0/1. must be a single row. number of columns must match those of known_x. at least one
'            entry must be 1. if jth entry is 0, jth column is ignored from the known_x matrix. if omitted, all entries
'            are assumed to be 1.
'   row_weight: included for backward compatibility. value not used.
'   f_test: included for backward compatibility. value not used.
'   cluster_id: included for backward compatibility. value not used.
'output:
'   output is a matrix having 7 rows. number of columns depends on input. there is one column for each known_x column
'   included in the regression, one column for constant if included, and one more column for row labels.
'   row 1: used for column labels. if known_y includes a header, the first entry contains that header followed by
'       the size of the output matrix in brackets. otherwise, the first entry has the name of the function (logisticregtrain),
'       followed by the size of the output matrix in brackets. if constant is included, the second entry says
'       "constant". next are entries corresponding to columns included from known_x in order. if known_x contains a header,
'       it is used as column labels here. otherwise, the column labels are var1, var2, ... and so on.
'   row 2: first entry says "coefficients". later entries are the logistic regression constant (if included) and coefficients
'       corresponding to known_x columns that are included.
'   row 3: first entry says "p-value". later entries are the p-values for the coefficients in row 2.
'   row 4: first entry says "std error". later entries are the std errors for the coefficients in row 2.
'   row 5: first entry says "log-likelihood". next entry is the log-likelihood of the model at termination. this and subsequent
'       rows have only two entries.
'   row 6: first entry says "number valid obs". next entry is the number of rows in known_x.
'   row 7: first entry says "total obs". next entry is the number of rows in known_x.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
on error goto unexpected_error_linr
    
    dim include_constant as long, x_header() as string, x_header_in as range, y_header as range, selectedvars as long, col_select_flags() as variant
    'count variables
    dim i as long, j as long, jj as long
    
    'check if we need to have constant
    if constant then
        include_constant = 1
    else
        include_constant = 0
    end if
    
    'handle header
    set x_header_in = util.getheaderrowsandresizerange(known_x)
    set y_header = util.getheaderrowsandresizerange(known_y)
    
    redim col_select_flags(1 to known_x.columns.count)
    if col_select is nothing then
        for j = 1 to known_x.columns.count
            col_select_flags(j) = 1
        next j
        selectedvars = known_x.columns.count
    else
        if col_select.columns.count <> known_x.columns.count then
            linearregtrain = "number of entries in col_select must equal the total columns in x"
            exit function
        end if
        if col_select.rows.count <> 1 then
            linearregtrain = "col_select must have exactly one row"
            exit function
        end if
        selectedvars = application.worksheetfunction.countif(col_select, "=1")
        if selectedvars = 0 then
            linearregtrain = "at least one column must be selected"
            exit function
        end if
        if application.worksheetfunction.countif(col_select, "=0") + selectedvars <> known_x.columns.count then
            linearregtrain = "all values in col_select must be zero or one"
            exit function
        end if
        for j = 1 to col_select.count
            col_select_flags(j) = col_select(j)
        next j
    end if
    
    if not x_header_in is nothing then
        redim x_header(1 to selectedvars)
        j = 1
        for i = 1 to known_x.columns.count
            if col_select_flags(i) = 1 then
                x_header(j) = x_header_in(i)
                j = j + 1
            end if
        next i
    end if

     
    'read data dimensions
    dim k as long, n as long
    
    n = known_y.rows.count ' n is total observation
    k = selectedvars + include_constant 'k is the number of total explanatory variables
    
    ' validate the input
    if n <> known_x.rows.count then
        linearregtrain = "number of x and y variables should be the same."
        exit function
    end if
    if application.worksheetfunction.count(known_x) <> known_x.columns.count * known_x.rows.count then
        linearregtrain = "all x variables should be numeric."
        exit function
    end if
    if known_y.columns.count <> 1 then
        linearregtrain = "known_y must have exactly one column"
        exit function
    end if
    
    ' get the data
    dim y() as variant
    redim y(1 to n)
    y = known_y.value2
    
    'removing unselected columns from x. rng_x=x from now on
    dim x() as double
    redim x(1 to n, 1 to k - include_constant)
    dim data_x() as variant
    data_x = known_x.value2
    for i = 1 to n
        jj = 1
        for j = 1 to known_x.columns.count
            if col_select_flags(j) = 1 then
                x(i, jj) = data_x(i, j)
                jj = jj + 1
            end if
        next j
    next i
    
    ' prepare the run the regression
    dim linestretval() as variant
    
    ' find the number of variables (k includes the intercept; we want the
    ' number of variables only)
    dim n_variables as long
    n_variables = ubound(x, 2)
    
    if n_variables > 64 then
        ' if we have more than 64 variables, we need to do regression manually; linest
        ' won't work
        
        ' linest returns the coefficients in reverse order, so we need to reverse our x
        ' matrix for everything to work
        i = 1
        while i < ((n_variables / 2) + 0.1)
            for j = 1 to n
                dim buffer as double
                buffer = x(j, i)
                x(j, i) = x(j, n_variables - i + 1)
                x(j, n_variables - i + 1) = buffer
            next j
            i = i + 1
        wend
        
        if include_constant then
            ' if we need a constant, add a column of 1s to the end of our x matrix
            ' (conveniently, linest returns the intercept as its last element, so
            ' the last column is right)
            redim preserve x(1 to n, 1 to n_variables + 1)
            
            for j = 1 to n
                x(j, n_variables + 1) = 1
            next j
        end if
        
        ' do the regression; first, find (x^t x)^-1
        dim xtx as variant
        dim xtxmone as variant
        xtx = worksheetfunction.mmult(worksheetfunction.transpose(x), x)
                
        if worksheetfunction.mdeterm(xtx) <= 0.000001 then
            ' linest doesn't report this error, but we do since we depend on inverting the determinant
            linearregtrain = "error: your data is highly co-linear; linear regression cannot be fit"
            exit function
        end if
        
        xtxmone = worksheetfunction.minverse(xtx)
        
        ' now find beta_hat and sigma^2
        dim beta_hat as variant
        beta_hat = worksheetfunction.mmult(worksheetfunction.mmult(xtxmone, worksheetfunction.transpose(x)), y)
        
        ' find y_hat, ssr, and sst
        dim y_hat as variant
        dim ssr as double
        dim sst as double
        
        y_hat = worksheetfunction.mmult(x, beta_hat)
        for j = 1 to n
            ssr = ssr + (y(j, 1) - y_hat(j, 1)) ^ 2
        next j
        sst = worksheetfunction.var_p(y) * n
              
        ' find sigma_square_hat and r_squared
        dim sigma_squared_hat as double
        dim r_squared as double
        
        sigma_squared_hat = ssr / (n - k)
        r_squared = 1 - (ssr / sst)
                
        ' construct the linest output
        redim linestretval(1 to 3, 1 to k)
        
        for i = 1 to k
            linestretval(1, i) = beta_hat(i, 1)
            linestretval(2, i) = math.sqr(sigma_squared_hat * xtxmone(i, i))
            linestretval(3, 1) = r_squared
        next i
    
    else
        linestretval = application.worksheetfunction.linest(y, x, include_constant, true)
    end if
    
    dim relogit()
    redim relogit(1 to 7, 1 to k + 1)
    
    if y_header is nothing then
        relogit(1, 1) = "linearregtrain (7," & k + 1 & ")"
    else
        relogit(1, 1) = y_header(1, 1) & " (7," & k + 1 & ")"
    end if
         
    'coefficients
    relogit(2, 1) = "coefficients"
    for j = 1 to k 'k variables
        if x_header_in is nothing then
            relogit(1, j + 1) = "var" & j - include_constant
        elseif j > include_constant then
            relogit(1, j + 1) = x_header(j - include_constant)
        end if
        ' output by linest has inverted columns compared to the input.
        relogit(2, j + 1) = linestretval(1, k - j + 1)
    next j
    if constant then
        relogit(1, 2) = "constant"
    end if

     
    'additional statistics if requested
    relogit(3, 1) = "std error"
    relogit(4, 1) = "p-value"
    for j = 1 to k
        dim std_err as double, z as double, pvalue as double
        
        std_err = linestretval(2, k - j + 1)
        if std_err <> 0 then
            ' compute p-value using t-distribution.
            z = linestretval(1, k - j + 1) / std_err
            pvalue = application.worksheetfunction.tdist(abs(z), n - k, 1) * 2
        else
            pvalue = 0
        end if
        relogit(3, j + 1) = std_err
        relogit(4, j + 1) = pvalue
        relogit(5, j + 1) = ""
        relogit(6, j + 1) = ""
        relogit(7, j + 1) = ""
    next j
       
    'additional statistics
    relogit(5, 1) = "r-sqr"
    relogit(5, 2) = linestretval(3, 1)
    relogit(6, 1) = "number valid obs"
    relogit(6, 2) = n
    relogit(7, 1) = "total obs"
    relogit(7, 2) = n
    
    linearregtrain = relogit
     
    exit function
     
    'error handler
unexpected_error_linr:
    linearregtrain = "fatal error " & err.number & ":" & err.description
end function
 
public function twoslsregtrain(known_y as range, optional exog_x as range = nothing, optional endog_x as range = nothing, _
                                optional instruments as range = nothing, optional constant as boolean = true, _
                                    optional not_implemented2 as long = 0, optional exog_col_select as range = nothing, _
                                    optional not_implemented3 as range = nothing, optional not_implemented4 as boolean)
' original signature: todo_dg
'runs a two-stage-least-squares regression. the first stage is run using linest, and
'the second by matrix inversion
'
'input:
'   known_y (required): a range containing y. must be a single column.
'                    top row(s) may include header(s). only the first header row will be used when displaying output.
'                    headers must be non-numeric or formatted differently from the data (see listheaderrows property for
'                    more on how excel identifies header rows.)
'   exog_x (optional): a range containing exogenous x. all entries must be numeric. top row(s) may include header(s).
'                    only the first header row will be used when displaying output. number of rows not counting header(s)
'                    must match those of known_y.
'   endog_x (required): a range containing endogenous x. all entries must be numeric. top row(s) may include header(s).
'                    only the first header row will be used when displaying output. number of rows not counting header(s)
'                    must match those of known_y.
'   instruments (required): a range containing instruments. all entries must be numeric. top row(s) may include header(s).
'                    only the first header row will be used when displaying output. number of rows not counting header(s)
'                    must match those of known_y.
'   constant: boolean indicating whether to include constant or not. assumed true if omitted.
'   na_action: included for backward compatibility. value not used.
'   exog_col_select: a range containing 0/1. must be a single row. number of columns must match those of exog_x. at least one
'            entry must be 1. if jth entry is 0, jth column is ignored from the exog_x matrix. if omitted, all entries
'            are assumed to be 1.
'   row_weight: included for backward compatibility. value not used.
'   show_first_stage: included for backward compatibility. value not used. first stage will not be shown
'output:
'   output is a matrix having 6 rows. number of columns depends on input. there is one column for each exog_x and endog_x
'   variable included in the regression, one column for constant, and one more column for row labels.
'   row 1: used for column labels. if known_y includes a header, the first entry contains that header followed by
'       the size of the output matrix in brackets. otherwise, the first entry has the name of the function (twoslsregtrain),
'       followed by the size of the output matrix in brackets. the second entry says "constant". next are entries
'       corresponding to columns included from known_x in order. if known_x contains a header, it is used as column labels here.
'       otherwise, the column labels are exog_var1, exog_var2, ..., endog_var1, endog_var2, and so on.
'   row 2: first entry says "coefficients". later entries are coefficients
'   row 3: first entry says "p-value". later entries are p-values
'   row 4: first entry says "std error". later entries are std errors
'   row 5: first entry says "number valid obs". next entry is the number of rows in known_x.
'   row 6: first entry says "total obs". next entry is the number of rows in known_x.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
on error goto unexpected_error_twoslstrain
    
    ' define variables
    ' ----------------
    
    dim include_constant as long
    
    ' header input ranges
    dim exog_x_header_in as range
    dim endog_x_header_in as range
    dim instrument_header_in as range
    dim y_header_in as range
    
    ' header arrays/strings
    dim exog_x_header() as string
    dim endog_x_header() as string
    dim y_header as string
    
    ' selected exogenous columns
    dim exog_col_select_flags() as variant
    
    ' number of variables
    dim exog_n_vars as long
    dim endog_n_vars as long
    dim instruments_n_vars as long
    
    ' number of data points
    dim n as long
    
    'count variables
    dim i as long, j as long, jj as long
    
    ' output variable
    dim out() as variant
    
    ' deal with exog_col_select and the number of exogenous vars
    ' ----------------------------------------------------------
    
    if exog_x is nothing then
        exog_n_vars = 0
    else
        redim exog_col_select_flags(1 to exog_x.columns.count)
        
        if exog_col_select is nothing then
            for j = 1 to exog_x.columns.count
                exog_col_select_flags(j) = 1
            next j
            exog_n_vars = exog_x.columns.count
        else
            if exog_col_select.columns.count <> exog_x.columns.count then
                twoslsregtrain = "number of entries in col_select must equal the total columns in x"
                exit function
            end if
            if exog_col_select.rows.count <> 1 then
                twoslsregtrain = "col_select must have exactly one row"
                exit function
            end if
            
            exog_n_vars = application.worksheetfunction.countif(exog_col_select, "=1")
            
            if application.worksheetfunction.countif(exog_col_select, "=0") + exog_n_vars <> exog_x.columns.count then
                twoslsregtrain = "all values in col_select must be zero or one"
                exit function
            end if
            
            for j = 1 to exog_col_select.count
                exog_col_select_flags(j) = exog_col_select(j)
            next j
        end if
    end if
    
    ' validate presence of y, endogenous variables and exogenous
    ' variables, and find how many of them there are
    ' ----------------------------------------------------------
    
    if endog_x is nothing then
        twoslsregtrain = "please include endogenous x variables"
        exit function
    else
        endog_n_vars = endog_x.columns.count
    end if
    
    if instruments is nothing then
        twoslsregtrain = "please include instrumental variables"
        exit function
    else
        instruments_n_vars = instruments.columns.count
    end if
    
    ' know_y cannot be nothing because it is not optional in the function
    if known_y.columns.count <> 1 then
        twoslsregtrain = "known_y must have exactly one column"
        exit function
    end if

    ' decide if a constant should be included
    ' ---------------------------------------
    
    if constant then
        include_constant = 1
    else
        include_constant = 0
    end if

    ' handle headers
    ' --------------
        
    ' remove the headers from the input ranges, and save the header ranges elsewhere
    if exog_n_vars > 0 then
        set exog_x_header_in = util.getheaderrowsandresizerange(exog_x)
    end if
    set endog_x_header_in = util.getheaderrowsandresizerange(endog_x)
    set instrument_header_in = util.getheaderrowsandresizerange(instruments)
    set y_header_in = util.getheaderrowsandresizerange(known_y)
    
    ' process the exogenous header ranges; the constant counts as an exogenous
    ' variable
    redim exog_x_header(1 to exog_n_vars + include_constant)
    
    if include_constant = 1 then
        exog_x_header(1) = "constant"
    end if
    
    if exog_x_header_in is nothing then
        for i = 1 to exog_n_vars
            exog_x_header(i + include_constant) = "exog var " & i
        next
    else
        j = 1
        for i = 1 to exog_x.columns.count
            if exog_col_select_flags(i) = 1 then
                exog_x_header(j + include_constant) = exog_x_header_in(i)
                j = j + 1
            end if
        next i
    end if
    
    ' process the endogenous header ranges
    redim endog_x_header(1 to endog_n_vars)
    
    if endog_x_header_in is nothing then
        for i = 1 to endog_n_vars
            endog_x_header(i) = "endog var " & i
        next i
    else
        for i = 1 to endog_x.columns.count
            endog_x_header(i) = endog_x_header_in(i)
        next i
    end if
            
    ' note: we don't need to process the instrument headers, since they won't be output
    
    ' process the y header range
    if y_header_in is nothing then
        y_header = "twoslsregtrain"
    else
        y_header = y_header_in(1)
    end if
    
    ' find and validate data dimensions
    ' ---------------------------------
    
    n = known_y.rows.count
    
    if not (exog_x is nothing) then
        if n <> exog_x.rows.count then
            twoslsregtrain = "number of exog_x and y rows should be the same."
            exit function
        end if
    end if
    
    if n <> endog_x.rows.count then
        twoslsregtrain = "number of endog_x and y rows should be the same."
        exit function
    end if
    
    if n <> instruments.rows.count then
        twoslsregtrain = "number of instruments rows and y rows should be the same."
        exit function
    end if
    
    ' ensure all data is numeric
    ' --------------------------
    
    if not (exog_x is nothing) then
        if application.worksheetfunction.count(exog_x) <> exog_x.columns.count * exog_x.rows.count then
            twoslsregtrain = "all exog_x values should be numeric."
            exit function
        end if
    end if
    
    if application.worksheetfunction.count(endog_x) <> endog_x.columns.count * endog_x.rows.count then
        twoslsregtrain = "all endog_x values should be numeric."
        exit function
    end if
    
    if application.worksheetfunction.count(instruments) <> instruments.columns.count * instruments.rows.count then
        twoslsregtrain = "all instrument values should be numeric."
        exit function
    end if
    
    if application.worksheetfunction.count(known_y) <> known_y.rows.count then
        twoslsregtrain = "all know_y values should be numeric."
        exit function
    end if

    ' construct the x, z, and y matrices
    ' ----------------------------------
    dim y() as variant
    dim x() as double
    dim z() as double
    
    redim y(1 to n)
    redim x(1 to n, 1 to exog_n_vars + endog_n_vars + include_constant)
    redim z(1 to n, 1 to exog_n_vars + instruments_n_vars + include_constant)
    
    ' deal with y
    y = known_y.value2
    
    ' create temporary variables to extract exog_x, endog_x, instruments
    dim exog_x_temp() as variant
    dim endog_x_temp() as variant
    dim instruments_temp() as variant
    
    if not (exog_x is nothing) then
        exog_x_temp = exog_x.value2
    end if
    endog_x_temp = endog_x.value2
    instruments_temp = instruments.value2
       
    ' deal with x, given by [ const? | x_exog | x_endog ]
    for i = 1 to n
        if include_constant = 1 then
            x(i, 1) = 1
        end if
        
        if not (exog_x is nothing) then
            jj = 1
            for j = 1 to exog_x.columns.count
                if exog_col_select_flags(j) = 1 then
                    x(i, include_constant + jj) = exog_x_temp(i, j)
                    jj = jj + 1
                end if
            next j
        end if
        
        for j = 1 to endog_x.columns.count
            x(i, include_constant + exog_n_vars + j) = endog_x_temp(i, j)
        next j
    next i
    
    ' deal wtih z, given by [ const? | x_exog | instruments ]
    for i = 1 to n
        if include_constant = 1 then
            z(i, 1) = 1
        end if
    
        if not (exog_x is nothing) then
            jj = 1
            for j = 1 to exog_x.columns.count
                if exog_col_select_flags(j) = 1 then
                    z(i, include_constant + jj) = exog_x_temp(i, j)
                    jj = jj + 1
                end if
            next j
        end if
        
        for j = 1 to instruments.columns.count
            z(i, include_constant + exog_n_vars + j) = instruments_temp(i, j)
        next j
    next i
    
    ' do the regression
    ' -----------------
    
    ' step 1
    dim ztz as variant
    dim ztzmone as variant
    dim zt as variant, zt2 as variant
        
    zt = largetranspose(z)
    ztz = largemmult(zt, z)
    
    'zt = worksheetfunction.transpose(z)
    'ztz = worksheetfunction.mmult(zt, z)
    
    if worksheetfunction.mdeterm(ztz) <= 0.00001 then
        twoslsregtrain = "error: your data is highly co-linear; cannot fit regression"
        exit function
    end if
    
    ztzmone = worksheetfunction.minverse(ztz)
    
    ' step 2
    dim ztx as variant
    dim xtz_ztzmone as variant
    
    ztx = largemmult(zt, x)
    'xtz_ztzmone = largemmult(largetranspose(ztx), ztzmone)
    
    'ztx = worksheetfunction.mmult(zt, x)
    xtz_ztzmone = worksheetfunction.mmult(worksheetfunction.transpose(ztx), ztzmone)
    
        
    ' step 3
    dim xtz_ztzmone_ztx as variant
    dim var_b_mat as variant
    
    'xtz_ztzmone_ztx = largemult(xtz_ztzmone, ztx)
    xtz_ztzmone_ztx = worksheetfunction.mmult(xtz_ztzmone, ztx)
    
    if worksheetfunction.mdeterm(xtz_ztzmone_ztx) <= 0.00001 then
        twoslsregtrain = "error: your data is highly co-linear; cannot fit regression"
        exit function
    end if
    
    var_b_mat = worksheetfunction.minverse(xtz_ztzmone_ztx)
    
    ' step 4
    dim xtz_ztzmone_zty as variant
    
    xtz_ztzmone_zty = largemmult(xtz_ztzmone, largemmult(zt, y))
    'xtz_ztzmone_zty = worksheetfunction.mmult(xtz_ztzmone, worksheetfunction.mmult(zt, y))
    
    
    ' step 5
    dim b_hat as variant
    
    b_hat = worksheetfunction.mmult(var_b_mat, xtz_ztzmone_zty)
    
    ' find the errors
    ' ---------------
    
    dim y_hat as variant
    dim ess as double
    dim s_squared as double
    
    y_hat = worksheetfunction.mmult(x, b_hat)
    
    for i = 1 to n
        ess = ess + (y(i, 1) - y_hat(i, 1)) ^ 2
    next i
    
    s_squared = ess / (n - (include_constant + exog_n_vars + endog_n_vars))
        
    ' build the result
    ' ----------------
    redim out(1 to 6, 1 to exog_n_vars + endog_n_vars + include_constant + 1)
    
    ' top left cell
    out(1, 1) = y_header & " (6," & exog_n_vars + endog_n_vars + include_constant + 1 & ")"
    
    ' headers
    for i = 1 to exog_n_vars + include_constant
        out(1, i + 1) = exog_x_header(i)
    next i
    
    for i = 1 to endog_n_vars
        out(1, i + exog_n_vars + include_constant + 1) = endog_x_header(i)
    next i
    
    ' coefficients
    out(2, 1) = "coefficients"
    
    for j = 1 to exog_n_vars + endog_n_vars + include_constant
        out(2, 1 + j) = b_hat(j, 1)
    next j
    
    ' standard errors
    out(3, 1) = "std error"
    
    for j = 1 to exog_n_vars + endog_n_vars + include_constant
        out(3, 1 + j) = math.sqr(s_squared * var_b_mat(j, j))
    next j
    
    ' p-values
    out(4, 1) = "p-value"
    
    for j = 1 to exog_n_vars + endog_n_vars + include_constant
        if out(3, 1 + j) = 0 then
            out(4, 1 + j) = 0
        else
            out(4, j + 1) = application.worksheetfunction.tdist(abs(out(2, j + 1) / out(3, j + 1)), n - (exog_n_vars + endog_n_vars + include_constant), 1) * 2
        end if
    next j
    
    ' other details
    out(5, 1) = "number valid obs"
    out(5, 2) = n
    
    out(6, 1) = "total obs"
    out(6, 2) = n
        
    ' pad the rest of rows 5 and 6 with spaces
    for i = 3 to (exog_n_vars + endog_n_vars + include_constant + 1)
        out(5, i) = ""
        out(6, i) = ""
    next i
        
    twoslsregtrain = out
     
    exit function
     
    'error handler
unexpected_error_twoslstrain:
    twoslsregtrain = "fatal error " & err.number & ":" & err.description
end function

 
' https://joannamsc.wordpress.com/2017/08/27/estimate-credit-score-with-logit-application-in-vba/
public function logisticregtrain(known_y as range, known_x as range, optional constant as boolean = true, _
optional not_implemented1 as long = 0, optional col_select as range = nothing, _
optional not_implemented2 as range = nothing)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: logisticregtrain(rng_y as range, rng_x as range, optional constant as boolean = true, _
optional na_action as long = 0, optional col_select as range = nothing, optional row_weight as range = nothing)
'runs logistic regression on data using newton's method for at most 50 iterations.
'reports error if no convergence is reached.
'input:
'   known_y (required): a range containing y. must be a single column. all entries must be 0/1.
'                    top row(s) may include header(s). only the first header row will be used when displaying output.
'                    headers must be non-numeric or formatted differently from the data (see listheaderrows property for
'                    more on how excel identifies header rows.)
'   known_x (required): a range containing x. all entries must be numeric. top row(s) may include header(s).
'                    only the first header row will be used when displaying output. number of rows not counting header(s)
'                    must match those of known_y.
'   constant: boolean indicating whether to include constant or not. assumed true if omitted.
'   na_action: included for backward compatibility. value not used.
'   col_select: a range containing 0/1. must be a single row. number of columns must match those of known_x. at least one
'            entry must be 1. if jth entry is 0, jth column is ignored from the known_x matrix. if omitted, all entries
'            are assumed to be 1.
'   row_weight: included for backward compatibility. value not used.
'
'output:
'   output is a matrix having 7 rows. number of columns depends on input. there is one column for each known_x column
'   included in the regression, one column for constant if included, and one more column for row labels.
'   row 1: used for column labels. if known_y includes a header, the first entry contains that header followed by
'       the size of the output matrix in brackets. otherwise, the first entry has the name of the function (logisticregtrain),
'       followed by the size of the output matrix in brackets. if constant is included, the second entry says
'       "constant". next are entries corresponding to columns included from known_x in order. if known_x contains a header,
'       it is used as column labels here. otherwise, the column labels are var1, var2, ... and so on.
'   row 2: first entry says "coefficients". later entries are the logistic regression constant (if included) and coefficients
'       corresponding to known_x columns that are included.
'   row 3: first entry says "p-value". later entries are the p-values for the coefficients in row 2.
'   row 4: first entry says "std error". later entries are the std errors for the coefficients in row 2.
'   row 5: first entry says "log-likelihood". next entry is the log-likelihood of the model at termination. this and subsequent
'       rows have only two entries.
'   row 6: first entry says "number valid obs". next entry is the number of rows in known_x.
'   row 7: first entry says "total obs". next entry is the number of rows in known_x.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    on error goto unexpected_error
    
    dim include_constant as long, x_header() as string, x_header_in as range, y_header as range, selectedvars as long, col_select_flags() as variant
    'count variables
    dim i as long, j as long, jj as long
    
    'check if we need to have constant
    if constant then
        include_constant = 1
    else
        include_constant = 0
    end if
    
    'handle header
    set x_header_in = util.getheaderrowsandresizerange(known_x)
    set y_header = util.getheaderrowsandresizerange(known_y)
    
    redim col_select_flags(1 to known_x.columns.count)
    if col_select is nothing then
        for j = 1 to known_x.columns.count
            col_select_flags(j) = 1
        next j
        selectedvars = known_x.columns.count
    else
        if col_select.columns.count <> known_x.columns.count then
            logisticregtrain = "number of entries in col_select must equal the total columns in x"
            exit function
        end if
        if col_select.rows.count <> 1 then
            logisticregtrain = "col_select must have exactly one row"
            exit function
        end if
        selectedvars = application.worksheetfunction.countif(col_select, "=1")
        if selectedvars = 0 then
            logisticregtrain = "at least one column must be selected"
            exit function
        end if
        if application.worksheetfunction.countif(col_select, "=0") + selectedvars <> known_x.columns.count then
            logisticregtrain = "all values in col_select must be zero or one"
            exit function
        end if
        for j = 1 to col_select.count
            col_select_flags(j) = col_select(j)
        next j
    end if
    
    if not x_header_in is nothing then
        redim x_header(1 to selectedvars)
        j = 1
        for i = 1 to known_x.columns.count
            if col_select_flags(i) = 1 then
                x_header(j) = x_header_in(i)
                j = j + 1
            end if
        next i
    end if

     
    'read data dimensions
    dim k as long, n as long
    
    n = known_y.rows.count ' n is total observation
    k = selectedvars + include_constant 'k is number of total explanatory variables
    
    ' validate the input
    if n <> known_x.rows.count then
        logisticregtrain = "number of x and y variables should be the same."
        exit function
    end if
    if application.worksheetfunction.count(known_x) <> known_x.columns.count * known_x.rows.count then
        logisticregtrain = "all x variables should be numeric."
        exit function
    end if
    if known_y.columns.count <> 1 then
        logisticregtrain = "known_y must have exactly one column"
        exit function
    end if
    if application.worksheetfunction.countif(known_y, "=0") + application.worksheetfunction.countif(known_y, "=1") <> n then
        logisticregtrain = "y variables should be zero or one."
        exit function
    end if
    
    ' get the data
    dim y() as variant
    redim y(1 to n)
    y = known_y.value2
    
    'adding a vector of ones to the x matrix if constant=1, name rng_x=x from now on
    dim x() as double
    redim x(1 to n, 1 to k)
    dim data_x() as variant
    data_x = known_x.value2
    for i = 1 to n
        x(i, 1) = 1
        jj = 1 + include_constant
        for j = 1 + include_constant to known_x.columns.count + include_constant
            if col_select_flags(j - include_constant) = 1 then
                x(i, jj) = data_x(i, j - include_constant)
                jj = jj + 1
            end if
        next j
    next i
    
    
    'initializing the coefficient vector (b) and the score (bx)
    dim b() as double, bx() as double, ybar as double
    redim b(1 to k) 'weights, coeffients
    redim bx(1 to n)
     
    ybar = application.worksheetfunction.average(y)
    if include_constant = 1 then b(1) = log(ybar / (1 - ybar))
    for i = 1 to n
        bx(i) = b(1)
    next i
     
    'defining the variables used in the newton procedure
    dim sensitivity as double, maxiter as integer, iter as integer, lnl as double
    dim lambda() as double, maxgradnorm as double, dlnl() as double, hesse() as double, hinv(), hinvg()
    redim lambda(1 to n)
    
    ' optimization parameters
    sensitivity = 1 * 10 ^ (-4): maxiter = 50
    
    ' inititalize the loop
    maxgradnorm = sensitivity + 1
    iter = 1: lnl = 0
     
    'loop for newton iteration
    do while iter < maxiter
        iter = iter + 1
        
        'reset derivative of log likelihood and hessian
        erase dlnl, hesse
        lnl = 0
        redim dlnl(1 to k): redim hesse(1 to k, 1 to k)
        
        'compute prediction lambda, gradient dlnl, hessian hesse, and log likelihood lnl
        for i = 1 to n 'n number of observation
            'doevents
            lambda(i) = logisticfunction(bx(i))
            'lambda(i) = 1 / (1 + exp(-bx(i)))
            for j = 1 to k
                dlnl(j) = dlnl(j) + (y(i, 1) - lambda(i)) * x(i, j)
                for jj = 1 to k
                    hesse(jj, j) = hesse(jj, j) - lambda(i) * (1 - lambda(i)) * x(i, jj) * x(i, j)
                next jj
            next j
            'lnl(iter) = lnl(iter) + y(i, 1) * log(1 / (1 + exp(-bx(i)))) + (1 - y(i, 1)) * log(1 - 1 / (1 + exp(-bx(i))))
            lnl = lnl + y(i, 1) * log(logisticfunction(bx(i))) + (1 - y(i, 1)) * log(1 - logisticfunction(bx(i)))
        next i
        
        ' compute the max-norm of the gradient
        maxgradnorm = 0
        for j = 1 to k
            if abs(dlnl(j)) > maxgradnorm then
                maxgradnorm = abs(dlnl(j))
            end if
        next j
      
        
        'compute inverse hessian (=hinv) and multiply hinv with gradient dlnl
        hinv = application.worksheetfunction.minverse(hesse)
        hinvg = application.worksheetfunction.mmult(dlnl, hinv)
               
        'if convergence achieved, exit now and keep the b corresponding with the estimated hessian
        if maxgradnorm <= sensitivity then exit do
        
        ' apply newton's scheme for updating coefficients b
        for j = 1 to k
            b(j) = b(j) - hinvg(j)
        next j
        
        'compute new score (bx)
        for i = 1 to n
            bx(i) = 0
            for j = 1 to k
                bx(i) = bx(i) + b(j) * x(i, j)
            next j
        next i
     
    loop
     
    'some error handling
    if iter >= maxiter then
        logisticregtrain = "maximum number of iteration exceeded. no convergence achieved."
        exit function
    end if
          
    'output
    dim relogit()
    redim relogit(1 to 7, 1 to k + 1)
    
    if y_header is nothing then
        relogit(1, 1) = "logisticregtrain (7," & k + 1 & ")"
    else
        relogit(1, 1) = y_header(1, 1) & " (7," & k + 1 & ")"
    end if
         
    'coefficients
    relogit(2, 1) = "coefficients"
    for j = 1 to k 'k variables
        if x_header_in is nothing then
            relogit(1, j + 1) = "var" & j - include_constant
        elseif j > include_constant then
            relogit(1, j + 1) = x_header(j - include_constant)
        end if
        relogit(2, j + 1) = b(j)
    next j
    if constant then
        relogit(1, 2) = "constant"
    end if

     
    'additional statistics if requested
    relogit(3, 1) = "std error"
    relogit(4, 1) = "p-value"
    for j = 1 to k
        dim std_err as double, z as double, pvalue as double
        if k > 1 then
            std_err = sqr(-hinv(j, j))
        else
            'if k=1, this is a 1d array. have to do this hack for one specific case when there is no constant, and only one column in x.
            std_err = sqr(-hinv(j))
        end if
        
        z = b(j) / std_err
        pvalue = (application.worksheetfunction.normsdist(-abs(z))) * 2
        relogit(3, j + 1) = std_err
        relogit(4, j + 1) = pvalue
        relogit(5, j + 1) = ""
        relogit(6, j + 1) = ""
        relogit(7, j + 1) = ""
    next j
       
    'additional statistics
    relogit(5, 1) = "log-likelihood"
    relogit(5, 2) = lnl
    relogit(6, 1) = "number valid obs"
    relogit(6, 2) = n
    relogit(7, 1) = "total obs"
    relogit(7, 2) = n
    
    logisticregtrain = relogit
     
    exit function
     
    'error handler
unexpected_error:
    logisticregtrain = "fatal error: " & err.description
end function


private sub registerlinrtrain()
'function signature:  logisticregtrain(rng_y as range, rng_x as range, optional include_constant as boolean = true, _
optional dummy as long = 0, optional col_select_in as range = nothing, optional dummy2 as range = nothing)
    'delclaring the necessary variables
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    'depending on the function arguments define the necessary variables on the array.
    dim argdesc(1 to 8) as string
    
    funcname = "linearregtrain"
    
    'here we add the function's description.
    funcdesc = "array function to train a linear regression model. user can optionally select labels in the top row." & _
    vbnewline & vbnewline & "output:" & vbnewline & "1st row: coefficients" & vbnewline & "2nd row: std error" & vbnewline & _
    "3rd row: p-value" & vbnewline & "4th row: r squared"
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    'you can also use instead of numbers the full category name, for example:
    'funccat = "engineering"
    'or you can define your own custom category:
    'funccat = "my vba functions"
    
    'here we add the description for the function's arguments.
    argdesc(1) = "known y's"
    argdesc(2) = "known x's"
    argdesc(3) = "if true or omitted, a constant is included in the regression."
    argdesc(4) = "included for backward compatibility. value ignored."
    argdesc(5) = "a vector of 0/1. if omitted, all columns in known_x are considered for regression." & _
                 "if specified, and if the jth entry is 0 in this vector, the jth column in known_x matrix is ignored."
    argdesc(6) = "included for backward compatibility. value ignored."
    argdesc(7) = "included for backward compatibility. value ignored."
    argdesc(8) = "included for backward compatibility. value ignored."

    util.functiondescription funcname, funcdesc, funccat, argdesc
end sub

private sub registertwoslsregtrain()
'function signature:  public function twoslsregtrain(known_y as range, optional exog_x as range = nothing, optional endog_x as range = nothing, _
                                optional instruments as range = nothing, optional constant as boolean = true, _
                                    optional not_implemented2 as long = 0, optional exog_col_select as range = nothing, _
                                    optional not_implemented3 as range = nothing, optional not_implemented4 as boolean)
                                    
    'delclaring the necessary variables
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    'depending on the function arguments define the necessary variables on the array.
    dim argdesc(1 to 9) as string
    
    funcname = "twoslsregtrain"
    
    'here we add the function's description.
    funcdesc = "array function to train a two-stage least squares (2sls) regression model. user can optionally select labels in the top row." & _
    vbnewline & vbnewline & "output:" & vbnewline & "1st row: coefficients" & vbnewline & "2nd row: std error" & vbnewline & _
    "3rd row: p-value"
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    'you can also use instead of numbers the full category name, for example:
    'funccat = "engineering"
    'or you can define your own custom category:
    'funccat = "my vba functions"
    
    'here we add the description for the function's arguments.
    argdesc(1) = "known y's"
    argdesc(2) = "exogenous x's (can be omitted)"
    argdesc(3) = "endogenous x's"
    argdesc(4) = "instruments"
    argdesc(5) = "if true or omitted, include a constant term in the first and second stage regression"
    argdesc(6) = "included for backward compatibility. value ignored."
    argdesc(7) = "a vector of 0/1. if omitted, all columns in exog_x are considered for regression." & _
                 "if specified, and if the jth entry is 0 in this vector, the jth column in exog_x is ignored."
    argdesc(8) = "included for backward compatibility. value ignored."
    argdesc(9) = "included for backward compatibility. value ignored."
    
    util.functiondescription funcname, funcdesc, funccat, argdesc
end sub


private sub registerlinrpredict()
'function signature:  linearregpredict(model as range, known_x as range, optional constant as boolean = true, _
optional col_select as range = nothing)
    'delclaring the necessary variables
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    'depending on the function arguments define the necessary variables on the array.
    dim argdesc2(1 to 4) as string
    
    funcname = "linearregpredict"
    
    'here we add the function's description.
    funcdesc = "array function to predict y using a linear regression model on known x." & _
    vbnewline & vbnewline & "output:" & vbnewline & "estimate of y for each row of known x"
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    'you can also use instead of numbers the full category name, for example:
    'funccat = "engineering"
    'or you can define your own custom category:
    'funccat = "my vba functions"
    
    'here we add the description for the function's arguments.
    argdesc2(1) = "model coefficients"
    argdesc2(2) = "known x's"
    argdesc2(3) = "should be set to true or omitted if a constant is included when generating the model."
    argdesc2(4) = "a vector of 0/1. should match col_select argument given at the time of training for correct prediction."
    
    util.functiondescription funcname, funcdesc, funccat, argdesc2
end sub

private sub registerlogrtrain()
'function signature:  logisticregtrain(rng_y as range, rng_x as range, optional include_constant as boolean = true, _
optional dummy as long = 0, optional col_select_in as range = nothing, optional dummy2 as range = nothing)
    'delclaring the necessary variables
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    'depending on the function arguments define the necessary variables on the array.
    dim argdesc(1 to 6) as string
    
    funcname = "logisticregtrain"
    
    'here we add the function's description.
    funcdesc = "array function to train a logistic regression model. user can optionally select labels in the top row." & _
    vbnewline & vbnewline & "output:" & vbnewline & "1st row: coefficients" & vbnewline & "2nd row: std error" & vbnewline & _
    "3rd row: p-value" & vbnewline & "4th row: log-likelihood"
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    'you can also use instead of numbers the full category name, for example:
    'funccat = "engineering"
    'or you can define your own custom category:
    'funccat = "my vba functions"
    
    'here we add the description for the function's arguments.
    argdesc(1) = "known y's"
    argdesc(2) = "known x's"
    argdesc(3) = "if true or omitted, a constant is included in the regression."
    argdesc(4) = "included for backward compatibility. value ignored."
    argdesc(5) = "a vector of 0/1. if omitted, all columns in known_x are considered for regression." & _
                 "if specified, and if the jth entry is 0 in this vector, the jth column in known_x matrix is ignored."
    argdesc(6) = "included for backward compatibility. value ignored."

    util.functiondescription funcname, funcdesc, funccat, argdesc
end sub


private sub registerlogrpredict()
'function signature:  logisticregpredict(model as range, known_x as range, optional include_constant as boolean = true, optional col_select_in as range = nothing, optional dummy as boolean = false)
    'delclaring the necessary variables
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    'depending on the function arguments define the necessary variables on the array.
    dim argdesc2(1 to 5) as string
    
    funcname = "logisticregpredict"
    
    'here we add the function's description.
    funcdesc = "array function to predict the probability of y being 1 using a logistic regression model on known x." & _
    vbnewline & vbnewline & "output:" & vbnewline & "probability of y being 1 for each row of known x"
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    'you can also use instead of numbers the full category name, for example:
    'funccat = "engineering"
    'or you can define your own custom category:
    'funccat = "my vba functions"
    
    'here we add the description for the function's arguments.
    argdesc2(1) = "model coefficients"
    argdesc2(2) = "known x's"
    argdesc2(3) = "should be set to true or omitted if a constant is included when generating the model."
    argdesc2(4) = "a vector of 0/1. should match col_select argument given at the time of training for correct prediction."
    argdesc2(5) = "included for backward compatibility. value ignored."
    
    util.functiondescription funcname, funcdesc, funccat, argdesc2
end sub



public sub registerall()
    registerlogrtrain
    registerlogrpredict
    registerlinrtrain
    registerlinrpredict
    registertwoslsregtrain
end sub





