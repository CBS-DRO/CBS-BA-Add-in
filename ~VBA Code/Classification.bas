option explicit


private function internal_confusionmatrix(actual() as variant, prediction() as variant, n as long)
' internal function to compute the confusiont matrix. no checking is done here and we operate on arrays instead of ranges.
dim confusion_matrix(1 to 5, 1 to 4) as variant
confusion_matrix(1, 1) = "confusion matrix (5,4)"
confusion_matrix(1, 2) = "predicted"
confusion_matrix(1, 3) = ""
confusion_matrix(1, 4) = ""
confusion_matrix(2, 1) = "actual"
confusion_matrix(3, 1) = 1
confusion_matrix(4, 1) = 0
confusion_matrix(5, 1) = "total"
confusion_matrix(2, 2) = 1
confusion_matrix(2, 3) = 0
confusion_matrix(2, 4) = "total"

dim j as long
for j = 1 to n
    if actual(j, 1) = 1 and prediction(j, 1) = 1 then
        ' true positives
        confusion_matrix(3, 2) = confusion_matrix(3, 2) + 1
    elseif actual(j, 1) = 1 and prediction(j, 1) = 0 then
        ' false negatives
        confusion_matrix(3, 3) = confusion_matrix(3, 3) + 1
    elseif actual(j, 1) = 0 and prediction(j, 1) = 1 then
        ' false positives
        confusion_matrix(4, 2) = confusion_matrix(4, 2) + 1
    elseif actual(j, 1) = 0 and prediction(j, 1) = 0 then
        ' true negatives
        confusion_matrix(4, 3) = confusion_matrix(4, 3) + 1
    end if
next j

' total actual positives
confusion_matrix(3, 4) = confusion_matrix(3, 2) + confusion_matrix(3, 3)
' total actual negatives
confusion_matrix(4, 4) = confusion_matrix(4, 2) + confusion_matrix(4, 3)
' total predicted positives
confusion_matrix(5, 2) = confusion_matrix(3, 2) + confusion_matrix(4, 2)
' total predicted negatives
confusion_matrix(5, 3) = confusion_matrix(3, 3) + confusion_matrix(4, 3)

' overall total
confusion_matrix(5, 4) = confusion_matrix(5, 2) + confusion_matrix(5, 3)
    

internal_confusionmatrix = confusion_matrix
end function


public function confusionmatrix(actual as range, prediction as range)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'function to compute confusion matrix based on actual and predicted results. this function does necessary input sanity
'checks and then calls internal_confusionmatrix to do actual computation.
'inputs:
'  actual: actual results, a column of 0/1
'  prediction: predicted results, a columns of 0/1
'output:
'  an array of size 5x4 valued as follows:
'   confusion matrix (5, 4) \tab predicted                  \tab <blank>                    \tab <blank>
'   actual                  \tab 1                          \tab 0                          \tab total
'   1                       \tab true positives             \tab false negatives            \tab total actual positives
'   0                       \tab false positives            \tab true negatives             \tab total actual negatives
'   total                   \tab total predicted positives  \tab total predicted negatives  \tab overall total
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    on error goto unexpected_error_cfm

    if actual.columns.count <> 1 or prediction.columns.count <> 1 then
        confusionmatrix = "actual and prediction, both should consist of a single column"
        exit function
    end if
    
    if actual.rows.count <> prediction.rows.count then
        confusionmatrix = "actual and prediction should both have the same number of rows"
        exit function
    end if
    if application.worksheetfunction.count(actual) <> actual.rows.count or application.worksheetfunction.count(prediction) <> prediction.rows.count then
        confusionmatrix = "all entries must be numeric"
        exit function
    end if
    if application.worksheetfunction.countif(actual, "=0") + application.worksheetfunction.countif(actual, "=1") <> actual.rows.count then
        confusionmatrix = "all actual entries should be zero or one."
        exit function
    end if
    if application.worksheetfunction.countif(prediction, "=0") + application.worksheetfunction.countif(prediction, "=1") <> prediction.rows.count then
        confusionmatrix = "all prediction entries should be zero or one."
        exit function
    end if
    
    dim act_arr() as variant, pred_arr() as variant, n as long
    n = actual.rows.count
    redim act_arr(1 to n)
    redim pred_arr(1 to n)
    
    act_arr = actual.value2
    pred_arr = prediction.value2
    
    ' sanity check is over. call internal_confusionmatrix to do actual computation on arrays
    confusionmatrix = internal_confusionmatrix(act_arr, pred_arr, n)
    exit function
unexpected_error_cfm:
    confusionmatrix = "fatal error: " & err.description
end function


private function trapezoid_area(x1 as long, x2 as long, y1 as long, y2 as long)
' function to compute area under a line from (x1, y1) to (x2, y2).
dim base as long, height as double
if x1 > x2 then
    base = x1 - x2
else
    base = x2 - x1
end if
height = (y1 + y2) / 2

trapezoid_area = base * height
end function


private function internal_auc_faster(score() as variant, actual() as variant, num_pos as long, num_neg as long)
' internal function to compute auc by computing area under line segments.
dim indices() as long, j as long, n as long, score2() as variant
n = num_pos + num_neg
redim indices(1 to n)
redim score2(1 to n)

' make indices array and convert score to a one dimensional array to prepare for sorting.
for j = 1 to n
    indices(j) = j
next j
score2 = util.makeoned(score)

' after this, score2(indices(.)) is a sorted version of score2(.)
util.mergesort score2, indices

dim fp_prev as long, fp as long, tp_prev as long, tp as long, area as double, score_prev as double, percentage as double
fp_prev = 0
fp = 0
tp_prev = 0
tp = 0
area = 0
score_prev = score(indices(1), 1) - 1

' go from highest score to lowest
for j = n to 1 step -1
    doevents
    
    if score_prev <> score(indices(j), 1) then
        ' if the score has changed (gone strictly lower), then update the area
        area = area + trapezoid_area(fp, fp_prev, tp, tp_prev)
        score_prev = score(indices(j), 1)
        fp_prev = fp
        tp_prev = tp
    end if
    ' decide whether we are moving up or to the right on the roc curve.
    if actual(indices(j), 1) = 1 then
        tp = tp + 1
    else
        fp = fp + 1
    end if
next j
' add residual area
area = area + trapezoid_area(num_neg, fp_prev, num_pos, tp_prev)

application.statusbar = false

' normalize
internal_auc_faster = area / (num_pos * num_neg)

end function


private function internal_auc(score() as variant, actual() as variant, n as long) as double
' internal function to compute auc using concordance. this function is not being used since version 0.0.5
dim i as long, j as long, concordance as double, total_unequal as long

for i = 1 to n
    debug.print i
    for j = 1 to i - 1
        doevents
        if actual(i, 1) <> actual(j, 1) then
            'one is 0 other is 1. if 1 has a higher score than 0 then count this pair
            if actual(i, 1) = 0 and score(i, 1) < score(j, 1) then
                concordance = concordance + 1
            elseif actual(j, 1) = 0 and score(j, 1) < score(i, 1) then
                concordance = concordance + 1
            elseif score(j, 1) = score(i, 1) then
                'if both have the same score, add 0.5
                concordance = concordance + 0.5
            end if
            total_unequal = total_unequal + 1
        end if
    next j
next i

if total_unequal = 0 then
    internal_auc = 0
else
    internal_auc = concordance / total_unequal
end if

end function


public function auc(score as range, actual as range, optional reverse as boolean = false)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'function to compute area under the roc curve based on scores given to data points and their actual values.
'input:
'   score: score given to a data point by our model
'   actual: actual realized value must be 0/1
'   reverse(optional): if true, higher score is an indication of actual realization to be 0. assumed false if omitted.
'output:
'   the area under the roc curve.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

on error goto unexpected_error_auc

    if actual.columns.count <> 1 or score.columns.count <> 1 then
        auc = "actual and score, both should consist of a single column"
        exit function
    end if
    
    dim num_pos as long, num_neg as long
    
    if actual.rows.count <> score.rows.count then
        auc = "actual and score should both have the same number of rows"
        exit function
    end if
    if application.worksheetfunction.count(actual) <> actual.rows.count or application.worksheetfunction.count(score) <> score.rows.count then
        auc = "all entries must be numeric"
        exit function
    end if
    num_pos = application.worksheetfunction.countif(actual, "=1")
    num_neg = application.worksheetfunction.countif(actual, "=0")
    if num_pos + num_neg <> actual.rows.count then
        auc = "all actual entries should be zero or one."
        exit function
    end if
    if num_pos = 0 then
        auc = "at least one entry must have actual outcome as 1"
        exit function
    end if
    if num_neg = 0 then
        auc = "at least one entry must have actual outcome as 0"
        exit function
    end if
    
    
    dim n as long
    n = score.rows.count
    dim score_arr() as variant, actual_arr() as variant
    redim score_arr(1 to n)
    redim actual_arr(1 to n)
    
    application.screenupdating = false
    application.displaystatusbar = true
    application.statusbar = "evaluating macro..."
    score_arr = score.value2
    actual_arr = actual.value2
    
    'input has been sanity checked and imported into array from ranges. now call the internal function to do the computation.
    'note that <auc if reverse=true> = 1 - <auc if reverse=false>. the internal functions always assume reverse=false
    if not reverse then
        'auc = internal_auc(score_arr, actual_arr, n)
        auc = internal_auc_faster(score_arr, actual_arr, num_pos, num_neg)
    else
        'auc = 1 - internal_auc(score_arr, actual_arr, n)
         auc = 1 - internal_auc_faster(score_arr, actual_arr, num_pos, num_neg)
    end if
    
    exit function
    
unexpected_error_auc:
    auc = "fatal error: " & err.description
    
end function

private function internal_roc(score() as variant, actual() as variant, threshold() as variant, cost() as variant, reversed as boolean, _
n as long, m as long, num_pos as long, num_neg as long, cost_included as boolean)

dim i as long, j as long, pos_count as long, neg_count as long, tp as long, fp as long
dim score2() as variant, indices() as long, thres2() as variant, thresindices() as long, rocarr() as double
redim score2(1 to n)
redim indices(1 to n)
redim thresindices(1 to m)
redim thres2(1 to n)

' size of the output depends on whether to include a column for cost or not.
if cost_included then
    redim rocarr(1 to m, 1 to 3)
else
    redim rocarr(1 to m, 1 to 2)
end if

' we need to sort scores and thresholds. create indices array and make them one-dimensional
for j = 1 to n
    indices(j) = j
next j
for i = 1 to m
    thresindices(i) = i
next i
score2 = util.makeoned(score)
thres2 = util.makeoned(threshold)

' after this, score2(indices(.)) will be a sorted version of score2(.)
util.mergesort score2, indices
' after this thres2(thresindices(.)) will be a sorted version of thres2(.)
util.mergesort thres2, thresindices

' remainingarr denotes whether there is more of score array to be traversed.
dim remainingarr as boolean
remainingarr = true
j = 1
pos_count = 0
neg_count = 0
for i = 1 to m
    ' for now, assuming reversed=false. more precisely, if score is strictly greater, model predicts 1.
    ' for equal or lesser score, model predicts 0
    if remainingarr then
        do while score2(indices(j)) <= thres2(thresindices(i))
            ' all these indices are predicted to be 0 according to this threshold.
            ' maintain count of actual positives and negatives in pos_count and neg_count respectively.
            if actual(indices(j), 1) = 1 then
                pos_count = pos_count + 1
            else
                neg_count = neg_count + 1
            end if
            j = j + 1
            if j > n then
                ' score2 has been fully traversed. the rest of the thresholds will always predict 0 for all points.
                remainingarr = false
                goto out_of_loop
            end if
        loop
    end if
out_of_loop:
    ' pos_count now contains false negatives. neg_count contains true negatives.
    tp = num_pos - pos_count
    fp = num_neg - neg_count
    ' find true and false positive rates by normalizing by number of total positives and total negatives.
    ' if reversed is true, we also need to subtract the rate from 1.
    if reversed then
        rocarr(i, 1) = 1 - fp / num_neg
        rocarr(i, 2) = 1 - tp / num_pos
    else
        rocarr(i, 1) = fp / num_neg
        rocarr(i, 2) = tp / num_pos
    end if
    if cost_included then
        ' compute cost from cost matrix provided.
        rocarr(i, 3) = cost(1, 1) * rocarr(i, 2) * num_pos + _
                            cost(2, 1) * rocarr(i, 1) * num_neg + cost(1, 2) * (1 - rocarr(i, 2)) * num_pos + _
                            cost(2, 2) * (1 - rocarr(i, 1)) * num_neg
    end if
next i

internal_roc = rocarr
end function


public function roc(score as range, actual as range, threshold as range, optional cost as range = nothing, optional reversed as boolean = false)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'function to compute roc curve given scores given by a model, actual realizations and a range of thresholds.
'input
'   score(required): score used to decide if the actual outcomes would be 1. higher score indicates a higher probability
'       of the outcome being 1. must be a single column.
'   actual(required): actual values (0 or 1). must be a column with the same number of rows as score.
'   threshold(required): threshold above which model predicts 1. must be a column.
'   cost: cost matrix. must be 2 by 2 matrix detailing costs for each of four cases -- true positive, false negative,
'       false positive and true negative.
'   reversed: if set to true, then higher score is indicative of a higher probability of the outcomes being 0 instead of 1.
'       if omitted, assumed to be false.
'output
'   two columns, each having one row for each row of threshold, representing true and false positive rates of the model
'   if operating on threshold. it is assumed that the model outputs 0 if score equals or is less than threshold
'   (if reversed is false. if reversed is true, equality outputs 1).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

on error goto unexpected_error_roc

    if actual.columns.count <> 1 or score.columns.count <> 1 or threshold.columns.count <> 1 then
        roc = "actual, threshold and score, all should consist of a single column"
        exit function
    end if
    
    if actual.rows.count <> score.rows.count then
        roc = "actual and score should both have the same number of rows"
        exit function
    end if
    if application.worksheetfunction.count(actual) <> actual.rows.count _
    or application.worksheetfunction.count(score) <> score.rows.count _
    or application.worksheetfunction.count(threshold) <> threshold.rows.count then
        roc = "all entries must be numeric"
        exit function
    end if
    if application.worksheetfunction.countif(actual, "=0") + application.worksheetfunction.countif(actual, "=1") <> actual.rows.count then
        roc = "all actual entries should be zero or one."
        exit function
    end if
    
    dim n as long, m as long, num_pos as long, num_neg as long
    n = score.rows.count
    m = threshold.rows.count
    
    ' if we have a single threshold, return an error
    if m = 1 then
        roc = "error: this function only works with multiple thresholds; you selected one threshold only"
        exit function
    end if
    
    
    dim score_arr() as variant, actual_arr() as variant
    dim thres_arr() as variant, cost_arr() as variant, cost_included as boolean
    redim score_arr(1 to n)
    redim actual_arr(1 to n)
    redim thres_arr(1 to m)
    
    cost_included = false
    if not cost is nothing then
        if cost.rows.count <> 2 then
            roc = "cost matrix must have exactly 2 rows."
            exit function
        end if
        if cost.columns.count <> 2 then
            roc = "cost matrix must have exactly 2 columns."
            exit function
        end if
        if application.worksheetfunction.count(cost) <> 4 then
            roc = "cost matrix must have all entries numeric"
            exit function
        end if
        redim cost_arr(1 to 2, 1 to 2)
        cost_arr = cost.value2
        cost_included = true
    end if
    
    score_arr = score.value2
    actual_arr = actual.value2
    thres_arr = threshold.value2

    num_pos = application.worksheetfunction.countif(actual, "=1")
    num_neg = application.worksheetfunction.countif(actual, "=0")
    
    ' sanity checks are done, input has been copied from range to arrays. now call the internal function for actual computation.
    roc = internal_roc(score_arr, actual_arr, thres_arr, cost_arr, reversed, n, m, num_pos, num_neg, cost_included)
    exit function
    
unexpected_error_roc:
    roc = "fatal error: " & err.description
end function

private sub registerauc()
'function signature:  auc(score as range, actual as range, optional reverse as boolean = false)
    'delclaring the necessary variables
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    'depending on the function arguments define the necessary variables on the array.
    dim argdesc2(1 to 3) as string
    
    funcname = "auc"
    
    'here we add the function's description.
    funcdesc = "function to compute the area under an roc curve." & vbnewline & vbnewline & _
    "output:" & vbnewline & "area under the roc curve"
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    'you can also use instead of numbers the full category name, for example:
    'funccat = "engineering"
    'or you can define your own custom category:
    'funccat = "my vba functions"
    
    'here we add the description for the function's arguments.
    argdesc2(1) = "estimated probability of actual values being 1"
    argdesc2(2) = "actual values (0 or 1)"
    argdesc2(3) = "set to true if score is the probability of value being 0. assumed to be false if omitted."
    
    util.functiondescription funcname, funcdesc, funccat, argdesc2
end sub

private sub registerconfusionmatrix()

'function signature:  confusionmatrix(actual as range, prediction as range)
    'delclaring the necessary variables
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    'depending on the function arguments define the necessary variables on the array.
    dim argdesc2(1 to 2) as string
    
    funcname = "confusionmatrix"
    
    'here we add the function's description.
    funcdesc = "array function to compute a confusion matrix from actual and predicted values." & _
    vbnewline & vbnewline & "input:" & vbnewline & "1. column of actual values (0 or 1)" & vbnewline & _
    "2. column of predicted values (0 or 1)" & vbnewline & vbnewline & _
    "output:" & vbnewline & "probability of y being 1 for each row of known x"
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    'you can also use instead of numbers the full category name, for example:
    'funccat = "engineering"
    'or you can define your own custom category:
    'funccat = "my vba functions"
    
    'here we add the description for the function's arguments.
    argdesc2(1) = "actual values (0 or 1)"
    argdesc2(2) = "predicted values (0 or 1)"
    
    util.functiondescription funcname, funcdesc, funccat, argdesc2
end sub


private sub registerroc()

'function signature:  roc(score as range, actual as range, threshold as range, optional cost as range = nothing, optional reversed as boolean = false)
    'delclaring the necessary variables
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    'depending on the function arguments define the necessary variables on the array.
    dim argdesc2(1 to 5) as string
    
    funcname = "roc"
    
    'here we add the function's description.
    funcdesc = "array function to compute an roc matrix from actual values and scores for a range of thresholds." & _
    vbnewline & "output:" & vbnewline & "for each threshold, two columns, one each for false and true positive rates." & _
    "an additional column for total cost, if cost is present." '& _
    '"first two columns are false positive and true positive rates respectively. third column is the total cost incurred for that threshold."
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    'you can also use instead of numbers the full category name, for example:
    'funccat = "engineering"
    'or you can define your own custom category:
    'funccat = "my vba functions"
    
    'here we add the description for the function's arguments.
    argdesc2(1) = "scores given to each case"
    argdesc2(2) = "actual values (0 or 1)"
    argdesc2(3) = "column of threshold. if score is strictly above this, model predicts 1 if reversed is false." 'if reversed is true, model predicts 0 if score is strictly above threshold."
    argdesc2(4) = "cost matrix (2 by 2)."
    argdesc2(5) = "whether model outputs 1 if score is greater than threshold. default is false."
    
    util.functiondescription funcname, funcdesc, funccat, argdesc2
end sub


public sub registerall()
    registerconfusionmatrix
    registerauc
    registerroc
end sub