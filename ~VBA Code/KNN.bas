option explicit


private sub clean_array(a() as variant)
    ' make sure that all entries are either numbers or set to empty if non-numeric
    dim i as long, j as long
    
    for i = lbound(a, 1) to ubound(a, 1)
        for j = lbound(a, 2) to ubound(a, 2)
            if application.worksheetfunction.isnumber(a(i, j)) = false then
                a(i, j) = empty
            end if
        next j
    next i
end sub


private function check_array_numeric(a() as variant)
    ' make sure that all entries are either numbers or set to empty if non-numeric
    dim i as long, j as long
    
    for i = lbound(a, 1) to ubound(a, 1)
        for j = lbound(a, 2) to ubound(a, 2)
            if application.worksheetfunction.isnumber(a(i, j)) = false and not isempty(a(i, j)) then
                if trim(a(i, j)) = "" then
                    a(i, j) = empty
                else
                    check_array_numeric = false
                    exit function
                end if
            end if
        next j
    next i

    check_array_numeric = true

end function


private function average(a() as variant, n as long, col as long, optional excluderow as long = 0)
' utility function to compute average of col column of a() while excluding excluderow row
dim i as long, count as long, sum as double
count = 0
sum = 0
for i = 1 to n
    ' exclude excluderow row and empty cells while computing sum and count
    if (not isempty(a(i, col))) and i <> excluderow then
        sum = sum + a(i, col)
        count = count + 1
    end if
next i

' return empty if there were no valid values.
if count = 0 then
    average = empty
else
    average = sum / count
end if

end function


private function std_dev(a() as variant, n as long, col as long, optional excluderow as long = 0)
' utility function to compute standard deviation of col column of a() while excluding excluderow row
dim mean as variant, i as long, variance as double, count as long

' first compute the mean. if it is empty, return empty
mean = average(a, n, col, excluderow)
if isempty(mean) then
    std_dev = empty
    exit function
end if

variance = 0
count = 0
for i = 1 to n
    ' compute sum of centered squares while excluding excluderow and empty cells from the sum
    if (not isempty(a(i, col))) and i <> excluderow then
        variance = variance + (a(i, col) - mean) * (a(i, col) - mean)
        count = count + 1
    end if
next i

' note that count must be positive at this stage since otherwise, mean would have been empty
' and we would have returned after computing it.
std_dev = sqr(variance / count)
end function


private function donormalize(x() as variant, m as long, n as long, d as long, optional givenavg as variant, _
optional givenstddev as variant, optional exclude_row as long = 0) as variant
' utility function to normalize an array, optionally using given average and std dev. if one is given, both must be given.
' exclude_row is kept unchanged and only rows from m to n are affected.
dim normalizedx() as variant
redim normalizedx(1 to n, 1 to d)
dim avrg as variant, std as variant, i as long, j as long

' iterate over all columns
for j = 1 to d
    ' compute avrg and std dev if not given.
    if ismissing(givenavg) then
        avrg = average(x, n, j)
        std = std_dev(x, n, j)
    else
        avrg = givenavg(j)
        std = givenstddev(j)
    end if
    
    for i = m to n
        if i = exclude_row or std = 0 then
            ' if row is to be excluded or if std dev is 0, leave the cell unchanged
            normalizedx(i, j) = x(i, j)
        elseif isempty(avrg) or isempty(x(i, j)) then
            ' leave empty cells empty.
            normalizedx(i, j) = empty
        else
            normalizedx(i, j) = (x(i, j) - avrg) / std
        end if
    next i
next j

donormalize = normalizedx
end function


private function knn_dist_row(a() as variant, b() as variant, row_a as long, row_b as long, _
scaling() as variant) as double
    ' utility function to compute scaled distance between row_a row of a() and row_b row of b().
    ' assumes all arrays are 2d.
    dim dist as double, sum_weights as double
    
    dist = 0
    sum_weights = 0
    
    dim i as long
    for i = lbound(a, 2) to ubound(a, 2)
        if not (isempty(a(row_a, i))) and not (isempty(b(row_b, i))) then
            ' add scaled squared difference between cells if they are non-empty. accumulate scaling factors.
            dist = dist + scaling(1, i) ^ 2 * ((a(row_a, i) - b(row_b, i)) ^ 2)
            sum_weights = sum_weights + scaling(1, i) ^ 2
        end if
    next
    
    if sum_weights > 0 then
        ' normalize distances by sum of scaling
        knn_dist_row = sqr(dist) / sqr(sum_weights)
    else
        ' if sum_weights = 0, return a large number.
        knn_dist_row = 1e+31
    end if
end function


private function validate_knn_functions(x as range, optional k as integer, optional y as range = nothing, _
optional reference as range = nothing, optional dist_scale_vec as range = nothing)
    ' one common function to validate all knn related functions.
    dim n as long, d as long
    
    ' n is the number of rows/observations
    n = x.rows.count
    ' d is the number of cols/attributes
    d = x.columns.count
    
    if k < 0 then
        validate_knn_functions = "k cannot be negative"
        exit function
    end if
    if k >= n then
        validate_knn_functions = "k must be smaller than the number of entries in known_x"
        exit function
    end if
 
    
    if not y is nothing then
        if n <> y.rows.count then
            validate_knn_functions = "known_x and known_y must have the same number of rows."
            exit function
        end if
        
        if y.columns.count <> 1 then
            validate_knn_functions = "known_y must have a single column"
            exit function
        end if
    end if
    
    if not dist_scale_vec is nothing then
        if dist_scale_vec.columns.count <> d then
            validate_knn_functions = "dist scale vec should have same number of columns as known_x"
            exit function
        end if
        if dist_scale_vec.rows.count <> 1 then
            validate_knn_functions = "dist scale vec should have exactly 1 row"
            exit function
        end if
    end if
    
    if not reference is nothing then
        if d <> reference.columns.count then
            validate_knn_functions = "known_x and reference should have the same number of columns."
            exit function
        end if
        if reference.rows.count <> 1 then
            validate_knn_functions = "reference range should have one row."
            exit function
        end if
    end if
    
    
    validate_knn_functions = ""
end function


private function get_dist_scaling(dist_scale_vec as range, d as long)
    ' utility function to get 2-d array scaling from dist_scale_vec, which is a range.
    dim scaling() as variant, i as long
    if dist_scale_vec is nothing then
        ' if not provided, default scaling is a vector of 1's.
        redim scaling(1 to 1, 1 to d)
        for i = 1 to d
            scaling(1, i) = 1
        next i
    else
        scaling = dist_scale_vec.value2
    end if
    
  get_dist_scaling = scaling
end function


private function internal_knn_dist(a() as variant, b() as variant, scaling() as variant, n as long, _
d as long, normalize as boolean, optional b_row as long = 1, optional b_in_sample as boolean = false)
    ' internal function to compute knn_dist.
    dim output() as double, dist_a() as variant, dist_b() as variant, avg() as variant, stddev() as variant
    dim i as long, j as long, in_sample_row as long
    
    ' if b_in_sample is set, then b_row is part of a, and should be excluded while computing average and std dev for normalization.
    if b_in_sample then
        in_sample_row = b_row
    end if
    
    redim output(1 to n, 1 to 1)
    if normalize then
        redim avg(1 to d)
        redim stddev(1 to d)
        ' compute average and std dev of each column of a, excluding in_sample_row.
        ' note that no row will be excluded if b_in_sample is false since in_sample_row will then be 0.
        for j = 1 to d
            avg(j) = average(a, n, j, in_sample_row)
            stddev(j) = std_dev(a, n, j, in_sample_row)
        next j
        ' use avg and stddev computed from a to normalize both a and b. a, from row 1 to row n, and only b_row row of b.
        dist_a = donormalize(a, 1, n, d, avg, stddev)
        dist_b = donormalize(b, b_row, b_row, d, avg, stddev)
        
        ' use the normalized arrays dist_a and dist_b to compute distances from dist_b to each row of dist_a.
        for i = 1 to n
            output(i, 1) = knn_dist_row(dist_a, dist_b, i, b_row, scaling)
        next i
    else
        ' no normalization. directly compute distances.
        for i = 1 to n
            output(i, 1) = knn_dist_row(a, b, i, b_row, scaling)
        next i
    end if
    
    internal_knn_dist = output
end function

public function knn_dist(known_x as range, reference as range, optional normalize as boolean, _
optional not_implemented1 as long, optional dist_scale_vec as range = nothing, optional not_implemented2 as boolean, _
optional ref_in_sample as boolean = false)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: knn_dist(known_x as range, reference as range, optional normalize as boolean, _
optional dist_metric as long, optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean, _
optional ref_in_sample as boolean = false)
'compute distances between each row of known_x and reference. the number of columns of the ranges should match
'input:
'   known_x(required): x values.
'   reference(required): point from which distance is supposed to be calculated. must have the same number of columns
'       as known_x.
'   normalize: normalize so that each attribute has mean 0 and variance 1 before computing distance. the mean and
'       variance of known_x is used to scale reference.
'   dist_metric: unused
'   dist_scale_vec: a vector having a single row and columns equal to known_x. each entry indicates the multiple by which
'       to scale each attribute of known_x before computing distance.
'   zero_as_missing: unused
'   ref_in_sample: whether reference is a part of known_x. false by default. if true, it must be a part of the known_x range,
'       and the row containing reference is excluded when normalizing. only relevant if normalizing.
'output:
'   distances to known_x points from reference.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    on error goto err_knn_dist
    ' validate the input
    dim validation as string
    validation = validate_knn_functions(known_x, 0, , reference, dist_scale_vec)
    if not validation = "" then
        knn_dist = validation
        exit function
    end if
    
    dim ref_row as long
    ref_row = 1
    if ref_in_sample then
        ' reference range must be a part of known_x range.
        dim roverlap as range
        set roverlap = intersect(known_x, reference)
        if roverlap is nothing then
            knn_dist = "reference doesn't overlap with known_x. ref_in_sample should be false."
            exit function
        elseif roverlap.count <> reference.count then
            knn_dist = "reference isn't contained within known_x. ref_in_sample should be false."
            exit function
        end if
        ' henceforth, we shall use known_x(ref_row, .) to refer to reference(1, .).
        ref_row = reference.row - known_x.row + 1
    end if
   
    
    dim a() as variant, b() as variant
    dim n as long, d as long, j as long
    n = known_x.rows.count
    d = known_x.columns.count
    dim scaling() as variant
    scaling = get_dist_scaling(dist_scale_vec, d)
    
    ' get data from excel
    a = known_x.value2
    b = reference.value2
    
    ' check that the data array has all entries numeric or set to
    
    if not check_array_numeric(a) or not check_array_numeric(b) then
        knn_dist = "input not numeric"
        exit function
    end if
        
    
    if ref_in_sample and normalize then
        ' if ref_in_sample, we need to exclude ref_row when normalizing. we pass b:=a, b_row=ref_row and b_in_sample=true. to internal function.
        knn_dist = internal_knn_dist(a, a, scaling, n, d, true, ref_row, true)
    else
        ' in this case, if b is part of a, it will be included while normalizing. if it isn't a part of a, it won't be included.
        knn_dist = internal_knn_dist(a, b, scaling, n, d, normalize)
    end if
    
    exit function
err_knn_dist:
    knn_dist = "fatal error: " & err.description
end function

private function internal_knn_nearest(distances() as double, n as long)
    ' internal function to neighbours sorted according to the distances array.
    ' distances should be from a fixed reference whose nearest neighbours we desire.
    dim indices() as long, dists() as double
    redim indices(1 to n)
    redim dists(1 to n)
    dim j as long
    
    ' convert data type double array to data type variant array
    ' and initialize the indices array to indices(j) = j
    for j = 1 to n
        indices(j) = j
        dists(j) = distances(j, 1)
    next j
    
    ' sort them, the indices array keeps the new order, that is,
    ' indicies(1) is the closest, indices(2) is the second closest, etc.
    util.mergesort dists, indices
    
    internal_knn_nearest = indices
end function

private function internal_knn_in_movie(data() as variant, dist_data() as variant, _
n as long, d as long, k as integer, weighted_voting as boolean, dist_scale_vec() as variant)
    ' internal function to compute knn_in_movie. this function does not do any sanity checks and only uses arrays instead of ranges.
    ' expects normalized input. this function won't do any normalization.
    
    ' initialize an array with the outpus
    dim output()
    redim output(1 to n, 1 to d)
    
    ' step 2: compute predictions using the k-nn algorithm
    dim i as long, j as long, r as long
    dim indices() as long, distances() as double
    dim prediction as variant, num_obs as long, sum_weights as double
    dim count_same as long
    
    redim distances(1 to n)
    ' for each row of the data
    for i = 1 to n
        
        'dist_data is already normalized in the outer function if normalized is true. so send false here.
        distances = internal_knn_dist(dist_data, dist_data, dist_scale_vec, n, d, false, i)
        indices = internal_knn_nearest(distances, n)
        
        ' compute a prediction for each attribute by taking the k nearest neighbors
        for r = 1 to d
            prediction = 0
            num_obs = 0
            sum_weights = 0
            count_same = 0
            if weighted_voting then
                'points with zero distance get infinite weight in their votes.
                'if there are any such points, just report their average value.
                j = 1
                do while distances(indices(j), 1) = 0 and j <= n
                    doevents
                    if indices(j) <> i and not (isempty(data(indices(j), r))) then
                        count_same = count_same + 1
                    end if
                    j = j + 1
                loop
            end if
            if (not weighted_voting) or count_same = 0 then
                ' go over the neighbors in order of proximity
                for j = 1 to n
                    ' we need to go deeper in the list of neighbors if these are empty
                    ' also make sure not to consider i as a neighbor
                    if indices(j) <> i and not (isempty(data(indices(j), r))) then
                        if weighted_voting then
                            ' weight the votes by the distance
                            prediction = prediction + data(indices(j), r) / distances(indices(j), 1)
                            sum_weights = sum_weights + 1 / distances(indices(j), 1)
                        else
                            prediction = prediction + data(indices(j), r)
                            sum_weights = sum_weights + 1
                        end if
                        num_obs = num_obs + 1
                        ' break if we reached k observations
                        if num_obs = k then
                            exit for
                        end if
                    end if
                next j
            else
                ' here, weighted_voting=true and count_same > 0. we need to compute average of the votes by neighbours who are at distance 0.
                for j = 1 to n
                    if indices(j) <> i and not (isempty(data(indices(j), r))) then
                        prediction = prediction + data(indices(j), r)
                        sum_weights = sum_weights + 1
                        num_obs = num_obs + 1
                    end if
                    
                    if num_obs = count_same then
                        ' only count_same neighbours matter as they have infinite weight.
                        exit for
                    end if
                next j
            end if
            ' determine our prediction
            if num_obs = 0 then
                ' all neighbors are empty
                prediction = "n/a"
            else
                ' compute prediction by taking average.
                prediction = prediction / sum_weights
            end if
            ' record the result
            output(i, r) = prediction
        next r
    next i
    
    ' step 3: output to excel
    internal_knn_in_movie = output
end function


public function knn_in_movie(known_x as range, k as integer, optional normalize as boolean, _
optional not_implemented1 as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional not_implemented2 as boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: knn_in_movie(known_x as range, k as integer, optional normalize as boolean, _
optional dist_metric as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional zero_as_missing as boolean)
'compute predictions using the k-nn in-movie algorithm
'input:
'   known_x(required): x values. one or more cells may be empty.
'   k(required): k in the k-nearest neighbor algorithm. must be strictly larger than the number of rows in known_x.
'   normalize: normalize so that each attribute has mean 0 and variance 1 before computing distance. x is assumed to be
'       out of sample. that is, the mean and variance of known_x is used to scale it.
'   dist_metric: unused
'   weighted_voting: each nearest neighbourõs vote is weighted in inverse proportion to their distance from the
'       respective row. if a non-zero number of neighbours are at distance 0 from a row, they are given infinite weight
'       when computing prediction for that row. that is, the unweighted average of such neighbours, whatever their number
'       may be, is reported as the prediction.
'   dist_scale_vec: a vector having a single row and columns equal to known_x. each entry indicates the multiple by which
'       to scale each attribute of known_x before computing distance.
'   zero_as_missing: unused
'output:
'   predictions for attributes of xõs according to the knn movie algorithm run on known_x. when computing distances
'   between two rows, only coordinates which have non-empty entries for both rows are considered. if, for an attribute,
'   the number of neighbours having non-empty entries in that attribute is less than k, output is òn/aó.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    on error goto err_knn_in_movie
    dim data() as variant, dist_data() as variant, scaling() as variant
    dim n as long, d as long, i as long
    
    ' n is the number of rows/observations
    n = known_x.rows.count
    ' d is the number of cols/attributes
    d = known_x.columns.count
    
    dim validation as string
    validation = validate_knn_functions(known_x, k, , , dist_scale_vec)
    
    if not validation = "" then
        knn_in_movie = validation
        exit function
    end if
    
    scaling = get_dist_scaling(dist_scale_vec, d)
    ' step 1: get data from excel
    data = known_x.value2

    ' check that the data array has all entries numeric or set to
    if not check_array_numeric(data) then
        knn_in_movie = "input not numeric"
        exit function
    end if
        
    if normalize then
        ' normalize before passing the data to internal function, which does not do any normalization.
        dist_data = donormalize(data, 1, n, d)
        knn_in_movie = internal_knn_in_movie(data, dist_data, n, d, k, weighted_voting, scaling)
    else
        knn_in_movie = internal_knn_in_movie(data, data, n, d, k, weighted_voting, scaling)
    end if
    
    exit function
err_knn_in_movie:
    knn_in_movie = "fatal error: " & err.description
end function


private function internal_knn_out(known_y() as variant, known_x() as variant, new_x() as variant, _
scaling() as variant, n as long, d as long, k as integer, normalize as boolean, _
weighted_voting as boolean, optional exclude_row as boolean = false, optional new_x_row as long = 1)
' internal function to compute output using the k-nn algorithm. this function is also used when new_x is in sample.
' exclude_row should be set to true if new_x is in sample.

    dim j as long
    dim indices() as long, distances() as double, count_same as long
    dim prediction as variant, num_obs as long, sum_weights as double
    
    ' compute distances from new_x, normalizing and scaling if needed.
    distances = internal_knn_dist(known_x, new_x, scaling, n, d, normalize, new_x_row, exclude_row)
    ' use computed distances to find neighbours in sorted order.
    indices = internal_knn_nearest(distances, n)
    
    prediction = 0
    num_obs = 0
    count_same = 0
    if weighted_voting then
        'points with zero distance get infinite weight in their votes.
        'if there are any such points, just report their average value.
        j = 1
        do while distances(indices(j), 1) = 0 and j <= n
            doevents
            if (not exclude_row) or (indices(j) <> new_x_row) then
                count_same = count_same + 1
            end if
            j = j + 1
        loop
    end if
    
    if (not weighted_voting) or count_same = 0 then
    ' go over the neighbors in order of proximity
        for j = 1 to n
            if (not exclude_row) or (indices(j) <> new_x_row) then
                if weighted_voting then
                    ' weigh votes according to their distance from new_x
                    prediction = prediction + known_y(indices(j), 1) / distances(indices(j), 1)
                    sum_weights = sum_weights + 1 / distances(indices(j), 1)
                else
                    prediction = prediction + known_y(indices(j), 1)
                    sum_weights = sum_weights + 1
                end if
                num_obs = num_obs + 1
                ' break if we reached k observations
                if num_obs = k then
                    exit for
                end if
            end if
        next j
    else
        ' weighted_voting = true and count_same > 0. our prediction should be the average of all points that are
        ' at 0 distance from new_x, regardless of what k is.
        for j = 1 to n
            if (not exclude_row) or (indices(j) <> new_x_row) then
                prediction = prediction + known_y(indices(j), 1)
                sum_weights = sum_weights + 1
                num_obs = num_obs + 1
            end if
            
            if num_obs = count_same then
                ' we have counted all count_same points.
                exit for
            end if
        next j
    end if
    ' determine our prediction
    if num_obs = 0 then
        ' all neighbors are empty
        prediction = "n/a"
    else
        ' compute prediction by taking average.
        prediction = prediction / sum_weights
    end if
    ' record the result
    internal_knn_out = prediction
    
end function


private function internal_knn_in(known_y() as variant, known_x() as variant, scaling() as variant, _
n as long, d as long, k as integer, normalize as boolean, weighted_voting as boolean)
    ' initialize an array with the outpus
    dim output(), i as long
    redim output(1 to n, 1 to 1)
    dim dist_data() as variant
    
    ' for each row of the data
    for i = 1 to n
       ' use knn_out to compute prediction while setting new_x=known_x, exclude_row=true and new_x_row=i
       output(i, 1) = internal_knn_out(known_y, known_x, known_x, scaling, n, d, k, normalize, weighted_voting, true, i)
       if output(i, 1) = "n/a" then
            output(i, 1) = "not enough neighbours"
       end if
    next i
    
    internal_knn_in = output
end function



public function knn_in(known_y as range, known_x as range, k as integer, optional normalize as boolean, _
optional not_implemented1 as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional not_implemented2 as boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature:knn_in(known_y as range, known_x as range, k as integer, optional normalize as boolean, _
optional dist_metric as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional zero_as_missing as boolean)
'function to compute y values for in-sample observations using the k-nn algorithm.
'input:
'   known_y (required): y values
'   known_x (required): x values corresponding to known_y. must have the same number of rows as known_y.
'   k (required): k in the k-nearest neighbor algorithm. must be less than or equal to the number of rows in known_x.
'   normalize: normalize so that each attribute has mean 0 and variance 1 before computing distance. for each row in known_x,
'       the mean and variance of known_x is computed excluding that row and used to scale the entire known_x matrix
'       (including the excluded row). known_y is not scaled.
'   dist_metric: unused
'   weighted_voting: each nearest neighbourõs vote is weighted in inverse proportion to their distance from the respective row.
'       if a non-zero number of neighbours are at distance 0 from a row, they are given infinite weight when computing
'       prediction for that row. that is, the unweighted average of such neighbours, whatever their number may be, is reported
'       as the prediction.
'   dist_scale_vec: a vector having a single row and columns equal to known_x. each entry indicates the multiple by which
'       to scale each attribute of known_x before computing distance.
'   zero_as_missing: unused
'output:
'   y predictions for known_xõs according to the knn algorithm.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    on error goto err_knn_in

    dim data() as variant, data_y() as variant, scaling() as variant
    dim n as long, d as long
    
    ' n is the number of rows/observations
    n = known_x.rows.count
    ' d is the number of cols/attributes
    d = known_x.columns.count
    
    dim validation as string
    validation = validate_knn_functions(known_x, k, known_y, , dist_scale_vec)
    if not validation = "" then
        knn_in = validation
        exit function
    end if
    scaling = get_dist_scaling(dist_scale_vec, d)
    
    ' step 1: get data from excel
    data = known_x.value2
    data_y = known_y.value2
    
   ' check that the data array has all entries numeric or set to
    
    if not check_array_numeric(data) or not check_array_numeric(data_y) then
        knn_in = "input not numeric"
        exit function
    end if
        
    
    knn_in = internal_knn_in(data_y, data, scaling, n, d, k, normalize, weighted_voting)
    
    exit function
err_knn_in:
    knn_in = "fatal error: " & err.description

end function


public function knn_out(x as range, known_y as range, known_x as range, k as integer, optional normalize as boolean, _
optional not_implemented1 as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional not_implemented2 as boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: public function knn_out(x as range, known_y as range, known_x as range, k as integer, optional normalize as boolean, _
optional dist_metric as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional zero_as_missing as boolean)
'input:
'   x(required): new x values.
'   known_y(required): y values.
'   known_x(required): x values corresponding to known_y. must have same number of columns as x and same number of rows as known_y
'   k(required): k in the k-nearest neighbor algorithm. must be larger than or equal to the number of rows in known_x.
'   normalize: normalize so that each attribute has mean 0 and variance 1 before computing distance. x is assumed to be out of sample.
'       that is, the mean and variance of known_x is used to scale it.
'   dist_metric: unused
'   weighted_voting: each nearest neighbourõs vote is weighted in inverse proportion to their distance from the respective row.
'       if a non-zero number of neighbours are at distance 0 from a row, they are given infinite weight when computing prediction for
'       that row. that is, the unweighted average of such neighbours, whatever their number may be, is reported as the prediction.
'   dist_scale_vec: a vector having a single row and columns equal to known_x. each entry indicates the multiple by which to scale
'       each attribute of known_x before computing distance.
'   zero_as_missing: unused
'output:
'   y predictions for xõs according to the knn algorithm run on known_x and known_y
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
on error goto err_knn_out

    dim data() as variant, data_y() as variant, unknown_x() as variant, scaling() as variant
    dim n as long, d as long, m as long
    
    ' n is the number of rows/observations
    n = known_x.rows.count
    ' m is the number of rows of new_x
    m = x.rows.count
    ' d is the number of cols/attributes
    d = known_x.columns.count
    
    if d <> x.columns.count then
        knn_out = "known_x and x must have the same number of columns."
        exit function
    end if
    
    if application.worksheetfunction.count(x) <> x.rows.count * x.columns.count then
        knn_out = "all values in x must be numeric"
        exit function
    end if
    
    dim validation as string
    validation = validate_knn_functions(known_x, k + 1, known_y, , dist_scale_vec)
    if not validation = "" then
        knn_out = validation
        exit function
    end if
    scaling = get_dist_scaling(dist_scale_vec, d)
    
    ' step 1: get data from excel
    data = known_x.value2
    data_y = known_y.value2
    unknown_x = x.value2
    
    ' step 2: check that the data array has all entries numeric or set to
    
    if not check_array_numeric(data) or not check_array_numeric(data_y) or not check_array_numeric(unknown_x) then
        knn_out = "input not numeric"
        exit function
    end if
        
    
    dim output() as variant, i as long
    redim output(1 to m, 1 to 1)
    
    for i = 1 to m
        ' use internal knn_out separately on each row of the data to find the y output.
        ' we set exclude_row:=false and new_x_row:=i
        output(i, 1) = internal_knn_out(data_y, data, unknown_x, scaling, n, d, k, normalize, _
        weighted_voting, false, i)
        if output(i, 1) = "n/a" then
            output(i, 1) = "not enough neighbours"
        end if
    next i
    
    knn_out = output
    
exit function
err_knn_out:
    knn_out = "fatal error: " & err.description

end function

private function rmse(x() as variant, y() as variant, n as long, optional d as long = 1)
' internal function to compute rmse between x and y.
dim output as double, i as long, j as long, count as long
output = 0
count = 0

for j = 1 to d
    for i = 1 to n
        if (not isempty(x(i, j))) and (not isempty(y(i, j))) then
            ' only include cells if both entries are non-empty.
            output = output + ((x(i, j) - y(i, j)) ^ 2)
            count = count + 1
        end if
    next i
next j

output = output / count
rmse = sqr(output)

end function

public function ba_rmse(predicted as range, actual as range, optional not_implemented as range)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: ba_rmse(data as range, truth as range, row_weight as range)
'input:
'   data/predicted: predicted values by our model
'   truth/actual: true values
'   row_weight: unused
'output:
'   rmse between data/predicted and truth/actual
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
on error goto err_ba_rmse

dim n as long
n = predicted.rows.count
' check the input sanity
if n <> actual.rows.count then
    ba_rmse = "predicted and actual must have the same number of rows."
    exit function
end if

if actual.columns.count <> 1 or predicted.columns.count <> 1 then
    ba_rmse = "predicted and actual both must have exactly one column each."
    exit function
end if

' extract input to arrays.
dim x() as variant, y() as variant
x = predicted.value2
y = actual.value2

' check that the data array has all entries numeric or set to
 
 if not check_array_numeric(x) or not check_array_numeric(y) then
     ba_rmse = "input not numeric"
     exit function
 end if
     
ba_rmse = rmse(x, y, n)

exit function
err_ba_rmse:
    ba_rmse = "fatal error: " & err.description
end function



public function knn_in_rmse(known_y as range, known_x as range, k as integer, optional normalize as boolean, _
optional not_implemented1 as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional not_implemented2 as boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: knn_in_rmse(known_y as range, known_x as range, k as integer, optional normalize as boolean, _
optional dist_metric as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional zero_as_missing as boolean)
'input:
'   known_y (required): y values
'   known_x(required): x values corresponding to known_y. must have the same number of rows as known_y.
'   k(required): k in the k-nearest neighbor algorithm. must be less than or equal to the number of rows in known_x.
'   normalize: normalize so that each attribute has mean 0 and variance 1 before computing distance. same behaviour as in knn_in
'   dist_metric: unused
'   weighted_voting: each nearest neighbourõs vote is weighted in inverse proportion to their distance from the respective row.
'       if a non-zero number of neighbours are at distance 0 from a row, they are given infinite weight when computing prediction
'       for that row. that is, the unweighted average of such neighbours, whatever their number may be, is reported as the prediction.
'   dist_scale_vec: a vector having a single row and columns equal to known_x. each entry indicates the multiple by which to
'       scale each attribute of known_x before computing distance.
'   zero_as_missing: unused
'output:
'   rmse between yõs predicted by the knn algorithm and known_yõs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

on error goto err_knn_in_rmse

    dim data() as variant, data_y() as variant, knn_output() as variant, scaling() as variant
    dim n as long, d as long
    
    ' n is the number of rows/observations
    n = known_x.rows.count
    ' d is the number of cols/attributes
    d = known_x.columns.count
    
    dim validation as string
    validation = validate_knn_functions(known_x, k, known_y, , dist_scale_vec)
    if not validation = "" then
        knn_in_rmse = validation
        exit function
    end if
    scaling = get_dist_scaling(dist_scale_vec, d)
    
    ' step 1: get data from excel
    data = known_x.value2
    data_y = known_y.value2
    
   ' check that the data array has all entries numeric or set to
    
    if not check_array_numeric(data) or not check_array_numeric(data_y) then
        knn_in_rmse = "input not numeric"
        exit function
    end if
        
    
    ' compute predictions with k-nn algorithm
    knn_output = internal_knn_in(data_y, data, scaling, n, d, k, normalize, weighted_voting)
    ' use the predicted knn_output with data_y to compute rmse.
    knn_in_rmse = rmse(knn_output, data_y, n)
    
    exit function
err_knn_in_rmse:
    knn_in_rmse = "fatal error: " & err.description

end function

public function knn_in_movie_rmse(known_x as range, k as integer, optional normalize as boolean, _
optional not_implemented1 as long, optional weighted_voting as boolean, _
optional dist_scale_vec as range = nothing, optional not_implemented2 as boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: public function knn_in_movie_rmse(known_x as range, k as integer, optional normalize as boolean, _
optional dist_metric as long, optional weighted_voting as boolean, _
optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean)
'input:
'   known_x(required): x values.
'   k(required): k in the k-nearest neighbor algorithm. must be less than or equal to the number of rows in known_x.
'   normalize: normalize so that each attribute has mean 0 and variance 1 before computing distance. same behaviour as in knn_in_movie
'   dist_metric: unused
'   weighted_voting: each nearest neighbourõs vote is weighted in inverse proportion to their distance from the respective row.
'       if a non-zero number of neighbours are at distance 0 from a row, they are given infinite weight when computing prediction
'       for that row. that is, the unweighted average of such neighbours, whatever their number may be, is reported as the prediction.
'   dist_scale_vec: a vector having a single row and columns equal to known_x. each entry indicates the multiple by which to
'       scale each attribute of known_x before computing distance.
'   zero_as_missing: unused
'output:
'   rmse between yõs predicted by the knn algorithm and known_yõs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

on error goto err_knn_in_movie_rmse

    dim data() as variant, dist_data() as variant, knn_output() as variant, scaling() as variant
    dim n as long, d as long
    
    ' n is the number of rows/observations
    n = known_x.rows.count
    ' d is the number of cols/attributes
    d = known_x.columns.count
    
    dim validation as string
    validation = validate_knn_functions(known_x, k, , , dist_scale_vec)
    if not validation = "" then
        knn_in_movie_rmse = validation
        exit function
    end if
    scaling = get_dist_scaling(dist_scale_vec, d)
    
    ' step 1: get data from excel
    data = known_x.value2
    
   ' check that the data array has all entries numeric or set to
    
    if not check_array_numeric(data) then
        knn_in_movie_rmse = "input not numeric"
        exit function
    end if
        
    if normalize then
        ' just like in knn_in_movie, conduct normalization here before calling internal_knn_in_movie.
        dist_data = donormalize(data, 1, n, d)
        knn_output = internal_knn_in_movie(data, dist_data, n, d, k, weighted_voting, scaling)
    else
        knn_output = internal_knn_in_movie(data, data, n, d, k, weighted_voting, scaling)
    end if
    ' use the predicted knn_output to compute rmse against data.
    knn_in_movie_rmse = rmse(knn_output, data, n, d)
    
    exit function
err_knn_in_movie_rmse:
    knn_in_movie_rmse = "fatal error: " & err.description

end function

public function knn_nearest(known_x as range, reference as range, optional normalize as boolean, _
optional not_implemented1 as long, optional dist_scale_vec as range = nothing, optional not_implemented2 as boolean, _
optional ref_in_sample as boolean = false)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' original signature: knn_nearest(known_x as range, reference as range, optional normalize as boolean, _
optional dist_metric as long, optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean, _
optional ref_in_sample as boolean = false)
'input:
'   known_x(required): x values.
'   reference(required): point from which distance is supposed to be calculated. must have the same number of columns as known_x.
'   normalize: normalize so that each attribute has mean 0 and variance 1 before computing distance. the mean and variance
'       of known_x is used to scale reference.
'   dist_metric: unused
'   dist_scale_vec: a vector having a single row and columns equal to known_x. each entry indicates the multiple by which
'       to scale each attribute of known_x before computing distance.
'   zero_as_missing: unused
'   ref_in_sample: whether reference is a part of known_x. false by default. if true, it must be a part of the known_x
'   range, and the row containing reference is excluded when normalizing. only relevant if normalizing.
'output:
'   indices to points in known_x array in order of nearest to farthest from reference.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    on error goto err_knn_nearest
    ' validate the input
    dim validation as string
    validation = validate_knn_functions(known_x, 0, , reference, dist_scale_vec)
    if not validation = "" then
        knn_nearest = validation
        exit function
    end if
    dim ref_row as long
    ref_row = 1
    if ref_in_sample then
        ' reference range must be a part of known_x range.
        dim roverlap as range
        set roverlap = intersect(known_x, reference)
        if roverlap is nothing then
            knn_nearest = "reference doesn't overlap with known_x. ref_in_sample should be false."
            exit function
        elseif roverlap.count <> reference.count then
            knn_nearest = "reference isn't contained within known_x. ref_in_sample should be false."
            exit function
        end if
        ' henceforth, we shall use known_x(ref_row, .) to refer to reference(1, .).
        ref_row = reference.row - known_x.row + 1
    end if
    
    dim a() as variant, b() as variant, output() as variant, first_out() as long, distances() as double
    dim n as long, i as long, d as long
    n = known_x.rows.count
    d = known_x.columns.count
    dim scaling() as variant
    
    scaling = get_dist_scaling(dist_scale_vec, d)
    redim output(1 to n, 1 to 1)
    
    ' get data from excel
    a = known_x.value2
    b = reference.value2
    
   ' check that the data array has all entries numeric or set to
    
    if not check_array_numeric(a) or not check_array_numeric(b) then
        knn_nearest = "input not numeric"
        exit function
    end if
        
    
    if ref_in_sample and normalize then
        ' if ref_in_sample, we need to exclude ref_row when normalizing. we pass b:=a, b_row=ref_row and b_in_sample=true. to internal function.
        distances = internal_knn_dist(a, a, scaling, n, d, true, ref_row, true)
    else
        ' in this case, if b is part of a, it will be included while normalizing. if it isn't a part of a, it won't be included.
        distances = internal_knn_dist(a, b, scaling, n, d, normalize)
    end if
    ' use distances computed above to compute knn_nearest.
    first_out = internal_knn_nearest(distances, n)
    
    ' convert to 2-d array.
    for i = 1 to n
        output(i, 1) = first_out(i)
    next i
    
    knn_nearest = output
    
    exit function
err_knn_nearest:
    knn_nearest = "fatal error: " & err.description
end function

private sub registerba_rmse()
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    'ba_rmse(predicted as range, actual as range, not_implemented as range)
    funccat = "cbs ba add-in functions"
    dim argdesc(1 to 3) as string
    argdesc(1) = "column containing predictions"
    argdesc(2) = "column containing actual outcomes"
    argdesc(3) = "included for compatibility. value not used."
    funcdesc = "function to compute the rmse between predicted and actual values."
    funcname = "ba_rmse"
    
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
end sub


public sub registerall()

    registerba_rmse
    
    dim funcname as string
    dim funcdesc as string
    dim funccat as variant
    
    dim x as string, known_y as string, known_x as string, k as string, normalize as string, dist_metric as string
    dim weighted_voting as string, dist_scale_vec as string, zero_as_missing as string, ref_in_sample as string
    
    'choose the built-in function category (it will no longer appear in udf category).
    'for example, 15 is the engineering category, 4 is the statistical category etc.
    funccat = "cbs ba add-in functions"
    
    x = "new x values"
    known_y = "known y values"
    known_x = "known x values"
    k = "k in the knn algorithm"
    normalize = "if true, normalize each attribute to mean 0 and variance 1. false by default"
    dist_metric = "for compatibility. value unused"
    weighted_voting = "each neighbour's vote is weighted in inverse proportion to its distance if true. false by default."
    dist_scale_vec = "factors to scale each attribute of known_x before computing distance."
    zero_as_missing = "for compatibility. value unused"
    ref_in_sample = "whether reference is a part of known_x. false by default. if true, the row containing reference is excluded when normalizing."
    
     'knn_out(x as range, known_y as range, known_x as range, k as long, optional normalize as boolean, optional dist_metric as long, _
optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean)
    dim argdesc() as string
    redim argdesc(0 to 8)
 'here we add the description for the function's arguments.
    argdesc(0) = x
    argdesc(1) = known_y
    argdesc(2) = known_x
    argdesc(3) = k
    argdesc(4) = normalize
    argdesc(5) = dist_metric
    argdesc(6) = weighted_voting
    argdesc(7) = dist_scale_vec
    argdesc(8) = zero_as_missing
    
    funcname = "knn_out"
    funcdesc = "array function to predict y values for unknown x values based on existing known_y values and known_x values using the knn algorithm." & _
    vbnewline & vbnewline & "output:" & vbnewline & "y predictions for x."
    
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    'function signature:  knn_in(known_y as range, known_x as range, k as long, optional normalize as boolean, optional dist_metric as long, _
optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean)

    'depending on the function arguments define the necessary variables on the array.
    redim argdesc(1 to 8)
    argdesc(1) = known_y
    argdesc(2) = known_x
    argdesc(3) = k
    argdesc(4) = normalize
    argdesc(5) = dist_metric
    argdesc(6) = weighted_voting
    argdesc(7) = dist_scale_vec
    argdesc(8) = zero_as_missing
    
    funcname = "knn_in"
    funcdesc = "array function to predict y values for known_x values based on existing known_y values using the knn algorithm." & _
    vbnewline & vbnewline & "output:" & vbnewline & "y predictions."
    
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    'knn_in_rmse(known_y as range, known_x as range, k as long, optional normalize as boolean, optional dist_metric as long, _
optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean)

    funcname = "knn_in_rmse"
    funcdesc = "function to compute rmse for output y values computed using knn algorithm"
    
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
   'knn_in_movie(known_x as range, k as integer, optional normalize as boolean, _
optional dist_metric as long, optional weighted_voting as boolean, optional dist_scale_vec as range = nothing, _
optional zero_as_missing as boolean)
    redim argdesc(2 to 8)
    argdesc(2) = known_x
    argdesc(3) = k
    argdesc(4) = normalize
    argdesc(5) = dist_metric
    argdesc(6) = weighted_voting
    argdesc(7) = dist_scale_vec
    argdesc(8) = zero_as_missing
    funcname = "knn_in_movie"
    funcdesc = "function to compute predictions for attributes using the knn_in_movie algorithm"
    
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    'knn_in_movie_rmse(known_x as range, k as integer, optional normalize as boolean, _
optional dist_metric as long, optional weighted_voting as boolean, _
optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean)
    funcname = "knn_in_movie_rmse"
    funcdesc = "function to compute rmse for in_movie ratings computed using knn_in_movie algorithm"
    
    util.functiondescription funcname, funcdesc, funccat, argdesc
    'knn_dist(known_x as range, reference as range, optional normalize as boolean, optional dist_metric as long, _
optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean)
    redim argdesc(1 to 7)
    argdesc(1) = known_x
    argdesc(2) = "point from which to compute distance"
    argdesc(3) = normalize
    argdesc(4) = dist_metric
    argdesc(5) = dist_scale_vec
    argdesc(6) = zero_as_missing
    argdesc(7) = ref_in_sample
    
    funcname = "knn_dist"
    funcdesc = "array function to compute distance from reference to each point in known_x" & vbnewline & _
    "output:" & vbnewline & "distance from each point to reference"
    
    util.functiondescription funcname, funcdesc, funccat, argdesc

    'knn_nearest(known_x as range, reference as range, optional normalize as boolean, optional dist_metric as long, _
optional dist_scale_vec as range = nothing, optional zero_as_missing as boolean)
    
    funcname = "knn_nearest"
    funcdesc = "array function to find points sorted according to their distance from reference" & vbnewline & _
    "output:" & vbnewline & "indices of points from known_x array in order of their distance from reference from nearest to farthest"
    
    util.functiondescription funcname, funcdesc, funccat, argdesc
    
    
end sub
























