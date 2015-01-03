#notes.md

##Processing time for creating lists of cells
All the Cost Center sheets in *testKojo_Atlantic Engineering CC Budgets for 2012_2017_JunQtr Update_9 16 13_Aug Actuals updated.xlsx*, `sheets = wb.get_sheet_names()[1:]`, wouldn't process at once. I get a `MemoryError: `. So, I tried smaller slices.

Processing sheets 1-7 `sheets = wb.get_sheet_names()[1:8]` of  takes ~36 seconds with memory cleared.
<pre>
    Processing time: 36.1719999313 seconds
    FC has 7 Cost Centers.
    57100 52
    52P04 48
    57165 49
    52P02 47
    55B99 61
    52P07 49
    57010 55
</pre>

Sheet #8 `sheets = wb.get_sheet_names()[8:9]`, Cost Center 57166, ALONE takes 17 seconds
<pre>
    Processing time: 17.245000124 seconds
    FC has 1 Cost Centers.
    57166 50
</pre>

The final sheets also take ~34 seconds, with memory cleared:
<pre>
    sheets = wb.get_sheet_names()[9:]
    Processing time: 34.1930000782 seconds
    FC has 9 Cost Centers.
    57171 51
    57999 47
    61623 46
    57250 50
    57620 50
    57240 49
    63002 50
    57260 48
    57619 48
</pre>


##duplicate_edit.py
Initial run results of the code to count rows & columns in each sheet shows why there were hangups with the *exception.py* code. Some of the sheets have the maximum number of columns. But there's no DAT there, it's all blank.
<pre>
    Time to load source worksheet 4.84899997711
    Sheet 52P02 has 202 rows and 46 columns
    Sheet 52P04 has 102 rows and 46 columns
    Sheet 52P07 has 211 rows and 46 columns
    Sheet 55B99 has 216 rows and 16384 columns
    Sheet 57010 has 210 rows and 46 columns
    Sheet 57100 has 264 rows and 46 columns
    Sheet 57165 has 103 rows and 46 columns
    Sheet 57166 has 104 rows and 16384 columns
    Sheet 57171 has 214 rows and 46 columns
    Sheet 57240 has 120 rows and 46 columns
    Sheet 57250 has 129 rows and 46 columns
    Sheet 57260 has 110 rows and 46 columns
    Sheet 57619 has 102 rows and 46 columns
    Sheet 57620 has 105 rows and 46 columns
    Sheet 57999 has 99 rows and 38 columns
    Sheet 61623 has 100 rows and 16384 columns
    Sheet 63002 has 105 rows and 16384 columns
</pre>

I cleared all the cells from Columns AU to the end. New results show the difference. Note Cost Center 57166. I also added a timer.

<pre>
    Time to load source worksheet 4.49500012398
    Sheet 52P02 has 202 rows and 46 columns
    Sheet 52P04 has 102 rows and 46 columns
    Sheet 52P07 has 211 rows and 46 columns
    Sheet 55B99 has 216 rows and 16384 columns
    Sheet 57010 has 210 rows and 46 columns
    Sheet 57100 has 264 rows and 46 columns
    Sheet 57165 has 103 rows and 46 columns
    Sheet 57166 has 104 rows and 46 columns
    Sheet 57171 has 214 rows and 46 columns
    Sheet 57240 has 120 rows and 46 columns
    Sheet 57250 has 129 rows and 46 columns
    Sheet 57260 has 110 rows and 46 columns
    Sheet 57619 has 102 rows and 46 columns
    Sheet 57620 has 105 rows and 46 columns
    Sheet 57999 has 99 rows and 38 columns
    Sheet 61623 has 100 rows and 16384 columns
    Sheet 63002 has 105 rows and 16384 columns
    Time to count rows & columns in all sheets: 0.105999946594 seconds.
</pre>

Now, with ALL the Max Column sheets "cleaned":
<pre>
    Time to load source worksheet 4.09900021553 seconds.
    Sheet 52P02 has 202 rows and 46 columns
    Sheet 52P04 has 102 rows and 46 columns
    Sheet 52P07 has 211 rows and 46 columns
    Sheet 55B99 has 216 rows and 46 columns
    Sheet 57010 has 210 rows and 46 columns
    Sheet 57100 has 264 rows and 46 columns
    Sheet 57165 has 103 rows and 46 columns
    Sheet 57166 has 104 rows and 46 columns
    Sheet 57171 has 214 rows and 46 columns
    Sheet 57240 has 120 rows and 46 columns
    Sheet 57250 has 129 rows and 46 columns
    Sheet 57260 has 110 rows and 46 columns
    Sheet 57619 has 102 rows and 46 columns
    Sheet 57620 has 105 rows and 46 columns
    Sheet 57999 has 99 rows and 38 columns
    Sheet 61623 has 100 rows and 41 columns
    Sheet 63002 has 105 rows and 46 columns
    Time to count rows & columns in all sheets: 0.0019998550415 seconds.
</pre>
0.105999946594 / 0.0019998550415 = 53.00 times faster. This should ALSO make *exceptions.py* run faster.