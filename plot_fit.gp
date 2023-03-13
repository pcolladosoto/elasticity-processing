# Data generation:
# jq -r '.Prob18[0][1].data | [.elongationN, .tensionMPa] | transpose[] | @csv' processed_data.json > foo.csv

set datafile separator ","
set ylabel "Tensión"
set xlabel "Elongación"
plot "prob18.csv" using 1:2 with points t "Prob18-5", 1179,16314071 * x t "Fit"
