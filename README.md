# Parsing elasticity data from raw text files
We have been tasked with parsing `;` separated files in order to retrieve information
generated by several elasticity experiments studying the properties of plastics suitable
for 3D-Printing.

We'll just parse the different raw files to then generate Excel (i.e. `*.xlsx`) file,
thus making handling them much easier.

We'll also use [`jq(1)`](https://stedolan.github.io/jq/manual/) together with
[`gnuplot(1)`](http://www.gnuplotting.org/manpage-gnuplot-4-6/) to generate any
graphs that might be requested.

## Useful commands
Both `jq` and `gnuplot` are very powerful tools. The thing is, they are also a bit complex
to use... This section contains several commands we can rely on for common tasks.

### Converting JSON data to CSV
We can leverage `jq` filters to this end:

	# Option breakdown:
		# -r: We want raw output instead of compliant JSON elements.
		# <filter>: The filter will access the root document member (`.data`) and then
			# we'll choose the series both for elgontation and tension. We'll end
			# up by transposing the results so that we get columns instead of just
			# two lines and then convert that to CSC with `@csv`.
		# Prob14.json: The JSON file we are processing.
		# Shell redirection: We'll save the result in `tensionElongation.csv`.
	jq -r '.data | [.elongationN, .tensionMPa] | transpose[] | @csv' Prob14.json > tensionElongation.csv

We can use variations of the above for a ton of operations: `jq` is extremely powerful!

### Plotting CSV files
We can use `gnuplot` to plot the two columns we exported just now:

	# Let's tell gnuplot to break up the lines on `,`
	set datafaile separator ","

	# We'll set the axes labels
	set xlabel "Tensión [MPa]"
	set ylabel "Elongación"

	# We can now plot the data:
		# "tensionElongation.csv": The file containing the data.
		# using 1:2: We'll use the first column on the X axis and the
			# second one on the Y axis. Note we can also use the
			# pseudocolumn `0` to plot a single data series too.
		# with lines: We'll use solid lines in the plot.
		# t "Prob14": The name to show on the graph's legend.
	plot "tensionElongation.csv" using 1:2 with lines t "Prob14"

Like before, variations of the above can become quite useful too!
