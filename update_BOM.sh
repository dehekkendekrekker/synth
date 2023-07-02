dirs=$(ls -d */)
for dir in $dirs
do
	fullpath=$(readlink -f ./$dir);
	dir=$(echo $dir | sed 's/.$//')
	python3 "/usr/share/kicad/plugins/bom_csv_grouped_by_value.py" "$fullpath/$dir.xml" "$fullpath/$dir.csv"
done
