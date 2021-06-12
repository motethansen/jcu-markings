# jcu-markings

The script clean-marksheet.py can be used to remove comments from blackboard marksheets.
With that it becomes easier to extract the marks into the JCU exam templates.

to use, do a:

```shell
$ pip install -r requirements.txt

python clean-marksheet.py -i yourinputfile.xlsx -o youroutputfilename.xlsx
```

the output file is optional as the script will generate a standard output file.