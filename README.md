# TA-xlsx-to-splunk
TA for XLSX files

# Usage
## Create a path file to wrap the script execution command

Add three arguments:
argv[1] = string : <path to file>
example: /path/to/file

argv[2] = string : <wildcard or filename>
example 1: \*.xlsx
example 2: reports*.xlsx

argv[3] = integer : <starting row in file>
example: the fields are expected to be on row 7, then value will be (7)

Full Example: 
filename: input1.path
content: `$SPLUNK_HOME/etc/apps/TA-xlsx-to-splunk/bin/xlsx2splunk.py /path/to/file/ '*.xlsx' 1`

## Configure inputs.conf
In example above, you can configure the inputs.conf using this example
```
[script://$SPLUNK_HOME/etc/apps/TA-xlsx-to-splunk/bin/input1.path]
disabled = false
index = main
interval = 60
sourcetype = xlsx2splunk:parsed
```

You can configure multiple inputs by creating different path file to pass the required arguments if files belong to different directories or filenames.