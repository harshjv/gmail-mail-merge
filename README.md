# ⚡️ Mail Merge using Spreadsheet and Gmail

> This script will use the ***most recent*** draft in Gmail as a ***template***.


## Features

* Handles duplicate entries (by email address)
* Works with HTML and plain text emails


## Usage

1. Open Spreadsheet
2. Go to `Tools > Script editor...`
  1. Paste contents of `mailmerge.gs` there
  2. Save it
3. After saving script, you will notice a new menu **Gmail Mail Merge** in the Spreadsheet
4. From Spreadsheet, Go to `Gmail Mail Merge > Start mail merge utility`


### Example Spreadsheet

| First name | Last name | Email       |
| ---------- | --------- | ----------- |
| Abc        | Xyz       | abc@def.com |
| Def        | Pqr       | def@pqr.com |


### Template

```
To: Email (here, put the title of email column)
Body:
Hello {{First name}} {{Last name}}...
```


### Email Outcome

```
To: abc@def.com
Body:
Hello Abc Xyz...
```


## License

MIT
