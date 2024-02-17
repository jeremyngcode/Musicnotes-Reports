Musicnotes Reports
==================

Intro
-----
I get my sheetmusic revenue reports quarterly from [Musicnotes](https://www.musicnotes.com/sheet-music/artist/jeremy-ng), and just about a week ago, I found myself tediously copying over data yet again into my master Excel file. So just as with my [Soundrop-Reports mini project](https://github.com/jeremyngcode/Soundrop-Reports), I decided to write a similar script to automate this process as well.

The Process
-----------
An Excel file template is created (sample available in template folder). This is the file that gets written on by the script, subsequent runs will simply overwrite.

The initial configuration requires entering the paths for `master_xl_file` and `xl_file` in [settings.py](settings.py):
- `master_xl_file`: The master Excel file that holds all my Musicnotes revenue data from the beginning, and is where the script will retrieve my sheetmusic titles.
- `xl_file`: The script writes to this template Excel file, which I will then copy over to my `master_xl_file` with one copy-paste.

The regular process every quarter then looks like this:
1. Save the given Musicnotes Excel file and rename it.
2. Change `musicnotes_xl_file` variable in settings.py to the path of the newly saved Musicnotes Excel file.
3. Change `year` and `quarter` variables in settings.py to the reporting year and quarter respectively. (eg. '2023' and 'Q4').
4. Run [Update_Sheetmusic_Sales_Report.py](Update_Sheetmusic_Sales_Report.py).

The script will turn this (`xl_file`)...

![xl-template](https://github.com/jeremyngcode/Musicnotes-Reports/assets/156220343/84569b3d-62bf-4e0f-a562-3d410b5212c9)

into this...

![xl-template-filled](https://github.com/jeremyngcode/Musicnotes-Reports/assets/156220343/76f08822-1bee-419b-8901-69641ec0b5fc)

Blank cells mean I didn't have any sales for the corresponding title.

5. Make one copy-paste from here to my `master_xl_file`.
6. Save the file and that's it. 😃

Extra Thoughts
--------------
- Same as before, I prefer having my script write to a template file first instead of the master Excel file directly because I didn't want to risk having my code mess up something in the master. It would only save me one copy-paste action anyway.

- Maybe it's a little overkill, I only needed to do this 4 times a year after all. But really, it's mostly just another excuse for me to write more code. And more practice. 😁

#### Notable libraries used for this project:
- [openpyxl](https://pypi.org/project/openpyxl/)
