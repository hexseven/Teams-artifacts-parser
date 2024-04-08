# About project
Program `teamsParser.py` allow to analyze Microsoft Teams artifacts in human-readable format. 
For input receives **.json** file of `https_teams.live.com_0.indexeddb.leveldb` 
from ``C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Teams``. To create json file please use `ms_teams_parser/ms_teams_parser.exe` by Alexander 'lxndrblz ' Bilz(licence attached).
For output program creates simply **.txt** file. 
</br>Possible to extract:
- Information's about users present in Teams.
- Information's about performed calls.
- Information's about messages included **content**.
- Information's about meetings.
- Reactions on messages.

Based on my research I selected some fields particular fields from particular groups of artifacts. 
Because it is open-source program on MIT licence feel free to modify code for your requirements.
Please only attach my LICENCE statement. 
# Usage
```
usage: teamsParser.py [-h] [-f FILE] [-o OUTPUT] [-u] [-c] [-m] [-t] [-r]

options:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  Input file for analyze.
  -o OUTPUT, --output OUTPUT
                        Name of output file. By default is "output.txt"
  -u, --users           Get users present in Teams.
  -c, --calls           Get information about calls.
  -m, --messages        Get message content.
  -t, --meetings        et meeting information.
  -r, --reactions       Get reactions data.
```

# TODO
I build this project after hours for training purposed. I was curiosity what can be extracted form Microsoft Teams.
Now project is in initial version that's allows to perform investigations.
</br>Probable future improvements:
- Prettier output 
- Look for more artifacts in **.json** input file