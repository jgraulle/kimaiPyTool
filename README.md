Purpose
=======

Console tool to have some functionality on Kimai Time tracking web application like :
- copy time event into google calendar
- export a .tsv file per customer by date

Setup
=====
For now is only compatible on Linux.

Developement environement
-------------------------

    python3 -m venv venv
    source venv/bin/activate
    python3 -m pip install -r requirements.txt

Runtime environement
--------------------

For google calendar you have to generate a google API OAuth 2.0 token file from
https://developers.google.com/google-apps/calendar/quickstart/python#prerequisites.

Run the setup.sh script to install in `~/.local/share/kimaiPyTool/` and the run command in
`~/.local/bin/`.

If you have never done run `activate-global-python-argcomplete [--user]` to have the autocompletion.
