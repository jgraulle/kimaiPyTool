#!/bin/bash

appname=kimaiPyTool

mkdir ~/.local/share/${appname}
cp ${appname}.py ~/.local/share/${appname}/
python3 -m venv ~/.local/share/${appname}/venv
echo -e \#\!/bin/bash"\n\nsource ~/.local/share/${appname}/venv/bin/activate\npython3 ~/.local/share/${appname}/${appname}.py \$@\ndeactivate" > ~/.local/bin/${appname}
chmod u+x ~/.local/bin/${appname}

source ~/.local/share/${appname}/venv/bin/activate
pip install -r requirements.txt
deactivate
