#!/bin/bash

appname=kimaiPyTool

if [ ! -d ~/.local/share/${appname} ]
then
    echo "Create folder ~/.local/share/${appname}"
    mkdir ~/.local/share/${appname}
fi

echo "Copy ${appname}.py into ~/.local/share/${appname}/"
cp ${appname}.py ~/.local/share/${appname}/


if [ ! -d ~/.local/share/${appname}/venv ]
then
    echo "Create Python virtual env ~/.local/share/${appname}"
    python3 -m venv ~/.local/share/${appname}/venv
    echo "Create ~/.local/bin/${appname}"
    echo -e \#\!/bin/bash"\n\nsource ~/.local/share/${appname}/venv/bin/activate\npython3 ~/.local/share/${appname}/${appname}.py \$@\ndeactivate" > ~/.local/bin/${appname}
    chmod u+x ~/.local/bin/${appname}
fi

if ! diff requirements.txt ~/.local/share/${appname}/venv/requirements.txt > /dev/null
then
    echo "Update Python virtual env ~/.local/share/${appname}"
    source ~/.local/share/${appname}/venv/bin/activate
    pip install -r requirements.txt > /dev/null
    deactivate
    cp requirements.txt ~/.local/share/${appname}/venv/requirements.txt
fi

if ! grep ${appname} ~/.bashrc > /dev/null > /dev/null
then
    echo "Add \$(register-python-argcomplete ${appname}) in ~/.bashrc"
    echo -e '\neval "$(register-python-argcomplete '${appname}')"' >> ~/.bashrc
    echo "You need to open new console to have autocompletion"
fi
