#!/bin/bash
if [[ $1 == "" ]] 
then
  echo -e "Git Commit comment missing "
  exit 0
fi

git add . 
targString=$( git commit -am"$0" | awk '{print $0}')

if [[ $targString == *"nothing to commit"* ]]
then
  echo -e "Nothing to do"
  exit 0
fi

git push origin master

ssh rogerionascimento.com << EOF
  sudo su -
  cd ~/blog
  ./deploy.sh
  exit  
EOF
