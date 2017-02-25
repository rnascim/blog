#!/bin/bash
git add . 
git commit -am"Test Message"
targString=$( git commit -am"Test Message" | awk '{print $0}')

if [[ $targString == *"nothing to commit"*]]; then
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
