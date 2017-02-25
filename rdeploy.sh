#!/bin/bash
git add . 
git commit -am"Test Message"
targString=$( git commit -am"Test Message" | awk '{print $0}')
git push origin master
echo 'push: '$targString
exit 0  

ssh rogerionascimento.com << EOF
  sudo su -
  cd ~/blog
  ./deploy.sh
  exit  
EOF
