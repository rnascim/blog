#!/bin/bash
git add . 
git commit -am"Test Message"
#git push origin master
targString=$( git push origin master | awk '{print $2}')
echo $targString
exit 0  

ssh rogerionascimento.com << EOF
  sudo su -
  cd ~/blog
  ./deploy.sh
EOF
