# RICCARDO_SI&O


## Comandi corretto upload 


git submodule update --init --recursive


git submodule foreach 'git checkout main || git checkout master'


git submodule foreach 'git pull origin $(git rev-parse --abbrev-ref HEAD)'


git add .


git commit -m "Aggiornati submodules all'ultimo commit remoto"


git push origin main


## Comando rimozione mark del web

Unblock-File "..."