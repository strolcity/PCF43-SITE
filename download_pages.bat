@echo off
echo Ouverture des pages dans Brave...
echo Enregistrez chaque page manuellement dans le dossier "pages/" avec les noms suivants :
echo - Législatives 2024 - Circonscription 1 : LEG_24_4301.html
echo - Législatives 2024 - Circonscription 2 : LEG_24_4302.html
echo - Présidentielles 2022 - Département 43 : PRES_22_4300.html

start brave "https://www.archives-resultats-elections.interieur.gouv.fr/resultats/legislatives2024/ensemble_geographique/84/43/4301/index.php"
start brave "https://www.archives-resultats-elections.interieur.gouv.fr/resultats/legislatives2024/ensemble_geographique/84/43/4302/index.php"
start brave "https://www.archives-resultats-elections.interieur.gouv.fr/resultats/presidentielle-2022/084/043/index.php"

timeout /t 2 >nul
echo Toutes les pages sont ouvertes. Enregistrez-les dans "pages/" avec les noms indiqués ci-dessus.
pause