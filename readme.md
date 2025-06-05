<div alt="header" style="display: flex; justify-content: center; flex-direction: row; width: 100%; border: solid grey; height: 100px;">
    <div alt="logo" style="width: 15%; height: 50px; align-items: center; border-right: solid grey; height: 100px;">
        <img src="./src/public/images/logo-d.png" style="padding-top: 1.5rem; text-align: justify; max-width: 100%; max-height: 100%; object-fit: contain; object-position: center; object-fit: fill;">
    </div>
    <div alt="title" style="width: 70%; height: 50px; display: flex; align-items: center; flex-direction: column">
        <p style="color:#b9b9dc; font-size: 2rem; margin-bottom: 0rem; margin-top: 0.5rem; padding-bottom: 0.5rem; font-weight: bold; width: 100%; text-align: center; border-bottom: solid grey;">Mode Opératoire</p>
        <p style="">Cession de créance factor</p>
    </div>
    <div alt="end" style="width: 15%; border-left: solid grey; height: 100px;"></div>
</div>
<br>

Ce fichier contient le mode opératoire de l'application de cession de créance factor. Il permet de décrire les différentes étapes à suivre pour effectuer une cession de créance factor.

### 1 - Description Général du processus

La cession de créance factor est un processus qui permet à une entreprise de céder ses créances à un factor. Le factor est une société spécialisée dans le financement et la gestion des créances. Il permet à l'entreprise de bénéficier d'une avance de trésorerie sur ses créances en échange d'une commission. La cession de créance factor est un moyen efficace pour l'entreprise de financer son activité et de réduire son risque de non-paiement.

l'application de cession de créance factor permet de gérer ce processus de manière automatisée. Elle permet à l'entreprise de saisir les informations relatives à la cession de créance, de générer les documents nécessaires et de suivre l'avancement du processus.

Le point de départ du processus est l'ouverture du fichier FACTOR_TEMPLATES.xlsm. Ce fichier contient les modèles de documents nécessaires à la cession de créance factor. Il permet à l'entreprise de saisir les informations relatives à la cession de créance et de générer les documents finaux.

### 2 - Ouvrir le fichier FACTOR_TEMPLATES.xlsm

1. ouvrir le fichier
2. se rendre dans l'onglet panel
    - L'onglet `PANEL` permet de piloter le processus cession de créance en quelques cliques

:warning: **Attention:** Avant de commencer tout action dans l'application, il va falloir renseigner certains champs.
Dans les onglets remises vous trouverez des cellules à renseigner pour les champs suivants:
    - Dossier de remise factor
    - Numero du client
    - Numéro de la remise
    - Date de la remise

### 3 - Description des onglets

#### 3.1 - Onglet PANEL

L'onglet PANEL permet de piloter l'ensemble de l'application. Il contient les boutons qui permettent à l'entreprise de naviguer entre les différentes fonctionnalités de l'application.
[Fonctionnement du panel](#panel)

#### 3.2 - Onglet IMPORT

L'onglet IMPORT permet à l'entreprise de visualiser et selectionner les données à envoyer en cession de créance factor. Il contient les informations relatives aux créances à céder, telles que le numéro de la facture, le montant de la facture, la date d'échéance, etc. A la fin de chaque ligne, vous trouverez une checkbox qui permet de selectionner les données à envoyer en cession de créance factor.

#### 3.3 - Onglets REMISES (Domestique/Export)

Les onglets REMISES (Domestique/Export) permettent à l'entreprise de visualiser les données des remises factor. Ils contiennent le récapitulatifs des lignes qui seront cédé au factor et les informations relatives aux remises factor, telles que le numéro de la remise, le montant de la remise, la date de la remise, etc. Ces onglets permettent également de visualiser les numéros d'écritures générées dans SAP une fois la remise factor effectuée.

### 4 - Utiliser les boutons de l'onglet panel {#panel}

L'onglet panel contient plusieurs boutons qui permettent à l'entreprise de piloter l'application. Chaque bouton a une fonction spécifique et doit être utilisé dans un ordre précis pour garantir le bon fonctionnement de l'application.

L'onglet panel est divisé en deux sections:

- La section REMISES qui permet de gérer les remises factor
    Chaque bouton doit etre cliqué dans l'ordre suivant:
    - BOUTON [IMPORT](#import-remise)
    - BOUTON [RELOAD](#reload-remise)
    - BOUTON [MAKE](#make-remise)
    - BOUTON [RESET](#reset-remise) (si necessaire)
- La section PAIEMENTS qui permet de gérer les paiements factor
    Chaque bouton doit etre cliqué dans l'ordre suivant:
    - BOUTON [IMPORT](#import-paiement)
    - BOUTON [SEND](#send-paiement)
    - BOUTON [RESET](#reset-paiement) (si necessaire)

:warning: **Attention:** Il est important de respecter l'ordre des boutons pour garantir le bon fonctionnement de l'application. chaque bouton ne doit etre utilisé qu'une seule fois.
[Utilisation des boutons](#boutons)

### 5 - Utilisation des boutons {#boutons}

#### 5.1 - BOUTON IMPORT (REMISE) {#import-remise}

Le bouton IMPORT permet à l'entreprise d'importer les données nécessaires à la cession de créance factor. Il ouvre une fenêtre de dialogue qui permet de sélectionner le fichier contenant les données à importer.

:warning: **Attention:** Le fichier à importer doit être au format CSV et contenir les informations nécessaires à la cession de créance factor.
Une fois le fichier importé, rendez-vous dans l'onglet Import pour vérifier que les données ont bien été importées.
Si les données sont correctemment chargées, et que les checkbox sont bien cochées, vous pouvez passer à l'étape suivante.

#### 5.2 - BOUTON RELOAD (REMISE) {#reload-remise}

Les bouton RELOAD permettent de recharger les données des onglets remises (Domestioque/Export) une fois la selection des données effectuée dans l'onglet Import.

:warning: **Attention:** Il est très important de ne pas cliquer sur ce bouton avant d'avoir importé les données.

#### 5.3 - BOUTON MAKE (REMISE) {#make-remise}

Les boutons MAKE permettent à la fois de générer les documents finaux de la cession de créance factor et d'enregistrer les écriteure de cession factor et statistique dans SAP. Si tout ces bien passé, vous devriez voir apparaitre un message de confirmation avec le nom du fichier généré ainsi que les numéros d'écritures générées dans la colonne "Référence écriture factor" des onglets remises (Domestique/Export).

:warning: **Attention:** Il est très important de ne pas cliquer sur ce bouton avant d'avoir importé les données et rechargé les données des onglets remises (Domestioque/Export).
Avant de cliquer sur le bouton MAKE, assurez-vous que les données importées sont correctes et que les champs obligatoires sont bien renseignés. Une fois le bouton MAKE cliqué, les données ne pourront plus être modifiées.

#### 5.4 - BOUTON RESET (REMISE) {#reset-remise}

Le bouton RESET permet de réinitialiser les données de l'onglet Import. Il permet à l'entreprise de recommencer le processus de cession de créance factor depuis le début.

:warning: **Attention:** Il est très important de ne pas cliquer sur ce bouton si vous avez eu un soucis lors du processus d'envoie des données dans SAP. Si vous avez un doute, verrifiez qu'il n'y a pas de données dans la colonne "Référence écriture factor" des onglets remises (Domestique/Export) avant de cliquer sur ce bouton.

#### 5.5 - BOUTON IMPORT (PAIEMENT) {#import-paiement}

Le bouton IMPORT permet à l'entreprise d'importer les données nécessaires au paiement factor. Il ouvre une fenêtre de dialogue qui permet de sélectionner le fichier contenant les données à importer.

:warning: **Attention:** Le fichier à importer doit être au format CSV et contenir les informations nécessaires au paiement factor. Une fois le fichier importé, rendez-vous dans l'onglet PAIMENT pour vérifier que les données ont bien été importées.

#### 5.8 BOUTON SEND (PAIEMENT) {#send-paiement}

Le bouton SEND permet d'enregistrer les écritures paiement dans SAP. Si tout c'est bien passé, Si tous c'est bien passez, toues les lignes traitées devrait etre passez au vert.

#### 5.7 - BOUTON RESET (PAIEMENT) {#reset-paiement}

Le bouton RESET permet de réinitialiser les données de l'onglet PAIEMENT. Il permet à l'entreprise de recommencer le processus de paiement factor depuis le début.

:warning: **Attention:** Il est très important de ne pas cliquer sur ce bouton si vous avez eu un soucis lors du processus de paiement factor. Si vous avez un doute, verrifiez qu'il n'y a pas de données dans la colonne "Référence paiement factor" de l'onglet PAIEMENT avant de cliquer sur ce bouton.
