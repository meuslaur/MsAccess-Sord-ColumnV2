# MsAccess-Sord-ColumnV2
## Création automatique d'un formulaire en continu avec tri sur les colonnes

![Formulaire de démarrage](Doc/Frm_img.gif)

### Modification avec la V1

### SUPP
-----
- - Suppréssion des Suffixe.
- - Utilise un preffixe par defaut "txt_" plus de len pref/suff.
- - Suppréssion TexteOn et TexteColor.
- - Supprime les paramètres de la fonction `SordColumn`.
- - Supprime choix image on/off CommandButton sur form `F_CreateForm`.
### ADD
----
- - Création d'un module standard contenant la function d'utilisation de la classe `CSordFormColumn`.
- - Insère le code d'utilisation de la classe dans ce module.
- - 
### MOD
----
- - Utilise VarClasse et function par defauts (`m_CSordForm` et `SordColumn`).
- - Utilise CurrentProject si le dossier des images est un sous-dossier de l'application.
- - Modification Fonction `CreateModule` (Création de la classe et du module standard).
- - Modification de la fonction `CreateFormColumn`.
- - - Utilise le nom des CommandButton sans le préfixe pour le nom des champs.
- - - N'insère créer plus de code dans le formulaire créer.
- - - Lance la function avec `=SordColumn()` sur event onClick des boutons.
- - - Ajoute `=CloseSordColumn()` sur event OnClose du formulaire.

- - Modification de la classe `CSordFormColumn` pour l'adapter aux modifications de CreateFormColumn et CCreateFormContinu.

- - Utilisable dans un SF.

### Différences avec la V1

- (-)Ne permet plus d'utiliser différentes images sur les boutons.
- (+)Le formulaire ne contient plus de code.
- (+)Utilisation plus simple sur un sous-formulaire.
- (+)Plus de table contenant le code.
- (-)Plus de suffixe possible sur les TextBox et CommandButton.
- Utilisation des noms de fonction et de variable de classe par défaut.
- Choix ogligation des images sur les CommandButton.
- (+)Renforcement de la gestions des erreurs.
