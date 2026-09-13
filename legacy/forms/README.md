# Formulaires hérités (non utilisés)

Ces fichiers `.frm` contiennent uniquement le **code** des anciens UserForms (Cockpit, Login, Settings).
La partie visuelle (contrôles, positions) vit dans des fichiers `.frx` binaires qui **n'existent pas** dans ce dépôt :
les importer dans Excel produit l'erreur « Instruction incorrecte à l'extérieur d'une procédure ».

L'interface officielle de S.A.F.A v10 est la feuille **MENU** à boutons construite par `src/vba/SAFA_Menu.bas`
(aucun UserForm requis), avec un mode console de secours (`SAFA_Console.Demarrer`).

Ces fichiers sont conservés à titre de référence si quelqu'un souhaite un jour reconstruire des UserForms
dans l'éditeur VBA (créer le formulaire à la main, puis coller le code).
