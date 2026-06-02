Correction GO SIYAHA DAP V4

1) Copier le fichier api/index.py de ce dossier dans votre repo GitHub au même emplacement : api/index.py
2) Remplacer requirements.txt par celui fourni ici.
3) Commit changes.
4) Dans Vercel, attendre Ready ou cliquer Redeploy.
5) Tester /health : il doit afficher code_version = DAP_SDT_XML_AFTER_SAVE_V4_2026_06_02.

Si code_version n'apparaît pas dans /health, le nouveau code n'est pas déployé.
