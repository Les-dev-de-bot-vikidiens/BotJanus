import pywikibot
import time
from pywikibot import textlib

# Variables

BOT_NAME = "le_nom_du_bot"  # Remplacez par le nom de votre bot

# Configuration pour Vikidia en français
FAMILY = 'vikidia' # vous pouvez aussi mettre pour wikipédia : 'wikipedia'
LANG = 'fr' # pour Vikidia en français

# Nombre de pages de discussion à analyser
N_PAGES = 1000

# Délai (en secondes) entre chaque édition pour éviter le flood
SLEEP_TIME = 5 # conseillé pour ne pas surcharger les serveurs

# Résumé d’édition
EDIT_SUMMARY = "Bot: substitution"

# ==============================
# 🚀  DÉBUT DU SCRIPT
# ==============================

def main():
    # Création du site et connexion
    site = pywikibot.Site(LANG, FAMILY)
    site.login()

    print(f"Connecté à {LANG}.{FAMILY}.org en tant que {site.user()} ✅")
    print(f"Recherche de {N_PAGES} pages de discussion utilisateur aléatoires...\n")

    # Récupération de pages aléatoires dans l'espace de noms 3 (Discussion utilisateur)
    pages = site.randompages(total=N_PAGES, namespaces=[3])

    pages_modified = 0

    for page in pages:
        print(f"➡️  Analyse de la page : {page.title()}")
        try:
            text = page.text

            # 1. Première passe : On cible et remplace les variantes d'IP (ex: {{Bienvenue ip}}, {{bienvenue IP}}, etc.)
            text_after_ip = textlib.replaceExcept(
                text,
                r'\{\{\s*[Bb]ienvenue\s+[Ii][Pp]\s*\}\}',
                '{{subst:bienvenue IP}}',
                exceptions=['nowiki', 'comment', 'pre', 'code']
            )

            # 2. Seconde passe : On cible et remplace les variantes standards (ex: {{bienvenue}}, {{Bienvenue}})
            text_after_bienvenue = textlib.replaceExcept(
                text_after_ip,
                r'\{\{\s*[Bb]ienvenue\s*\}\}',
                '{{subst:bienvenue}}',
                exceptions=['nowiki', 'comment', 'pre', 'code']
            )

            # 3. Troisième passe : On cible les modèles "averto" SAUF s'ils sont suivis d'un slash /
            # Le (?!/) bloque la capture si un slash est collé à "averto" ou "Averto"
            new_text = textlib.replaceExcept(
                text_after_bienvenue,
                r'\{\{\s*([Aa]verto(?!/)[^}]*)\}\}',
                r'{{subst:\1}}',
                exceptions=['nowiki', 'comment', 'pre', 'code']
            )

            # Sauvegarde si une modification a eu lieu
            if new_text != text:
                page.text = new_text

                # Correction du WARNING : on utilise 'bot=True' au lieu de 'botflag=True'
                page.save(summary=EDIT_SUMMARY, minor=True, bot=True)

                pages_modified += 1
                print(f"   ✨ Modèle(s) substitué(s) avec succès !")
                print(f"   ⏳ Attente de {SLEEP_TIME} sec avant la prochaine page...\n")
                time.sleep(SLEEP_TIME)
            else:
                print("   ✅ Pas de modèle à substituer (ou protégé par <nowiki>).")

        except pywikibot.exceptions.Error as e:
            print(f"   ❌ Erreur Pywikibot sur {page.title()} : {e}")
        except Exception as e:
            print(f"   ⚠️ Autre erreur sur {page.title()} : {e}")
            continue

    print(f"🏁 Script terminé ! Pages modifiées par BOT_NAME : {pages_modified} 🎉")

if __name__ == "__main__":
    main()