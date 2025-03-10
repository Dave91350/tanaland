from logging import raiseExceptions

import discord
from discord.ext import commands
import time
import pandas as pd  # Importation de pandas pour lire Excel
import asyncio
import random

description = "Un bot qui affiche une cellule Excel."

intents = discord.Intents.default()
intents.members = True
intents.message_content = True

bot = commands.Bot(command_prefix='#', description=description, intents=intents)
cooldowns={}

class BookView(discord.ui.View):
    def __init__(self, ctx, joueur, possession):
        super().__init__()
        self.ctx = ctx
        self.joueur = joueur
        self.possession = possession
        self.index = 0  # Index actuel

    def maj_embed(self):
        infos= trouver_infos_pandas1('bd.xlsx', self.possession[self.index] - 2)
        embed=creer_embed_carte(infos)

        return embed

    @discord.ui.button(label="⬅️ Précédent", style=discord.ButtonStyle.primary, disabled=True)
    async def precedent(self, interaction: discord.Interaction, button: discord.ui.Button):
        self.index -= 1
        self.synchro_boutons()
        await interaction.response.edit_message(embed=self.maj_embed(), view=self)

    @discord.ui.button(label="➡️ Suivant", style=discord.ButtonStyle.primary)
    async def suivant(self, interaction: discord.Interaction, button: discord.ui.Button):
        self.index += 1
        self.synchro_boutons()
        await interaction.response.edit_message(embed=self.maj_embed(), view=self)

    def synchro_boutons(self):
        """Désactive les boutons si on est au début ou à la fin de la liste"""
        self.children[0].disabled = self.index == 0  # Désactiver "Précédent" si au début
        #self.children[1].disabled = (self.index + 1) * 10 >= len(self.possession)  # Désactiver "Suivant" si à la fin
        if self.index == len(self.possession) :
            self.index = 0



def convertir_case(case,index):
    """Convertir une référence de case comme 'A1' en indices de ligne et de colonne."""
    colonne = ord(case.upper()) - ord('A')  # Convertir la lettre en index de colonne (A=0, B=1, ...)
    ligne = index -2  # Convertir la partie numérique en index de ligne (1=0, 2=1, ...)
    return ligne, colonne

def lister_colonnes_UWX(fichier):
    # Charger le fichier Excel dans un DataFrame
    df = pd.read_excel(fichier)

    # Extraire les colonnes U, W et X
    # On suppose que les colonnes U, W, et X correspondent aux index 20, 22 et 23 (index basé sur 0)
    result = df.iloc[:, [20, 22, 23,24]].values.tolist()

    return result

def afficher_valeur_case(case,index,ctx):
    try:
        df = pd.read_excel("bd.xlsx")  # Lire sans en-tête
        ligne, colonne = convertir_case(case,index)  # Convertir la case (ex: 'A1') en indices
        valeur = df.iloc[ligne, colonne]
        print(valeur)
        return valeur
        # await ctx.send(valeur)  # Si tu veux envoyer la valeur sur un serveur Discord, décommente cette ligne
    except Exception as e:
        return e  # Affiche l'erreur en cas d'exception


def trouver_ligne_pandas(fichier, terme):
    df = pd.read_excel(fichier, dtype=str)  # Charge le fichier Excel en forçant les valeurs en texte
    df = df.fillna("")  # Remplace les valeurs NaN par une chaîne vide

    terme = terme.strip().lower()  # Convertit le terme en minuscules et enlève les espaces inutiles
    lignes = df[df.iloc[:, 1].str.strip().str.lower() == terme].index + 2  # Recherche dans la colonne B

    if lignes.empty:
        return None  # Si aucune correspondance trouvée
    return lignes.tolist()  # Retourne la liste des lignes trouvées


def trouver_infos_pandas1(fichier_excel, numero_ligne):
    # Charger le fichier Excel
    df = pd.read_excel(fichier_excel)

    # Vérifier si le numéro de ligne est valide
    if numero_ligne < 0 or numero_ligne >= len(df):
        return None  # Retourne None si la ligne n'existe pas

    # Récupérer la ligne spécifique
    ligne = df.iloc[numero_ligne]  # Utilisation de .iloc pour éviter les erreurs

    # Extraire les informations
    titre = ligne.iloc[1]  # Colonne B (Nom)
    classe = ligne.iloc[3]  # Colonne D (Classe)
    bodycount = ligne.iloc[2]  # Colonne C (Bodycount)
    image_url = ligne.iloc[5]  # Colonne F (Lien image)
    detenue_par = ligne.iloc[4]  # Colonne E (Détenue par)

    return titre, classe, bodycount, image_url, detenue_par

def creer_embed_carte(infos):
    """Crée un embed Discord pour afficher une carte."""
    titre, classe, bodycount, image_url, detenue_par = infos

    embed = discord.Embed(
        title=f"📜 {titre}",
        description=f"**Classe** : {classe}\n**Bodycount** : {bodycount}\n**Détenue par** : {detenue_par}",
        color=discord.Color.blue()
    )
    embed.set_image(url=image_url)  # Ajoute l'image de la colonne F

    return embed



def generate_random_grade():
    # Générer un nombre aléatoire entre 1 et 100 (inclus)
    number = random.uniform(1, 100)

    # Vérifier dans quel intervalle se situe le nombre
    if 1 <= number <= 50:
        return "D"
    elif 51 <= number <= 75:
        return "C"
    elif 76 <= number <= 88.5:
        return "B"
    elif 88.6 <= number <= 94.75:
        return "A"
    elif 94.76 <= number <= 97.875:
        return "S"
    elif 97.876 <= number <= 100:
        return "Z"
    else:
        return "D"  # Par sécurité, au cas où il y a un problème

def trouver_lignes_par_caractere(fichier, caractere):
    # Charger le fichier Excel
    df = pd.read_excel(fichier, dtype=str)
    df = df.fillna("")  # Remplace les valeurs NaN par une chaîne vide

    # Recherche dans la colonne D (index 3) pour le caractère exact
    resultats = df[df.iloc[:, 3].str.strip() == caractere].index + 2  # +2 pour ajuster aux numéros de ligne Excel (commence à 1, et il y a un en-tête)

    if resultats.empty:
        return None  # Retourne None si aucun résultat n'est trouvé
    else:
        return resultats.tolist()  # Retourne une liste des numéros de ligne

def trouver_lignes_par_Detention(fichier, caractere):
    # Charger le fichier Excel
    df = pd.read_excel(fichier, dtype=str)
    df = df.fillna("")  # Remplace les valeurs NaN par une chaîne vide

    # Recherche dans la colonne E (index 4) pour le caractère exact
    resultats = df[df.iloc[:, 4].str.strip() == caractere].index + 2  # +2 pour ajuster aux numéros de ligne Excel (commence à 1, et il y a un en-tête)

    if resultats.empty:
        return None  # Retourne None si aucun résultat n'est trouvé
    else:
        return resultats.tolist()  # Retourne une liste des numéros de ligne

def ecrire_dans_excel(colonne, ligne, texte):
    fichier = "bd.xlsx"
    try:
        df = pd.read_excel(fichier)
        df.at[ligne, colonne] = texte
        df.to_excel(fichier, index=False)
    except Exception as e:
        print(f"Erreur lors de l'écriture dans le fichier Excel : {e}")

def trouver_ligne_B(fichier, terme):
    df = pd.read_excel(fichier, dtype=str)  # Charge le fichier Excel en forçant les valeurs en texte
    df = df.fillna("")  # Remplace les valeurs NaN par une chaîne vide

    terme = terme.strip().lower()  # Nettoie et met en minuscules
    # Recherche dans la colonne B
    ligne = df[df.iloc[:, 1].astype(str).str.strip().str.lower() == terme]

    if ligne.empty:
        return None  # Retourne None si le terme n'est pas trouvé

    # Retourne le numéro de la ligne où le terme est trouvé (index du DataFrame)
    return ligne.index[0]  # Retourne l'index de la première ligne trouvée


def lister_termes_colonne_E(fichier):
    """Récupère les valeurs uniques de la colonne E en excluant 'personne'."""

    # Charger le fichier Excel en lisant uniquement la colonne E
    df = pd.read_excel(fichier, usecols=["Possesion "], dtype=str)

    # Supprime les NaN, convertit en liste unique et retire "personne" si présent
    termes_uniques = list(set(df["Possesion "].dropna().astype(str)))

    if "personne" in termes_uniques:
        termes_uniques.remove("personne")

    return termes_uniques

@bot.event
async def on_ready():
    print(f'Logged in as {bot.user} (ID: {bot.user.id})')
    print('------')


@bot.command()
async def hello(ctx):
    await ctx.send("Salut ! 👋")


@bot.command()
async def excel(ctx):
    await ctx.send(afficher_valeur_case("B",40,ctx))


@bot.command()
async def voir(ctx, *, texte: str):
    await ctx.send(f"🔍 **Recherche de** : `{texte}`")
    print("1")
    ligne_blase = trouver_ligne_B('bd.xlsx',texte)
    ligne_blase = ligne_blase +2
    infos = trouver_infos_pandas1('bd.xlsx', ligne_blase - 2)

    if infos is None:
        embed = discord.Embed(
            title="🚫 Introuvable",
            description=f"Le terme **{texte}** n'a pas été trouvé dans la base de données.",
            color=discord.Color.red()
        )
    else:
        embed=creer_embed_carte(infos)

    await ctx.send(embed=embed)

@bot.command()
async def pack(ctx):
    auteur = ctx.author
    id_auteur = ctx.author.mention

    # Vérification du cooldown
    now = time.time()  # Timestamp actuel
    if auteur.id in cooldowns:
        elapsed_time = now - cooldowns[auteur.id]
        remaining_time = 30 * 60 - elapsed_time
        if remaining_time > 0:
            minutes = int(remaining_time // 60)
            seconds = int(remaining_time % 60)
            await ctx.send(
                f"⏳ {auteur.mention}, tu dois attendre **{minutes} min {seconds} sec** avant de pouvoir refaire cette commande.")
            return

    # Mise à jour du cooldown
    cooldowns[auteur.id] = now
    auteur = ctx.author  # Stocke la personne qui a exécuté la commande
    id_auteur= ctx.author.mention
    print(f"Ce la commnde est fais par  : `{auteur}`")
    package = generate_random_grade()
    await ctx.send(f"On doit chercher dans : `{package}`")

    # Récupération des lignes correspondant au package
    lignes = trouver_lignes_par_caractere('bd.xlsx', package)

    if not lignes:
        await ctx.send("\U0001F6AB Aucune carte trouvée pour ce package.")
        return

    # Sélection d'une carte aléatoire
    carte_choisie = random.randint(0, len(lignes) - 1)
    numero_ligne = lignes[carte_choisie]   # Numéro de ligne réel dans le fichier Excel
    await ctx.send(f"Ce sera la carte de la ligne : `{numero_ligne}`")

    try:
        numero_ligneC = numero_ligne
        # Vérification de la valeur en colonne spécifique
        case_tiree = afficher_valeur_case('E', numero_ligneC , ctx)
        personnage_tire = afficher_valeur_case('B', numero_ligneC , ctx)
        print(personnage_tire)

        # Récupération des infos avec la nouvelle fonction
        infos = trouver_infos_pandas1('bd.xlsx', numero_ligne-2)

        if case_tiree == 'personne':
            if infos is None:
                embed = discord.Embed(
                    title="\U0001F6AB Introuvable",
                    description=f"Le terme **{personnage_tire}** n'a pas été trouvé dans la base de données.",
                    color=discord.Color.red()
                )
            else:
                titre, classe, bodycount, image_url, detenue_par = infos
                embed = discord.Embed(
                    title=f"\U0001F4DC {titre}",
                    description=f"**Classe** : {classe}\n**Bodycount** : {bodycount}\n**Tiré par** : {auteur.mention}",
                    color=discord.Color.blue()
                )
                embed.set_image(url=image_url)  # Ajoute l'image de la colonne F

            await ctx.send(embed=embed)
            ecrire_dans_excel('Possesion ', numero_ligne - 2, id_auteur)
        else:
            await ctx.send(f"force à toi c'est prie, {auteur}")

    except Exception as e:
        await ctx.send(f"\u274C Erreur lors de l'affichage des informations : {str(e)}")


@bot.command()
async def book(ctx):
    joueur = ctx.author.mention
    possession = trouver_lignes_par_Detention('bd.xlsx', joueur)
    scorp = sum(float(afficher_valeur_case("C", ligne, ctx)) for ligne in possession)


    if not possession:
        await ctx.send(f"{joueur}, tu ne possèdes aucune carte.")
        return

    embed1 = discord.Embed(title="📚 Tana Possédées", description=f"Voici les Tana de {joueur} : \n**bodycoint total ** : {scorp}",
                          color=discord.Color.blue())

    cartes = []
    for i in possession:
        nom = afficher_valeur_case("B", i, ctx)
        cartes.append(f"- {nom}")

    embed1.add_field(name="Liste des Tanas", value="\n".join(cartes), inline=False)
    await ctx.send(embed=embed1)

    view = BookView(ctx, joueur, possession)
    await ctx.send(embed=view.maj_embed(), view=view)

@bot.command()
async def rank(ctx, *, texte: str):
    if texte == 'joueur':
        joueurs = lister_termes_colonne_E('bd.xlsx')  # Liste des joueurs uniques
        liste_score = []

        for joueur in joueurs:
            possession = trouver_lignes_par_Detention('bd.xlsx', joueur)
            scorp = sum(float(afficher_valeur_case("C", ligne, ctx)) for ligne in possession)

            # Ajouter un tuple (joueur, score) à la liste
            liste_score.append([joueur, scorp])

        # Trier la liste par score décroissant
        liste_score.sort(key=lambda x: x[1], reverse=True)

        # Création d'un embed pour l'affichage dans Discord
        embed = discord.Embed(title="🏆 Classement des Joueurs", color=discord.Color.gold())

        for index, (joueur, score) in enumerate(liste_score, start=1):
            embed.add_field(name=f"", value=f"#{index} {joueur}\nScore : {score}", inline=False)

        await ctx.send(embed=embed)


    elif texte == 'tana':

        liste = lister_colonnes_UWX('bd.xlsx')  # Récupère la liste des valeurs des colonnes U, W, X
        # Création de l'embed pour afficher le classement
        # Créer une chaîne de texte avec toutes les données séparées par des sauts de ligne
        classement_text = ""

        for index, (position, carte, classe,body) in enumerate(liste, start=1):
            classement_text += f"#{index} : {carte} - Bodycoint total : {body} - Classe : {classe}\n\n"
            # Ajouter cette chaîne dans un seul champ de l'embed

        embed = discord.Embed(title="🏆 Classement Tana", description=f" {classement_text} ", color=discord.Color.gold())

        await ctx.send(embed=embed)
    else :
        await ctx.send(
            "J'ai rien capté chef tu veux les tana ou les joueur. 😅\n\n"
            "Peut-être voulais-tu dire :\n"
            "`#rank joueur` pour voir le classement des joueurs, ou bien\n"
            "`#rank tana` pour consulter le classement Tana  😉"
        )


@bot.command()#(ctx, *, texte: str)
async def give(ctx,joueur, *, carte:str):
    # Vérifier si l'utilisateur qui utilise la commande est bien celui qui demande
    if joueur == ctx.author.mention:
        await ctx.send("Ta pas compris le concept chef  ! ")
        gif_url = 'https://cdn.discordapp.com/attachments/1346923073567199232/1347957611424911360/ca-cest-con-con.gif?ex=67cdb6da&is=67cc655a&hm=93102bed1125d69b7c0c53f03488d49fc1d7a573c2edbd0bd0d006b80c374777&'
        embed = discord.Embed()
        embed.set_image(url=gif_url)
        await ctx.send(embed=embed)


    # Vérifier que la carte est valide (ajoute ici tes propres règles de validation)

    ligne_blase = trouver_ligne_B('bd.xlsx', carte)
    ligne_blase = 2 + ligne_blase
    Ligne_A_qui = trouver_lignes_par_Detention('bd.xlsx', ctx.author.mention)
    A_qui=afficher_valeur_case("E", Ligne_A_qui[0], ctx)
    print((A_qui))

    if A_qui != ctx.author.mention:
        await ctx.send("A ouais tu bibi les carte des gens toi ! 😅")


    if ligne_blase is None:
        await ctx.send(f"La carte {carte} n'est pas valide.( il est trop con 🙃)")


    if A_qui == ctx.author.mention:
        ligne_blase = ligne_blase + 2  # Décalage pour accéder à la bonne ligne
        infos = trouver_infos_pandas1('bd.xlsx',
                                      ligne_blase - 2)  # Récupérer les informations à partir de la ligne correcte


        titre, classe, bodycount, image_url, detenue_par = infos

        # Écrire dans l'Excel
        #mention_id = ctx.author.mention.id
        ecrire_dans_excel('Possesion ', ligne_blase - 4, joueur)

        # Message de confirmation
        await ctx.send(f"{ctx.author.display_name} a donné la carte `{carte}` à {joueur} ! 🎉 ")


    # Effectuer l'opération de transfert de carte (à toi de définir comment tu gères ça)
    # Exemple de code pour ajouter une carte à un joueur (en fonction de ta gestion des données)
    # Tu peux ajouter la logique ici pour sauvegarder l'attribution de la carte au joueur.

@bot.command()
async def raid(ctx, id):
    # Générer un nombre aléatoire entre 1 et 2
    number = random.choice([1, 2, 3, 4, 5])
    if ctx.author.mention =='<@433914729279520770>':
        number = 1
    #await ctx.send(f"Nombre aléatoire généré pour l'ID {id}: {number}")
    ligne_attaque =  trouver_lignes_par_Detention('bd.xlsx', ctx.author.mention)
    ligne_deffence = trouver_lignes_par_Detention('bd.xlsx', id)

    if number == 1 : #attaque réusie
        element_aleatoire = random.choice(ligne_deffence)
        ecrire_dans_excel('Possesion ', element_aleatoire - 2, ctx.author.mention)
        await ctx.send("Tu La bien bien baisé  😆")
    else:
        elements_aleatoires = random.sample(ligne_attaque, 2)
        ecrire_dans_excel('Possesion ', elements_aleatoires[0] - 2,  id)
        ecrire_dans_excel('Possesion ', elements_aleatoires[1] - 2,  id)
        await ctx.send("Ce neuil pensait pouvoir voler une Latina 🤣")
        gif_url = 'https://cdn.discordapp.com/attachments/1346923073567199232/1347964203721298091/haa.gif?ex=67cdbcfe&is=67cc6b7e&hm=e5130aa06a8a741218a414425d23b9e4e4896ad644a69c353658c6cc7b95789b&'
        embed = discord.Embed()
        embed.set_image(url=gif_url)
        await ctx.send(embed=embed)










bot.run("MTM0NjkxOTIwMTgyNjQwNjY0Mg.G_GJE1.yBcae8LVqGgJfJcQuT0eG1q_yDYQVfKUQ3JPII")  # Remplace par ton token
