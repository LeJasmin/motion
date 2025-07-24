import plistlib


def create_hazel_rule(name, conditions, actions):
    return {
        'name': name,
        'enabled': True,
        'conditions': conditions,
        'actions': actions,
        'matchAllConditions': True
    }


def make_condition(attribute, operator, value):
    return {
        'attribute': attribute,
        'operator': operator,
        'value': value
    }


def make_move_action(folder):
    return {
        'action': 'Move',
        'folder': folder
    }


def make_shell_script_action(script):
    return {
        'action': 'RunShellScript',
        'script': script
    }


def make_rename_action(pattern):
    return {
        'action': 'Rename',
        'pattern': pattern
    }


hazel_rules = []

hazel_rules.append(create_hazel_rule(
    "Supprimer fichiers vides",
    [make_condition("Size", "less than", 1024)],
    [make_move_action("Trash")]
))

hazel_rules.append(create_hazel_rule(
    "Supprimer fichiers système inutiles",
    [make_condition("Name", "contains", ".DS_Store")],
    [make_move_action("Trash")]
))

hazel_rules.append(create_hazel_rule(
    "Renommer – Ajout date + catégorie",
    [make_condition("Extension", "is", "pdf")],
    [make_rename_action("%date created%-%kind%-[original name]")]
))

hazel_rules.append(create_hazel_rule(
    "Convertir en Markdown",
    [make_condition("Extension", "is", "docx")],
    [make_shell_script_action('/usr/local/bin/pandoc "$1" -o "${1%.*}.md"')]
))

keywords_rules = [
    ("Déplacer Logement", "logement", "03_Logement"),
    ("Déplacer MDPH", "mdph", "09_Sante"),
    ("Déplacer Carmen", "carmen", "02_Famille/Carmen"),
    ("Déplacer Billy", "billy", "02_Famille/Billy"),
    ("Déplacer Maman", "maman", "02_Famille/Maman"),
    ("Déplacer Farisa", "farisa", "02_Famille/Farisa"),
    ("Déplacer Createch", "createch", "08_Createch"),
    ("Déplacer CAF", "caf", "05_Finances"),
    ("Déplacer Travail", "travail", "04_Travail"),
]

for rule_name, keyword, folder in keywords_rules:
    hazel_rules.append(create_hazel_rule(
        rule_name,
        [make_condition("Name", "contains", keyword)],
        [make_move_action(f"~/Documents/01_Vault_Organisation/{folder}")]
    ))

hazel_export = {
    'HazelRules': hazel_rules
}

with open("gestion_administrative.hazelrules", "wb") as fp:
    plistlib.dump(hazel_export, fp, fmt=plistlib.FMT_XML)
