# Le projet

Le projet a pour but de créer plusieurs fichiers excel à partir d'un fichier .csv contenant des données, les fichiers excel doivent être générés à l'aide d'un scrip en Python.

# Description 

Le fichier drug_consumption.csv contient les résultats d'un sondage portant sur la consomation de drogues pour chaque répondants, chacun des individus sondés à également fournis des informations sur leur age, genre, nationalitée, niveau d'éducation, mais également des résultats correspondant à des tests de personnalités. 

Les données tels que présentent dans le fichier source .csv sont tous codés et a necessité un recodage des valeurs numériques vers leurs valeurs textuelles correspondantes. 

## Les drogues concernés

Au sein de ce sondage il est abordé la consomation de divers substances psychoactives allant de la cafféine et chocolat, au LSD ou benzodiazepine jusqu'au crack et à la méthadone. 

Ce large specte de divers substances permet de rendre compte d'un usage d'une grande quantitée de substances psychoactives.

Il est intéressant de noter que lors de l'élaboration du sondage les chercheurs ont également demandé aux individus si ils ont consomé une drogue appellé le "Semeron", totalement fictive cette substance à pour but de déceler si des répondants mentirait sur leur consomation pour une raison ou une autre. 

La liste des drogues présentent dans le sondage : 
- Alcohol
- Amphet
- Amyl Nitrite 
- Benzo Diazepine
- Caffeine
- Cannabis 
- Chocolate 
- Cocaine
- Crack Cocaine
- Ecstasy
- Heroin
- Ketamine
- Legal High
- LSD
- Methadone
- Magic Mushrooms
- Nicotine 
- Semeron
- Volatile Substance Abuse 


## Les données socio-démographiques

Les sondés ont renseigné leur tranche d'âge, leur genre, leur niveau de dîplomes, leur ethnie.

## Les données de personnalités

Ils ont également plusieurs score issus de tests de personnalités ayant pour but de rendre compte de potentiels profils types présent dans les usages. 

**Nscore** : Nscore est le NEO-FFI-R Neuroticisme. Le neuroticisme est l'un des cinq grands traits de personnalité de haut niveau dans l'étude de la psychologie. Les individus qui obtiennent un score élevé en neuroticisme sont plus susceptibles que la moyenne d'être de mauvaise humeur et de ressentir des sentiments tels que l'anxiété, l'inquiétude, la peur, la colère, la frustration, l'envie, la jalousie, la culpabilité, la dépression et la solitude.

**EScore** : EScore (Réel) est l'Extraversion du NEO-FFI-R. L'extraversion est l'un des cinq traits de personnalité de la théorie des Big Five. Elle indique à quel point une personne est extravertie et sociale. Une personne qui obtient un score élevé en extraversion lors d'un test de personnalité est le centre de l'attention. Elle aime être avec les gens, participer à des rassemblements sociaux et déborde d'énergie.

**Oscore** : Oscore (Réel) est l'Ouverture à l'expérience du NEO-FFI-R. L'ouverture est l'un des cinq traits de personnalité de la théorie des Big Five. Elle indique à quel point une personne est ouverte d'esprit. Une personne ayant un niveau élevé d'ouverture à l'expérience lors d'un test de personnalité aime essayer de nouvelles choses. Elle est imaginative, curieuse et ouverte d'esprit. Les individus ayant un faible niveau d'ouverture à l'expérience préfèrent ne pas essayer de nouvelles choses. Ils sont fermés d'esprit, littéraux et aiment avoir une routine.

**Ascore** : Ascore (Réel) est l'Agréabilité du NEO-FFI-R. L'agréabilité est l'un des cinq traits de personnalité de la théorie des Big Five. Une personne ayant un niveau élevé d'agréabilité lors d'un test de personnalité est généralement chaleureuse, amicale et diplomate. Elle a généralement une vision optimiste de la nature humaine et s'entend bien avec les autres.

**Cscore** : Cscore (Réel) est la Conscience du NEO-FFI-R. La conscience est l'un des cinq traits de personnalité de la théorie des Big Five. Une personne obtenant un score élevé en conscience a généralement un haut niveau d'autodiscipline. Ces individus préfèrent suivre un plan plutôt que d'agir spontanément. Leur planification méthodique et leur persévérance les rendent généralement très performants dans leur profession choisie.

**Impulsivité** : Impulsivité (Réel) est l'impulsivité mesurée par le BIS-11. En psychologie, l'impulsivité est une tendance à agir sur un coup de tête, affichant un comportement caractérisé par peu ou pas de prévoyance, de réflexion ou de considération des conséquences. Si vous décrivez quelqu'un comme impulsif, cela signifie qu'il agit soudainement sans y réfléchir attentivement au préalable.

**Sensation** : SS (Réel) est la recherche de sensations mesurée par l'ImpSS. La sensation est l'information sur le monde physique obtenue par nos récepteurs sensoriels, et la perception est le processus par lequel le cerveau sélectionne, organise et interprète ces sensations. En d'autres termes, les sens sont la base physiologique de la perception.
# La structure du projet


- data.py
- drug_consumption.csv
- excel_python.ipynb
- .gitignore
- README.md

## Documentation

- Le fichier drug_consumption.csv est le fichier contenant les données source
- Le fichier data.py est le script Python à lancer qui créera les différents fichiers excel
- Le fichier excel_python.ipynb est le Jupyter Notebook avec lequel j'ai expérimenté la création des fichiers excel.
