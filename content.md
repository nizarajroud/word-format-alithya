### Inventaire Scripts VM Windows

<aside>
📝 Pour bien dimensionner la cible AWS, j'aurais besoin d'un inventaire des scripts sur la VM Windows sous cette forme :

| Nom du script | Langage | Version | Frequence (cron) | Temps d'execution | Traitement (real-time/batch/mixte) | CloudWatch | PagerDuty | Critique? | A decomissionner? |

</aside>

1. **[Inventaire]** Serait-il possible de fournir la liste complete des scripts avec leur nom, langage (Python? autre?), version, et frequence d'execution ?
2. **[Temps execution]** Quel est le temps d'execution moyen et max de chaque script ? Certains depassent-ils 15 minutes ?
3. **[Real-time/Batch]** Quels scripts sont en traitement temps reel (event-driven) vs batch (cron) vs mixte ?
4. **[Observabilite]** Quels scripts sont connectes a CloudWatch ou PagerDuty ? Y a-t-il du monitoring/alerting en place pour chacun ?
5. **[Dependances]** Y a-t-il des dependances entre scripts qui imposent un ordre d'execution precis ?
6. **[Decomission]** Quels scripts sont prevus pour etre decommissionnes ou agreges dans d'autres ?

### 🛠️ Excellence Operationnelle

1. **[Monitoring]** CloudWatch est en place pour les logs. Quelles metriques manquent actuellement pour bien piloter les pipelines ?
2. **[Backup VM]** Y a-t-il une strategie de backup pour la VM Windows ? Quel est le plan de recovery en cas de perte ?
3. **[Matillion EC2]** L'instance EC2 Matillion est allumee/eteinte par une Lambda. Quels sont les horaires exacts et le cout mensuel ?
4. **[Frequence]** Matillion roule toutes les 6h pour le reporting. Quelle est la frequence la plus elevee parmi les scripts operationnels ?

### 🔐 Securite

1. **[MFA]** L'authentification est en SSO Microsoft avec key pair pour les services. Le MFA est-il prevu dans la cible AWS ?
2. **[Audit acces]** Y a-t-il un audit des acces utilisateurs prevu avant la migration ? Des comptes inactifs a nettoyer ?
3. **[VPN vs public]** Certaines APIs Lyft sont accessibles uniquement via VPN. Apres la migration, l'acces sera public. Quel mecanisme de securite est prevu pour remplacer le VPN ?
4. **[PII]** Les donnees qui transitent par le S3 Lyft contiennent-elles des PII ?

### ⚙️ Fiabilite

1. **[HA actuelle]** Le failover entre les 2 data centers est manuel (actif/passif). Quel est le temps de bascule moyen ? Combien de fois a-t-il ete utilise ?
2. **[Coexistence]** Pour la coexistence on-prem + cloud, quelle duree est acceptable ? Quel budget supplementaire ?
3. **[Rollback]** Si un script migre vers Lambda/ECS echoue, peut-on rebasculer sur la VM Windows immediatement ?

### 🚀 Performance

1. **[VM specs]** Quelles sont les specs de la VM Windows ? (CPU, RAM, disque) - Necessaire pour dimensionner Lambda/ECS
2. **[Panorama stack]** Panorama est developpe en interne. Quelle stack technique ? (Node.js/Vue.js?) Est-il dockerise ?
3. **[Data mart]** Des queries SQL sont executees dans un data mart. Ou est-il heberge et quel SGBD ?

### 💰 Couts

1. **[Cout actuel]** Quel est le cout mensuel de l'infra on-premise ? (VM, data centers, licences Windows)
2. **[Matillion licence]** Matillion est deploye via AWS Marketplace sur EC2. Quel est le cout de licence actuel ?
3. **[Budget coexistence]** Quel budget supplementaire est acceptable pour la periode de coexistence on-prem + cloud ?

### 🌱 Durabilite

1. **[Region AWS]** Confirmez-vous ca-central-1 (Montreal) comme region cible ?

### 🔗 Environnements & DevOps

1. **[Env AWS existants]** Une structure d'environnements AWS a ete mise en place par une autre compagnie mais n'est pas exploitee. Qui l'a mise en place et quelle est la structure ?
2. **[Env test Lyft]** Actuellement il n'y a pas d'environnement de test chez Lyft. La nouvelle plateforme en offrira-t-elle un ?
3. **[Docker]** Des sites web Docker existent deja. Lesquels ? Sont-ils candidats a une migration vers ECS/Fargate ?
4. **[Git]** Les scripts Python sont-ils tous versionnes dans GitHub ? Y a-t-il du code non versionne sur la VM ?