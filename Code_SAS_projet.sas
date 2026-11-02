/* ============================================================================
   PROJET SAS : SINISTRALITE D'UN PORTEFEUILLE DE COMPLEMENTAIRE SANTE
   Soins 2018-2019

   Entrees : contrat_h, individu_h, cotisation_h, presta_agr_h_f (bibliotheque TT)
   Sortie  : Resultats_projet.xlsx (un onglet par table de resultats
             + la base d'analyse detaillee et la table des prestations,
             section 21). Les tableaux croises dynamiques et les graphiques
             sont faits ensuite dans Excel, sur ces onglets.

   Plan du programme
     0.  Parametres et bibliotheque
     1.  Import des tables
     2.  Suppression des doublons exacts
     3.  Exposition (contrat x individu x garantie x mois)
     4.  Individus (sexe, date de naissance par mois)
     5.  Prestations agregees par mois de soins
     6.  Cotisations nettes par mois d'appel
     7.  Panel : exposition + individus + prestations + cotisations
     8.  Variables d'analyse (age, tranche, annee, sinistre)
     9.  Garanties sans cotisation et cohorte d'entree (base finale)
     10. Controles de jointure et de qualite
     11. Indicateurs globaux
     12. Resultats par annee
     13. Effet des nouveaux entrants 2019
     14. Profils : sexe, tranche d'age, qualite
     15. Evolution mensuelle
     16. Prestations par type de soins
     17. Resultats par garantie et effet de composition
     18. Concentration des prestations
     19. Sensibilite du S/P
     20. Chain ladder (provision pour sinistres tardifs)
     21. Export Excel
   ============================================================================ */


/* ---------------------------------------------------------------------------
   0. Parametres et bibliotheque
   --------------------------------------------------------------------------- */
%let chemin = /home/u64593552/sasuser.v94/Projet SAS;
%let fichier_excel = &chemin/Resultats_projet.xlsx;

libname TT "&chemin";

/* Verification du format des dates : les variables doivent etre des datetime */
proc contents data=TT.contrat_h varnum;
run;


/* ---------------------------------------------------------------------------
   1. Import dans la bibliotheque de travail
   --------------------------------------------------------------------------- */
data Cot;   set TT.cotisation_h;   run;
data Ind;   set TT.individu_h;     run;
data Contr; set TT.contrat_h;      run;
data Prest; set TT.presta_agr_h_f; run;


/* ---------------------------------------------------------------------------
   2. Suppression des doublons exacts
      Tri sur toutes les variables : les lignes identiques deviennent
      adjacentes et NODUPRECS les supprime.
      Attendu : 7 doublons dans contrat_h, 153 dans individu_h.
   --------------------------------------------------------------------------- */
proc sort data=Contr out=ContrU noduprecs dupout=Contr_dup;
  by _all_;
run;

proc sort data=Ind out=IndU noduprecs dupout=Ind_dup;
  by _all_;
run;

/* Nombre de doublons supprimes, garde en macro-variables pour les controles */
proc sql noprint;
  select count(*) into :nb_dup_contrat trimmed from Contr_dup;
  select count(*) into :nb_dup_individu trimmed from Ind_dup;
quit;
%put NOTE: doublons contrat_h = &nb_dup_contrat, doublons individu_h = &nb_dup_individu;


/* ---------------------------------------------------------------------------
   3. Exposition : une ligne = contrat x individu x garantie x mois
      Le mois est ramene au 1er du mois pour servir de cle de jointure.
   --------------------------------------------------------------------------- */
data Expo;
  set ContrU;
  Mois = intnx('month', datepart(dateDebutSituation), 0, 'B');
  format Mois monyy7.;
  keep idContrat idIndividu codeGarantie qualite Mois;
run;

/* Controle : une seule ligne par cle (attendu : 46 857 = 46 857) */
title "Controle : unicite de la cle d'exposition";
proc sql;
  select count(*) as Nb_lignes,
         count(distinct catx('|', idContrat, idIndividu, codeGarantie, Mois)) as Nb_cles
  from Expo;
quit;


/* ---------------------------------------------------------------------------
   4. Individus : sexe et date de naissance a chaque mois
      Apres suppression des doublons il reste une seule fiche par individu
      et par mois. DUPOUT permet de verifier que le NODUPKEY ne supprime rien.
   --------------------------------------------------------------------------- */
data Indiv;
  set IndU;
  Mois      = intnx('month', datepart(dateDebutSituation), 0, 'B');
  Date_nais = datepart(dateNaissance);
  format Mois monyy7. Date_nais ddmmyy10.;
  keep idIndividu Mois sexe Date_nais;
run;

proc sort data=Indiv nodupkey dupout=Indiv_dup;
  by idIndividu Mois;
run;

title "Controle : fiches individu en double sur le meme mois (attendu : 0)";
proc sql;
  select count(*) as Nb_doublons_individu_mois from Indiv_dup;
quit;


/* ---------------------------------------------------------------------------
   5. Prestations agregees par mois de soins
      presta_agr_h est deja agregee (une ligne = contrat x individu x garantie
      x type de soins x mois de soins x mois de reglement) : nb_lignes compte
      des lignes de prestations, pas des actes ni des sejours.
      prest_ecr = prestations ecretees a 5 000 euros par ligne (sensibilite).
   --------------------------------------------------------------------------- */
data Presta;
  set Prest;
  Mois = mdy(mois_soins, 1, annee_soins);
  format Mois monyy7.;
run;

proc sql;
  create table PrestAgg as
  select idContrat, idIndividu, codeGarantie, Mois,
         sum(Mtt_TOTALB)            as prest_m,
         sum(min(Mtt_TOTALB, 5000)) as prest_ecr,
         count(*)                   as nb_lignes
  from Presta
  group by idContrat, idIndividu, codeGarantie, Mois;
quit;


/* ---------------------------------------------------------------------------
   6. Cotisations nettes par mois d'appel
      On somme toutes les natures comptables. Les lignes APPAN et ANAPP
      s'annulent sur chaque cle (au centime pres) : le net est egal aux APPEL.
      Les cotisations appelees sont utilisees comme proxy des cotisations
      acquises (appel mensuel rattache au mois de couverture).
   --------------------------------------------------------------------------- */
data CotM;
  set Cot;
  Mois = intnx('month', datepart(moisAppel), 0, 'B');
  format Mois monyy7.;
run;

proc sql;
  create table CotAgg as
  select idContrat, idIndividu, codeGarantie, Mois,
         sum(montantMensuDetail) as cot_m
  from CotM
  group by idContrat, idIndividu, codeGarantie, Mois;
quit;

/* Controle : APPAN + ANAPP = 0 sur chaque cle (attendu : 0 cle en ecart) */
title "Controle : cles ou APPAN et ANAPP ne s'annulent pas (attendu : 0)";
proc sql;
  select count(*) as Nb_cles_en_ecart
  from (select idContrat, idIndividu, codeGarantie, Mois,
               sum(case when codeNatureComptable in ('APPAN','ANAPP')
                        then montantMensuDetail else 0 end) as Solde
        from CotM
        group by idContrat, idIndividu, codeGarantie, Mois)
  where abs(Solde) > 0.005;
quit;


/* ---------------------------------------------------------------------------
   7. Panel : exposition completee par individu, prestations et cotisations
      Left join depuis l'exposition : aucun mois couvert n'est perdu,
      les mois sans soin ou sans cotisation recoivent 0.
   --------------------------------------------------------------------------- */
proc sql;
  create table Panel as
  select e.idContrat, e.idIndividu, e.codeGarantie, e.qualite, e.Mois,
         i.sexe, i.Date_nais,
         coalesce(p.prest_m, 0)   as prest_m,
         coalesce(p.prest_ecr, 0) as prest_ecr,
         coalesce(p.nb_lignes, 0) as nb_lignes,
         coalesce(c.cot_m, 0)     as cot_m
  from Expo as e
       left join Indiv    as i
         on  e.idIndividu = i.idIndividu and e.Mois = i.Mois
       left join PrestAgg as p
         on  e.idContrat = p.idContrat and e.idIndividu = p.idIndividu
         and e.codeGarantie = p.codeGarantie and e.Mois = p.Mois
       left join CotAgg   as c
         on  e.idContrat = c.idContrat and e.idIndividu = c.idIndividu
         and e.codeGarantie = c.codeGarantie and e.Mois = c.Mois;
quit;


/* ---------------------------------------------------------------------------
   8. Variables d'analyse
      Age a la fin du mois d'observation : evite 29 ages negatifs
      (nouveau-nes dans le mois si on prenait le 1er du mois).
      Sinistre = prestation strictement positive sur le mois-garantie.
   --------------------------------------------------------------------------- */
data Analyse;
  set Panel;
  length Tranche $5;
  Annee    = year(Mois);
  Fin_mois = intnx('month', Mois, 0, 'E');
  Age      = floor(yrdif(Date_nais, Fin_mois, 'AGE'));
  if missing(Age)   then Tranche = 'NA';
  else if Age <= 17 then Tranche = '0-17';
  else if Age <= 29 then Tranche = '18-29';
  else if Age <= 44 then Tranche = '30-44';
  else if Age <= 59 then Tranche = '45-59';
  else if Age <= 69 then Tranche = '60-69';
  else                   Tranche = '70+';
  Sinistre = (prest_m > 0);
  format Fin_mois ddmmyy10.;
run;


/* ---------------------------------------------------------------------------
   9. Base finale : garanties sans cotisation et cohorte d'entree
      - Sans_cot = 1 si la garantie n'a aucune cotisation sur toute la periode
        (32 codes : cotisation portee ailleurs ou hors extraction).
      - Nouveau_2019 = 1 si le premier mois d'exposition de l'assure est en 2019.
      - Anciennete = nombre de mois depuis ce premier mois.
   --------------------------------------------------------------------------- */
proc sql;
  create table Gar_sans_cot as
  select codeGarantie
  from Analyse
  group by codeGarantie
  having sum(cot_m) <= 0;

  create table Entree as
  select idIndividu, min(Mois) as Mois_entree format=monyy7.
  from Analyse
  group by idIndividu;

  create table Base as
  select a.*,
         e.Mois_entree,
         case when e.Mois_entree >= '01JAN2019'd then 1 else 0 end as Nouveau_2019,
         intck('month', e.Mois_entree, a.Mois)                   as Anciennete,
         case when g.codeGarantie is not null then 1 else 0 end  as Sans_cot
  from Analyse as a
       inner join Entree as e
         on a.idIndividu = e.idIndividu
       left join Gar_sans_cot as g
         on a.codeGarantie = g.codeGarantie;
quit;


/* ---------------------------------------------------------------------------
   10. Controles de jointure et de qualite
       Une ligne par controle, exportee dans l'onglet Controles.
   --------------------------------------------------------------------------- */
proc sql;
  create table Controles as
  select 'Doublons exacts supprimes dans contrat_h' as Controle length=70,
         count(*) as Valeur
  from Contr_dup
  union all
  select 'Doublons exacts supprimes dans individu_h', count(*) from Ind_dup
  union all
  select 'Lignes d''exposition (mois-garantie)', count(*) from Base
  union all
  select 'Individu-mois distincts', count(*)
  from (select distinct idIndividu, Mois from Base)
  union all
  select 'Lignes du panel sans individu (sexe manquant)', sum(missing(sexe)) from Base
  union all
  select 'Prestations source (euros)', round(sum(Mtt_TOTALB), 0.01) from Prest
  union all
  select 'Prestations rattachees au panel (euros)', round(sum(prest_m), 0.01) from Base
  union all
  select 'Part des prestations rattachees', round(sum(prest_m) / (select sum(Mtt_TOTALB) from Prest), 0.0001) from Base
  union all
  select 'Cotisations nettes source (euros)', round(sum(montantMensuDetail), 0.01) from Cot
  union all
  select 'Cotisations rattachees au panel (euros)', round(sum(cot_m), 0.01) from Base
  union all
  select 'Part des cotisations rattachees', round(sum(cot_m) / (select sum(montantMensuDetail) from Cot), 0.0001) from Base
  union all
  select 'Ages negatifs au 1er du mois (corriges par age fin de mois)',
         sum(Date_nais > Mois) from Base
  union all
  select 'Lignes avec age superieur a 100 ans (conservees)', sum(Age > 100) from Base
  union all
  select 'Lignes de prestations a 0 euro', sum(Mtt_TOTALB = 0) from Prest
  union all
  select 'Lignes de prestations superieures a 5 000 euros', sum(Mtt_TOTALB > 5000) from Prest
  union all
  select 'Montant des prestations superieures a 5 000 euros',
         round(sum(case when Mtt_TOTALB > 5000 then Mtt_TOTALB else 0 end), 0.01) from Prest
  union all
  select 'Garanties sans aucune cotisation', count(*) from Gar_sans_cot
  union all
  select 'Mois-garantie a cotisation nulle', sum(cot_m = 0) from Base
  union all
  select 'Prestations sur des mois-garantie a cotisation nulle (euros)',
         round(sum(case when cot_m = 0 then prest_m else 0 end), 0.01) from Base;
quit;

title "Controles de jointure et de qualite";
proc print data=Controles noobs;
run;

title "Cotisations par nature comptable";
proc sql;
  create table Res_nature as
  select codeNatureComptable,
         count(*)                as Nb_lignes,
         sum(montantMensuDetail) as Montant
  from Cot
  group by codeNatureComptable;
quit;
proc print data=Res_nature noobs;
run;


/* ---------------------------------------------------------------------------
   11. Indicateurs globaux (base retenue)
   --------------------------------------------------------------------------- */
proc sql;
  create table Res_global as
  select count(*)                       as Mois_garantie,
         count(distinct idIndividu)     as Assures,
         count(distinct idContrat)      as Contrats,
         count(distinct codeGarantie)   as Garanties,
         round(sum(cot_m), 0.01)        as Cotisations,
         round(sum(prest_m), 0.01)      as Prestations,
         sum(prest_m) / sum(cot_m)      as SP,
         mean(Sinistre)                 as Taux_sinistres
  from Base;
quit;

title "Indicateurs globaux";
proc print data=Res_global noobs;
  format Cotisations Prestations comma15.2 SP Taux_sinistres percent8.2;
run;


/* ---------------------------------------------------------------------------
   12. Resultats par annee
       Le taux de sinistres est calcule sur deux bases :
       - par mois-garantie (une personne avec 2 garanties compte 2 fois)
       - par individu-mois (une personne = une exposition par mois)
   --------------------------------------------------------------------------- */
proc sql;
  create table IndMois as
  select idIndividu, Mois, Annee,
         sum(prest_m)            as prest,
         (calculated prest > 0) as Sinistre
  from Base
  group by idIndividu, Mois, Annee;

  create table Res_annee_mg as
  select b.Annee,
         count(*)                         as Mois_garantie,
         count(distinct b.idIndividu)     as Assures,
         round(sum(b.cot_m), 0.01)        as Cotisations,
         round(sum(b.prest_m), 0.01)      as Prestations,
         sum(b.prest_m) / sum(b.cot_m)    as SP,
         mean(b.Sinistre)                 as Taux_sin_mois_garantie,
         sum(b.prest_m) / count(*)        as Prest_par_mois_garantie,
         sum(b.cot_m) / count(*)          as Cot_par_mois_garantie,
         sum(b.prest_m) / sum(b.nb_lignes) as Montant_moyen_ligne
  from Base as b
  group by b.Annee;

  create table Res_annee_ind as
  select Annee, count(*) as Individu_mois, mean(Sinistre) as Taux_sin_individu_mois
  from IndMois
  group by Annee;

  create table Res_annee as
  select a.*, i.Individu_mois, i.Taux_sin_individu_mois
  from Res_annee_mg as a inner join Res_annee_ind as i
    on a.Annee = i.Annee
  order by a.Annee;
quit;

title "Resultats par annee";
proc print data=Res_annee noobs;
run;


/* ---------------------------------------------------------------------------
   13. Effet des nouveaux entrants 2019
       On compare, en 2019, les assures deja presents en 2018 et ceux qui
       entrent en 2019. Puis le S/P des nouveaux selon leur anciennete.
   --------------------------------------------------------------------------- */
proc sql;
  create table Res_cohorte as
  select Annee,
         case when Nouveau_2019 = 1 then 'Nouveaux 2019'
              else 'Presents en 2018' end    as Cohorte length=20,
         count(distinct idIndividu)          as Assures,
         count(*)                            as Mois_garantie,
         round(sum(cot_m), 0.01)             as Cotisations,
         round(sum(prest_m), 0.01)           as Prestations,
         sum(prest_m) / sum(cot_m)           as SP,
         sum(cot_m) / count(*)               as Cot_par_mois_garantie,
         sum(prest_m) / count(*)             as Prest_par_mois_garantie
  from Base
  group by 1, 2      /* 1 = Annee, 2 = Cohorte (alias calcule) */
  order by 1, 2;

  create table Res_anciennete as
  select case when Anciennete = 0 then '1. Mois d''entree'
              when Anciennete = 1 then '2. 1 mois'
              when Anciennete = 2 then '3. 2 mois'
              when Anciennete <= 5 then '4. 3 a 5 mois'
              else '5. 6 mois et plus' end   as Anciennete_cl length=20,
         count(*)                            as Mois_garantie,
         round(sum(cot_m), 0.01)             as Cotisations,
         round(sum(prest_m), 0.01)           as Prestations,
         sum(prest_m) / sum(cot_m)           as SP,
         mean(cot_m = 0)                     as Part_cot_nulle
  from Base
  where Nouveau_2019 = 1
  group by 1
  order by 1;

  /* Garanties choisies par les nouveaux entrants */
  create table Res_cohorte_garantie as
  select codeGarantie,
         count(*)                  as Mois_garantie_nouveaux,
         round(sum(cot_m), 0.01)   as Cotisations,
         round(sum(prest_m), 0.01) as Prestations,
         case when sum(cot_m) > 0 then sum(prest_m) / sum(cot_m) end as SP
  from Base
  where Nouveau_2019 = 1
  group by codeGarantie
  order by Mois_garantie_nouveaux desc;
quit;

/* Verification : les nouveaux sont-ils vraiment nouveaux ?
   - la date d'effet de leur garantie (dateEffetGar) doit tomber sur leur
     premier mois observe (sinon il leur manquerait des mois d'exposition).
     On ne prend pas la date d'effet du contrat : un conjoint ou un enfant
     ajoute a un contrat existant a une date de contrat plus ancienne ;
   - Part_cot_nulle (ci-dessus) doit rester stable selon l'anciennete
     (sinon le pic de S/P viendrait de cotisations manquantes). */
proc sql;
  create table Verif_entree as
  select count(*) as Nouveaux_2019,
         sum(intnx('month', datepart(c.Effet), 0, 'B') = e.Mois_entree)
           as Effet_au_mois_entree,
         calculated Effet_au_mois_entree / count(*) as Part format=percent8.1
  from Entree as e
       inner join (select idIndividu, min(dateEffetGar) as Effet
                   from ContrU group by idIndividu) as c
         on e.idIndividu = c.idIndividu
  where e.Mois_entree >= '01JAN2019'd;
quit;

title "Verification : date d'effet des nouveaux entrants";
proc print data=Verif_entree noobs; run;

title "S/P par annee et par cohorte d'entree";
proc print data=Res_cohorte noobs; run;
title "Nouveaux entrants 2019 : S/P selon l'anciennete";
proc print data=Res_anciennete noobs; run;


/* ---------------------------------------------------------------------------
   14. Profils : sexe, tranche d'age, qualite
       La macro cree une table Res_<nom> par variable.
       Frequence = lignes de prestations par mois-garantie,
       cout moyen = montant par ligne de prestation.
   --------------------------------------------------------------------------- */
%macro profil(var, nom);
  proc sql;
    create table Res_&nom as
    select &var,
           count(*)                           as Mois_garantie,
           round(sum(cot_m), 0.01)            as Cotisations,
           round(sum(prest_m), 0.01)          as Prestations,
           sum(prest_m) / sum(cot_m)          as SP,
           mean(Sinistre)                     as Taux_sinistres,
           sum(prest_m) / count(*)            as Prest_par_mois_garantie,
           sum(cot_m) / count(*)              as Cot_par_mois_garantie,
           sum(nb_lignes) / count(*)          as Frequence_lignes,
           sum(prest_m) / sum(nb_lignes)      as Montant_moyen_ligne
    from Base
    group by &var
    order by &var;
  quit;

  title "Resultats par &nom";
  proc print data=Res_&nom noobs;
  run;
%mend profil;

%profil(sexe,    sexe)
%profil(Tranche, age)
%profil(qualite, qualite)


/* ---------------------------------------------------------------------------
   15. Evolution mensuelle
   --------------------------------------------------------------------------- */
proc sql;
  create table Res_mois as
  select Mois format=monyy7.,
         count(*)                  as Mois_garantie,
         round(sum(cot_m), 0.01)   as Cotisations,
         round(sum(prest_m), 0.01) as Prestations,
         sum(prest_m) / sum(cot_m) as SP
  from Base
  group by Mois
  order by Mois;
quit;

title "Evolution mensuelle";
proc print data=Res_mois noobs;
run;


/* ---------------------------------------------------------------------------
   16. Prestations par type de soins (base complete presta_agr_h)
       Les montants sont aussi rapportes a l'exposition de chaque annee
       (mois-garantie) pour neutraliser la croissance du portefeuille,
       et separes entre assures presents en 2018 et nouveaux 2019.
   --------------------------------------------------------------------------- */
proc sql noprint;
  select count(*) into :expo_2018 trimmed from Base where Annee = 2018;
  select count(*) into :expo_2019 trimmed from Base where Annee = 2019;
quit;

proc sql;
  create table Res_soins as
  select p.TypeSoins,
         count(*)                                     as Nb_lignes,
         round(sum(p.Mtt_TOTALB), 0.01)               as Montant,
         sum(p.Mtt_TOTALB) / count(*)                 as Montant_moyen_ligne,
         sum(case when p.annee_soins = 2018 then 1 else 0 end)            as Lignes_2018,
         sum(case when p.annee_soins = 2019 then 1 else 0 end)            as Lignes_2019,
         sum(case when p.annee_soins = 2018 then p.Mtt_TOTALB else 0 end) as Montant_2018,
         sum(case when p.annee_soins = 2019 then p.Mtt_TOTALB else 0 end) as Montant_2019,
         calculated Montant_2018 / &expo_2018         as Par_mois_garantie_2018,
         calculated Montant_2019 / &expo_2019         as Par_mois_garantie_2019,
         sum(case when p.annee_soins = 2019 and e.Mois_entree >= '01JAN2019'd
                  then p.Mtt_TOTALB else 0 end)       as Montant_2019_nouveaux
  from Presta as p
       left join Entree as e
         on p.idIndividu = e.idIndividu
  group by p.TypeSoins
  order by Montant desc;
quit;

title "Prestations par type de soins";
proc print data=Res_soins noobs;
run;


/* ---------------------------------------------------------------------------
   17. Resultats par garantie et effet de composition
       Effectif_faible = moins de 5 000 euros de cotisations sur la periode.
       Res_garantie_annee : cotisation moyenne a garantie egale, pour
       separer un effet prix d'un effet de composition du portefeuille.
   --------------------------------------------------------------------------- */
proc sql;
  create table Res_garantie as
  select codeGarantie,
         count(*)                         as Mois_garantie,
         round(sum(cot_m), 0.01)          as Cotisations,
         round(sum(prest_m), 0.01)        as Prestations,
         case when sum(cot_m) > 0 then sum(prest_m) / sum(cot_m) end as SP,
         max(Sans_cot)                    as Sans_cotisation,
         (sum(cot_m) < 5000)              as Effectif_faible
  from Base
  group by codeGarantie
  order by Cotisations desc;

  create table Res_garantie_annee as
  select codeGarantie, Annee,
         count(*)                  as Mois_garantie,
         round(sum(cot_m), 0.01)   as Cotisations,
         round(sum(prest_m), 0.01) as Prestations,
         case when sum(cot_m) > 0 then sum(prest_m) / sum(cot_m) end as SP,
         sum(cot_m) / count(*)     as Cot_par_mois_garantie
  from Base
  where codeGarantie in ('ATP1', 'ATP2', 'ATP3', 'AT1+', 'ARH3')
  group by codeGarantie, Annee
  order by codeGarantie, Annee;
quit;

title "Resultats par garantie";
proc print data=Res_garantie(obs=15) noobs; run;
title "Cotisation moyenne a garantie egale";
proc print data=Res_garantie_annee noobs; run;


/* ---------------------------------------------------------------------------
   18. Concentration des prestations entre assures
       Assures tries par prestations decroissantes, puis part des prestations
       portee par les 1 %, 5 %, 10 % et 20 % les plus couteux.
       Convention : le top p % = les floor(p x nombre d'assures) premiers.
   --------------------------------------------------------------------------- */
proc sql;
  create table Ind_tot as
  select idIndividu, sum(prest_m) as Prest_ind
  from Base
  group by idIndividu;
quit;

proc sql noprint;
  select count(*)       into :nb_ind    trimmed from Ind_tot;
  /* format 32.2 : sans format, SAS arrondit la valeur a 8 chiffres */
  select sum(Prest_ind) format=32.2 into :tot_prest trimmed from Ind_tot;
quit;

proc sort data=Ind_tot out=Ind_rang;
  by descending Prest_ind;
run;

data Ind_rang;
  set Ind_rang;
  Rang = _N_;
  Cumul + Prest_ind;              /* somme cumulee (instruction somme) */
  Part_prest = Cumul / &tot_prest;
run;

%macro top(p);
  proc sql;
    insert into Res_concentration
    select "Top &p pourcent" as Groupe,
           count(*)                       as Nb_assures,
           sum(Prest_ind) / &tot_prest    as Part_prestations
    from Ind_rang
    where Rang <= floor(&p / 100 * &nb_ind);
  quit;
%mend top;

proc sql;
  create table Res_concentration
    (Groupe char(30), Nb_assures num, Part_prestations num);
quit;

%top(1)
%top(5)
%top(10)
%top(20)

proc sql;
  insert into Res_concentration
  select 'Assures sans aucune prestation', count(*), 0
  from Ind_tot
  where Prest_ind = 0;
quit;

/* Courbe de concentration : un point environ tous les 1 % d'assures
   (axe X = part cumulee des assures, axe Y = part cumulee des prestations) */
data Res_courbe;
  set Ind_rang end=fin;
  Part_assures = Rang / &nb_ind;
  if mod(Rang, ceil(&nb_ind / 100)) = 0 or fin then output;
  keep Rang Part_assures Part_prest;
run;

title "Concentration des prestations";
proc print data=Res_concentration noobs;
run;


/* ---------------------------------------------------------------------------
   19. Sensibilite du S/P
       Meme indicateur sur cinq perimetres pour tester la robustesse.
       Les conditions sont ecrites sans signe = (gt, eq) pour ne pas etre
       confondues avec un parametre nomme de la macro.
   --------------------------------------------------------------------------- */
%macro sp(libelle, condition, montant=prest_m);
  select "&libelle" as Perimetre length=50,
         sum(case when Annee = 2018 then &montant else 0 end)
           / sum(case when Annee = 2018 then cot_m else 0 end) as SP_2018,
         sum(case when Annee = 2019 then &montant else 0 end)
           / sum(case when Annee = 2019 then cot_m else 0 end) as SP_2019,
         sum(&montant) / sum(cot_m)                            as SP_total
  from Base
  where &condition
%mend sp;

proc sql;
  create table Res_sensibilite as
  %sp(Base retenue, Annee gt 0)
  union all
  %sp(Hors garanties sans cotisation, Sans_cot eq 0)
  union all
  %sp(Assures presents en 2018 uniquement, Nouveau_2019 eq 0)
  union all
  %sp(Hors garantie ATP1, codeGarantie ne 'ATP1')
  union all
  %sp(Prestations ecretees a 5 000 euros par ligne, Annee gt 0, montant=prest_ecr);
quit;

title "Sensibilite du S/P";
proc print data=Res_sensibilite noobs;
  format SP_2018 SP_2019 SP_total percent8.1;
run;
title;


/* ---------------------------------------------------------------------------
   20. Chain ladder : faut-il une provision pour sinistres tardifs ?
       Nos S/P sont en date de soins, avec des reglements connus jusqu'en
       2024 : les soins 2018-2019 sont presque tous regles. Pour le
       verifier, on construit le triangle des reglements :
         - en ligne, le trimestre de soins (Trim_soins : 0 = 2018-T1 ... 7 = 2019-T4)
         - en colonne, le delai de reglement en trimestres (Delai_trim :
           0 = regle dans le trimestre des soins ... 8 = 8 trimestres ou plus).
       On se place au 31/12/2019 : on ne connait que les cases ou
       Trim_soins + Delai_trim <= 7. Le chain ladder estime la charge finale
       de chaque trimestre, que l'on compare ensuite a la charge reelle.
       On garde les lignes rattachees a un mois-garantie de la base
       (Mtt_rattache), pour avoir les memes prestations que dans le S/P.
   --------------------------------------------------------------------------- */
proc sql;
  create table Presta_tri as
  select p.*,
         cats(p.annee_soins, '-T', floor((p.mois_soins - 1) / 3) + 1)
                                                    as Trimestre_soins length=7,
         (p.annee_soins - 2018) * 4 + floor((p.mois_soins - 1) / 3) as Trim_soins,
         min((p.annee_Regl - 2018) * 4 + floor((p.mois_Regl - 1) / 3)
             - calculated Trim_soins, 8)             as Delai_trim,
         case when b.idContrat is not null then p.Mtt_TOTALB else 0 end
                                                    as Mtt_rattache
  from Presta as p
       left join (select distinct idContrat, idIndividu, codeGarantie, Mois
                  from Base) as b
         on  p.idContrat = b.idContrat and p.idIndividu = b.idIndividu
         and p.codeGarantie = b.codeGarantie and p.Mois = b.Mois;
quit;

/* Triangle complet (toutes les cases, 0 si aucun reglement) */
proc sql;
  create table Grille as
  select t.Trim_soins, t.Trimestre_soins, d.Delai_trim
  from (select distinct Trim_soins, Trimestre_soins from Presta_tri) as t,
       (select distinct Delai_trim from Presta_tri) as d;

  create table Triangle as
  select g.Trim_soins, g.Trimestre_soins, g.Delai_trim,
         coalesce(sum(p.Mtt_rattache), 0) as Montant
  from Grille as g
       left join Presta_tri as p
         on p.Trim_soins = g.Trim_soins and p.Delai_trim = g.Delai_trim
  group by g.Trim_soins, g.Trimestre_soins, g.Delai_trim
  order by g.Trim_soins, g.Delai_trim;
quit;

/* Montants cumules et cases connues au 31/12/2019 */
data Triangle;
  set Triangle;
  by Trim_soins;
  if first.Trim_soins then Cumul = 0;
  Cumul + Montant;
  Connu_fin_2019 = (Trim_soins + Delai_trim <= 7);
run;

/* Facteurs de developpement : sur les trimestres ou les deux delais sont
   connus fin 2019, f(j) = somme des cumuls a j+1 / somme des cumuls a j */
proc sql;
  create table Res_facteurs as
  select a.Delai_trim as Delai,
         sum(b.Cumul) / sum(a.Cumul) as Facteur
  from Triangle as a inner join Triangle as b
    on  a.Trim_soins = b.Trim_soins
    and b.Delai_trim = a.Delai_trim + 1
  where b.Connu_fin_2019 = 1
  group by a.Delai_trim
  order by a.Delai_trim;
quit;

/* Projection : dernier cumul connu fin 2019 x produit des facteurs restants
   (aucun facteur au-dela du delai 7 : on suppose qu'il n'y a plus rien a payer) */
proc sql;
  create table Res_chain_ladder as
  select c.Trimestre_soins, c.Trim_soins,
         floor(c.Trim_soins / 4) + 2018          as Annee,
         c.Cumul                                 as Paye_fin_2019,
         coalesce((select exp(sum(log(f.Facteur))) from Res_facteurs as f
                   where f.Delai >= c.Delai_trim), 1) as Facteur_restant,
         calculated Paye_fin_2019 * calculated Facteur_restant as Charge_estimee,
         calculated Charge_estimee - calculated Paye_fin_2019    as Provision,
         (select sum(t.Montant) from Triangle as t
          where t.Trim_soins = c.Trim_soins)      as Charge_reelle,
         calculated Charge_estimee / calculated Charge_reelle - 1 as Ecart_reel
  from Triangle as c
  where c.Trim_soins + c.Delai_trim = 7
  order by c.Trim_soins;

  create table Res_chain_ladder_annee as
  select c.Annee,
         sum(c.Paye_fin_2019)  as Paye_fin_2019,
         sum(c.Charge_estimee) as Charge_estimee,
         sum(c.Provision)      as Provision,
         sum(c.Charge_reelle)  as Charge_reelle,
         a.Cotisations,
         calculated Paye_fin_2019  / a.Cotisations as SP_sans_provision,
         calculated Charge_estimee / a.Cotisations as SP_chain_ladder,
         calculated Charge_reelle  / a.Cotisations as SP_reel
  from Res_chain_ladder as c inner join Res_annee as a
    on c.Annee = a.Annee
  group by c.Annee, a.Cotisations
  order by c.Annee;
quit;

title "Chain ladder : facteurs de developpement (vus au 31/12/2019)";
proc print data=Res_facteurs noobs;
run;
title "Chain ladder au 31/12/2019 compare aux reglements reels";
proc print data=Res_chain_ladder noobs;
  format Paye_fin_2019 Charge_estimee Provision Charge_reelle comma12.2
         Ecart_reel percent8.1;
run;
proc print data=Res_chain_ladder_annee noobs;
  format Paye_fin_2019 Charge_estimee Provision Charge_reelle Cotisations comma12.2
         SP_sans_provision SP_chain_ladder SP_reel percent8.1;
run;
title;


/* ---------------------------------------------------------------------------
   21. Export Excel : un onglet par table de resultats + la base d'analyse
       + la table des prestations.
       Dans Excel, chaque graphique est construit a partir d'un tableau
       croise dynamique (TCD) fait sur l'onglet Base (ou Prestations pour
       le type de soins). Les tables de resultats SAS servent a controler
       les chiffres obtenus dans Excel (ecart nul).
       Attention : PROC EXPORT reecrit les onglets du meme nom. Ne pas
       exporter dans un fichier qui contient deja les graphiques (garder
       une copie a part).
   --------------------------------------------------------------------------- */

/* Base d'analyse : une ligne = un assure x une garantie x un mois.
   Exportee pour faire les tableaux croises dynamiques dans Excel.
   Le RETAIN fixe l'ordre des colonnes. */
data Base_export;
  retain idContrat idIndividu codeGarantie qualite sexe Age Tranche
         Mois Annee Mois_entree Nouveau_2019 Anciennete Sans_cot
         cot_m prest_m prest_ecr nb_lignes Sinistre;
  set Base;
  keep idContrat idIndividu codeGarantie qualite sexe Age Tranche
       Mois Annee Mois_entree Nouveau_2019 Anciennete Sans_cot
       cot_m prest_m prest_ecr nb_lignes Sinistre;
run;

/* Table des prestations (une ligne par ligne de presta_agr_h_f) :
   le type de soins n'est pas dans la base mois-garantie, le TCD du
   graphique sur les types de soins est fait sur cet onglet, ainsi que
   le triangle du chain ladder (section 20). */
data Presta_export;
  retain idContrat idIndividu codeGarantie TypeSoins annee_soins
         mois_soins annee_Regl mois_Regl Trimestre_soins Delai_trim
         Mtt_TOTALB Mtt_rattache;
  set Presta_tri;
  keep idContrat idIndividu codeGarantie TypeSoins annee_soins
       mois_soins annee_Regl mois_Regl Trimestre_soins Delai_trim
       Mtt_TOTALB Mtt_rattache;
run;

%macro exporte(table, onglet);
  proc export data=&table
              outfile="&fichier_excel"
              dbms=xlsx replace;
    sheet="&onglet";
  run;
%mend exporte;

%exporte(Controles,            Controles)
%exporte(Res_nature,           Nature_comptable)
%exporte(Res_global,           Global)
%exporte(Res_annee,            Par_annee)
%exporte(Res_cohorte,          Cohortes)
%exporte(Res_anciennete,       Anciennete_nouveaux)
%exporte(Res_cohorte_garantie, Garanties_nouveaux)
%exporte(Res_sexe,             Sexe)
%exporte(Res_age,              Age)
%exporte(Res_qualite,          Qualite)
%exporte(Res_mois,             Mensuel)
%exporte(Res_soins,            Type_soins)
%exporte(Res_garantie,         Garanties)
%exporte(Res_garantie_annee,   Garanties_par_annee)
%exporte(Res_concentration,    Concentration)
%exporte(Res_courbe,           Courbe_concentration)
%exporte(Verif_entree,         Verif_entree)
%exporte(Res_sensibilite,      Sensibilite)
%exporte(Res_facteurs,         Facteurs_dev)
%exporte(Res_chain_ladder,     Chain_ladder)
%exporte(Res_chain_ladder_annee, Chain_ladder_annee)
%exporte(Base_export,          Base)
%exporte(Presta_export,        Prestations)

/* Fin du programme */
