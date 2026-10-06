/**
 * ============================================================
 * FOYER RURAL D'ISNEAUVILLE — Inscriptions 2026/2027
 * Google Apps Script — v8.82
 * ============================================================
 * Corrections v8.3 :
 *  1. genererFacturePDF   → recalcul propre depuis tarif_brut
 *                           remise 15% affichée en ligne dédiée (montant exact)
 *                           FNSMR affiché par membre unique
 *                           FFTT affiché si > 0
 *                           Avoirs / Pass'sport / ANCV affichés
 *                           totalReel = Σbruts − remise + FNSMR + FFTT
 *  2. validerPaiementSheet → emailRows construits APRÈS recalculerRemiseDossier
 *                            + relit la col 28 corrigée depuis le Sheet
 *  3. envoyerEmailAdmin    → total recalculé depuis les lignes (pas col 30 stale)
 * ============================================================
 */

// ── ENVIRONNEMENT : PRODUCTION ou TEST ────────────────────────
// Le même Code.gs sert aux deux projets Apps Script. Dans la COPIE de test, définir
// dans Paramètres du projet › Propriétés du script :
//   FRI_MODE      = test
//   FRI_SHEET_ID  = ID de la copie de test de la feuille Google
//   FRI_EMAIL_TEST = adresse qui reçoit TOUS les emails du site de test (sinon : votre compte)
// En mode test : feuille de test, dossiers Drive préfixés « TEST - », emails détournés
// vers FRI_EMAIL_TEST et HelloAsso en bac à sable (api.helloasso-sandbox.com).
var SHEET_ID_PRODUCTION = '1KMVBYHReafOYgwolaeCb_yWHfKBdN05gskwHijmYtp4';
var MODE_TEST = PropertiesService.getScriptProperties().getProperty('FRI_MODE') === 'test';
var ERREUR_CONFIG = ''; // message renvoyé à toutes les requêtes si le serveur est mal configuré
if (MODE_TEST) {
  var _sheetTest = PropertiesService.getScriptProperties().getProperty('FRI_SHEET_ID') || '';
  if (!_sheetTest || _sheetTest === SHEET_ID_PRODUCTION) {
    // Garde-fou : un projet de test ne doit jamais écrire dans la feuille de production
    ERREUR_CONFIG = 'Serveur de test mal configuré : définir FRI_SHEET_ID (copie de test de la feuille) dans les propriétés du script.';
    _sheetTest = '';
  }
}

// ── À CONFIGURER ──────────────────────────────────────────────
var SHEET_ID         = MODE_TEST ? _sheetTest : SHEET_ID_PRODUCTION;
var SHEET_LICENCES_2526 = '1mIhP_kTUhMs7Mqs9urmNOJKSkuZzpAgIdztDAxwL5zE'; // Google Sheet licences FFTT 25-26
var SHEET_ADHERENTS_2526 = '1N6brFQdVIWMUDboJ82Neb6TvwByfH4RuqUDWb44t-jA'; // Google Sheet adhérents FRI 25-26
var EMAIL_ADMIN      = 'fri.inscri@gmail.com';
var EMAIL_TRESORIER  = 'bloquet.t@orange.fr';
var NOM_ASSO         = "Foyer Rural d\'Isneauville";
var DEFAULT_CAPACITY = 20;
var HELLOASSO_URL    = MODE_TEST
  ? (PropertiesService.getScriptProperties().getProperty('HA_URL_TEST') || 'https://www.helloasso-sandbox.com')
  : 'https://www.helloasso.com/associations/foyer-rural-d-isneauville/paiements/reglement-adhesion-fri';
// ── Secrets — stockés dans Propriétés du script (jamais en clair) ──
// Pour les définir : Apps Script → Paramètres du projet → Propriétés du script
// Clés attendues : FRI_SECRET_TOKEN, HA_CLIENT_ID, HA_CLIENT_SECRET
var _props           = PropertiesService.getScriptProperties();
var FRI_SECRET_TOKEN = _props.getProperty('FRI_SECRET_TOKEN') || '';
if (!FRI_SECRET_TOKEN) Logger.log('⚠️ FRI_SECRET_TOKEN absent des propriétés du script !');

// ── HelloAsso Checkout API ────────────────────────────────────
var HA_CLIENT_ID     = _props.getProperty('HA_CLIENT_ID')     || '';
var HA_CLIENT_SECRET = _props.getProperty('HA_CLIENT_SECRET') || '';
// En test : bac à sable HelloAsso (identifiants et organisation du compte sandbox)
var HA_HOST          = MODE_TEST ? 'https://api.helloasso-sandbox.com' : 'https://api.helloasso.com';
var HA_ORG_SLUG      = (MODE_TEST && _props.getProperty('HA_ORG_SLUG')) || 'foyer-rural-d-isneauville';
var HA_API_BASE      = HA_HOST + '/v5';
var HA_RETURN_URL    = ScriptApp.getService ? (function(){ try { return ScriptApp.getService().getUrl(); } catch(e){ return ''; } })() : '';
// ────────────────────────────────────────────────────────────

// ── Aiguillage Drive et emails selon l'environnement ─────────
// Les dossiers Drive sont retrouvés par leur nom : sans préfixe, un projet de test
// lancé avec le même compte Google écrirait (et supprimerait !) dans les dossiers réels.
function nomDossierDrive(nom) { return MODE_TEST ? 'TEST - ' + nom : nom; }
function dossiersDriveParNom(nom) { return DriveApp.getFoldersByName(nomDossierDrive(nom)); }
function creerDossierDrive(nom) { return DriveApp.createFolder(nomDossierDrive(nom)); }

// Remplace GmailApp.sendEmail (mêmes paramètres). En test, tout part vers FRI_EMAIL_TEST.
function envoyerEmail(destinataire, sujet, corps, options) {
  if (!MODE_TEST) return GmailApp.sendEmail(destinataire, sujet, corps, options);
  var dest = _props.getProperty('FRI_EMAIL_TEST') || Session.getEffectiveUser().getEmail();
  var opts = {};
  Object.keys(options || {}).forEach(function(k) { if (k !== 'cc' && k !== 'bcc') opts[k] = options[k]; });
  opts.name = '[TEST] ' + (opts.name || NOM_ASSO);
  Logger.log('MODE TEST — email pour ' + destinataire + ' détourné vers ' + dest);
  return GmailApp.sendEmail(dest, '[TEST → ' + destinataire + '] ' + sujet, corps, opts);
}

var SHEET_INSCRIPTIONS  = 'Inscriptions';
var SHEET_RECAPITULATIF = 'Recapitulatif';
var SHEET_PLACES        = 'Places';
var SHEET_CHEQUE_1      = 'Cheques 1';
var SHEET_CHEQUE_2      = 'Cheques 2';
var SHEET_CHEQUE_3      = 'Cheques 3';
var SHEET_AVOIRS        = 'Avoirs générés';   // onglet avoirs générés (suppression activité)
var SHEET_REMBOURSEMENTS = 'Remboursements';  // onglet remboursements (positionné juste après Avoirs générés)
var SHEET_AVOIRS_UTILISES = 'Avoirs utilisés';    // onglet avoirs consommés au paiement
var SHEET_AIDE_ANCV     = 'Aides ANCV';
var SHEET_AIDE_ATOUT    = 'Aides Atout Normandie';
var SHEET_AIDE_PASS_J   = 'Aides Pass Jeunes';
var SHEET_AIDE_PASS_S   = 'Aides Pass Sport';
var SHEET_ESPECES       = 'Espèces';
var SHEET_HELLOASSO     = 'HelloAsso';

// ============================================================
// STRUCTURE DES 39 COLONNES
// 1  Numéro de licence   2  Civilité            3  Nom
// 4  Prénom              5  Date naissance       6  Appartement
// 7  Bâtiment            8  Voie                 9  Lieu-dit
// 10 Code postal         11 Ville               12 Cedex
// 13 Code pays           14 Tél. fixe           15 Tél. portable
// 16 Email               17 Commentaire         18 Tél. 2
// 19 Email 2             20 N° Dossier          21 Date inscription
// 22 Statut paiement     23 Activité            24 Jour
// 25 Heure               26 Lieu                27 Animateur
// 28 Tarif EUR           29 FNSMR EUR           30 Total famille EUR
// 31 Mode paiement       32 Avoir               33 QS Santé
// 34 Pass / Aide         35 Sexe                36 ID Activité
// 37 Responsable         38 Statut inscription  39 Licence FFTT (€)
// ============================================================

var logoBase64 = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAQDAwMDAgQDAwMEBAQFBgoGBgUFBgwICQcKDgwPDg4MDQ0PERYTDxAVEQ0NExoTFRcYGRkZDxIbHRsYHRYYGRj/2wBDAQQEBAYFBgsGBgsYEA0QGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBj/wAARCACCAHcDASIAAhEBAxEB/8QAHQABAAEEAwEAAAAAAAAAAAAAAAgFBgcJAQMEAv/EAEgQAAECBQIDBAUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwCn9KEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgBCEIAQhCAEIQgD/2Q==';

// ============================================================
// GÉNÉRATION SEMAINIER PDF
// ============================================================

function genererFacturePDF(rows, modeLabel, titreOverride) {
  try {
    var r0       = rows[0];
    var code     = r0.code_dossier || 'FRI-????';
    var dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
    var prenom   = r0.responsable_prenom || '';
    var nom      = r0.responsable_nom    || '';
    var adresse  = (r0.adresse||'')+(r0.cp?', '+r0.cp:'')+(r0.ville?' '+r0.ville:'');

    var parMembre={};
    rows.forEach(function(r){
      var k=(r.membre_prenom||'')+' '+(r.membre_nom||'');
      if(!parMembre[k])parMembre[k]=[];
      parMembre[k].push(r);
    });

    // Commune pour eligibilite TT (PING*)
    var _villeF = r0.ville || '';
    var _communeF = (r0.commune === 'isno' || r0.commune === 'Isneauville')
      ? 'Isneauville' : communeFromVille(_villeF);
    var isIsneauvilleF = (_communeF === 'Isneauville');
    function estEligibleRemise(pid) {
      if (!pid) return false;
      if (pid === 'MNOME/S10') return false;  // Marche Nordique
      if (pid.indexOf('PING') >= 0) return isIsneauvilleF;  // TT : Isneauville uniquement
      return true;
    }

    // ── v8.3 : Calcul depuis tarif_brut ──────────────────────
    // Déterminer si le dossier a la remise 15% (≥3 activités éligibles)
    var seenPids = {}, nbEligibles = 0;
    rows.forEach(function(r) {
      var pid = r.activite_id || '';
      var isWaitPdf = String(r.statut_inscription||'').toLowerCase().indexOf('attente') >= 0;
      if (!isWaitPdf && !seenPids[pid]) { seenPids[pid] = true; if (estEligibleRemise(pid)) nbEligibles++; }
    });
    nbEligibles += Number((rows[0] && rows[0].nb_elig_existants) || 0);
    var aRemise15 = nbEligibles >= 3;

    // Membres uniques → FNSMR (sauf adhésion déjà réglée : ajout d'activités à un dossier existant)
    var nbMembres = Object.keys(parMembre).filter(function(k) {
      return (parMembre[k] || []).some(function(r) { return !(r && r.adhesion_deja_reglee); });
    }).length;
    var totalFnsmr = nbMembres * 15;

    // FFTT par membre (dédupliqué)
    var ffttMembres = {};
    rows.forEach(function(r) {
      var prix = parseFloat(r.fftt_price) || 0;
      if (prix > 0) {
        var k = (r.membre_prenom||'')+' '+(r.membre_nom||'');
        if (!ffttMembres[k]) ffttMembres[k] = prix;
      }
    });
    var totalFFTT = Object.keys(ffttMembres).reduce(function(acc,k){return acc+ffttMembres[k];}, 0);

    // Calcul tarif brut et remise par ligne
    var totalBrut     = 0; // somme des tarifs bruts (avant remise)
    var totalRemise   = 0; // montant total de la remise
    var totalNet      = 0; // somme des tarifs après remise

    rows.forEach(function(r) {
      var tarifBrut = parseFloat(r.tarif_brut) || parseFloat(r.tarif) || 0;
      var noteTarifTot = String(r.note_tarif||'');
      var isForfaitPDF = noteTarifTot && noteTarifTot.toLowerCase().indexOf('forfait') >= 0;
      var elig = !isForfaitPDF && aRemise15 && estEligibleRemise(r.activite_id || '');
      var remiseLigne = elig ? Math.round(tarifBrut * 0.15 * 100) / 100 : 0;
      var tarifFinal  = Math.round((tarifBrut - remiseLigne) * 100) / 100;
      totalBrut   += tarifBrut;
      totalRemise += remiseLigne;
      totalNet    += tarifFinal;
    });
    totalBrut   = Math.round(totalBrut   * 100) / 100;
    totalRemise = Math.round(totalRemise * 100) / 100;
    totalNet    = Math.round(totalNet    * 100) / 100;

    // Déductions — utiliser _calcFinancier pour cohérence avec les emails
    var _fPDF = _calcFinancier(rows);
    var deducPassSport  = _fPDF.deducPassSport;
    var deducAncv       = _fPDF.deducAncv;
    var deducAvoir      = _fPDF.deducAvoir;
    var deducPassJeunes = _fPDF.deducPassJeunes;
    var deducAtout      = _fPDF.deducAtout;
    var totalDeductions = _fPDF.totalDeductions;

    // Total général = tarifs nets + FNSMR + FFTT
    var totalGeneral = Math.round((totalNet + totalFnsmr + totalFFTT) * 100) / 100;
    // Solde après déductions (utiliser _fPDF pour cohérence)
    var soldeNet = _fPDF.solde;

    var passAideLabel = String(r0.pass_aide || '');

    // ── Création du document ──────────────────────────────────
    var docTitle='FACTURE_FRI_'+code+'_'+dateJour.replace(/\//g,'-');
    var doc=DocumentApp.create(docTitle);
    var body=doc.getBody();
    body.setMarginTop(28).setMarginBottom(28).setMarginLeft(42).setMarginRight(42);

    // Logo FRI en haut du document
    try {
      var logoFact = getLogoBlob();
      if (logoFact) {
        var logoImg = body.appendImage(logoFact);
        logoImg.setWidth(70).setHeight(70);
        body.getParagraphs()[body.getParagraphs().length-1]
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('');
      }
    } catch(eLogo) { Logger.log('Logo facture KO: ' + eLogo); }

    // En-tête
    var titreDoc=body.appendParagraph('FOYER RURAL D\'ISNEAUVILLE');
    titreDoc.editAsText().setFontSize(16).setBold(true).setForegroundColor('#1b5e20');
    titreDoc.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var sousTitreDoc=body.appendParagraph('Saison 2026 / 2027');
    sousTitreDoc.editAsText().setFontSize(11).setForegroundColor('#555555');
    sousTitreDoc.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var adresseAsso=body.appendParagraph('Salle des fêtes — Place A. Cramilly — 76230 Isneauville');
    adresseAsso.editAsText().setFontSize(9).setForegroundColor('#888888');
    adresseAsso.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var contactAsso=body.appendParagraph('frisneauville@orange.fr  |  02.35.59.01.01  |  www.frisneauville.fr');
    contactAsso.editAsText().setFontSize(9).setForegroundColor('#888888');
    contactAsso.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendHorizontalRule();

    var titreFACTURE=body.appendParagraph('REÇU DE PAIEMENT');
    titreFACTURE.editAsText().setFontSize(20).setBold(true).setForegroundColor('#1a2e22');
    titreFACTURE.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');
    var infoDossier=body.appendParagraph('N° Dossier : '+code+'     |     Date : '+dateJour+'     |     Réglé par : '+modeLabel);
    infoDossier.editAsText().setFontSize(9).setBold(true).setForegroundColor('#1a2e22');
    infoDossier.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');
    var titreAdh=body.appendParagraph('ADHÉRENT / RESPONSABLE LÉGAL');
    titreAdh.editAsText().setFontSize(9).setBold(true).setForegroundColor('#888888');
    var nomAdh=body.appendParagraph(prenom+' '+nom);
    nomAdh.editAsText().setFontSize(13).setBold(true).setForegroundColor('#1a2e22');
    if(adresse.trim()!==','){var adrAdh=body.appendParagraph(adresse);adrAdh.editAsText().setFontSize(10).setForegroundColor('#444444');}
    if(r0.email1){var emailAdh=body.appendParagraph('Email : '+r0.email1);emailAdh.editAsText().setFontSize(10).setForegroundColor('#444444');}
    body.appendParagraph('');

    // ── Tableau des activités ─────────────────────────────────
    var table=body.appendTable();
    var rowEntete=table.appendTableRow();

    // Colonnes : Membre | Activité | Jour / Heure | Lieu | Tarif brut | Remise | Tarif net
    // Largeurs en points (A4 utilisable ~451 pts marges 42pt)
    // Avec remise : 75+120+90+38+38+38+52 = 451
    // Sans remise : 75+160+110+52+54 = 451
    var colWidthsRemise = [75, 110, 90, 38, 42, 42, 54];
    var colWidthsSimple = [75, 160, 110, 52, 54];
    var entetes = aRemise15
      ? ['Membre','Activité','Jour / Heure','Lieu','T. Brut','Remise','T. Net']
      : ['Membre','Activité','Jour / Heure','Lieu','Tarif'];
    var colWidths = aRemise15 ? colWidthsRemise : colWidthsSimple;
    var nbCols = entetes.length;
    // Formatage montant : supprime les .00 inutiles (ex: 145.00 → 145, 43.50 → 43.50)
    function fmt(n) {
      var s = n.toFixed(2);
      return s.endsWith('.00') ? s.slice(0,-3) + ' €' : s + ' €';
    }

    entetes.forEach(function(txt, ci) {
      var cell = rowEntete.appendTableCell(txt);
      cell.editAsText().setFontSize(8).setBold(true).setForegroundColor('#ffffff');
      cell.setBackgroundColor('#1b5e20');
      cell.setWidth(colWidths[ci]);
      if (ci >= 4) cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    });

    var ligneIndex = 0;
    Object.keys(parMembre).forEach(function(membre) {
      parMembre[membre].forEach(function(r) {
        var tarifBrut  = parseFloat(r.tarif_brut) || parseFloat(r.tarif) || 0;
        var noteTarifF   = String(r.note_tarif||'');
        var isForfaitLigne = noteTarifF && noteTarifF.toLowerCase().indexOf('forfait') >= 0;
        var elig       = !isForfaitLigne && aRemise15 && estEligibleRemise(r.activite_id || '');
        var remiseLigne = elig ? Math.round(tarifBrut * 0.15 * 100) / 100 : 0;
        var tarifFinal  = Math.round((tarifBrut - remiseLigne) * 100) / 100;
        var bgLigne    = elig ? '#d8f3dc' : (ligneIndex % 2 === 0 ? '#f0f7f3' : '#ffffff');
        // Affichage colonnes remise
        var tarifBrutLabel  = fmt(tarifBrut);
        var remiseLabel     = elig ? '- '+fmt(remiseLigne).replace(' €','')+' €'
                            : isForfaitLigne ? noteTarifF
                            : '—';
        var tarifFinalLabel = isForfaitLigne && tarifBrut === 0 ? 'Inclus' : fmt(tarifFinal);
        var cellValues = aRemise15 || noteTarifF
          ? [membre, (r.activite||'').replace(/\n/g,' — '), (r.jour||'')+'\n'+(r.heure||''), r.lieu||'',
             tarifBrutLabel, remiseLabel, tarifFinalLabel]
          : [membre, (r.activite||'').replace(/\n/g,' — '), (r.jour||'')+'\n'+(r.heure||''), r.lieu||'',
             tarifBrut.toFixed(2)+' €'];

        var tr=table.appendTableRow();
        cellValues.forEach(function(txt, ci) {
          var cell=tr.appendTableCell(txt);
          cell.editAsText().setFontSize(8).setForegroundColor(elig ? '#1a5c2a' : '#222222');
          cell.setBackgroundColor(bgLigne);
          cell.setWidth(colWidths[ci]);
          if (ci >= 4) cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
        });
        ligneIndex++;
      });
    });

    // ── Ligne sous-total activités brut ──────────────────────
    if (aRemise15) {
      var trBrut = table.appendTableRow();
      ['','Sous-total activités (brut)','','','','',fmt(totalBrut)].forEach(function(txt,ci){
        var cell=trBrut.appendTableCell(txt);
        cell.editAsText().setFontSize(8).setBold(true).setForegroundColor('#1a2e22');
        cell.setBackgroundColor('#f0f7f3');
        cell.setWidth(colWidths[ci]);
        if(ci>=4) cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      });

      // Ligne remise famille totale
      var trRemise=table.appendTableRow();
      ['','🎉 Remise famille -15% ('+nbEligibles+' activité(s))','','','',' - '+fmt(totalRemise).replace(' €','')+' €',''].forEach(function(txt,ci){
        var cell=trRemise.appendTableCell(txt);
        cell.editAsText().setFontSize(8).setBold(true).setForegroundColor('#2d6a4f');
        cell.setBackgroundColor('#d8f3dc');
        cell.setWidth(colWidths[ci]);
        if(ci>=4) cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      });

      // Ligne sous-total après remise
      var trNet=table.appendTableRow();
      ['','Sous-total activités (après remise)','','','','',fmt(totalNet)].forEach(function(txt,ci){
        var cell=trNet.appendTableCell(txt);
        cell.editAsText().setFontSize(8).setBold(true).setForegroundColor('#1b5e20');
        cell.setBackgroundColor('#e8f5e9');
        cell.setWidth(colWidths[ci]);
        if(ci>=4) cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      });
    } else {
      // Sans remise : juste le sous-total
      var trSousTotal=table.appendTableRow();
      ['','Sous-total activités','','','',fmt(totalBrut)].slice(0,nbCols).forEach(function(txt,ci){
        var cell=trSousTotal.appendTableCell(txt);
        cell.editAsText().setFontSize(8).setBold(true).setForegroundColor('#1a2e22');
        cell.setBackgroundColor('#f0f7f3');
        cell.setWidth(colWidths[ci]);
        if(ci===nbCols-1) cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      });
    }

    // ── Ligne FNSMR (par membre) ──────────────────────────────
    var trFnsmr=table.appendTableRow();
    var fnsmrCells = aRemise15
      ? ['','Adhésion FNSMR ('+nbMembres+' pers. × 15,00 €)','','','','',totalFnsmr.toFixed(2)+' €']
      : ['','Adhésion FNSMR ('+nbMembres+' pers. × 15,00 €)','','',fmt(totalFnsmr)];
    fnsmrCells.forEach(function(txt,ci){
      var cell=trFnsmr.appendTableCell(txt);
      cell.editAsText().setFontSize(8).setItalic(true).setForegroundColor('#555555');
      cell.setBackgroundColor('#f5f5f5');
      cell.setWidth(colWidths[ci]);
      if(ci===fnsmrCells.length-1)
        cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    });

    // ── Lignes FFTT (par membre) ──────────────────────────────
    Object.keys(ffttMembres).forEach(function(membreNom) {
      var prixFFTT = ffttMembres[membreNom];
      var trFFTT=table.appendTableRow();
      var ffttCells = aRemise15
        ? ['','🏓 Licence FFTT — '+membreNom,'','','','',fmt(prixFFTT)]
        : ['','🏓 Licence FFTT — '+membreNom,'','',fmt(prixFFTT)];
      ffttCells.forEach(function(txt,ci){
        var cell=trFFTT.appendTableCell(txt);
        cell.editAsText().setFontSize(8).setBold(true).setForegroundColor('#e65100');
        cell.setBackgroundColor('#fff3e0');
        cell.setWidth(colWidths[ci]);
        if(ci===ffttCells.length-1)
          cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      });
    });

    // ── Ligne Total général (avant déductions) ────────────────
    var trTotalGen=table.appendTableRow();
    var totalGenCells = aRemise15
      ? ['','TOTAL GÉNÉRAL (activités + FNSMR + FFTT)','','','','',fmt(totalGeneral)]
      : ['','TOTAL GÉNÉRAL','','',fmt(totalGeneral)];
    totalGenCells.forEach(function(txt,ci){
      var cell=trTotalGen.appendTableCell(txt);
      cell.editAsText().setFontSize(9).setBold(true).setForegroundColor('#1b5e20');
      cell.setBackgroundColor('#d8f3dc');
      cell.setWidth(colWidths[ci]);
      if(ci===totalGenCells.length-1)
        cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    });

    // ── Les déductions et le solde sont affichés uniquement dans la section
    //    "DÉTAIL DU RÈGLEMENT" ci-dessous (pas de doublon dans le tableau activités)

    // ── Récapitulatif financier détaillé ─────────────────────
    body.appendParagraph('');
    var titreRegl = body.appendParagraph('DÉTAIL DU RÈGLEMENT');
    titreRegl.editAsText().setFontSize(10).setBold(true).setForegroundColor('#1a2e22');
    titreRegl.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

    var tableRegl = body.appendTable();
    tableRegl.setBorderWidth(0.5);

    function ligneRegl(label, montant, couleur, bold) {
      var tr = tableRegl.appendTableRow();
      var c1 = tr.appendTableCell(label);
      var c2 = tr.appendTableCell(montant);
      c1.editAsText().setFontSize(9).setForegroundColor(couleur||'#333333').setBold(bold||false);
      c2.editAsText().setFontSize(9).setForegroundColor(couleur||'#333333').setBold(bold||false);
      c1.setWidth(320); c2.setWidth(120);
      c2.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      c1; c1; c1;
      c2; c2; c2;
      return tr;
    }

    // Total activités + FNSMR
    ligneRegl('Total activités + FNSMR + Licence FFTT', fmt(totalGeneral), '#1a2e22', true);

    // Remise famille si applicable
    if (aRemise15 && totalRemise > 0) {
      ligneRegl('  Dont remise famille -15% ('+nbEligibles+' activité(s))', '- '+fmt(totalRemise).replace(' €','')+' €', '#2d6a4f', false);
    }

    // Déductions aides/avoirs/ANCV
    if (deducPassSport > 0) ligneRegl('  Déduction Pass\'sport État', '- '+fmt(deducPassSport).replace(' €','')+' €', '#2d6a4f', false);
    if (deducAncv > 0)      ligneRegl('  Déduction Coupon ANCV', '- '+fmt(deducAncv).replace(' €','')+' €', '#2d6a4f', false);
    if (deducAvoir > 0)     ligneRegl('  Avoir dossier', '- '+fmt(deducAvoir).replace(' €','')+' €', '#2d6a4f', false);

    // Pass'jeunes 76 / Handipass'sport
    if (deducPassJeunes > 0) ligneRegl('  Pass\'jeunes 76 / Handipass\'sport', '- '+fmt(deducPassJeunes).replace(' €','')+' €', '#6a1b9a', false);
    // Atout Normandie
    if (deducAtout > 0)      ligneRegl('  Atout Normandie', '- '+fmt(deducAtout).replace(' €','')+' €', '#e65100', false);

    // Séparateur et solde final (uniquement si des déductions existent)
    if (totalDeductions > 0) {
      // Ligne séparateur sur 2 colonnes (tableRegl a 2 colonnes c1=320 c2=120)
      var trSep2 = tableRegl.appendTableRow();
      var sep1 = trSep2.appendTableCell(''); sep1.setWidth(320);
      var sep2 = trSep2.appendTableCell(''); sep2.setWidth(120);
      sep1.editAsText().setFontSize(1); sep1.setBackgroundColor('#cccccc');
      sep2.editAsText().setFontSize(1); sep2.setBackgroundColor('#cccccc');
      var rSolde = ligneRegl('SOLDE À RÉGLER', fmt(soldeNet), '#1565c0', true);
      rSolde.getCell(0).setBackgroundColor('#e8f4fd');
      rSolde.getCell(1).setBackgroundColor('#e8f4fd');
    }

    // ── Séparateur avant mode de paiement ──
    var trSep3 = tableRegl.appendTableRow();
    var sep3a = trSep3.appendTableCell(''); sep3a.setWidth(320);
    var sep3b = trSep3.appendTableCell(''); sep3b.setWidth(120);
    sep3a.editAsText().setFontSize(1); sep3a.setBackgroundColor('#e0e0e0');
    sep3b.editAsText().setFontSize(1); sep3b.setBackgroundColor('#e0e0e0');

    // Mode de paiement avec détail
    var modeLabelMap = {helloasso:'HelloAsso', cheque:'Cheque', cheque3:'Cheques 3x', especes:'Especes', ancv:'Coupon ANCV'};
    var modeLabelDisplay = modeLabelMap[modeLabel] || modeLabel || 'Non precis';

    // Détail chèques si disponible
    var detailPaie = modeLabelDisplay;
    if ((modeLabel === 'cheque' || modeLabel === 'cheque3') && r0.cheque_1_num) {
      detailPaie += ' n' + r0.cheque_1_num;
      if (r0.cheque_1_banque) detailPaie += ' (' + r0.cheque_1_banque + ')';
      if (r0.cheque_2_num)    detailPaie += ' + n' + r0.cheque_2_num;
      if (r0.cheque_3_num)    detailPaie += ' + n' + r0.cheque_3_num;
    }
    var rMode = ligneRegl('Mode de paiement', detailPaie, '#333333', true);
    rMode.getCell(0).setBackgroundColor('#f5f5f5');
    rMode.getCell(1).setBackgroundColor('#f5f5f5');
    var rDate = ligneRegl('Date de paiement', dateJour, '#555555', false);
    rDate.getCell(0).setBackgroundColor('#f5f5f5');
    rDate.getCell(1).setBackgroundColor('#f5f5f5');
    // Montant réglé (rappel final)
    var montantRegle = totalDeductions > 0 ? soldeNet : totalGeneral;
    var rMontant = ligneRegl('Montant regle', fmt(montantRegle).replace(' €',' EUR'), '#1a2e22', true);
    rMontant.getCell(0).setBackgroundColor('#d8f3dc');
    rMontant.getCell(1).setBackgroundColor('#d8f3dc');

    body.appendParagraph('');
    body.appendHorizontalRule();
    var pied=body.appendParagraph('Ce reçu vous est adressé par le '+NOM_ASSO+' — Saison 2026/2027. Conservez ce document.');
    pied.editAsText().setFontSize(8).setForegroundColor('#aaaaaa').setItalic(true);
    pied.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    doc.saveAndClose();
    var docFile=DriveApp.getFileById(doc.getId());
    var pdfBlob=docFile.getAs('application/pdf');
    pdfBlob.setName('Recu_FRI_'+code+'_'+dateJour.replace(/\//g,'-')+'.pdf');
    try{
      var dossierFact=creerDossierSecurise('4-Factures acquittées');
      var pdfBlobDrive=docFile.getAs('application/pdf');
      pdfBlobDrive.setName(pdfBlob.getName());
      dossierFact.createFile(pdfBlobDrive);
    }catch(driveErr){Logger.log('Sauvegarde Drive facture KO : '+driveErr.toString());}
    docFile.setTrashed(true);
    Logger.log('Facture PDF v8.3 générée : '+pdfBlob.getName()
      +' | brut:'+totalBrut+' remise:'+totalRemise+' net:'+totalNet
      +' FNSMR:'+totalFnsmr+' FFTT:'+totalFFTT+' total:'+totalGeneral);
    return pdfBlob;
  }catch(err){Logger.log('ERREUR genererFacturePDF : '+err.toString());return null;}
}

// ============================================================
// HELLOASSO — inchangé (v8.2)
// ============================================================

function helloassoGetToken() {
  var props = PropertiesService.getScriptProperties();
  var cached = props.getProperty('HA_TOKEN');
  var expiry  = parseInt(props.getProperty('HA_TOKEN_EXPIRY') || '0');
  if (cached && Date.now() < expiry) return cached;
  var maxRetries = 3;
  var response, httpCode;
  for (var attempt = 0; attempt < maxRetries; attempt++) {
    if (attempt > 0) {
      var delaiMs = Math.pow(2, attempt) * 2000;
      Logger.log('HelloAsso token 429 — attente ' + (delaiMs/1000) + 's avant retry ' + attempt + '/' + maxRetries);
      Utilities.sleep(delaiMs);
    }
    response = UrlFetchApp.fetch(HA_HOST + '/oauth2/token', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: 'grant_type=client_credentials&client_id='+encodeURIComponent(HA_CLIENT_ID)+'&client_secret='+encodeURIComponent(HA_CLIENT_SECRET),
      muteHttpExceptions: true
    });
    httpCode = response.getResponseCode();
    if (httpCode !== 429) break;
    Logger.log('HelloAsso token 429 (tentative ' + (attempt+1) + ')');
  }
  if (httpCode !== 200) {
    Logger.log('HelloAsso token KO ('+httpCode+'): '+response.getContentText());
    return null;
  }
  var data = JSON.parse(response.getContentText());
  var token = data.access_token;
  var expiresIn = (parseInt(data.expires_in) || 1800) - 60;
  props.setProperty('HA_TOKEN', token);
  props.setProperty('HA_TOKEN_EXPIRY', String(Date.now() + expiresIn * 1000));
  Logger.log('✅ HelloAsso token obtenu — expire dans ' + expiresIn + 's');
  return token;
}

function helloassoCreerLienPaiement(totalCentimes, prenom, nom, email, codeDossier, description, coordonnees) {
  try {
    var token = helloassoGetToken();
    if (!token) { Logger.log('HelloAsso token absent — lien non généré'); return null; }
    function nettoyerNom(s) { return String(s || '').replace(/[^a-zA-ZàâäéèêëîïôöùûüçÀÂÄÉÈÊËÎÏÔÖÙÛÜÇ'\- ]/g, '').trim().substring(0, 50); }
    function nettoyerTel(s) { return String(s || '').replace(/[^0-9+\-\s().]/g, '').trim().substring(0, 20); }
    function nettoyerAdresse(s) { return String(s || '').replace(/[^\w\sàâäéèêëîïôöùûüçÀÂÄÉÈÊËÎÏÔÖÙÛÜÇ'\-,.]/g, '').trim().substring(0, 100); }
    var prenomPropre = nettoyerNom(prenom) || 'Adherent';
    var nomPropre    = nettoyerNom(nom)    || 'FRI';
    if (prenomPropre.toLowerCase() === nomPropre.toLowerCase()) nomPropre = nomPropre + '.';
    var scriptUrl = '';
    try { scriptUrl = ScriptApp.getService().getUrl(); } catch(e) {}
    var returnUrl = scriptUrl ? scriptUrl+'?action=haPaiement&code='+encodeURIComponent(codeDossier)+'&status=ok'    : HELLOASSO_URL;
    var errorUrl  = scriptUrl ? scriptUrl+'?action=haPaiement&code='+encodeURIComponent(codeDossier)+'&status=error' : HELLOASSO_URL;
    var backUrl   = scriptUrl ? scriptUrl+'?action=haPaiement&code='+encodeURIComponent(codeDossier)+'&status=back'  : HELLOASSO_URL;
    var coord = coordonnees || {};
    var payer = { firstName: prenomPropre, lastName: nomPropre, email: email || '' };
    var adresseStr = nettoyerAdresse(coord.adresse || '');
    var cpStr      = String(coord.cp || '').replace(/[^0-9]/g, '').substring(0, 10);
    var villeStr   = nettoyerNom(coord.ville || '');
    var telStr     = nettoyerTel(coord.tel || '');
    if (adresseStr) payer.address = adresseStr;
    if (cpStr)      payer.zipCode = cpStr;
    if (villeStr)   payer.city   = villeStr;
    if (telStr)     payer.phone  = telStr;
    payer.country = 'FRA';
    var body = {
      totalAmount: totalCentimes, initialAmount: totalCentimes,
      itemName: (description || 'Inscription FRI 2026/2027').substring(0, 255),
      backUrl: backUrl, errorUrl: errorUrl, returnUrl: returnUrl,
      containsDonation: false, payer: payer,
      metadata: { dossier: codeDossier, saison: '2026-2027', asso: HA_ORG_SLUG }
    };
    var url = HA_API_BASE + '/organizations/' + HA_ORG_SLUG + '/checkout-intents';
    var response = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': 'Bearer '+token },
      payload: JSON.stringify(body), muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    var text = response.getContentText();
    if (code !== 200 && code !== 201) { Logger.log('❌ HelloAsso checkout KO: '+text); return null; }
    var result = JSON.parse(text);
    return result.redirectUrl || null;
  } catch(err) { Logger.log('helloassoCreerLienPaiement ERREUR : '+err.toString()); return null; }
}

// ── Lien de paiement HelloAsso à mettre dans un email ──
// Un paiement HelloAsso (checkout) expire en ~15 min : l'email contient donc un lien vers
// ce script (?action=payer&t=…), qui crée un paiement neuf au moment du clic.
// Le montant est stocké côté serveur (jamais dans l'URL).
function creerLienPaiementEmail(code, montant, email, prenom, nom, description) {
  try {
    var url = ScriptApp.getService().getUrl();
    if (!url || !(montant > 0)) return '';
    var jeton = Utilities.getUuid().replace(/-/g, '');
    var props = PropertiesService.getScriptProperties();
    props.setProperty('pay_' + jeton, JSON.stringify({
      code: code, montant: Math.round(montant * 100) / 100, email: email || '',
      prenom: prenom || '', nom: nom || '', description: description || '', cree: Date.now()
    }));
    // Ménage : jetons de plus de 90 jours
    try {
      var limite = Date.now() - 90 * 24 * 3600 * 1000, toutes = props.getProperties();
      Object.keys(toutes).forEach(function(k) {
        if (k.indexOf('pay_') !== 0) return;
        try { if (JSON.parse(toutes[k]).cree < limite) props.deleteProperty(k); } catch(e) { props.deleteProperty(k); }
      });
    } catch(e) {}
    return url + '?action=payer&t=' + jeton;
  } catch(e) { Logger.log('creerLienPaiementEmail KO : ' + e); return ''; }
}

// Page ouverte depuis le lien de l'email : crée le paiement HelloAsso et propose le bouton
function pagePaiementEmail(jeton) {
  var info = null;
  try { info = JSON.parse(PropertiesService.getScriptProperties().getProperty('pay_' + String(jeton || '').replace(/[^a-f0-9]/gi, '')) || 'null'); } catch(e) {}
  var page = function(contenu) {
    return HtmlService.createHtmlOutput('<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
      + '<body style="font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f0f4f2">'
      + '<div style="background:white;border-radius:14px;padding:32px 24px;max-width:480px;width:100%;text-align:center;box-shadow:0 10px 40px rgba(0,0,0,.1)">'
      + '<div style="font-size:40px">🏡</div><h2 style="color:#1a2e22">' + NOM_ASSO + '</h2>' + contenu + '</div></body></html>')
      .setTitle('Règlement — ' + NOM_ASSO);
  };
  if (!info) return page('<p>Ce lien de paiement n\'est plus valable.</p><p style="color:#666;font-size:14px">Contactez-nous : frisneauville@orange.fr — 02.35.59.01.01</p>');
  var montantCts = Math.round(Number(info.montant) * 100);
  var lienHA = helloassoCreerLienPaiement(montantCts, info.prenom, info.nom, info.email, info.code,
    (info.description || 'Règlement') + ' — ' + info.code);
  var lien = lienHA || HELLOASSO_URL;
  return page('<p>Dossier <strong>' + info.code + '</strong></p>'
    + '<p style="font-size:22px;font-weight:900;color:#1a2e22;margin:8px 0">' + Number(info.montant).toFixed(2) + ' €</p>'
    + '<a href="' + lien + '" target="_top" style="display:inline-block;background:#4c40cf;color:white;text-decoration:none;font-weight:700;padding:14px 26px;border-radius:10px;margin:10px 0">Payer avec HelloAsso</a>'
    + (lienHA ? '<p style="font-size:12px;color:#666">Formulaire pré-rempli à votre nom, paiement sécurisé.</p>'
              : '<p style="font-size:12px;color:#666">Indiquez le montant ci-dessus et votre code dossier <strong>' + info.code + '</strong> sur le formulaire HelloAsso.</p>'));
}

function helloassoBoutonHtml(lienCheckout, montantLabel, noteOverride) {
  var lien = lienCheckout || HELLOASSO_URL;
  var svgLock = '<svg width="9" height="10" viewBox="0 0 11 12" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M3.875 3V4.5H7.625V3C7.625 1.969 6.781 1.125 5.75 1.125C4.695 1.125 3.875 1.969 3.875 3ZM2.75 4.5V3C2.75 1.359 4.086 0 5.75 0C7.391 0 8.75 1.359 8.75 3V4.5H9.5C10.32 4.5 11 5.18 11 6V10.5C11 11.344 10.32 12 9.5 12H2C1.156 12 0.5 11.344 0.5 10.5V6C0.5 5.18 1.156 4.5 2 4.5H2.75ZM1.625 6V10.5C1.625 10.711 1.789 10.875 2 10.875H9.5C9.688 10.875 9.875 10.711 9.875 10.5V6C9.875 5.813 9.688 5.625 9.5 5.625H2C1.789 5.625 1.625 5.813 1.625 6Z" fill="#2e2f5e"/></svg>';
  var montantHtml = montantLabel ? '<div style="font-size:18px;font-weight:900;color:#1a2e22;margin:6px 0 10px;">' + montantLabel + '</div>' : '';
  var noteHtml = noteOverride
    ? '<div style="font-size:11px;color:#666;margin-top:4px;">' + noteOverride + '</div>'
    : lienCheckout
    ? '<div style="font-size:11px;color:#666;margin-top:4px;">🔒 Formulaire pré-rempli à votre nom — valable 15 min</div>'
    : '<div style="font-size:11px;color:#666;margin-top:4px;">Ou lors des permanences : mardis 16h30-18h30</div>';
  var bouton = '<a href="'+lien+'" style="text-decoration:none;display:inline-block;">'
    +'<div style="display:inline-flex;align-items:stretch;border-radius:8px;overflow:hidden;font-family:\'Open Sans\',Arial,sans-serif;">'
    +'<div style="background:#ffffff;border:1px solid #4c40cf;border-top-left-radius:8px;border-bottom-left-radius:8px;padding:10px 14px;display:flex;align-items:center;">'
    +'<img src="https://api.helloasso.com/v5/img/logo-ha.svg" alt="HelloAsso" style="width:28px;height:auto;display:block;" /></div>'
    +'<div style="background:#4c40cf;border:1px solid #4c40cf;border-top-right-radius:8px;border-bottom-right-radius:8px;padding:0 18px;display:flex;align-items:center;gap:6px;color:white;font-size:15px;font-weight:800;">'
    +'<span>Payer avec HelloAsso</span></div></div></a>';
  var secureBar = '<div style="display:flex;align-items:center;gap:6px;padding:6px 14px;font-size:11px;font-weight:600;color:#2e2f5e;font-family:\'Open Sans\',Arial,sans-serif;">'
    +svgLock+'<span>Paiement sécurisé</span>'
    +'<img src="https://helloassodocumentsprod.blob.core.windows.net/public-documents/bouton_payer_avec_helloasso/logo-visa.svg" alt="Visa" style="height:14px;" />'
    +'<img src="https://helloassodocumentsprod.blob.core.windows.net/public-documents/bouton_payer_avec_helloasso/logo-mastercard.svg" alt="Mastercard" style="height:14px;" />'
    +'<img src="https://helloassodocumentsprod.blob.core.windows.net/public-documents/bouton_payer_avec_helloasso/logo-cb.svg" alt="CB" style="height:14px;" />'
    +'<img src="https://helloassodocumentsprod.blob.core.windows.net/public-documents/bouton_payer_avec_helloasso/logo-pci.svg" alt="PCI" style="height:14px;" /></div>';
  return '<div style="background:#f8f9ff;border:1.5px solid #4c40cf;border-radius:12px;padding:16px 20px;text-align:center;margin:16px 0;">'
    +(lienCheckout ? '<div style="font-size:12px;color:#4c40cf;font-weight:600;margin-bottom:6px;">🎯 Formulaire pré-rempli à votre nom</div>' : '')
    +montantHtml+'<div style="display:flex;justify-content:center;">'+bouton+'</div>'
    +'<div style="display:flex;justify-content:center;">'+secureBar+'</div>'+noteHtml+'</div>';
}

// ============================================================
// doGet — POINT D'ENTRÉE PRINCIPAL
// ============================================================

// ══════════════════════════════════════════════════════════════
// HELLOASSO API — Configuration
// Renseigner dans les propriétés du script :
//   HA_CLIENT_ID     → votre client_id HA
//   HA_CLIENT_SECRET → votre client_secret HA
// ══════════════════════════════════════════════════════════════
// (HA_API_BASE et HA_ORG_SLUG : définis en tête de fichier, production ou bac à sable)
var HA_FORM_SLUG = 'paiement-de-l-adhesion-au-foyer-rural-isneauville-2';

function getHAClientId() {
  var v = PropertiesService.getScriptProperties().getProperty('HA_CLIENT_ID') || '';
  Logger.log('HA_CLIENT_ID lu: "' + v.substring(0,8) + '..." (len:' + v.length + ')');
  return v;
}
function getHAClientSecret() {
  var v = PropertiesService.getScriptProperties().getProperty('HA_CLIENT_SECRET') || '';
  Logger.log('HA_CLIENT_SECRET lu: len=' + v.length);
  return v;
}

// ── Obtenir un access_token HA (stocké en cache 25 min) ──
function getHAAccessToken() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('HA_ACCESS_TOKEN');
  if (cached) { Logger.log('Token HA depuis cache'); return cached; }

  // Fallback : token récent dans les propriétés (< 25 min)
  var props = PropertiesService.getScriptProperties();
  var storedToken = props.getProperty('HA_TOKEN_CACHE');
  var storedTs    = Number(props.getProperty('HA_TOKEN_TS') || '0');
  if (storedToken && (Date.now() - storedTs) < 1500000) {
    Logger.log('Token HA depuis propriétés (fallback)');
    cache.put('HA_ACCESS_TOKEN', storedToken, 1500);
    return storedToken;
  }

  var clientId     = getHAClientId();
  var clientSecret = getHAClientSecret();
  if (!clientId)     clientId     = PropertiesService.getUserProperties().getProperty('HA_CLIENT_ID')     || '';
  if (!clientSecret) clientSecret = PropertiesService.getUserProperties().getProperty('HA_CLIENT_SECRET') || '';
  Logger.log('Auth HA — clientId len:' + clientId.length + ' secret len:' + clientSecret.length);
  if (!clientId || !clientSecret) throw new Error('Credentials HA manquants (HA_CLIENT_ID / HA_CLIENT_SECRET)');

  // Retry avec backoff en cas de rate limit 429
  var waits = [0, 3000, 7000];
  for (var i = 0; i < waits.length; i++) {
    if (waits[i] > 0) { Logger.log('Rate limit HA — attente ' + waits[i] + 'ms'); Utilities.sleep(waits[i]); }

    var resp = UrlFetchApp.fetch(HA_HOST + '/oauth2/token', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: 'grant_type=client_credentials&client_id=' + encodeURIComponent(clientId)
             + '&client_secret=' + encodeURIComponent(clientSecret),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code === 429) { Logger.log('Auth HA 429 tentative ' + (i+1)); continue; }
    if (code !== 200) throw new Error('Auth HA KO (' + code + '): ' + resp.getContentText().substring(0, 200));

    var token = JSON.parse(resp.getContentText()).access_token;
    cache.put('HA_ACCESS_TOKEN', token, 1500);
    props.setProperty('HA_TOKEN_CACHE', token);
    props.setProperty('HA_TOKEN_TS', String(Date.now()));
    Logger.log('✅ Token HA obtenu (tentative ' + (i+1) + ')');
    return token;
  }
  throw new Error('Auth HA KO : rate limit 429 persistant après 3 tentatives');
}

// ── Créer un checkout intent HA ──
// Retourne { redirectUrl, checkoutIntentId }
function creerCheckoutIntentHA(params) {
  // params : { totalAmount (centimes), itemName, prenom, nom, email,
  //            codeDossier, returnUrl, backUrl, errorUrl }
  var token = getHAAccessToken();

  var total = params.totalAmount; // en centimes

  // ── Calcul des versements ──
  var body;
  if (params.echelonne) {
    // Paiement en 3 fois : 40% + 30% + 30% (solde)
    // Centimes arrondis — la somme doit être exactement = total
    var v1 = Math.round(total * 0.40); // 40% à l\'inscription
    var v2 = Math.round(total * 0.30); // 30% à 3 mois
    var v3 = total - v1 - v2;          // solde à 6 mois (évite les erreurs d'arrondi)

    // Dates fixes de prélèvement saison 2026/2027
    var fmt = function(d) {
      return d.getFullYear() + '-'
        + String(d.getMonth()+1).padStart(2,'0') + '-'
        + String(d.getDate()).padStart(2,'0') + 'T00:00:00.000Z';
    };
    var date2 = new Date('2026-10-15'); // 2e versement : 15/10/2026
    var date3 = new Date('2027-01-15'); // 3e versement : 15/01/2027

    body = {
      totalAmount:      total,
      initialAmount:    v1,
      itemName:         (params.itemName || 'Adhésion FRI 2026-2027 — paiement 3x').substring(0, 250),
      backUrl:          params.backUrl  || '',
      errorUrl:         params.errorUrl || '',
      returnUrl:        params.returnUrl || '',
      containsDonation: false,
      terms: [
        { amount: v2, date: fmt(date2) },
        { amount: v3, date: fmt(date3) }
      ],
      payer: {
        firstName: params.prenom || '',
        lastName:  params.nom    || '',
        email:     params.email  || ''
      },
      metadata: {
        codeDossier:  params.codeDossier || '',
        source:       'FRI-Inscriptions',
        paiement3x:   'true'
      }
    };
    Logger.log('Checkout HA 3x (40/30/30) — v1:'+v1+' v2:'+v2+' v3:'+v3+' total:'+total+' cts');
  } else {
    // Paiement en une fois
    body = {
      totalAmount:      total,
      initialAmount:    total,
      itemName:         (params.itemName || 'Adhésion FRI 2026-2027').substring(0, 250),
      backUrl:          params.backUrl  || '',
      errorUrl:         params.errorUrl || '',
      returnUrl:        params.returnUrl || '',
      containsDonation: false,
      payer: {
        firstName: params.prenom || '',
        lastName:  params.nom    || '',
        email:     params.email  || ''
      },
      metadata: {
        codeDossier: params.codeDossier || '',
        source:      'FRI-Inscriptions'
      }
    };
  }

  var url = HA_API_BASE + '/organizations/' + HA_ORG_SLUG + '/checkout-intents';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var text = resp.getContentText();
  Logger.log('Checkout HA ' + code + ' : ' + text.substring(0, 200));

  if (code !== 200 && code !== 201) {
    throw new Error('Checkout HA KO (' + code + '): ' + text.substring(0, 300));
  }

  var result = JSON.parse(text);
  return { redirectUrl: result.redirectUrl, checkoutIntentId: result.id };
}

// ── Logo FRI (base64 JPEG) ──
function getLogoBlob() {
  try {
    var LOGO_B64 = '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';
    var decoded = Utilities.base64Decode(LOGO_B64);
    return Utilities.newBlob(decoded, 'image/jpeg', 'Logo_FRI.jpg');
  } catch(e) { Logger.log('Logo blob KO: ' + e); return null; }
}

// URL data pour les emails HTML (inline)
function getLogoDataUrl() {
  return 'data:image/jpeg;base64,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';
}


function doGet(e) {
  if(ERREUR_CONFIG) return repondreAvecCb({status:'error',message:ERREUR_CONFIG},null,null,(e&&e.parameter&&e.parameter.callback)||null);
  try{var allParams=JSON.stringify((e&&e.parameter)?e.parameter:{});Logger.log('doGet params: '+allParams.substring(0,300));}catch(logErr){}
  var params=(e&&e.parameter)?e.parameter:{};
  if(params.action==='fds'&&params.token)return traiterActionFinDeSaison(params.token);
  if(params.action==='payer'&&params.t)return pagePaiementEmail(params.t);
  if(params.action==='haPaiement'&&params.code){
    var codeHA=params.code||'',statusHA=params.status||'';
    Logger.log('HelloAsso retour — code: '+codeHA+' status: '+statusHA);
    // SÉCURITÉ : cette URL est publique, on ne valide donc RIEN ici.
    // Le paiement est enregistré par le webhook HelloAsso puis validé par l'admin.
    codeHA=/^FRI-[A-Z0-9]{4}$/.test(codeHA)?codeHA:'';
    return HtmlService.createHtmlOutput(
      '<html><head><meta charset="UTF-8"><meta http-equiv="refresh" content="3;url=https://www.frisneauville.fr">'
      +'<style>body{font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f0f4f2}'
      +'.box{background:white;border-radius:14px;padding:40px;max-width:480px;text-align:center;box-shadow:0 10px 40px rgba(0,0,0,.1)}</style></head>'
      +'<body><div class="box"><div style="font-size:48px;margin-bottom:16px">'+(statusHA==='ok'?'✅':'❌')+'</div>'
      +'<h2 style="color:#1a2e22">'+(statusHA==='ok'?'Paiement transmis !':'Paiement annulé')+'</h2>'
      +'<p style="color:#555">Dossier N° <strong>'+codeHA+'</strong></p>'
      +(statusHA==='ok'?'<p style="color:#2d6a4f">Merci ! Votre paiement va être vérifié par le Foyer Rural, puis vous recevrez un email de confirmation.</p>':'<p style="color:#888">Vous pouvez relancer le paiement depuis votre email de confirmation.</p>')
      +'<p style="font-size:12px;color:#aaa;margin-top:16px">Redirection vers frisneauville.fr dans 3 secondes…</p>'
      +'</div></body></html>');
  }
  return traiterRequete(e);
}

function doPost(e) {
  if(ERREUR_CONFIG) return ContentService.createTextOutput(JSON.stringify({status:'error',message:ERREUR_CONFIG})).setMimeType(ContentService.MimeType.JSON);
  try{
    var body=(e&&e.postData&&e.postData.contents)?e.postData.contents:'';
    // form-urlencoded no-cors : le body peut arriver dans e.parameter.payload ou e.postData
    var paramPL=(e&&e.parameter&&e.parameter.payload)?e.parameter.payload:'';
    // Priorite : e.postData si non vide, sinon e.parameter.payload
    if(!body&&paramPL) body=paramPL;
    // Si body commence par 'payload=' c'est du form-urlencoded -> decoder
    if(body&&body.indexOf('payload=')===0){
      try{body=decodeURIComponent(body.substring(8));}catch(eDec){}
    }
    Logger.log('doPost body length: '+body.length+' preview: '+body.substring(0,80));
    if(!verifierRateLimit(e))return ContentService.createTextOutput(JSON.stringify({status:'error',message:'Trop de requêtes'})).setMimeType(ContentService.MimeType.JSON);
    if(!body||body.length===0){Logger.log('doPost: body vide');return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);}
    var payload;
    var trimmed=body.trim();
    if(trimmed.charAt(0)==='{'){payload=JSON.parse(trimmed);}
    else if(trimmed.indexOf('payload=')===0){payload=JSON.parse(decodeURIComponent(trimmed.substring('payload='.length)));}
    else{payload=JSON.parse(trimmed);}
    Logger.log('doPost action: '+payload.action);
    // Appel direct de saveBordereauFFTT depuis doPost
    // (evite le double JSON.stringify qui corrompt la signature base64)
    if (payload.action === 'uploadJustificatif') {
      try {
        var resJ = uploadJustificatifGAS(
          String(payload.code     || ''),
          String(payload.nom      || ''),
          String(payload.prenom   || ''),
          String(payload.filename || 'justificatif.pdf'),
          String(payload.mimeType || 'application/octet-stream'),
          String(payload.fileBase64 || '')
        );
        return repondreAvecCb(resJ, null, null, callback);
      } catch(eJ) {
        Logger.log('uploadJustificatif KO : ' + eJ);
        return repondreAvecCb({status:'error',message:eJ.toString()}, null, null, callback);
      }
    }

    if (payload.action === 'uploadCertificatMedical') {
      if (!verifierToken(payload)) {
        return ContentService.createTextOutput(JSON.stringify({status:'error',message:'Token invalide'})).setMimeType(ContentService.MimeType.JSON);
      }
      var resCertif = uploadCertificatMedicalGAS(
        String(payload.code     || ''),
        String(payload.nom      || ''),
        String(payload.prenom   || ''),
        String(payload.filename || 'certificat.pdf'),
        String(payload.mimeType || 'application/octet-stream'),
        String(payload.fileBase64 || '')
      );
      return ContentService.createTextOutput(JSON.stringify(resCertif)).setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'saveBordereauFFTT') {
      if (!verifierToken(payload)) {
        return ContentService.createTextOutput(JSON.stringify({status:'error',message:'Token invalide'})).setMimeType(ContentService.MimeType.JSON);
      }
      var rowsBFF  = payload.rows  ? JSON.parse(payload.rows)  : [];
      var extraBFF = payload.extra ? JSON.parse(payload.extra) : {};
      var resBFF   = saveBordereauFFTT(
        String(payload.code      || ''),
        String(payload.signature || ''),
        extraBFF,
        rowsBFF
      );
      return ContentService.createTextOutput(JSON.stringify(resBFF)).setMimeType(ContentService.MimeType.JSON);
    }
    if(payload.action==='envoyerAttestationPDF'){
      if(!verifierToken(payload))return ContentService.createTextOutput(JSON.stringify({status:'error',message:'Token invalide'})).setMimeType(ContentService.MimeType.JSON);
      try{
        var pdfB64=payload.pdfBase64||'';
        if(!pdfB64)return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
        pdfB64=pdfB64.replace(/\s/g,'');
        while(pdfB64.length%4!==0)pdfB64+='=';
        var nomFichier=String(payload.filename||'QS-Sante.pdf').replace(/[\\/:*?"<>|]/g,'_');
        var pdfBytes=Utilities.base64Decode(pdfB64,Utilities.Charset.UTF_8);
        var pdfBlob=Utilities.newBlob(pdfBytes,'application/pdf',nomFichier);
        var nomDossierQS = String(payload.dossierOverride||'2-QS Santé Adhérents');
        var dossierQS=creerDossierSecurise(nomDossierQS);
        var fichierQS=dossierQS.createFile(pdfBlob);
        securiserFichier(fichierQS);
        return ContentService.createTextOutput(JSON.stringify({status:'ok',saved:nomFichier})).setMimeType(ContentService.MimeType.JSON);
      }catch(errQS){
        return ContentService.createTextOutput(JSON.stringify({status:'error',message:errQS.toString()})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    // ── Webhook HelloAsso ──
    // HA envoie : { "eventType": "Payment", "data": { "order": { ... }, "payer": {...}, "amount": {...} } }
    Logger.log('Webhook reçu — eventType:' + payload.eventType + ' checkoutId:' + (payload.checkoutIntentId||'?'));
    Logger.log('Webhook HA payload preview: ' + JSON.stringify(payload).substring(0, 500));
    if(payload.eventType === 'Order' || payload.eventType === 'Payment' || payload.eventType === 'Checkout'){
      try {
        var haData   = payload.data ? payload.data : payload;
        var haOrder  = haData.order  || haData || {};

        // ── Filtrer : ignorer les paiements qui ne sont pas des inscriptions FRI ──
        var formSlug = String(haOrder.formSlug || haData.formSlug || '').toLowerCase();
        var formType = String(haOrder.formType || haData.formType || '').toLowerCase();
        // Formulaires FRI acceptés : adhésion uniquement (pas events, dons, etc.)
        var isFRIForm = formSlug.indexOf('adhesion') >= 0
                     || formSlug.indexOf('reglement') >= 0
                     || formSlug.indexOf('inscription') >= 0
                     || formSlug === HA_FORM_SLUG.toLowerCase();
        if (!isFRIForm) {
          Logger.log('Webhook HA ignoré — formulaire non FRI : ' + formSlug + ' (type:' + formType + ')');
          return ContentService.createTextOutput(JSON.stringify({status:'ok',skipped:true,reason:'formulaire non FRI: '+formSlug}))
            .setMimeType(ContentService.MimeType.JSON);
        }
        var haPayer  = haData.payer  || {};
        var haAmount = haData.amount || {};
        var emailPayer = String(haPayer.email||'').trim();

        // Pour le type Order, installmentNumber est dans payments[0]
        var haPayments = haData.payments || [];
        var firstPayment = haPayments.length > 0 ? haPayments[0] : {};

        // Chercher le code dossier — plusieurs sources possibles
        var codeDossier = '';

        // 1. items[].name contient "Adhésion FRI 2026-2027 — FRI-XXXX"
        var haItems = haData.items || [];
        haItems.forEach(function(item) {
          if (!codeDossier) {
            var m = String(item.name||'').match(/FRI-[A-Z0-9]{4}/);
            if (m) codeDossier = m[0];
          }
        });

        // 2. metadata.codeDossier (checkout intent)
        if (!codeDossier) {
          var meta = haData.metadata || haOrder.metadata || {};
          codeDossier = String(meta.codeDossier || '').trim();
        }

        // 3. customFields
        if (!codeDossier) {
          var customFields = haOrder.customFields || haData.customFields || [];
          customFields.forEach(function(f) {
            if (!codeDossier) {
              var fname = String(f.name||'').toLowerCase();
              if (fname.indexOf('dossier') >= 0 || fname.indexOf('code') >= 0) {
                codeDossier = String(f.answer||'').trim();
              }
            }
          });
        }

        // 4. Commentaire libre
        if (!codeDossier) {
          var comment = String(haOrder.comment||haData.comment||'');
          var mc = comment.match(/FRI-[A-Z0-9]{4}/);
          if (mc) codeDossier = mc[0];
        }

        // 5. items[].customFields (paiement direct)
        if (!codeDossier) {
          haItems.forEach(function(item) {
            if (!codeDossier) {
              var cfs = item.customFields || [];
              cfs.forEach(function(cf) {
                if (!codeDossier) {
                  var ans = String(cf.answer||cf.value||'');
                  var mf = ans.match(/FRI-[A-Z0-9]{4}/);
                  if (mf) codeDossier = mf[0];
                }
              });
            }
          });
        }

        // 6. Email du payeur → chercher dans le Sheet un dossier HelloAsso avec cet email
        if (!codeDossier && emailPayer) {
          try {
            var ssSearch = SpreadsheetApp.openById(SHEET_ID);
            var shInsc2 = ssSearch.getSheetByName('Inscriptions');
            if (shInsc2 && shInsc2.getLastRow() > 1) {
              var dataS = shInsc2.getRange(2, 1, shInsc2.getLastRow()-1, 34).getValues();
              for (var si = 0; si < dataS.length; si++) {
                var emailRow = String(dataS[si][15]||'').trim().toLowerCase(); // col P email1
                var modeRow  = String(dataS[si][32]||'').toLowerCase(); // col AG mode
                var statRow  = String(dataS[si][21]||'').toLowerCase(); // col V statut
                var codeRow  = String(dataS[si][19]||'').trim(); // col T code dossier
                if (emailRow === emailPayer.toLowerCase()
                    && modeRow.indexOf('hello') >= 0
                    && statRow.indexOf('attente') >= 0
                    && codeRow.match(/^FRI-[A-Z0-9]{4}$/)) {
                  codeDossier = codeRow;
                  Logger.log('Webhook HA — code trouvé par email payeur: ' + codeDossier);
                  break;
                }
              }
            }
          } catch(eSearch) { Logger.log('Recherche code par email KO: ' + eSearch); }
        }

        Logger.log('Webhook HA — code trouvé: "' + codeDossier + '" (email payeur: ' + emailPayer + ')');

        var montantCts = Number(haAmount.total || firstPayment.amount || 0); // en centimes
        var montantEur = montantCts > 0 ? (montantCts / 100).toFixed(2) : '?';

        Logger.log('Webhook HA — eventType:'+payload.eventType+' code:'+codeDossier+' email:'+emailPayer+' montant:'+montantEur+'€');

        if (codeDossier && codeDossier.match(/^FRI-[A-Z0-9]{4}$/)) {
          // Paiement 3x : ne valider qu'au 1er versement, juste logguer les suivants
          var installNum = Number(haData.installmentNumber || firstPayment.installmentNumber || 1);
          if (installNum > 1) {
            Logger.log('✅ Webhook HA — versement ' + installNum + ' reçu pour ' + codeDossier + ' (' + montantEur + '€) — pas de re-validation');
          } else {
            // Vérifier si le dossier est déjà validé (anti-doublon)
            var ssHA = SpreadsheetApp.openById(SHEET_ID);
            var sheetInsc = ssHA.getSheetByName('Inscriptions');
            var dejaValide = false;
            if (sheetInsc && sheetInsc.getLastRow() > 1) {
              var dataCheck = sheetInsc.getRange(2, 1, sheetInsc.getLastRow()-1, 22).getValues();
              for (var dc = 0; dc < dataCheck.length; dc++) {
                if (String(dataCheck[dc][19]||'').trim() === codeDossier) {
                  var statutV = String(dataCheck[dc][21]||'').toLowerCase();
                  if (statutV.indexOf('pay') >= 0 || statutV.indexOf('valid') >= 0) {
                    dejaValide = true; break;
                  }
                }
              }
            }
            // Verrou supplémentaire basé sur checkoutIntentId (évite doublons si HA renvoie)
            var checkoutId = String(payload.checkoutIntentId || haData.checkoutIntentId || '');
            var verroKey = 'ha_checkout_' + (checkoutId || codeDossier);
            var verroProps = PropertiesService.getScriptProperties();
            var dejaTraite = verroProps.getProperty(verroKey);

            var dossierPresent = false;
            if (sheetInsc && sheetInsc.getLastRow() > 1) {
              dossierPresent = dataCheck.some(function(rc) { return String(rc[19] || '').trim() === codeDossier; });
            }
            if (!dossierPresent) {
              envoyerEmail(EMAIL_ADMIN, '[FRI] ALERTE — paiement HelloAsso reçu, dossier ' + codeDossier + ' absent de la feuille',
                'Un paiement HelloAsso de ' + montantEur + ' € a été reçu pour le dossier ' + codeDossier
                + ' (payeur : ' + emailPayer + '), mais aucune ligne de ce dossier n\'existe dans l\'onglet Inscriptions.\n\n'
                + 'Contactez la famille pour qu\'elle refasse son inscription (sans payer à nouveau), ou utilisez « Récupérer dossiers » dans la console admin.\n\n'
                + 'Détail HelloAsso :\n' + JSON.stringify(haData, null, 2).substring(0, 4000), { name: 'FRI Admin' });
            }
            if (dejaValide && !dejaTraite && checkoutId) {
              // Complément (ex. activité ajoutée) payé sur un dossier déjà réglé :
              // on l'enregistre dans l'onglet HelloAsso sans toucher au statut du dossier.
              verroProps.setProperty(verroKey, String(Date.now()));
              majOngletHelloAssoWebhook(SpreadsheetApp.openById(SHEET_ID), codeDossier, haPayer, montantCts / 100,
                String(haOrder.formType || ''), String(haOrder.formName || 'Complément'));
              envoyerEmail(EMAIL_ADMIN, '[FRI] Complément HelloAsso reçu — ' + codeDossier,
                'Paiement HelloAsso de ' + montantEur + ' € reçu pour le dossier ' + codeDossier + ' (dossier déjà réglé : complément).',
                { name: 'FRI Admin' });
              Logger.log('✅ Webhook HA — complément enregistré : ' + codeDossier + ' ' + montantEur + '€');
            } else if (dejaValide || dejaTraite) {
              Logger.log('⚠️ Webhook HA — skip (déjàValidé:' + dejaValide + ' déjàTraité:' + !!dejaTraite + ') code:' + codeDossier + ' checkout:' + checkoutId);
            } else {
              verroProps.setProperty(verroKey, String(Date.now()));
              var montantHAEur = montantCts / 100;
              var ssWebhook = SpreadsheetApp.openById(SHEET_ID);
              // NE PAS valider automatiquement — enregistrer le paiement reçu
              // et laisser l'admin valider après vérification des pièces justificatives
              var haFT = String(haOrder.formType || '');
              var haFN = String(haOrder.formName || '');
              majOngletHelloAssoWebhook(ssWebhook, codeDossier, haPayer, montantHAEur, haFT, haFN);
              // Mettre à jour col 22 (statut paiement) : paiement reçu, en attente de validation admin
              var shInscW = ssWebhook.getSheetByName('Inscriptions');
              if (shInscW && shInscW.getLastRow() > 1) {
                var dataW = shInscW.getRange(2,1,shInscW.getLastRow()-1,22).getValues();
                for (var wi=0; wi<dataW.length; wi++) {
                  if (String(dataW[wi][19]||'').trim() === codeDossier) {
                    var statutActW = String(dataW[wi][21]||'').toLowerCase();
                    // Ne pas écraser si déjà validé par l'admin
                    if (statutActW.indexOf('valid') < 0 && statutActW.indexOf('pay') < 0) {
                      shInscW.getRange(wi+2, 22)
                        .setValue('✅ Paiement HelloAsso reçu — ' + montantHAEur.toFixed(2) + ' € — en attente validation admin')
                        .setBackground('#fff8e1').setFontColor('#856404').setFontWeight('bold');
                    }
                  }
                }
              }
              Logger.log('✅ Webhook HA — paiement enregistré (non validé auto) : ' + codeDossier + ' ' + montantHAEur + '€');
            }
          } // fin else installNum
        } else {
          // Pas de code dossier trouvé — envoyer un email à l'admin
          Logger.log('⚠️ Webhook HA — code dossier introuvable pour payer: ' + emailPayer);
          envoyerEmail(EMAIL_ADMIN,
            '[FRI] ALERTE Paiement HelloAsso sans code dossier',
            'Un paiement HelloAsso de ' + montantEur + ' € a été reçu de ' + emailPayer + ' mais aucun code dossier FRI-XXXX n\'a été trouvé.\n\nDétail HA :\n' + JSON.stringify(haData, null, 2),
            { name: 'FRI Admin' });
        }
      } catch(eHA) {
        Logger.log('Webhook HA ERREUR : ' + eHA.toString());
      }
      return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
    }

    if(payload.action==='addRegistration'&&payload.rows&&payload.rows.length>0){
      if(payload.cheques&&payload.rows[0])payload.rows[0].cheques=payload.cheques;
      var result=addRegistration(payload.rows,payload.status||'paid',payload);
      return ContentService.createTextOutput(JSON.stringify(reponseAddRegistration(result))).setMimeType(ContentService.MimeType.JSON);
    }
    if(payload.action==='ajouterActiviteDossier'){
      if(!verifierTokenAdmin(payload._adminToken))return ContentService.createTextOutput(JSON.stringify(REPONSE_ADMIN_REQUISE)).setMimeType(ContentService.MimeType.JSON);
      var result4=ajouterActiviteDossierSheet(payload);
      return ContentService.createTextOutput(JSON.stringify({status:'ok',inserted:result4.inserted})).setMimeType(ContentService.MimeType.JSON);
    }
    return traiterRequete({parameter:{payload:body}});
  }catch(err){
    Logger.log('doPost ERREUR: '+err.toString());
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function traiterActionDirecte(payload, callback) {
  // Re-entrée dans traiterRequete avec payload déjà parsé
  try {
    var e2 = { parameter: { payload: JSON.stringify(payload), callback: callback || '' } };
    return traiterRequete(e2);
  } catch(e) {
    return repondreAvecCb({status:'error', message:e.toString()}, null, null, callback);
  }
}

function traiterRequete(e) {
  try{
    var params=(e&&e.parameter)?e.parameter:{};
    var rawPayload=params.payload||null;
    var callback=params.callback||null;
    Logger.log('=== REQUETE === rawPayload length: '+(rawPayload?rawPayload.length:0)+' callback: '+callback);
    if(!rawPayload||rawPayload===''){
      // getToken via JSONP GET
      if(params.action==='getToken'){
        if(!FRI_SECRET_TOKEN) return repondreAvecCb({status:'error',message:'Token non configuré'},null,null,callback);
        var dynTok=genererTokenDynamique();
        return repondreAvecCb({status:'ok',token:dynTok,expires:Date.now()+4*60*1000},null,null,callback);
      }
      // Action rechercherAdherent passée directement en query string (JSONP)
      if(params.action==='rechercherAdherent'&&params.tel){
        var resTelDirect=rechercherAdherentParTel(String(params.tel||''),String(params.nom||''));
        return repondreAvecCb(resTelDirect,null,null,callback);
      }
      if(params.action==='validerPaiement'&&params.code){
        if(!verifierTokenAdmin(params._adminToken)){Logger.log('⛔ Session admin absente sur validerPaiement GET');return repondreAvecCb(REPONSE_ADMIN_REQUISE,null,null,callback);}
        var result=validerPaiementSheet([],params.code);
        return repondreAvecCb({status:'ok',updated:result.updated},null,null,callback);
      }
      // getDossierDetail sans payload : code passé directement en query string
      if(params.action==='getDossierDetail' && params.code){
        var payload2 = {action:'getDossierDetail', code:params.code, _token:params._token||'', _adminToken:params._adminToken||''};
        return traiterActionDirecte(payload2, callback);
      }
      // modifierActivite sans payload
      if(params.action==='modifierActivite' && params.code){
        try{ var payload3=JSON.parse(params.data||'{}'); payload3.action='modifierActivite'; payload3.code=params.code; payload3._token=params._token||''; payload3._adminToken=params._adminToken||'';
          return traiterActionDirecte(payload3, callback);
        }catch(e){}
      }
      return repondreAvecCb({status:'ok',message:'FRI Apps Script v8.3 operationnel'},null,null,callback);
    }
    var payload=JSON.parse(rawPayload);
    if(payload.action) payload.action = String(payload.action).trim();
    // Normaliser adminLogin quelle que soit la casse reçue
    if(payload.action && payload.action.toLowerCase() === 'adminlogin') payload.action = 'adminLogin';
    Logger.log('action normalisée: "' + payload.action + '"');

    // ── Actions réservées à l'équipe (admin / secrétariat) : session admin obligatoire ──
    if(ACTIONS_ADMIN.indexOf(payload.action)>=0){
      var sessionAdmin=verifierTokenAdmin(payload._adminToken||params._adminToken||'');
      if(!sessionAdmin){
        Logger.log('⛔ Action admin refusée sans session valide : '+payload.action);
        return repondreAvecCb(REPONSE_ADMIN_REQUISE,null,null,callback);
      }
      Logger.log('Action admin '+payload.action+' par '+sessionAdmin.user+' ('+sessionAdmin.role+')');
    }

    // ── getToken : génère un token dynamique signé, valable 5 minutes ──
    if(payload.action==='getToken'){
      if(!FRI_SECRET_TOKEN) return repondreAvecCb({status:'error',message:'Token non configuré'},null,null,callback);
      var dynToken = genererTokenDynamique();
      return repondreAvecCb({status:'ok',token:dynToken,expires:Date.now()+4*60*1000},null,null,callback);
    }

    // Actions publiques sans verification de token
    if(payload.action==='validerElementPaiement'){
      var resPay=validerElementPaiementGAS(String(payload.code||''),String(payload.onglet||''),String(payload.ligne||''));
      return repondreAvecCb(resPay,null,null,callback);
    }
    if(payload.action==='getElementsPaiement'){
      var resGet=getElementsPaiementGAS(String(payload.code||''));
      return repondreAvecCb(resGet,null,null,callback);
    }
    if(payload.action==='rechercherAdherent'){
      var resTel=rechercherAdherentParTel(String(payload.tel||''),String(payload.nom||''));
      return repondreAvecCb(resTel,null,null,callback);
    }
    if(payload.action==='getLicenceFFTT'){
      var resLic=getLicenceFFTT(String(payload.nom||''),String(payload.prenom||''),String(payload.numLicence||''));
      return repondreAvecCb(resLic,null,null,callback);
    }
    // adminLogin : action publique (pas de token requis — c'est la fonction qui crée la session)
    if(payload.action==='adminLogin'){
      try{
        var props=PropertiesService.getScriptProperties();
        var loginUser=String(payload.user||'').toLowerCase().trim();
        var loginPass=String(payload.pass||'');
        // Log de diagnostic (longueurs seulement, jamais les valeurs)
        Logger.log('adminLogin reçu: user="'+loginUser+'"');
        // Serveur sans configuration (ex. copie de projet : les propriétés ne sont pas copiées)
        if(!FRI_SECRET_TOKEN){
          return repondreAvecCb({status:'error',message:'Serveur non configuré : propriété FRI_SECRET_TOKEN absente (Paramètres du projet › Propriétés du script).'},null,null,callback);
        }
        var aDesComptes=Object.keys(props.getProperties()).some(function(k){return k.indexOf('ADMIN_CRED_')===0;});
        if(!aDesComptes){
          return repondreAvecCb({status:'error',message:'Aucun compte administrateur sur ce serveur : exécuter initAdminCredentials dans l\'éditeur Apps Script.'},null,null,callback);
        }
        var storedPassHex=props.getProperty('ADMIN_CRED_'+loginUser);
        var storedRole=props.getProperty('ADMIN_ROLE_'+loginUser);
        if(!storedPassHex||!storedRole){
          Logger.log('⛔ Admin login - identifiant inconnu : "'+loginUser+'"');
          return repondreAvecCb({status:'error',message:'Identifiant ou mot de passe incorrect'},null,null,callback);
        }
        // HMAC-SHA256(pass, FRI_SECRET_TOKEN) en hex — identique à initAdminCredentials
        var sig=Utilities.computeHmacSha256Signature(loginPass,FRI_SECRET_TOKEN);
        var sigHex=sig.map(function(b){return('0'+(b&0xff).toString(16)).slice(-2);}).join('');
        if(sigHex!==storedPassHex){
          Logger.log('⛔ Admin login - hash mismatch pour : "'+loginUser+'"');
          return repondreAvecCb({status:'error',message:'Identifiant ou mot de passe incorrect'},null,null,callback);
        }
        var sessionSecret=props.getProperty('ADMIN_SESSION_SECRET')||FRI_SECRET_TOKEN;
        if(!sessionSecret){
          Logger.log('⛔ adminLogin : aucun secret de session configuré (ADMIN_SESSION_SECRET / FRI_SECRET_TOKEN)');
          return repondreAvecCb({status:'error',message:'Serveur non configuré'},null,null,callback);
        }
        var expiry=String(Math.floor(Date.now()/1000)+28800);
        var payload2=loginUser+':'+storedRole+':'+expiry;
        var sig2=Utilities.computeHmacSha256Signature(payload2,sessionSecret);
        var sig2Hex=sig2.map(function(b){return('0'+(b&0xff).toString(16)).slice(-2);}).join('');
        var sessionToken=payload2+':'+sig2Hex;
        Logger.log('✅ Admin login OK : '+loginUser+' ('+storedRole+')');
        return repondreAvecCb({status:'ok',token:sessionToken,role:storedRole,user:loginUser},null,null,callback);
      }catch(eLogin){
        Logger.log('ERREUR adminLogin : '+eLogin);
        return repondreAvecCb({status:'error',message:'Erreur serveur'},null,null,callback);
      }
    }
    if(payload.action==='verifierCodeAvoir'){
      var resVer2=verifierCodeAvoirGAS(String(payload.codeAvoir||''),parseFloat(payload.montant||0));
      return repondreAvecCb(resVer2,null,null,callback);
    }
    if(payload.action==='getPlacesRestantes'){
      try {
        // ── Servir depuis le cache PropertiesService si < 5 min ──
        var propsP = PropertiesService.getScriptProperties();
        var cacheP = propsP.getProperty('fri_places_cache');
        var cacheTs = parseInt(propsP.getProperty('fri_places_cache_ts')||'0');
        var cacheAge = Date.now() - cacheTs;
        if(cacheP && cacheAge < 7200000) { // 2 heures
          Logger.log('getPlacesRestantes: cache hit (age='+Math.round(cacheAge/1000)+'s)');
          return repondreAvecCb({status:'ok',places:JSON.parse(cacheP),fromCache:true},null,null,callback);
        }
        // ── Lire le Sheet et mettre en cache ──
        var ssP=SpreadsheetApp.openById(SHEET_ID);
        var shP=getOrCreatePlacesSheet(ssP);
        var lrP=shP.getLastRow();var pmP={};
        if(lrP>1){var dP=shP.getRange(2,1,lrP-1,5).getValues();dP.forEach(function(r){var pid=String(r[0]||'').trim();if(pid)pmP[pid]={capacite:Number(r[2]||0),inscrits:Number(r[3]||0),restantes:Number(r[4]||0)};});}
        // Mettre en cache
        try {
          propsP.setProperty('fri_places_cache', JSON.stringify(pmP));
          propsP.setProperty('fri_places_cache_ts', String(Date.now()));
        } catch(ec){}
        Logger.log('getPlacesRestantes: '+Object.keys(pmP).length+' places lues depuis Sheet');
        return repondreAvecCb({status:'ok',places:pmP},null,null,callback);
      } catch(eP) {
        Logger.log('getPlacesRestantes KO: '+eP);
        return repondreAvecCb({status:'error',message:eP.toString()},null,null,callback);
      }
    }

    if(payload.action==='invaliderCachePlaces'){
      try {
        PropertiesService.getScriptProperties().deleteProperty('fri_places_cache');
        PropertiesService.getScriptProperties().deleteProperty('fri_places_cache_ts');
        return repondreAvecCb({status:'ok'},null,null,callback);
      } catch(e){ return repondreAvecCb({status:'error'},null,null,callback); }
    }
    // Actions admin protégées nécessitant un token valide
    // getStatsTresorier, getJournalSauvegardes, verifierDossiersPerdus : token requis mais pas de blocage
    // car ces actions sont déjà gérées plus bas dans le code
    if(payload.action!=='ping'
      && payload.action!=='addRegistration'
      && payload.action!=='getStatsTresorier'
      && payload.action!=='getJournalSauvegardes'
      && payload.action!=='viderJournalSauvegardes'
      && payload.action!=='verifierDossiersPerdus'
      && payload.action!=='getDossierDetail'
      && !verifierToken(payload)){
      return repondreAvecCb({status:'error',message:'Accès non autorisé'},null,null,callback);
    }
    if(payload.action==='ping')return repondreAvecCb({status:'ok',message:'PONG v8.82'+(MODE_TEST?' (TEST)':''),env:MODE_TEST?'test':'production'},null,null,callback);


    if(payload.action==='rechercherLicencieFFTT'){
      var resFF = rechercherLicencieFFTT(
        String(payload.nom    || ''),
        String(payload.prenom || ''),
        String(payload.ddn    || ''),
        String(payload.numLic || '')
      );
      return repondreAvecCb(resFF, null, null, callback);
    }

    if(payload.action==='saveBordereauFFTT'){
      try {
        var rowsBFF = payload.rows ? JSON.parse(payload.rows) : [];
        var extraBFF = payload.extra ? JSON.parse(payload.extra) : {};
        var resBFF = saveBordereauFFTT(payload.code||'', payload.signature||'', extraBFF, rowsBFF);
        return repondreAvecCb(resBFF, null, null, callback);
      } catch(eBFF) {
        Logger.log('saveBordereauFFTT KO: ' + eBFF);
        return repondreAvecCb({status:'error', message: String(eBFF)}, null, null, callback);
      }
    }

    if(payload.action==='utiliserCodeAvoir'){
      if(!verifierToken(payload))return repondreAvecCb({status:'error',message:'Token invalide'},null,null,callback);
      var resUtil = utiliserAvoirGAS(String(payload.codeAvoir||''), parseFloat(payload.montant||0));
      return repondreAvecCb(resUtil, null, null, callback);
    }

    if(payload.action==='creerAvoirManuel'){
      if(!verifierToken(payload))return repondreAvecCb({status:'error',message:'Token invalide'},null,null,callback);
      // type: 'avoir' (défaut, comportement historique) ou 'remboursement'
      var typeCav = String(payload.type||'avoir').toLowerCase();
      var resCav;
      if (typeCav === 'remboursement') {
        resCav = creerRemboursementManuelGAS(
          String(payload.code||''),
          String(payload.motif||'Remboursement manuel'),
          parseFloat(payload.montant||0),
          String(payload.modeRemboursement||'cheque'),
          String(payload.email||''),
          String(payload.nom||''),
          String(payload.prenom||'')
        );
      } else {
        resCav = creerAvoirManuelGAS(
          String(payload.code||''),
          String(payload.motif||'Avoir manuel'),
          parseFloat(payload.montant||0),
          String(payload.email||''),
          String(payload.nom||''),
          String(payload.prenom||'')
        );
      }
      return repondreAvecCb(resCav, null, null, callback);
    }

    if(payload.action==='envoyerRappelManuel'){
      try {
        var resRappel = envoyerRappelManuelGAS(payload.code || '');
        return repondreAvecCb(resRappel, null, null, callback);
      } catch(eRappelM) {
        Logger.log('envoyerRappelManuel KO: ' + eRappelM.toString());
        return repondreAvecCb({status:'error', message: eRappelM.toString()}, null, null, callback);
      }
    }

    if(payload.action==='getJournalSauvegardes'){
      try{
        var props2 = PropertiesService.getScriptProperties();
        var journal2 = JSON.parse(props2.getProperty('fri_journal_saves') || '[]');
        return repondreAvecCb({status:'ok', journal: journal2},null,null,callback);
      }catch(eJ2){
        return repondreAvecCb({status:'error',message:eJ2.toString()},null,null,callback);
      }
    }

    if(payload.action==='viderJournalSauvegardes'){
      try{
        PropertiesService.getScriptProperties().deleteProperty('fri_journal_saves');
        return repondreAvecCb({status:'ok'},null,null,callback);
      }catch(eJ3){
        return repondreAvecCb({status:'error',message:eJ3.toString()},null,null,callback);
      }
    }

    // NOTE : l'ancien gestionnaire 'getStatsTresorier' (lecture directe non
    // agrégée, une par activité) a été retiré d'ici — il interceptait la
    // requête AVANT le nouveau gestionnaire (voir plus bas, qui appelle
    // getStatsTresorierGAS()) et l'empêchait donc de s'exécuter, ce qui
    // expliquait à la fois la lenteur (~25-27s, pas d'agrégation ni de
    // filtrage) et le format de réponse inattendu côté client (pas
    // d'enveloppe {stats:...}). Le gestionnaire actif est désormais celui
    // ci-dessous (recherchez 'getStatsTresorierGAS').

    if(payload.action==='verifierDossiersPerdus'){
      try{
        var codesVP = payload.codes || [];
        if(!Array.isArray(codesVP)||codesVP.length===0)
          return repondreAvecCb({status:'ok',codes_presents:[]},null,null,callback);
        var ssVP = SpreadsheetApp.openById(SHEET_ID);
        var shVP = ssVP.getSheetByName(SHEET_INSCRIPTIONS);
        if(!shVP||shVP.getLastRow()<2)
          return repondreAvecCb({status:'ok',codes_presents:[]},null,null,callback);
        var colCodes = shVP.getRange(2,20,shVP.getLastRow()-1,1).getValues();
        var presents = [];
        colCodes.forEach(function(row){
          var c = String(row[0]||'').trim();
          if(c && codesVP.indexOf(c)>=0 && presents.indexOf(c)<0) presents.push(c);
        });
        Logger.log('verifierDossiersPerdus: '+codesVP.length+' codes vérifiés, '+presents.length+' présents');
        return repondreAvecCb({status:'ok',codes_presents:presents},null,null,callback);
      }catch(eVP){
        return repondreAvecCb({status:'error',message:eVP.toString()},null,null,callback);
      }
    }

    if(payload.action==='getDossierDetail'){
      try{
        var codeDD = String(payload.code||'').trim().toUpperCase();
        var ss2 = SpreadsheetApp.openById(SHEET_ID);
        var shDD = ss2.getSheetByName(SHEET_INSCRIPTIONS);
        if(!shDD) return repondreAvecCb({status:'error',message:'Onglet introuvable'},null,null,callback);
        var dataDD = shDD.getRange(2,1,Math.max(shDD.getLastRow()-1,1),42).getValues();
        var lignes = [];
        var totalFamille = 0, commune = '', ville = '', responsable = {}, modePaiement = '';
        dataDD.forEach(function(r){
          if(String(r[19]||'').trim() !== codeDD) return;
          var statut = lireStatutInscription(r);
          if(statut.toLowerCase().indexOf('supprim')>=0) return;
          commune = commune || (String(r[10]||'').toUpperCase().indexOf('ISNEAUVILLE')>=0?'isno':'');
          ville = ville || String(r[10]||'');
          if(!responsable.nom) responsable = {
            nom: String(r[2]||''), prenom: String(r[3]||''),
            email: String(r[15]||''), tel: String(r[14]||'')
          };
          var ddnRaw = r[4]; var ddnStr = '';
          if(ddnRaw instanceof Date) ddnStr = Utilities.formatDate(ddnRaw,'Europe/Paris','dd/MM/yyyy');
          else ddnStr = String(ddnRaw||'');
          var actId = lireActiviteId(r);
          totalFamille = Math.max(totalFamille, parseFloat(r[31])||0);
          if (!modePaiement) modePaiement = String(r[32]||'');
          lignes.push({
            activite_id: actId,
            activite: String(r[22]||''),
            membre_nom: String(r[2]||'').trim(),
            membre_prenom: String(r[3]||'').trim(),
            ddn: ddnStr,
            tarif_brut: parseFloat(r[27])||0,
            tarif: parseFloat(r[27])||0,
            statut: statut,
            jour: String(r[23]||''),
            heure: String(r[24]||''),
            lieu: String(r[25]||'')
          });
        });
        return repondreAvecCb({status:'ok',lignes:lignes,total:totalFamille,commune:commune,ville:ville,responsable:responsable,mode_paiement:modePaiement||''},null,null,callback);
      }catch(eDD){ return repondreAvecCb({status:'error',message:eDD.toString()},null,null,callback); }
    }

    if(payload.action==='modifierActivite'){
      try{
        var ss3 = SpreadsheetApp.openById(SHEET_ID);
        var shMod = ss3.getSheetByName(SHEET_INSCRIPTIONS);
        var codeM      = String(payload.code||'').trim().toUpperCase();
        var oldActId   = String(payload.oldActId||'').trim();
        var oldActNom  = String(payload.oldActNom||'').trim();
        var newActId   = String(payload.newActId||'').trim();
        var newActNom  = String(payload.newActNom||'').trim();
        var memNom     = String(payload.membreNom||'').trim().toUpperCase();
        var memPrenom  = String(payload.membrePrenom||'').trim();
        var newTarif   = parseFloat(payload.newTarif)||0;
        var oldTarif   = parseFloat(payload.oldTarif)||0;
        var diff       = parseFloat(payload.diff)||(newTarif - oldTarif);
        var modeP      = String(payload.modePaiement||'');
        var commentaireAdminMod = String(payload.commentaireAdmin||'').trim();
        var montantModifieMod   = payload.montantModifie === true || payload.montantModifie === 'true';
        // isPaid = vrai seulement si le statut inscription indique paiement validé
        var isPaid = false; // sera mis à jour après avoir trouvé la ligne

        Logger.log('modifierActivite: '+codeM+' '+oldActId+'\u2192'+newActId+' diff:'+diff);

        // Structure colonnes (index 0-based / col base-1) :
        // [22]=col23 Activit\u00e9  [23]=col24 Jour  [24]=col25 Heure  [25]=col26 Lieu  [26]=col27 Animateur
        // [27]=col28 AB Tarif brut  [28]=col29 AC Remise%  [29]=col30 AD Tarif net
        // [31]=col32 AF Total famille  [32]=col33 AG Mode paiement
        // [37]=col38 AL ID Activit\u00e9  [39]=col40 AN Statut inscription

        var dataM = shMod.getRange(2,1,Math.max(shMod.getLastRow()-1,1),41).getValues();
        var rowsUpdated = 0;
        var emailRow = null;
        var shRow = -1;

        for(var mi=0; mi<dataM.length; mi++){
          var rm = dataM[mi];
          if(String(rm[19]||'').trim() !== codeM) continue;
          if(String(rm[2]||'').trim().toUpperCase() !== memNom) continue;
          if(String(rm[3]||'').trim().toLowerCase() !== memPrenom.toLowerCase()) continue;
          var statut = lireStatutInscription(rm);
          if(statut.toLowerCase().indexOf('supprim')>=0) continue;
          var actIdM = lireActiviteId(rm);
          var actNomM = String(rm[22]||'').trim();
          var matchAct = (actIdM === oldActId) || (oldActNom && actNomM === oldActNom);
          if(!matchAct) continue;

          shRow = mi + 2;
          // ── Mise à jour onglet Inscriptions ──
          // Col Q(17) = activité dupliquée
          shMod.getRange(shRow, 17).setValue(newActNom);
          // Col 23 = Activité principale
          shMod.getRange(shRow, 23).setValue(newActNom);
          shMod.getRange(shRow, 24).setValue(String(payload.jour||''));
          shMod.getRange(shRow, 25).setValue(String(payload.heure||''));
          shMod.getRange(shRow, 26).setValue(String(payload.lieu||''));
          shMod.getRange(shRow, 27).setValue(String(payload.animateur||''));
          // Col 28 AB = Tarif brut
          shMod.getRange(shRow, 28).setValue(newTarif)
            .setBackground('#d8f3dc').setFontColor('#1b5e20').setFontWeight('bold');
          // Col 29 AC = Éligible remise
          // (estEligRemise n'existe pas en global — c'est estEligibleRemise(pid, commune) qu'il faut appeler)
          var communeM = communeFromVille(String(rm[10]||''));
          var newIsElig = estEligibleRemise(newActId, communeM) ? 1 : 0;
          shMod.getRange(shRow, 29).setValue(newIsElig)
            .setBackground(newIsElig ? '#e8f5e9' : '#fce4ec')
            .setFontColor(newIsElig ? '#1b5e20' : '#b71c1c').setFontWeight('bold');
          // Col 30 AD = Tarif net (formule =AB*(1-0,15*AC))
          shMod.getRange(shRow, 30).setFormula(formuleAD(shRow))
            .setBackground('#d8f3dc').setFontColor('#1b5e20');
          // Col 32 AF = Total famille (recalculé par calcTotalFamille plus bas)
          // Col 35 AI = QS Santé (conserver valeur existante)
          shMod.getRange(shRow, 35).setValue(String(rm[34]||''));
          // Col 38 AL ou 36 AJ = ID Activité
          var colActId = isNewStructure(rm) ? 38 : 36;
          shMod.getRange(shRow, colActId).setValue(newActId);
          // Couleur mauve claire sur toute la ligne
          shMod.getRange(shRow, 1, 1, 41).setBackground('#e1bee7');
          // Col 40 AN = Statut inscription
          shMod.getRange(shRow, 40).setValue('✅ Modifié — ' + Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy'))
            .setFontColor('#6a1b9a').setFontWeight('bold');
          rowsUpdated++;
          // isPaid : basé sur le STATUT DE PAIEMENT réel (col 22 / index 21),
          // PAS sur le statut d'inscription structurel (col AN/AL, lireStatutInscription).
          // Ce dernier passe à "✅ Inscrit" dès qu'une bascule liste d'attente
          // a eu lieu, indépendamment du règlement -- l'utiliser ici donnait de
          // faux positifs "réglé".
          var statutPaiementActuel = String(rm[21]||'').toLowerCase();
          isPaid = statutPaiementActuel.indexOf('valid\u00e9') >= 0
                || statutPaiementActuel.indexOf('pay\u00e9') >= 0;
          Logger.log('isPaid='+isPaid+' statutPaiement=['+statutPaiementActuel+']');
          emailRow = {
            code:      codeM,
            nom:       String(rm[2]||''),
            prenom:    String(rm[3]||''),
            email:     String(rm[15]||''),
            memNom:    memNom,
            memPrenom: memPrenom,
            ancienne:  actNomM || String(rm[22]||''),
            nouvelle:  newActNom,
            jour:      String(payload.jour||''),
            heure:     String(payload.heure||''),
            oldTarif:  oldTarif,
            newTarif:  newTarif,
            diff:      diff,
            mode:      modeP
          };
          break;
        }

        if(rowsUpdated===0) return repondreAvecCb({status:'error',message:'Ligne introuvable: '+codeM+'/'+oldActId},null,null,callback);

        SpreadsheetApp.flush();

        // ── Recalcul total famille via calcTotalFamille (met \u00e0 jour AF=32) ──
        var totalNet = 0;
        try { totalNet = calcTotalFamille(ss3, codeM); }
        catch(etf) { Logger.log('calcTotalFamille KO: '+etf); }

        // ── Mettre \u00e0 jour le R\u00e9capitulatif (col I=9 et col M=13) ──
        try {
          var recapSh = ss3.getSheetByName(SHEET_RECAPITULATIF);
          if(recapSh && recapSh.getLastRow()>1){
            var recapData = recapSh.getRange(2,1,recapSh.getLastRow()-1,15).getValues();
            for(var rr=0; rr<recapData.length; rr++){
              var hasCode = recapData[rr].some(function(c){ return String(c||'').trim()===codeM; });
              if(!hasCode) continue;
              recapSh.getRange(rr+2, 9).setValue(totalNet).setFontColor('#1565c0').setFontWeight('bold');  // col I = Total
              // col M = Activit\u00e9s (concat des activit\u00e9s du dossier)
              var actsConcat = dataM.filter(function(r){
                return String(r[19]||'').trim()===codeM && lireStatutInscription(r).toLowerCase().indexOf('supprim')<0;
              }).map(function(r){ return String(r[22]||''); }).join(', ');
              recapSh.getRange(rr+2, 13).setValue(actsConcat).setFontColor('#2d6a4f');
              Logger.log('\u2705 Recap mis \u00e0 jour: '+codeM+' total='+totalNet);
              break;
            }
          }
        } catch(eR){ Logger.log('Recap MAJ KO: '+eR); }

        // ── Supprimer la ligne de l'ancienne activité ──
        try {
          var oldSh = ss3.getSheetByName(oldActId);
          if(oldSh && oldSh.getLastRow()>1){
            var oldData2 = oldSh.getRange(2,1,oldSh.getLastRow()-1,5).getValues();
            for(var oi=0; oi<oldData2.length; oi++){
              if(String(oldData2[oi][0]||'').trim()!==codeM) continue;
              var oldMemNomRow = (String(oldData2[oi][3]||'')+' '+String(oldData2[oi][4]||'')).toUpperCase();
              if(oldMemNomRow.indexOf(memNom)<0) continue;
              oldSh.deleteRow(oi+2); // Suppression physique
              Logger.log('✅ Ligne supprimée de onglet: '+oldActId);
              break;
            }
          }
        } catch(eOld){ Logger.log('Onglet ancienne act KO: '+eOld); }

        // ── Mettre à jour les places ──
        try {
          var plSh = ss3.getSheetByName(SHEET_PLACES);
          if(plSh && plSh.getLastRow()>1){
            var plData = plSh.getRange(2,1,plSh.getLastRow()-1,6).getValues();
            for(var pli=0; pli<plData.length; pli++){
              var plId = String(plData[pli][0]||'').trim();
              if(plId===oldActId){
                plSh.getRange(pli+2,3).setValue(Math.max(0,parseInt(plData[pli][2]||0))+1); // dispo +1
                plSh.getRange(pli+2,4).setValue(Math.max(0,parseInt(plData[pli][3]||0))-1); // inscrits -1
                Logger.log('✅ Place libérée: '+oldActId);
              }
              if(plId===newActId){
                plSh.getRange(pli+2,3).setValue(Math.max(0,parseInt(plData[pli][2]||0))-1); // dispo -1
                plSh.getRange(pli+2,4).setValue(parseInt(plData[pli][3]||0)+1); // inscrits +1
                Logger.log('✅ Place prise: '+newActId);
              }
            }
          }
        } catch(ePlaces){ Logger.log('Places MAJ KO: '+ePlaces); }

        // ── Ajouter le membre dans l'onglet de la nouvelle activité ──
        try {
          var newSh = ss3.getSheetByName(newActId);
          if(!newSh) {
            newSh = ss3.insertSheet(newActId);
            var hdrN = ['N° Dossier','Date','Statut','Nom','Prénom','DDN','Sexe','Responsable','Tél','Email','Commune','QS Santé','Mode','Pass/Aide'];
            newSh.getRange(1,1,1,hdrN.length).setValues([hdrN]).setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
            newSh.setFrozenRows(1);
          }
          // Relire la ligne modifiée directement (dataM a l'ancien actId)
          var rmData = shMod.getRange(shRow, 1, 1, 41).getValues()[0];
          var dateJourN = Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy à HH:mm');
          var ddnN = rmData[4] instanceof Date ? Utilities.formatDate(rmData[4],'Europe/Paris','dd/MM/yyyy') : String(rmData[4]||'');
          var newActRow = [
            codeM, dateJourN, '✅ Inscrit (modifié)',
            String(rmData[2]||''), String(rmData[3]||''), ddnN,
            lireSexe(rmData), String(rmData[38]||''),
            String(rmData[14]||''), String(rmData[15]||''),
            String(rmData[10]||''), String(rmData[34]||''),
            String(rmData[32]||''), String(rmData[35]||'')
          ];
          var nextN = Math.max(newSh.getLastRow()+1, 2);
          newSh.getRange(nextN,1,1,newActRow.length).setValues([newActRow]).setBackground('#e8eaf6');
          Logger.log('✅ Membre ajouté à onglet: '+newActId);
        } catch(eNew){ Logger.log('Onglet nouvelle act KO: '+eNew); }

        // ── Mettre \u00e0 jour l'onglet du moyen de paiement ──
        try {
          if(modeP && (modeP.toLowerCase().indexOf('cheque')>=0 || modeP.toLowerCase().indexOf('esp')>=0)){
            var payShName = modeP.toLowerCase().indexOf('cheque')>=0 ? 'Cheques 1' : '\u0045sp\u00e8ces';
            var paySh = ss3.getSheetByName(payShName);
            if(paySh && paySh.getLastRow()>1){
              var payData = paySh.getRange(2,1,paySh.getLastRow()-1,7).getValues();
              for(var pi=0; pi<payData.length; pi++){
                if(String(payData[pi][0]||'').trim()!==codeM) continue;
                // Espèces: col E(index 4)=montant, Chèques: col F(index 5)=montant
                var isEsp = payShName.toLowerCase().indexOf('esp')>=0;
                var montColIdx = isEsp ? 4 : 5;
                var montColNum = isEsp ? 5 : 6;
                var oldPayMont = parseFloat(payData[pi][montColIdx]||0)||0;
                var newPayMont = Math.max(0, Math.round((oldPayMont+diff)*100)/100);
                paySh.getRange(pi+2, montColNum).setValue(newPayMont).setFontWeight('bold').setFontColor('#c0392b');
                paySh.getRange(pi+2,1,1,7).setBackground('#e1bee7');
                Logger.log('✅ Paiement: '+payShName+' '+codeM+' '+oldPayMont+'→'+newPayMont);
                break;
            }
          }
          }
        } catch(ePay){ Logger.log('Paiement MAJ KO: '+ePay); }

        // ── Avoir (réglé + diff<0) ou correction montant (non réglé + diff!=0) ──
        var codeAvoir = '';
        if(diff !== 0 && emailRow) {
          try {
            var avoirMontant = Math.abs(diff);
            if(isPaid && diff < 0) {
              // Dossier réglé + ancienne activité plus chère (nouvelle moins chère) → avoir
              codeAvoir = genererCodeAvoir(ss3);
              var avoirSh = getOrCreateAvoirSheet(ss3);
              var dateAvoir = Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy');
              var ligneAvoir = [
                codeM, emailRow.nom, emailRow.prenom,
                emailRow.ancienne + ' → ' + newActNom,
                avoirMontant, dateAvoir, codeAvoir, avoirMontant, 'Disponible'
              ];
              avoirSh.getRange(Math.max(avoirSh.getLastRow()+1,2),1,1,ligneAvoir.length).setValues([ligneAvoir]);
              Logger.log('Avoir généré: '+codeAvoir+' — '+avoirMontant+' €');
            } else if(!isPaid) {
              // Dossier non réglé → pas d'avoir : le nouveau total (totalNet, recalculé
              // par calcTotalFamille) est déjà le bon montant à régler en une fois.
              Logger.log('Montant corrigé (non réglé): '+codeM+' diff='+diff+' € — pas d\'avoir, total à régler = '+totalNet+' €');
            }
          } catch(eAvoir){ Logger.log('Avoir KO: '+eAvoir); }
        }

        SpreadsheetApp.flush();

        // ── Email : 4 situations distinctes, croisement (réglé / non réglé) × (plus cher / moins cher) ──
        // 1. Réglé      + nouvelle plus chère  → complément à régler (juste la différence)
        // 2. Réglé      + nouvelle moins chère → avoir généré (code avoir, différence)
        // 3. Non réglé  + nouvelle plus chère  → total à régler (totalNet, montant complet)
        // 4. Non réglé  + nouvelle moins chère → total à régler (totalNet, montant complet, réduit)
        // (diff === 0 : cas neutre, pas l'une des 4 situations — message minimal, pas de bloc paiement)
        if(emailRow){
          var NON_SPORT_NOMS = ['theatre','peinture','couture','guitare'];
          var ancienneEtaitCulture = NON_SPORT_NOMS.some(function(k){ return (emailRow.ancienne||'').toLowerCase().indexOf(k)>=0; });
          var nouvelleEstSport    = !NON_SPORT_NOMS.some(function(k){ return (emailRow.nouvelle||'').toLowerCase().indexOf(k)>=0; });
          var qsHtml = (ancienneEtaitCulture && nouvelleEstSport)
            ? '<div style="background:#e3f2fd;border:2px solid #1a237e;border-radius:8px;padding:14px;margin:12px 0;">'
              + '<strong style="color:#1a237e;">⚕️ Questionnaire de Santé obligatoire</strong><br>'
              + '<p style="font-size:13px;color:#333;margin:8px 0;">Vous passez d’une activité culturelle à une activité sportive. Un <strong>QS FNSMR</strong> est désormais obligatoire.</p>'
              + '<p style="font-size:13px;color:#333;margin:0;">Connectez-vous sur le site FRI pour le signer en ligne, ou apportez un <strong>certificat médical</strong> à la première séance.</p>'
              + '</div>' : '';

          // Lien HelloAsso : complément (réglé) ou totalité (non réglé)
          var montantHA = 0;
          if(isPaid && diff > 0) montantHA = Math.round(diff * 100);
          else if(!isPaid && totalNet > 0) montantHA = Math.round(totalNet * 100);
          var lienHA = HELLOASSO_URL;
          try {
            if(montantHA > 0 && emailRow.email)
              lienHA = helloassoCreerLienPaiement(montantHA, emailRow.prenom, emailRow.nom,
                emailRow.email, codeM, 'Règlement ' + codeM, {adresse:'',cp:'',ville:''}) || HELLOASSO_URL;
          } catch(eHA){ lienHA = HELLOASSO_URL; }
          var montantHALabel = montantHA > 0 ? (montantHA/100).toFixed(2) + ' €' : '';
          var boutonHA = montantHA > 0 ? helloassoBoutonHtml(lienHA, montantHALabel) : '';

          // Tableau commun (ancienne / nouvelle activité)
          var tableActivites = "<table style='width:100%;border-collapse:collapse;margin:16px 0'>"
            + "<tr style='background:#f5f5f5'><td style='padding:8px 12px'><strong>Ancienne activité</strong></td><td style='padding:8px 12px'>"+emailRow.ancienne+" — <strong>"+oldTarif.toFixed(2)+" €</strong></td></tr>"
            + "<tr><td style='padding:8px 12px'><strong>Nouvelle activité</strong></td><td style='padding:8px 12px'>"+newActNom+" — "+emailRow.jour+" "+emailRow.heure+" — <strong>"+newTarif.toFixed(2)+" €</strong></td></tr>"
            + "</table>";

          var sujetSuffixe, situationHtml, diffLigneHtml;

          if(diff === 0) {
            // Cas neutre : changement d'activité à tarif identique
            sujetSuffixe = 'Activité modifiée';
            diffLigneHtml = '<span style="color:#555;">Aucune différence de tarif</span>';
            situationHtml = '';
          } else if(isPaid && diff > 0) {
            // 1. Réglé + nouvelle plus chère → complément à régler
            sujetSuffixe = 'Complément à régler — ' + diff.toFixed(2) + ' €';
            diffLigneHtml = '<span style="color:#c0392b;font-weight:bold;">+ '+diff.toFixed(2)+' € à régler</span>';
            situationHtml = '<div style="background:#fff3e0;border:1px solid #ffcc02;border-radius:8px;padding:14px;margin:12px 0;">'
              + '<strong>⚠️ Complément à régler : '+diff.toFixed(2)+' €</strong><br>'
              + '<span style="font-size:13px;color:#555;">Votre dossier est déjà réglé — merci de régler uniquement la différence via le bouton ci-dessous.</span>'
              + '</div>' + boutonHA;
          } else if(isPaid && diff < 0) {
            // 2. Réglé + nouvelle moins chère → avoir généré
            sujetSuffixe = 'Avoir généré — ' + Math.abs(diff).toFixed(2) + ' €';
            diffLigneHtml = '<span style="color:#2d6a4f;font-weight:bold;">Avoir de '+Math.abs(diff).toFixed(2)+' €</span>';
            situationHtml = codeAvoir
              ? '<div style="background:#e8f5e9;border:2px solid #2d6a4f;border-radius:8px;padding:14px;margin:12px 0;">'
                + '<strong style="color:#2d6a4f;">✅ Avoir généré : '+Math.abs(diff).toFixed(2)+' €</strong><br>'
                + 'Code avoir : <strong style="font-family:monospace;font-size:16px;letter-spacing:2px;">'+codeAvoir+'</strong><br>'
                + '<span style="font-size:12px;color:#555;">Ce code vous permettra d’utiliser cet avoir lors d’une prochaine inscription.</span>'
                + '</div>' : '';
          } else if(!isPaid && diff > 0) {
            // 3. Non réglé + nouvelle plus chère → total à régler (montant complet)
            sujetSuffixe = 'Règlement à effectuer — ' + totalNet.toFixed(2) + ' €';
            diffLigneHtml = '<span style="color:#e65100;font-weight:bold;">+ '+diff.toFixed(2)+' € (nouvelle activité plus chère)</span>';
            situationHtml = '<div style="background:#fff3e0;border:1px solid #e65100;border-radius:8px;padding:14px;margin:12px 0;">'
              + '<strong>⚠️ Règlement à effectuer : '+totalNet.toFixed(2)+' €</strong><br>'
              + '<span style="font-size:13px;color:#555;">Votre dossier n’est pas encore réglé. Merci de régler la <strong>totalité</strong> de votre cotisation (et non la seule différence) via le bouton ci-dessous.</span>'
              + '</div>' + boutonHA;
          } else {
            // 4. Non réglé + nouvelle moins chère → total à régler (montant complet, réduit)
            sujetSuffixe = 'Règlement à effectuer — ' + totalNet.toFixed(2) + ' €';
            diffLigneHtml = '<span style="color:#2d6a4f;font-weight:bold;">'+diff.toFixed(2)+' € (nouvelle activité moins chère)</span>';
            situationHtml = '<div style="background:#fff3e0;border:1px solid #e65100;border-radius:8px;padding:14px;margin:12px 0;">'
              + '<strong>⚠️ Règlement à effectuer : '+totalNet.toFixed(2)+' €</strong><br>'
              + '<span style="font-size:13px;color:#555;">Votre dossier n’est pas encore réglé. Le changement d’activité réduit le montant dû — merci de régler la <strong>totalité</strong> (montant déjà mis à jour) via le bouton ci-dessous.</span>'
              + '</div>' + boutonHA;
          }

          var corps = "<div style='font-family:sans-serif;max-width:600px;margin:auto'>"
            + "<h2 style='color:#2d6a4f'>Modification d’activité — "+emailRow.code+"</h2>"
            + "<p>Bonjour "+emailRow.prenom+" "+emailRow.nom+",</p>"
            + "<p>Votre activité a été modifiée pour <strong>"+emailRow.memPrenom+" "+emailRow.memNom+"</strong> :</p>"
            + tableActivites
            + "<table style='width:100%;border-collapse:collapse;margin:16px 0'>"
            + "<tr style='background:#e8f5e9'><td style='padding:8px 12px'><strong>Différence</strong></td><td style='padding:8px 12px'>"+diffLigneHtml+"</td></tr>"
            + "<tr style='background:#f9f9f9'><td style='padding:8px 12px'><strong>Nouveau total cotisation</strong></td><td style='padding:8px 12px;font-size:16px;font-weight:bold;color:#2d6a4f'>"+totalNet.toFixed(2)+" €</td></tr>"
            + "</table>"
            + situationHtml + qsHtml
            + blocCommentaireAdminHtml(commentaireAdminMod)
            + "<p style='color:#888;font-size:12px'>Foyer Rural d’Isneauville — frisneauville@orange.fr — 02.35.59.01.01</p>"
            + "</div>";
          var sujet = '[FRI] Modification activité — '+emailRow.code+' — '+emailRow.memPrenom+' '+emailRow.memNom+' — '+sujetSuffixe;
          envoyerEmail(emailRow.email, sujet, '', {htmlBody:corps, name:'FRI Inscriptions'});
          envoyerEmail(EMAIL_ADMIN, '[ADMIN] '+sujet, '', {htmlBody:corps, name:'FRI Inscriptions'});
          logCommentaireAdmin(ss3, 'Modification activité', codeM, memNom, memPrenom, newActNom,
            (diff !== 0 ? diff : newTarif), commentaireAdminMod, montantModifieMod);
        }

        var msg = rowsUpdated+' ligne(s) modifi\u00e9e(s) \u2014 total: '+totalNet.toFixed(2)+' \u20ac';
        if(diff!==0) msg += ' \u2014 diff: '+(diff>0?'+':'')+diff.toFixed(2)+' \u20ac';
        if(codeAvoir) msg += ' \u2014 avoir: '+codeAvoir;
        return repondreAvecCb({status:'ok',message:msg,totalNet:totalNet,diff:diff,codeAvoir:codeAvoir},null,null,callback);
      }catch(eMod){
        Logger.log('modifierActivite KO: '+eMod+' '+eMod.stack);
        return repondreAvecCb({status:'error',message:eMod.toString()},null,null,callback);
      }
    }

    if(payload.action==='basculerListeAttente'){
      try {
        var tarifBrutRecu = parseFloat(payload.tarifBrut||0);
        Logger.log('basculerListeAttente reçu — code:' + payload.code + ' actId:' + payload.actId + ' tarifBrut:' + tarifBrutRecu);
        var resBascule = basculerListeAttenteGAS(
          payload.code        || '',
          payload.actId       || '',
          payload.actNom      || '',
          payload.prenomMembre|| '',
          payload.nomMembre   || '',
          tarifBrutRecu,
          payload.commentaireAdmin || '',
          payload.montantModifie === true || payload.montantModifie === 'true'
        );
        return repondreAvecCb(resBascule, null, null, callback);
      } catch(eBascule) {
        Logger.log('basculerListeAttente KO: ' + eBascule);
        return repondreAvecCb({status:'error', message: eBascule.toString()}, null, null, callback);
      }
    }

    if(payload.action==='saveReglement'){
      try {
        var ssRI = SpreadsheetApp.openById(SHEET_ID);
        var res = sauvegarderReglementDrive(
          payload.code    || '',
          payload.nom     || '',
          payload.prenom  || '',
          payload.email   || '',
          payload.date    || '',
          payload.signature || '',
          payload.dossierOverride || ''
        );
        return repondreAvecCb({status:'ok', fichier: res.fichierNom}, null, null, callback);
      } catch(eRI) {
        Logger.log('saveReglement KO: ' + eRI.toString());
        return repondreAvecCb({status:'error', message: eRI.toString()}, null, null, callback);
      }
    }
    if(payload.action==='createCheckoutIntent'){
      try {
        var haResult = creerCheckoutIntentHA({
          totalAmount:     Math.round(Number(payload.totalAmount||0) * 100), // € → centimes
          itemName:        payload.itemName || 'Adhésion FRI 2026-2027',
          prenom:          payload.prenom   || '',
          nom:             payload.nom      || '',
          email:           payload.email    || '',
          codeDossier:     payload.codeDossier || '',
          returnUrl:       payload.returnUrl   || '',
          backUrl:         payload.backUrl     || '',
          errorUrl:        payload.errorUrl    || '',
          echelonne:       payload.echelonne   || false
        });
        return repondreAvecCb({status:'ok', redirectUrl: haResult.redirectUrl, checkoutIntentId: haResult.checkoutIntentId}, null, null, callback);
      } catch(eHA) {
        var errMsg = eHA.toString();
        Logger.log('createCheckoutIntent KO: ' + errMsg);
        // Si credentials manquants → message clair
        if (errMsg.indexOf('manquants') >= 0 || errMsg.indexOf('HA_CLIENT') >= 0) {
          return repondreAvecCb({status:'error', message:'Credentials HelloAsso non configurés. Ajoutez HA_CLIENT_ID et HA_CLIENT_SECRET dans les propriétés du script GAS.'}, null, null, callback);
        }
        return repondreAvecCb({status:'error', message: errMsg}, null, null, callback);
      }
    }
    if(payload.action==='addRegistration'){
      if(!payload.rows||payload.rows.length===0)return repondreAvecCb({status:'error',message:'Aucune donnee'},null,null,callback);
      if(payload.cheques&&payload.rows[0])payload.rows[0].cheques=payload.cheques;
      var result=addRegistration(payload.rows,payload.status||'paid',payload);
      return repondreAvecCb(reponseAddRegistration(result),null,null,callback);
    }
    if(payload.action==='getDossiers'){
      var result=getDossiersSheet();
      return repondreAvecCb({status:'ok',dossiers:result.dossiers},null,null,callback);
    }
    if(payload.action==='getStatsTresorier'){
      try {
        var statsT = getStatsTresorierGAS();
        return repondreAvecCb({status:'ok',stats:statsT},null,null,callback);
      } catch(eST) {
        Logger.log('getStatsTresorier KO: '+eST.toString());
        return repondreAvecCb({status:'error',message:eST.toString()},null,null,callback);
      }
    }
    if(payload.action==='envoyerAttestationPDF'){
      try{
        var pdfB64=payload.pdfBase64||'';
        if(!pdfB64)return repondreAvecCb({status:'ok'},null,null,callback);
        var nomFichier=String(payload.filename||'QS-Sante.pdf').replace(/[\\/:*?"<>|]/g,'_');
        var pdfBytes=Utilities.base64Decode(pdfB64);
        var pdfBlob=Utilities.newBlob(pdfBytes,'application/pdf',nomFichier);
        var nomDossierQS = String(payload.dossierOverride||'2-QS Santé Adhérents');
        var dossierQS=creerDossierSecurise(nomDossierQS);
        var fichierQS=dossierQS.createFile(pdfBlob);
        securiserFichier(fichierQS);
        Logger.log('✅ Attestation PDF enregistrée : ' + nomFichier + ' (' + pdfBytes.length + ' octets)');
        return repondreAvecCb({status:'ok',saved:nomFichier},null,null,callback);
      }catch(errQS){Logger.log('❌ envoyerAttestationPDF erreur : ' + errQS.toString()); return repondreAvecCb({status:'error',message:errQS.toString()},null,null);}
    }
    if(payload.action==='ajouterActiviteDossier'){
      var result3=ajouterActiviteDossierSheet(payload);
      return repondreAvecCb({status:'ok',inserted:result3.inserted,nouveauTotal:result3.nouveauTotal||0},null,null,callback);
    }
    if(payload.action==='supprimerActiviteDossier'){
      var result=supprimerActiviteDossierSheet(payload.code||'',payload.actNom||'',payload.actId||'',payload.membreNom||'',parseFloat(payload.avoir)||0,payload.placesId||'',payload.commentaireAdmin||'',payload.montantModifie===true||payload.montantModifie==='true');
      return repondreAvecCb({status:'ok',deleted:result.deleted},null,null,callback);
    }
    if(payload.action==='supprimerActiviteNonRegle'){
      var result2=supprimerActiviteNonRegleSheet(payload.code||'',payload.actNom||'',payload.actId||'',payload.membreNom||'',payload.placesId||'',parseFloat(payload.nouveauTotal)||0,payload.commentaireAdmin||'',payload.montantModifie===true||payload.montantModifie==='true');
      return repondreAvecCb({status:'ok',deleted:result2.deleted,nouveauTotal:result2.nouveauTotal||0},null,null,callback);
    }
    if(payload.action==='supprimerDossier'){
      var result=supprimerDossierSheet(payload.code||'');
      return repondreAvecCb({status:'ok',deleted:result.deleted},null,null,callback);
    }
    if(payload.action==='validerPaiementBascule'){
      var resBP=validerPaiementBasculeGAS(String(payload.code||''),String(payload.actId||''),String(payload.modePaiement||'cheque'),parseFloat(payload.montant||0),String(payload.commentaireAdmin||''),payload.montantModifie===true||payload.montantModifie==='true');
      return repondreAvecCb(resBP,null,null,callback);
    }
    if(payload.action==='validerPaiement'){
      var code=payload.code||params.code||'';
      var result=validerPaiementSheet(payload.rows||[],code);
      return repondreAvecCb({status:'ok',updated:result.updated},null,null,callback);
    }
    // PDF inscription admin — dans traiterRequete (doGet/JSONP)
    if(payload.action==='exportGestafillSheet'){
      try{
        var ssEx = SpreadsheetApp.openById(SHEET_ID);
        var shEx = ssEx.getSheetByName(SHEET_INSCRIPTIONS);
        if(!shEx||shEx.getLastRow()<2) return repondreAvecCb({status:'ok',rows:[]},null,null,callback);
        // Lire 42 colonnes (A-AP) — col AP (index 41) = export Gestafill
        var dataEx = shEx.getRange(2,1,shEx.getLastRow()-1,42).getValues();
        var seen = {};
        var rows = [];
        dataEx.forEach(function(r){
          var statut22 = String(r[21]||'').toLowerCase();
          var statut40 = String(r[39]||'').toLowerCase();
          // Exporter uniquement les dossiers validés (col V = Règlement validé, col AN = Inscrit)
          if(statut22.indexOf('supprimée')>=0) return;
          var estValide = statut22.indexOf('valid')>=0 || statut40.indexOf('inscrit')>=0;
          if(!estValide) return;
          // Col AP (index 41) = compteur export Gestafill — exclure si déjà exporté
          var exportCount = Number(r[41]||0);
          if(exportCount >= 1) return; // déjà exporté — ne pas réexporter
          var nom    = String(r[2]||'').trim();
          var prenom = String(r[3]||'').trim();
          // DDN : peut être un objet Date (Google Sheets) → formater en dd/MM/yyyy
          var ddnRaw = r[4];
          var ddn = '';
          if (ddnRaw instanceof Date) {
            var dd = ('0'+ddnRaw.getDate()).slice(-2);
            var mm = ('0'+(ddnRaw.getMonth()+1)).slice(-2);
            var yyyy = ddnRaw.getFullYear();
            ddn = dd+'/'+mm+'/'+yyyy;
          } else if (ddnRaw) {
            ddn = String(ddnRaw).trim();
          }
          var sexe   = String(r[36]||'').trim(); // col AK (index 36) = sexe membre
          var adresse= String(r[7]||'').trim();
          var cp     = String(r[9]||'').trim();
          var ville  = String(r[10]||'').trim();
          // Formater téléphone XX XX XX XX XX
          function fmtTel(t) {
            var d = String(t||'').replace(/[^0-9]/g,'');
            if(d.length===9) d = '0'+d; // compléter le 0 manquant
            if(d.length===10) return d.replace(/(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/,'$1 $2 $3 $4 $5');
            return String(t||'').trim();
          }
          var tel1   = fmtTel(r[14]);
          var tel2   = fmtTel(r[17]);
          var email  = String(r[15]||'').trim();
          // Col AL (index 37) = activite_id (placesId)
          var actId  = String(r[37]||'').trim();
          var code   = String(r[19]||'').trim();
          var qsSante = String(r[34]||'').trim();
          // Formater le statut QS pour Gestafill
          var qsLabel = '';
          if (qsSante === 'Attestation OK')    qsLabel = 'Attestation QS';
          else if (qsSante === 'Certificat requis') qsLabel = 'Certificat médical requis';
          else if (qsSante)                    qsLabel = qsSante;
          var key    = (nom+'_'+prenom+'_'+ddn).toLowerCase();
          if(seen[key]) {
            if(actId && seen[key].q.indexOf(actId)<0) seen[key].q += ' | '+actId;
            return;
          }
          var civilite = sexe.toUpperCase()==='F' ? 'Mme' : 'M.';
          var qsPart = qsLabel ? ' | '+qsLabel : '';
          var obj = {
            a:'', b:civilite, c:nom.toUpperCase(), d:prenom, e:ddn,
            f:'', g:'', h:adresse, i:'', j:cp, k:ville, l:'', m:'FR',
            n:'', o:tel1||tel2, p:email,
            q:code+(actId?' | '+actId:'')+qsPart
          };
          seen[key] = obj;
          rows.push(obj);
        });
        // ── Annoter col AP (index 42 dans getRange 1-based = col 42) pour chaque ligne exportée ──
        var exportDate = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
        for(var ei=0; ei<dataEx.length; ei++){
          var r2 = dataEx[ei];
          // Vérifier que ce membre a été exporté (présent dans seen)
          var nomE    = String(r2[2]||'').trim();
          var prenomE = String(r2[3]||'').trim();
          var ddnRaw2 = r2[4];
          var ddnE = ddnRaw2 instanceof Date ?
            ('0'+ddnRaw2.getDate()).slice(-2)+'/'+ ('0'+(ddnRaw2.getMonth()+1)).slice(-2)+'/'+ddnRaw2.getFullYear() :
            String(ddnRaw2||'').trim();
          var keyE = (nomE+'_'+prenomE+'_'+ddnE).toLowerCase();
          if(seen[keyE]) {
            var curCount = Number(r2[41]||0) + 1;
            shEx.getRange(ei+2, 42).setValue('Gestafill '+curCount+' — '+exportDate);
          }
        }
        Logger.log('exportGestafillSheet — '+rows.length+' membres exportés, col AP annotée');
        return repondreAvecCb({status:'ok', rows:rows},null,null,callback);
      }catch(eEx){
        Logger.log('exportGestafillSheet KO: '+eEx);
        return repondreAvecCb({status:'error',message:eEx.toString()},null,null,callback);
      }
    }

    if(payload.action==='supprimerFichiersOrphelins'){
      try{
        var codeOrph = String(payload.code||'');
        if(!codeOrph || !codeOrph.match(/^FRI-[A-Z0-9]{4}$/)) {
          return repondreAvecCb({status:'error',message:'Code invalide'},null,null,callback);
        }
        // Vérifier que le dossier N'EST PAS dans le Sheet (sinon ne pas supprimer)
        var ssOrph = SpreadsheetApp.openById(SHEET_ID);
        var shOrph = ssOrph.getSheetByName(SHEET_INSCRIPTIONS);
        var estInscrit = false;
        if(shOrph && shOrph.getLastRow() > 1) {
          var dataOrph = shOrph.getRange(2,20,shOrph.getLastRow()-1,1).getValues();
          for(var oi=0; oi<dataOrph.length; oi++) {
            if(String(dataOrph[oi][0]||'').trim() === codeOrph) { estInscrit = true; break; }
          }
        }
        if(estInscrit) {
          Logger.log('supprimerFichiersOrphelins — dossier inscrit, pas de suppression : ' + codeOrph);
          return repondreAvecCb({status:'ok',message:'Dossier inscrit — fichiers conservés'},null,null,callback);
        }
        // Supprimer les fichiers QS et RI orphelins
        var supprimés = [];
        var dossiers = [
          {nom:'2-QS Santé Adhérents'},
          {nom:'3-Règlements intérieurs'}
        ];
        dossiers.forEach(function(d) {
          var folders = dossiersDriveParNom(d.nom);
          if(!folders.hasNext()) return;
          var folder = folders.next();
          var files = folder.getFiles();
          while(files.hasNext()) {
            var f = files.next();
            var fname = f.getName();
            if(fname.indexOf(codeOrph) >= 0) {
              f.setTrashed(true);
              supprimés.push(fname);
              Logger.log('✅ Fichier orphelin supprimé : ' + fname);
            }
          }
        });
        return repondreAvecCb({status:'ok', supprimes: supprimés, nb: supprimés.length}, null, null, callback);
      }catch(eOrph){
        Logger.log('supprimerFichiersOrphelins KO: '+eOrph);
        return repondreAvecCb({status:'error',message:eOrph.toString()},null,null,callback);
      }
    }

    // ── « Créer un compte » : la famille a-t-elle déjà un dossier cette saison ? ──
    if(payload.action==='chercherDossierFamille'){
      return repondreAvecCb(chercherDossierFamilleGAS(payload.email, payload.nom),null,null,callback);
    }
    // ── Espace famille (menu d'accueil) : consulter son dossier ──
    if(payload.action==='consulterDossier'){
      var resCD = lireDossierFamille(payload.code, payload.email);
      return repondreAvecCb(resCD,null,null,callback);
    }
    if(payload.action==='piecesFamille'){
      return repondreAvecCb(piecesFamilleGAS(payload.code, payload.email),null,null,callback);
    }
    // ── Espace famille : demande de modification des activités (traitée par l'admin) ──
    if(payload.action==='demanderModification'){
      return repondreAvecCb(demanderModificationGAS(payload),null,null,callback);
    }
    // ── Console admin : demandes de modification des familles ──
    // ── Coches de suivi de la console admin, partagées entre postes et comptes ──
    if(payload.action==='getSuiviAdmin'){
      return repondreAvecCb(lireSuiviAdmin(payload.code ? String(payload.code).trim().toUpperCase() : ''),null,null,callback);
    }
    if(payload.action==='majSuiviAdmin'){
      return repondreAvecCb(majSuiviAdminGAS(payload.changements, sessionAdmin),null,null,callback);
    }
    if(payload.action==='getPiecesDossier'){
      return repondreAvecCb(controlerPiecesDossier(String(payload.code||'').trim().toUpperCase()),null,null,callback);
    }
    if(payload.action==='getDemandesModification'){
      return repondreAvecCb(lireDemandesModification(),null,null,callback);
    }
    if(payload.action==='traiterDemandeModification'){
      return repondreAvecCb(traiterDemandeModificationGAS(payload, sessionAdmin),null,null,callback);
    }
    // ── Espace famille : message libre au Foyer Rural ──
    if(payload.action==='envoyerMessage'){
      return repondreAvecCb(envoyerMessageContactGAS(payload),null,null,callback);
    }

    // ── Activités d'un dossier existant (formulaire public « J'ai déjà un dossier ») ──
    // Remplace la lecture publique de la feuille (gviz) : seules les activités sont renvoyées.
    if(payload.action==='getActivitesDossier'){
      try{
        var codeAD = String(payload.code||'').trim().toUpperCase();
        if(!/^FRI-[A-Z0-9]{4}$/.test(codeAD))
          return repondreAvecCb({status:'error',message:'Format invalide'},null,null,callback);
        var shAD = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_INSCRIPTIONS);
        var activitesAD = [];
        if(shAD && shAD.getLastRow() > 1){
          var nbColAD = Math.min(Math.max(shAD.getLastColumn(), 40), 45);
          var dataAD = shAD.getRange(2,1,shAD.getLastRow()-1,nbColAD).getValues();
          dataAD.forEach(function(r){
            if(String(r[19]||'').trim() !== codeAD) return;
            var statutAD = lireStatutInscription(r);
            if(statutAD.toLowerCase().indexOf('supprim') >= 0) return;
            activitesAD.push({
              activite_id:        lireActiviteId(r),
              activite:           String(r[22]||'').trim(),
              tarif:              Number(r[27]||0),
              statut_inscription: statutAD
            });
          });
        }
        return repondreAvecCb({status:'ok', code:codeAD, activites:activitesAD},null,null,callback);
      }catch(eAD){
        Logger.log('getActivitesDossier KO: '+eAD);
        return repondreAvecCb({status:'error',message:'Erreur serveur'},null,null,callback);
      }
    }

    if(payload.action==='verifierDossier'){
      try{
        var codeVD = String(payload.code||'').trim().toUpperCase();
        if(!codeVD.match(/^FRI-[A-Z0-9]{4}$/))
          return repondreAvecCb({status:'error',message:'Format invalide'},null,null,callback);
        var ssVD = SpreadsheetApp.openById(SHEET_ID);
        var shVD = ssVD.getSheetByName(SHEET_INSCRIPTIONS);
        if(!shVD||shVD.getLastRow()<2)
          return repondreAvecCb({status:'error',message:'Aucune inscription'},null,null,callback);
        var dataVD = shVD.getRange(2,1,shVD.getLastRow()-1,40).getValues();
        // Structure colonnes (0-based index) :
        // 0=Licence 1=Civilité 2=Nom membre 3=Prénom membre 4=DDN 9=CP 10=Ville
        // 14=Tél 15=Email 19=N°Dossier 21=Statut paiement 36=Responsable(Prenom NOM)
        var found = null;
        for(var vi=0; vi<dataVD.length; vi++){
          if(String(dataVD[vi][19]||'').trim()===codeVD){
            // Responsable : col AK (index 36) format "Prénom NOM"
            var respRaw   = String(dataVD[vi][36]||'').trim();
            var respParts = respRaw.split(' ');
            var respPrenom = respParts[0] || '';
            var respNom    = respParts.slice(1).join(' ') || '';
            // DDN : col E (index 4)
            var ddnRaw = dataVD[vi][4];
            var ddnStr = '';
            if(ddnRaw instanceof Date) {
              ddnStr = Utilities.formatDate(ddnRaw, 'Europe/Paris', 'dd/MM/yyyy');
            } else if(ddnRaw) {
              ddnStr = String(ddnRaw).trim();
            }
            // Action publique : on ne renvoie que de quoi confirmer le dossier
            // (prénom + initiale du nom), jamais l'email, le téléphone ou la date de naissance.
            var initiale = function(n) { n = String(n||'').trim(); return n ? n.charAt(0).toUpperCase() + '.' : ''; };
            found = {
              code:        codeVD,
              membreNom:   initiale(dataVD[vi][2]),
              membrePrenom:String(dataVD[vi][3]||'').trim(),
              respPrenom:  respPrenom,
              respNom:     initiale(respNom)
            };
            break;
          }
        }
        if(!found)
          return repondreAvecCb({status:'error',message:'Dossier '+codeVD+' introuvable'},null,null,callback);
        Logger.log('verifierDossier OK : '+codeVD);
        return repondreAvecCb({status:'ok', dossier: found},null,null,callback);
      }catch(eVD){
        Logger.log('verifierDossier KO: '+eVD);
        return repondreAvecCb({status:'error',message:eVD.toString()},null,null,callback);
      }
    }

    if(payload.action==='genererPDFInscriptionAdmin'){
      try{
        var codeP = String(payload.code||'');
        if(!codeP) return repondreAvecCb({status:'error',message:'Code manquant'},null,null,callback);
        var ssP = SpreadsheetApp.openById(SHEET_ID);
        var shP = ssP.getSheetByName(SHEET_INSCRIPTIONS);
        if(!shP||shP.getLastRow()<2) return repondreAvecCb({status:'error',message:'Aucune inscription'},null,null,callback);
        var dataP = shP.getRange(2,1,shP.getLastRow()-1,41).getValues();
        var rowsP = dataP.filter(function(r){ return String(r[19]||'').trim()===codeP; });
        if(rowsP.length===0) return repondreAvecCb({status:'error',message:'Dossier '+codeP+' introuvable'},null,null,callback);
        var emailRowsP = rowsP.map(function(r){
          return {
            code_dossier:codeP, responsable_nom:String(r[1]||''), responsable_prenom:String(r[2]||''),
            email1:String(r[15]||''), tel1:String(r[14]||''), ville:String(r[11]||''), cp:String(r[9]||''),
            activite:String(r[22]||''), jour:String(r[23]||''), heure:String(r[24]||''),
            lieu:String(r[25]||''), tarif_brut:Number(r[27]||0), tarif_net:Number(r[29]||0),
            statut_inscription:String(r[39]||''), mode_paiement:String(r[32]||'helloasso'),
            membre_nom:String(r[3]||''), membre_prenom:String(r[4]||''), ddn:formaterDdn(r[5]),
            date:formaterDateHeure(r[20]), qs_sante:String(r[34]||''),
            fftt_price:Number(r[36]||0), total_famille:Number(r[30]||0),
            commune:String(r[11]||'').toLowerCase().indexOf('isneauville')>=0?'isno':'hc'
          };
        });
        var modeLabelP = {helloasso:'HelloAsso',cheque:'Chèque',cheque3:'Chèques échelonnés',
                          especes:'Espèces',ancv:'Coupon ANCV',aide:'Soldé par aides'}[emailRowsP[0].mode_paiement]||'HelloAsso';
        var pdfBlob = genererFacturePDF(emailRowsP, modeLabelP);
        if(!pdfBlob) return repondreAvecCb({status:'error',message:'Génération PDF échouée'},null,null,callback);
        var fname = codeP+'_'+emailRowsP[0].responsable_prenom+'-'+emailRowsP[0].responsable_nom+'_Inscription.pdf';
        // Sauvegarder dans Drive et renvoyer l'URL + base64 si taille raisonnable
        var pdfBytes = pdfBlob.getBytes();
        Logger.log('✅ PDF admin généré : '+fname+' ('+pdfBytes.length+' bytes)');
        // Toujours sauvegarder dans Drive (backup)
        var dossierFact = obtenirDossierFRI('2-Dossiers');
        pdfBlob.setName(fname);
        var fileP = dossierFact.createFile(pdfBlob);
        securiserFichier(fileP);
        // Si PDF < 600Ko encodé base64 : renvoyer base64 pour ouverture locale
        // Sinon : renvoyer uniquement l'URL Drive
        var response = {status:'ok', message:'PDF généré', filename: fname};
        if (pdfBytes.length < 450000) {
          response.pdfBase64 = Utilities.base64Encode(pdfBytes);
        } else {
          response.pdfUrl = fileP.getUrl();
          response.message = 'PDF trop volumineux pour ouverture locale — ouvert depuis Drive';
        }
        return repondreAvecCb(response, null, null, callback);
      }catch(ePDF){
        Logger.log('genererPDFInscriptionAdmin KO: '+ePDF);
        return repondreAvecCb({status:'error',message:ePDF.toString()},null,null,callback);
      }
    }

    if(payload.action==='renvoyerEmailInscription'){
      try{
        var codeRI = String(payload.code||'');
        if(!codeRI) return repondreAvecCb({status:'error',message:'Code dossier manquant'},null,null,callback);
        var ssRI = SpreadsheetApp.openById(SHEET_ID);
        var shRI = ssRI.getSheetByName(SHEET_INSCRIPTIONS);
        if(!shRI||shRI.getLastRow()<2) return repondreAvecCb({status:'error',message:'Aucune inscription'},null,null,callback);
        var dataRI = shRI.getRange(2,1,shRI.getLastRow()-1,41).getValues();
        // Récupérer toutes les lignes du dossier
        var rowsRI = dataRI.filter(function(r){ return String(r[19]||'').trim()===codeRI; });
        if(rowsRI.length===0) return repondreAvecCb({status:'error',message:'Dossier '+codeRI+' introuvable'},null,null,callback);
        // Reconstruire les rows au format attendu par envoyerEmailAdherent
        var emailRows = rowsRI.map(function(r){
          return {
            code_dossier:      codeRI,
            responsable_nom:   String(r[1]||''),
            responsable_prenom:String(r[2]||''),
            email1:            String(r[15]||''),
            tel1:              String(r[14]||''),
            ville:             String(r[11]||''),
            cp:                String(r[9]||''),
            activite:          String(r[22]||''),
            jour:              String(r[23]||''),
            heure:             String(r[24]||''),
            lieu:              String(r[25]||''),
            tarif_brut:        Number(r[27]||0),
            tarif_net:         Number(r[29]||0),
            statut_inscription:String(r[39]||''),
            mode_paiement:     String(r[32]||'helloasso'),
            membre_nom:        String(r[3]||''),
            membre_prenom:     String(r[4]||''),
            ddn:               formaterDdn(r[5]),
            sexe:              String(r[6]||''),
            animateur:         String(r[26]||''),
            date:              formaterDateHeure(r[20]),
            qs_sante:          String(r[34]||''),   // col AI index 34
            pass_aide:         String(r[35]||''),   // col AJ index 35
            fftt_price:        Number(r[36]||0),    // col AK index 36
            fftt_type:         String(r[37]||''),   // col AL index 37
            total_famille:     Number(r[30]||0),    // col AE index 30
            commune:           String(r[10]||'').toLowerCase().indexOf('isneauville')>=0?'isno':'hc'
          };
        });
        var modeLabel = {helloasso:'HelloAsso',cheque:'Chèque',cheque3:'Chèques échelonnés',
                         especes:'Espèces',ancv:'Coupon ANCV',aide:'Soldé par aides'}[emailRows[0].mode_paiement]||'HelloAsso';
        try{
          envoyerEmailAdherent(emailRows[0].email1, emailRows, modeLabel);
          Logger.log('✅ Email inscription renvoyé : '+codeRI+' → '+emailRows[0].email1);
        }catch(eM){ Logger.log('Email adherent KO: '+eM); }
        return repondreAvecCb({status:'ok',message:'Email renvoyé à '+emailRows[0].email1},null,null,callback);
      }catch(eREI){
        Logger.log('renvoyerEmailInscription KO: '+eREI);
        return repondreAvecCb({status:'error',message:eREI.toString()},null,null,callback);
      }
    }

    if(payload.action==='envoyerRappelPieces'){
      try{
        var code      = String(payload.code   || '');
        var email     = String(payload.email  || '');
        var nom       = String(payload.nom    || code);
        var pieces    = payload.piecesAttente || [];
        if(!email){ return repondreAvecCb({status:'error',message:'Email adhérent introuvable'},null,null,callback); }
        if(!pieces.length){ return repondreAvecCb({status:'error',message:'Aucune pièce en attente'},null,null,callback); }

        // Construire la liste des pieces manquantes (entites HTML, pas d'emojis)
        var listHtml = pieces.map(function(p){ return '<li style="margin:4px 0;color:#c62828;">&bull; '+p+'</li>'; }).join('');
        var listTxt  = pieces.map(function(p){ return '  - '+p; }).join('\n');

        // Corps HTML du mail - charset UTF-8 explicite, entites HTML
        var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>'
          + '<div style="font-family:sans-serif;max-width:600px;margin:auto;">'
          + '<div style="background:#1b5e20;padding:18px 24px;border-radius:8px 8px 0 0;">'
          + '<h2 style="color:#fff;margin:0;font-size:18px;">Dossier '+code+' - Pièce(s) manquante(s)</h2>'
          + '<p style="color:#a5d6a7;margin:6px 0 0;font-size:13px;">'+NOM_ASSO+' - Saison 2026/2027</p>'
          + '</div>'
          + '<div style="background:#fff8e1;border:1px solid #ffe082;padding:16px 24px;">'
          + '<p style="color:#333;font-size:14px;">Bonjour '+nom+',</p>'
          + '<p style="color:#333;font-size:14px;">Votre dossier d&#39;inscription <strong>'+code+'</strong> est bien enregistré.<br>'
          + 'Cependant, il nous manque encore la ou les pièce(s) suivante(s) pour valider votre inscription :</p>'
          + '<ul style="margin:12px 0;padding-left:20px;font-size:14px;">'+listHtml+'</ul>'
          + '<p style="color:#333;font-size:14px;">Merci de nous les faire parvenir :</p>'
          + '<ul style="font-size:13px;color:#555;">'
          + '<li>En permanence : <strong>mardi 16h30-18h30</strong> (période scolaire) - Salle des fêtes, Place A. Cramilly, Isneauville</li>'
          + '<li>Par email : <a href="mailto:frisneauville@orange.fr">frisneauville@orange.fr</a></li>'
          + '</ul>'
          + '<p style="color:#555;font-size:13px;margin-top:16px;">Votre inscription ne sera définitivement validée qu&#39;à réception de l&#39;ensemble des pièces et du règlement.</p>'
          + '</div>'
          + '<div style="background:#f5f5f5;padding:12px 24px;border-radius:0 0 8px 8px;font-size:11px;color:#888;text-align:center;">'
          + NOM_ASSO+' - Tel : 02.35.59.01.01 - frisneauville@orange.fr'
          + '</div></div></body></html>';

        var sujet = '[FRI] Dossier '+code+' - Piece(s) manquante(s)';
        var txtBody = 'Bonjour '+nom+',\n\nVotre dossier '+code+' est enregistre mais il manque :\n'+listTxt
          + '\n\nMerci de nous les apporter en permanence (mardi 16h30-18h30) ou par email a frisneauville@orange.fr.\n\n'+NOM_ASSO;

        envoyerEmail(email, sujet, txtBody, {
          htmlBody: htmlBody, name: NOM_ASSO, replyTo: EMAIL_ADMIN
        });
        // Copie admin
        envoyerEmail(EMAIL_ADMIN, '[ADMIN] Rappel envoyé — '+code+' — '+nom, txtBody, { name: NOM_ASSO });
        Logger.log('✅ Rappel pièces envoyé : '+code+' → '+email+' ('+pieces.length+' pièce(s))');
        return repondreAvecCb({status:'ok', message:'Email envoyé à '+email},null,null,callback);
      }catch(eRappel){
        Logger.log('❌ envoyerRappelPieces KO : '+eRappel);
        return repondreAvecCb({status:'error',message:eRappel.toString()},null,null,callback);
      }
    }

    if(payload.action==='envoyerRappelPieces'){
      try{
        var code2     = String(payload.code   || '');
        var email2    = String(payload.email  || '');
        var nom2      = String(payload.nom    || code2);
        var pieces2   = payload.piecesAttente || [];
        if(!email2){ return repondreAvecCb({status:'error',message:'Email adhérent introuvable'},null,null,callback); }
        if(!pieces2.length){ return repondreAvecCb({status:'error',message:'Aucune pièce en attente'},null,null,callback); }
        var listHtml2 = pieces2.map(function(p){ return '<li style="margin:4px 0;color:#c62828;">&bull; '+p+'</li>'; }).join('');
        var listTxt2  = pieces2.map(function(p){ return '  - '+p; }).join('\n');
        var htmlBody2 = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>'
          + '<div style="font-family:sans-serif;max-width:600px;margin:auto;">'
          + '<div style="background:#1b5e20;padding:18px 24px;border-radius:8px 8px 0 0;">'
          + '<h2 style="color:#fff;margin:0;font-size:18px;">Dossier '+code2+' - Pièce(s) manquante(s)</h2>'
          + '<p style="color:#a5d6a7;margin:6px 0 0;font-size:13px;">'+NOM_ASSO+' - Saison 2026/2027</p>'
          + '</div>'
          + '<div style="background:#fff8e1;border:1px solid #ffe082;padding:16px 24px;">'
          + '<p style="color:#333;font-size:14px;">Bonjour '+nom2+',</p>'
          + '<p style="color:#333;font-size:14px;">Votre dossier d\'inscription <strong>'+code2+'</strong> est bien enregistré.<br>'
          + 'Cependant, il nous manque encore la ou les pièce(s) suivante(s) :</p>'
          + '<ul style="margin:12px 0;padding-left:20px;font-size:14px;">'+listHtml2+'</ul>'
          + '<p style="color:#333;font-size:14px;">Merci de nous les faire parvenir :</p>'
          + '<ul style="font-size:13px;color:#555;">'
          + '<li>En permanence : <strong>mardi 16h30-18h30</strong> (période scolaire) - Salle des fêtes, Place A. Cramilly, Isneauville</li>'
          + '<li>Par email : <a href="mailto:frisneauville@orange.fr">frisneauville@orange.fr</a></li>'
          + '</ul></div>'
          + '<div style="background:#f5f5f5;padding:12px 24px;border-radius:0 0 8px 8px;font-size:11px;color:#888;text-align:center;">'
          + NOM_ASSO+' - Tel : 02.35.59.01.01 - frisneauville@orange.fr'
          + '</div></div></body></html>';
        var sujet2 = '[FRI] Dossier '+code2+' - Piece(s) manquante(s)';
        var txtBody2 = 'Bonjour '+nom2+',\n\nVotre dossier '+code2+' est enregistr\u00e9 mais il manque :\n'+listTxt2
          + '\n\nMerci de nous les apporter en permanence (mardi 16h30-18h30) ou par email a frisneauville@orange.fr.\n\n'+NOM_ASSO;
        envoyerEmail(email2, sujet2, txtBody2, { htmlBody: htmlBody2, name: NOM_ASSO, replyTo: EMAIL_ADMIN });
        Logger.log('✅ Rappel pièces envoyé (doPost) : '+code2+' → '+email2);
        return repondreAvecCb({status:'ok', message:'Email envoyé à '+email2},null,null,callback);
      }catch(eR2){
        Logger.log('❌ envoyerRappelPieces doPost KO : '+eR2);
        return repondreAvecCb({status:'error',message:eR2.toString()},null,null,callback);
      }
    }

    if(payload.action==='ecrireCheques'){
      try{
        var ss2=SpreadsheetApp.openById(SHEET_ID);
        var chData=payload.cheques||[];
        var code2=payload.code||'',nom2=payload.nom||'',prenom2=payload.prenom||'';
        if(chData.length===1){
          var sh=getOrCreateChequeSheet(ss2,SHEET_CHEQUE_1);
          var nr1=Math.max(sh.getLastRow()+1,4);
          sh.getRange(nr1,1,1,7).setValues([[code2,nom2,prenom2,String(chData[0].banque||''),String(chData[0].numCheque||''),Number(chData[0].montant||0),'⏳ En cours de validation']]);
          sh.getRange(nr1,7).setFontColor('#856404').setFontWeight('bold');
          sh.autoResizeColumns(1,6);
          majTotalCheque(ss2,SHEET_CHEQUE_1);
        } else {
          var noms2=[SHEET_CHEQUE_1,SHEET_CHEQUE_2,SHEET_CHEQUE_3];
          chData.forEach(function(ch,i){
            if(i>=3)return;
            var sh2=getOrCreateChequeSheet(ss2,noms2[i]);
            var nr2=Math.max(sh2.getLastRow()+1,4);
            sh2.getRange(nr2,1,1,7).setValues([[code2,nom2,prenom2,String(ch.banque||''),String(ch.numCheque||''),Number(ch.montant||0),'⏳ En cours de validation']]);
            sh2.getRange(nr2,7).setFontColor('#856404').setFontWeight('bold');
            sh2.autoResizeColumns(1,6);
            majTotalCheque(ss2,noms2[i]);
          });
        }
        return repondreAvecCb({status:'ok'},null,null,callback);
      }catch(e3){return repondreAvecCb({status:'error',message:String(e3)},null,null);}
    }

        return repondreAvecCb({status:'error',message:'Action inconnue: '+payload.action},null,null,callback);
  }catch(err){
    Logger.log('ERREUR traiterRequete: '+err.toString());
    return repondreAvecCb({status:'error',message:err.toString()},null,null);
  }
}

// ── Vérification token de session admin ──
// Actions réservées à l'équipe FRI : elles exigent le jeton de session obtenu par adminLogin
// (payload._adminToken). Le jeton dynamique "getToken" est public et ne suffit pas.
var ACTIONS_ADMIN = [
  'getDossiers', 'getDossierDetail', 'modifierActivite', 'basculerListeAttente',
  'creerAvoirManuel', 'creerRemboursementManuel', 'envoyerRappelManuel', 'envoyerRappelPieces',
  'exportGestafillSheet', 'genererPDFInscriptionAdmin', 'getStatsTresorier',
  'getJournalSauvegardes', 'viderJournalSauvegardes', 'verifierDossiersPerdus',
  'renvoyerEmailInscription', 'supprimerActiviteDossier', 'supprimerActiviteNonRegle',
  'supprimerDossier', 'validerPaiement', 'validerPaiementBascule',
  'getElementsPaiement', 'validerElementPaiement', 'ajouterActiviteDossier', 'ecrireCheques',
  'getDemandesModification', 'traiterDemandeModification', 'getPiecesDossier',
  'getSuiviAdmin', 'majSuiviAdmin'
];
var REPONSE_ADMIN_REQUISE = {
  status: 'error', code: 'ADMIN_AUTH',
  message: 'Session administrateur requise ou expirée — reconnectez-vous.'
};

function verifierTokenAdmin(token) {
  if (!token || typeof token !== 'string') return null;
  var parts = token.split(':');
  if (parts.length !== 4) return null;
  var user = parts[0], role = parts[1], expiry = parts[2], sig = parts[3];
  var now = Math.floor(Date.now() / 1000);
  if (parseInt(expiry) < now) { Logger.log('⛔ Token admin expiré'); return null; }
  var props = PropertiesService.getScriptProperties();
  var sessionSecret = props.getProperty('ADMIN_SESSION_SECRET') || FRI_SECRET_TOKEN;
  if (!sessionSecret) { Logger.log('⛔ Aucun secret de session configuré'); return null; }
  var payload2 = user + ':' + role + ':' + expiry;
  var sigExpected = Utilities.computeHmacSha256Signature(payload2, sessionSecret);
  var sigHex = sigExpected.map(function(b){return('0'+(b&0xff).toString(16)).slice(-2);}).join('');
  if (sig !== sigHex) { Logger.log('⛔ Token admin signature invalide'); return null; }
  // Compte supprimé ou rôle modifié depuis la connexion → session invalide
  if (props.getProperty('ADMIN_ROLE_' + user) !== role) { Logger.log('⛔ Compte admin inconnu ou rôle modifié : ' + user); return null; }
  return { user: user, role: role };
}

// ── Initialisation des credentials admin (à exécuter UNE FOIS manuellement) ──
// Instructions : modifier les valeurs ci-dessous puis exécuter cette fonction
// depuis l'éditeur Apps Script > Exécuter > initAdminCredentials
function diagnosticAdminCredentials() {
  // Exécuter depuis Apps Script > Exécuter > diagnosticAdminCredentials
  // pour voir ce qui est stocké dans PropertiesService
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var adminKeys = Object.keys(all).filter(function(k){ return k.indexOf('ADMIN_') === 0; });
  if (adminKeys.length === 0) {
    Logger.log('⚠️ AUCUNE propriété ADMIN_ trouvée — initAdminCredentials() n\'a pas encore été exécuté');
  } else {
    Logger.log('✅ Propriétés ADMIN_ trouvées (' + adminKeys.length + ') :');
    adminKeys.forEach(function(k){ Logger.log('  ' + k + ' = ' + String(all[k]).substring(0,12) + '...'); });
  }
  Logger.log('FRI_SECRET_TOKEN présent : ' + (FRI_SECRET_TOKEN ? 'OUI ('+FRI_SECRET_TOKEN.substring(0,6)+'...)' : 'NON'));
}

function initAdminCredentials() {
  var props = PropertiesService.getScriptProperties();
  // ── MODIFIER CES VALEURS AVANT D'EXÉCUTER ──
  var comptes = [
    { user: 'IDENTIFIANT_ADMIN', pass: 'CHANGER_MOT_DE_PASSE_1', role: 'admin' },
    { user: 'IDENTIFIANT_SECRETARIAT_1', pass: 'CHANGER_MOT_DE_PASSE_2', role: 'secretariat' },
    { user: 'IDENTIFIANT_SECRETARIAT_2', pass: 'CHANGER_MOT_DE_PASSE_3', role: 'secretariat' },
  ];
  // ── Secret de session unique (regénéré à chaque appel) ──
  var sessionSecret = Utilities.base64Encode(
    Utilities.computeHmacSha256Signature(String(Date.now()), FRI_SECRET_TOKEN)
  );
  props.setProperty('ADMIN_SESSION_SECRET', sessionSecret);
  comptes.forEach(function(c) {
    // Hash du mot de passe : HMAC-SHA256(pass, FRI_SECRET_TOKEN) en hex
    // Même algorithme utilisé dans adminLogin pour la comparaison
    var sig = Utilities.computeHmacSha256Signature(c.pass, FRI_SECRET_TOKEN);
    var passHex = sig.map(function(b){return('0'+(b&0xff).toString(16)).slice(-2);}).join('');
    var userKey = c.user.toLowerCase();
    props.setProperty('ADMIN_CRED_' + userKey, passHex);
    props.setProperty('ADMIN_ROLE_' + userKey, c.role);
    Logger.log('✅ Compte ' + c.user + ' (' + c.role + ') initialisé — hash: ' + passHex.substring(0,8) + '...');
  });
  Logger.log('✅ initAdminCredentials terminé — ' + comptes.length + ' compte(s) configuré(s)');
  Logger.log('⚠️ Effacez maintenant les mots de passe en clair de cette fonction !');
}

// Limite globale de requêtes POST par heure. L'ancien seuil (200/h, tous visiteurs confondus)
// bloquait le site entier dès qu'une vingtaine de familles s'inscrivaient en même temps.
var RATE_LIMIT_PAR_HEURE = 1500;
function verifierRateLimit(e) {
  try{
    var cache=CacheService.getScriptCache();
    var clef='rl_'+Math.floor(Date.now()/3600000);
    var compteur=parseInt(cache.get(clef)||'0',10);
    if(compteur>=RATE_LIMIT_PAR_HEURE){Logger.log('⛔ Rate limit atteint : '+compteur);return false;}
    cache.put(clef,String(compteur+1),3700); // expire tout seul après l'heure écoulée
    return true;
  }catch(err){return true;}
}

// Nettoyage manuel à exécuter une fois dans l'éditeur GAS pour vider les rl_ accumulés
function nettoyerRateLimitProps() {
  var props=PropertiesService.getScriptProperties();
  var allProps=props.getProperties();
  var heure=Math.floor(Date.now()/3600000);
  var deleted=0;
  Object.keys(allProps).forEach(function(k){
    if(k.indexOf('rl_')===0){
      var h=parseInt(k.substring(3));
      if(!isNaN(h)&&h<heure-1){props.deleteProperty(k);deleted++;}
    }
  });
  Logger.log('✅ '+deleted+' propriété(s) rl_ supprimée(s)');
}

// ============================================================
// SÉCURISATION DRIVE
// ============================================================
function securiserFichier(fichier) {
  try{fichier.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.NONE);fichier.getEditors().forEach(function(u){try{fichier.removeEditor(u);}catch(x){}});fichier.getViewers().forEach(function(u){try{fichier.removeViewer(u);}catch(x){}});}catch(e){Logger.log('securiserFichier KO : '+e.toString());}
}
function securiserDossier(dossier) {
  try{dossier.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.NONE);var fichiers=dossier.getFiles(),nb=0;while(fichiers.hasNext()){securiserFichier(fichiers.next());nb++;}Logger.log('Dossier "'+dossier.getName()+'" sécurisé — '+nb+' fichier(s)');}catch(e){Logger.log('securiserDossier KO : '+e.toString());}
}
function creerDossierSecurise(nom) {
  var dossiers=dossiersDriveParNom(nom);
  var dossier;
  if(dossiers.hasNext()){
    dossier=dossiers.next();
    // Ne PAS re-sécuriser tout le dossier à chaque fois (trop lent si beaucoup de fichiers)
    // Juste s'assurer que le dossier lui-même est privé
    try { dossier.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch(e) {}
  } else {
    dossier=creerDossierDrive(nom);
    try { dossier.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch(e) {}
  }
  return dossier;
}
// ── Génération d'un token dynamique signé (timestamp:signature) ──
function genererTokenDynamique() {
  var ts = String(Math.floor(Date.now() / 1000)); // secondes
  var sig = Utilities.computeHmacSha256Signature(ts, FRI_SECRET_TOKEN);
  var sigHex = sig.map(function(b){ return ('0'+(b&0xff).toString(16)).slice(-2); }).join('');
  return ts + ':' + sigHex;
}

// ── Vérification du token dynamique ──
function verifierToken(payload) {
  if(!FRI_SECRET_TOKEN || FRI_SECRET_TOKEN==='') return true;
  var token = payload._token || payload.token || '';
  // Compatibilité : accepter aussi l'ancien token statique pendant la transition
  if(token === FRI_SECRET_TOKEN) return true;
  // Valider le token dynamique ts:signature
  var parts = token.split(':');
  if(parts.length !== 2) { Logger.log('⛔ Token invalide (format): ' + token.slice(0,20)); return false; }
  var ts = parts[0];
  var sigRecu = parts[1];
  // Vérifier expiration : 20 min pour actions PDF lourdes, 5 min pour les autres
  var now = Math.floor(Date.now() / 1000);
  var age = now - parseInt(ts, 10);
  var action = payload.action || '';
  var maxAge = (action === 'envoyerAttestationPDF' || action === 'envoyerFacturePDF' || action === 'saveReglement') ? 1200 : 300;
  if(age < 0 || age > maxAge) { Logger.log('⛔ Token expiré (age: ' + age + 's, max: ' + maxAge + 's)'); return false; }
  // Vérifier signature HMAC
  var sig = Utilities.computeHmacSha256Signature(ts, FRI_SECRET_TOKEN);
  var sigAttendu = sig.map(function(b){ return ('0'+(b&0xff).toString(16)).slice(-2); }).join('');
  if(sigRecu !== sigAttendu) { Logger.log('⛔ Token invalide (signature)'); return false; }
  return true;
}
function repondreAvecCb(obj,msgId,origin,callback) {
  var json=JSON.stringify(obj);
  if(callback&&/^[a-zA-Z0-9_]+$/.test(callback)){var output=ContentService.createTextOutput(callback+'('+json+')');output.setMimeType(ContentService.MimeType.JAVASCRIPT);return output;}
  var output=ContentService.createTextOutput(json);output.setMimeType(ContentService.MimeType.JSON);return output;
}
function repondre(obj){return repondreAvecCb(obj,null,null);}

// ============================================================
// ÉCRITURE DANS GOOGLE SHEETS + EMAILS
// ============================================================

// ── Helper : formules AD et AF pour une ligne donnée ──
// Formule AD : toujours AB*(1 - 0.15*AC)
// AC=1 (éligible) → AD=AB*0.85 ; AC=0 (non éligible) → AD=AB

// ══════════════════════════════════════════════════════════════
// CALCULER ET ÉCRIRE LE TOTAL FAMILLE (col AF=32)
// Logique :
//   Si ≥3 activités éligibles (AC=1) dans le dossier → somme AD
//   Sinon → somme AB
//   + FNSMR (15€ × nb membres uniques)
//   + FFTT (col AO=41, par membre unique)
// ══════════════════════════════════════════════════════════════

// ── Helper : lire le prix FFTT — robuste ancienne (39 cols) et nouvelle (41 cols) structure ──


function calcTotalFamille(ss, code) {
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  SpreadsheetApp.flush();
  var nbCols = Math.min(sheet.getLastColumn(), 41);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, nbCols).getValues();

  var nbElig = 0, sumAD = 0, sumAB = 0;
  var membresVus = {}, ffttVus = {}, fnsmr = 0, fftt = 0;

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][19] || '').trim() !== code) continue;
    var statut = lireStatutInscription(data[i]).toLowerCase();
    if (statut.indexOf('supprim') >= 0) continue;
    if (statut.indexOf('attente') >= 0) continue; // liste d'attente : tarif=0

    var ab = Number((data[i][27] !== undefined ? data[i][27] : 0) || 0); // col 28 AB
    var ac = Number((data[i][28] !== undefined ? data[i][28] : 0) || 0); // col 29 AC
    var ad = Math.round(ab * (1 - 0.15 * ac) * 100) / 100;
    var memKey = String(data[i][3] || '') + ' ' + String(data[i][2] || '');
    var ffttVal = 0; // FFTT supprimé // col 41 AO (ou 39 ancienne structure)

    if (ac === 1) nbElig++;
    sumAD += ad;
    sumAB += ab;

    if (!membresVus[memKey]) {
      membresVus[memKey] = true;
      fnsmr += 15;
    }
    if (ffttVal > 0 && !ffttVus[memKey]) {
      ffttVus[memKey] = true;
      fftt += ffttVal;
    }
  }

  var totalActif = (nbElig >= 3) ? sumAD : sumAB;
  var total = Math.round((totalActif + fnsmr + fftt) * 100) / 100;

  Logger.log('calcTotalFamille — code:'+code+' nbElig:'+nbElig+' sumAD:'+sumAD+' sumAB:'+sumAB+' fnsmr:'+fnsmr+' fftt:'+fftt+' total:'+total);

  // Écrire le total sur toutes les lignes actives du dossier (col 32 AF)
  for (var j = 0; j < data.length; j++) {
    if (String(data[j][19] || '').trim() !== code) continue;
    var statutJ = String((data[j][39] !== undefined ? data[j][39] : '') || '').toLowerCase();
    if (statutJ.indexOf('supprim') >= 0) {
      // Ligne supprimée : remettre AF à 0 pour ne pas fausser les totaux
      sheet.getRange(j + 2, 32).setValue(0)
        .setBackground('#ffe0b2').setFontColor('#e65100');
      continue;
    }
    sheet.getRange(j + 2, 32).setValue(total)
      .setBackground('#e8f4fd').setFontColor('#1565c0').setFontWeight('bold');
  }
  return total;
}

// ── Préchauffage cache places (à appeler via trigger quotidien) ──
function prechauffercachePlaces() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var shP = getOrCreatePlacesSheet(ss);
    var lrP = shP.getLastRow();
    var pmP = {};
    if(lrP > 1) {
      var dP = shP.getRange(2, 1, lrP-1, 5).getValues();
      dP.forEach(function(r){ var pid=String(r[0]||'').trim(); if(pid) pmP[pid]={capacite:Number(r[2]||0),inscrits:Number(r[3]||0),restantes:Number(r[4]||0)}; });
    }
    var pr = PropertiesService.getScriptProperties();
    pr.setProperty('fri_places_cache', JSON.stringify(pmP));
    pr.setProperty('fri_places_cache_ts', String(Date.now()));
    Logger.log('prechauffercachePlaces: '+Object.keys(pmP).length+' places mises en cache');
  } catch(e) { Logger.log('prechauffercachePlaces KO: '+e); }
}

// Installer le trigger (à appeler une fois manuellement depuis l'éditeur GAS)
function installerTriggerCache() {
  // Supprimer les anciens triggers prechauffage
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === 'prechauffercachePlaces') ScriptApp.deleteTrigger(t);
  });
  // Créer un trigger toutes les 2h
  ScriptApp.newTrigger('prechauffercachePlaces')
    .timeBased().everyHours(2).create();
  Logger.log('Trigger cache places installé (toutes les 2h)');
}

function _sommeMontantsTR(ss, sheetName, colMontant) {
  try {
    var sh = ss.getSheetByName(sheetName);
    if(!sh || sh.getLastRow() < 2) return 0;
    var d = sh.getRange(2, 1, sh.getLastRow()-1, colMontant).getValues();
    return d.reduce(function(s, r) {
      var v = parseFloat(r[colMontant-1] || 0);
      return s + (isNaN(v) || v > 1e10 ? 0 : v);
    }, 0);
  } catch(e) { return 0; }
}

function invaliderCachePlaces() {
  try {
    var pr = PropertiesService.getScriptProperties();
    pr.deleteProperty('fri_places_cache');
    pr.deleteProperty('fri_places_cache_ts');
  } catch(e) {}
}

function formuleAD(row) {
  var r = String(row);
  return '=AB'+r+'*(1-0,15*AC'+r+')';
}
// Formule AF : total famille calculé par GAS (voir calcTotalFamille)
// Valeur écrite par GAS après chaque inscription/modification
function formuleAF(row) {
  // AF est une valeur fixe écrite par GAS — pas de formule Excel
  // (FNSMR par membre unique ne peut pas être calculé correctement en SUMIF)
  return null;
}


// ── Éligibilité remise 15% — TT éligible pour Isneauville ──
function estEligibleRemise(pid, commune) {
  if (!pid) return false;
  var NR = [
    'JAZME1015','JAZME1115','JAZME1315','JAZME1415','JAZME1515','JAZME1615','JAZME1715','JAZJ1745',
    'MNOME/S10','COUNV1830','COUNV1930','COUNV2030','COUNV2130'];
  if (NR.indexOf(pid) >= 0) return false;
  // TT (PING*) : éligible uniquement pour les habitants d\'Isneauville
  if (pid.indexOf('PING') >= 0) {
    return (commune === 'Isneauville' || commune === 'isno');
  }
  return true;
}

function communeFromVille(ville) {
  return String(ville||'').toLowerCase().indexOf('isneauville') >= 0 ? 'Isneauville' : 'Hors commune';
}

// ── Dates : une cellule Date lue dans la feuille, passée à String(), donne
//    « Sun Mar 02 1969 00:00:00 GMT+0100 (heure normale d'Europe centrale) ».
//    Ces helpers renvoient toujours un texte lisible.
// Date de naissance → 'yyyy-MM-dd' (format utilisé par le formulaire et la feuille)
function formaterDdn(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, 'Europe/Paris', 'yyyy-MM-dd');
  var t = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
  var m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  var d = new Date(t);                       // ex. « Sun Mar 02 1969 00:00:00 GMT+0100 »
  return isNaN(d.getTime()) ? t : Utilities.formatDate(d, 'Europe/Paris', 'yyyy-MM-dd');
}
// Date (et heure si présente) → 'dd/MM/yyyy' ou 'dd/MM/yyyy à HH:mm'
function formaterDateHeure(v) {
  if (v === null || v === undefined || v === '') return '';
  var d = v instanceof Date ? v : null;
  if (!d) {
    var t = String(v).trim();
    if (!/GMT|UTC|^\w{3} \w{3} \d/.test(t)) return t; // déjà lisible (ex. « 05/10/2026 à 10:12 »)
    d = new Date(t);
    if (isNaN(d.getTime())) return t;
  }
  if (isNaN(d.getTime())) return '';
  var heure = Utilities.formatDate(d, 'Europe/Paris', 'HH:mm');
  return Utilities.formatDate(d, 'Europe/Paris', 'dd/MM/yyyy') + (heure !== '00:00' ? ' à ' + heure : '');
}

// Outil de réparation (à exécuter une fois depuis l'éditeur) : remet au bon format les dates de
// naissance écrites en texte brut (« Sun Mar 02 1969… ») dans Inscriptions (col E) et les onglets d'activité (col F).
function corrigerDatesNaissance() {
  var ss = SpreadsheetApp.openById(SHEET_ID), nb = 0;
  var corriger = function(sh, col) {
    if (!sh || sh.getLastRow() < 2) return;
    var rg = sh.getRange(2, col, sh.getLastRow() - 1, 1), vals = rg.getValues(), modif = false;
    vals.forEach(function(r) {
      if (typeof r[0] === 'string' && /GMT|UTC/.test(r[0])) { r[0] = formaterDdn(r[0]); modif = true; nb++; }
    });
    if (modif) rg.setValues(vals);
  };
  corriger(ss.getSheetByName(SHEET_INSCRIPTIONS), 5);
  ss.getSheets().forEach(function(sh) {
    var entete = String(sh.getRange(1, 6).getValue() || '') + String(sh.getRange(2, 6).getValue() || '');
    if (sh.getName() !== SHEET_INSCRIPTIONS && /naiss|ddn/i.test(entete)) corriger(sh, 6);
  });
  Logger.log('✅ ' + nb + ' date(s) de naissance corrigée(s)');
}

// Outil de réparation (à exécuter une fois depuis l'éditeur) : remplace dans toutes les cellules
// texte de la feuille les codes HTML écrits par erreur (ex. « &#x1F5D1; Activité supprimée »)
// par le vrai caractère (« 🗑 Activité supprimée »).
function corrigerEmoticonesFeuille() {
  var ss = SpreadsheetApp.openById(SHEET_ID), nb = 0;
  var decoder = function(t) {
    return t.replace(/&#(x[0-9a-f]+|\d+);/gi, function(m, v) {
      var c = v.charAt(0).toLowerCase() === 'x' ? parseInt(v.substring(1), 16) : parseInt(v, 10);
      try { return String.fromCodePoint(c); } catch(e) { return m; }
    }).replace(/&amp;/g, '&');
  };
  ss.getSheets().forEach(function(sh) {
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 1 || lc < 1) return;
    var rg = sh.getRange(1, 1, lr, lc), vals = rg.getValues(), formules = rg.getFormulas(), modif = false;
    for (var i = 0; i < vals.length; i++) for (var j = 0; j < vals[i].length; j++) {
      var v = vals[i][j];
      if (typeof v === 'string' && v.indexOf('&#') >= 0 && !formules[i][j]) {
        var d = decoder(v);
        if (d !== v) { sh.getRange(i + 1, j + 1).setValue(d); nb++; }
      }
    }
  });
  Logger.log('✅ ' + nb + ' cellule(s) corrigée(s)');
}

// Retire les caractères < et > des textes saisis dans le formulaire public :
// ces valeurs sont ensuite affichées dans la console admin et dans les emails HTML,
// un nom contenant du HTML pourrait sinon y exécuter du code (XSS).
function nettoyerTexteSaisi(v) {
  if (typeof v === 'string') return v.replace(/[<>]/g, '');
  if (Array.isArray(v)) return v.map(nettoyerTexteSaisi);
  if (v && typeof v === 'object') {
    var o = {};
    Object.keys(v).forEach(function(k) { o[k] = nettoyerTexteSaisi(v[k]); });
    return o;
  }
  return v;
}

// ══════════════════════════════════════════════════════════════
// ESPACE FAMILLE (menu d'accueil du site)
// Actions publiques : la famille s'identifie avec son code dossier ET l'email
// du dossier. Aucune donnée n'est modifiée : les demandes partent par email à l'admin.
// ══════════════════════════════════════════════════════════════

// Compteur simple (CacheService) pour limiter les abus d'une action publique
function _compteurDepasse(clef, max, dureeSec) {
  try {
    var cache = CacheService.getScriptCache();
    var n = parseInt(cache.get(clef) || '0', 10);
    if (n >= max) return true;
    cache.put(clef, String(n + 1), dureeSec);
  } catch (e) {}
  return false;
}

function _emailValide(e) { return /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(String(e || '')); }

function _texteCourt(v, max) { return String(nettoyerTexteSaisi(v == null ? '' : String(v))).trim().substring(0, max); }

// Recherche d'un dossier existant (même email ET même nom de famille) pour basculer
// l'inscription en ajout d'activités. Ne renvoie que ce qui sert à pré-remplir les membres.
function chercherDossierFamilleGAS(emailBrut, nomBrut) {
  var email = String(emailBrut || '').trim().toLowerCase();
  var norm = function(v) { return String(v || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z]/g, ''); };
  var nom = norm(nomBrut);
  if (!_emailValide(email) || nom.length < 2) return {status:'ok', trouve:false};
  if (_compteurDepasse('cdf_' + email, 30, 3600)) return {status:'ok', trouve:false};
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_INSCRIPTIONS);
  if (!sh || sh.getLastRow() < 2) return {status:'ok', trouve:false};
  var nbCol = Math.min(Math.max(sh.getLastColumn(), 40), 45);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, nbCol).getValues();
  var code = '';
  for (var i = data.length - 1; i >= 0 && !code; i--) {      // dossier le plus récent d'abord
    var r = data[i];
    var em = [String(r[15] || '').trim().toLowerCase(), String(r[18] || '').trim().toLowerCase()];
    if (em.indexOf(email) < 0) continue;
    if (lireStatutInscription(r).toLowerCase().indexOf('supprim') >= 0) continue;
    var resp = norm(lireResponsable(r) || r[36]);
    if (norm(r[2]) === nom || (resp && resp.indexOf(nom) >= 0)) code = String(r[19] || '').trim();
  }
  if (!/^FRI-[A-Z0-9]{4}$/.test(code)) return {status:'ok', trouve:false};
  var membres = {}, ordre = [];
  data.forEach(function(r) {
    if (String(r[19] || '').trim() !== code) return;
    if (lireStatutInscription(r).toLowerCase().indexOf('supprim') >= 0) return;
    var k = norm(r[3]) + '|' + norm(r[2]);
    if (membres[k]) return;
    membres[k] = { prenom: String(r[3] || '').trim(), nom: String(r[2] || '').trim(), ddn: formaterDdn(r[4]), sexe: lireSexe(r) };
    ordre.push(k);
  });
  var nbElig = 0;
  try { nbElig = lireLignesRestantes(SpreadsheetApp.openById(SHEET_ID), code, null).nbEligibles || 0; } catch(e) {}
  return {status:'ok', trouve:true, code:code, membres: ordre.map(function(k) { return membres[k]; }), nbActivitesEligibles: nbElig};
}

// Lit le dossier d'une famille. Renvoie {status:'ok', dossier} ou {status:'error', message}.
function lireDossierFamille(codeBrut, emailBrut) {
  var code  = String(codeBrut || '').trim().toUpperCase();
  var email = String(emailBrut || '').trim().toLowerCase();
  var refus = {status:'error', message:'Code dossier ou email incorrect.'};
  if (!/^FRI-[A-Z0-9]{4}$/.test(code) || !_emailValide(email)) return refus;
  // Au-delà de 10 essais ratés par heure sur un même code : on bloque (anti-devinette)
  var clefEchecs = 'cd_echec_' + code;
  try {
    if (parseInt(CacheService.getScriptCache().get(clefEchecs) || '0', 10) >= 10)
      return {status:'error', message:'Trop de tentatives, réessayez dans une heure.'};
  } catch (e) {}

  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_INSCRIPTIONS);
  if (!sh || sh.getLastRow() < 2) { _compteurDepasse(clefEchecs, 10, 3600); return refus; }
  var nbCol = Math.min(Math.max(sh.getLastColumn(), 40), 45);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, nbCol).getValues();
  var lignes = data.filter(function(r) { return String(r[19] || '').trim() === code; });
  var emailOk = lignes.some(function(r) {
    return String(r[15] || '').trim().toLowerCase() === email || String(r[18] || '').trim().toLowerCase() === email;
  });
  if (!lignes.length || !emailOk) { _compteurDepasse(clefEchecs, 10, 3600); return refus; }

  var membres = {}, ordre = [];
  lignes.forEach(function(r) {
    var statut = lireStatutInscription(r);
    if (statut.toLowerCase().indexOf('supprim') >= 0) return;
    var prenom = String(r[3] || '').trim(), nom = String(r[2] || '').trim();
    var clef = (prenom + ' ' + nom).toUpperCase();
    if (!membres[clef]) { membres[clef] = {prenom: prenom, nom: nom, sexe: lireSexe(r), activites: []}; ordre.push(clef); }
    membres[clef].activites.push({
      id:     lireActiviteId(r),
      nom:    String(r[22] || '').trim(),
      jour:   String(r[23] || '').trim(),
      heure:  String(r[24] || '').trim(),
      lieu:   String(r[25] || '').trim(),
      tarif:  Number(r[27] || 0),
      statut: statut || 'Inscrit'
    });
  });
  var r0 = lignes[0];
  return {status:'ok', dossier: {
    code: code,
    responsable: lireResponsable(r0) || String(r0[36] || '').trim(),
    email: String(r0[15] || '').trim(),
    statutPaiement: String(r0[21] || '').trim(),
    isno: communeFromVille(r0[10]) === 'Isneauville',
    membres: ordre.map(function(k) { return membres[k]; }),
    justificatif: _donneesJustificatif(code, lignes)
  }};
}

// ── Justificatif de règlement (facture d'adhésion pour l'employeur / CE) ──
// Disponible quand toutes les activités actives du dossier sont réglées.
// Coût des activités (remise famille comprise), licence/adhésion FNSMR, aides publiques
// (Pass'Jeunes, Atout Normandie, Pass'Sport) ; les coupons ANCV et avoirs sont des moyens de règlement.
function _estStatutRegle(st) {
  var s = String(st || '').toLowerCase();
  if (s.indexOf('attente') >= 0 || s.indexOf('cours') >= 0 || s.indexOf('à valider') >= 0) return false;
  return s.indexOf('validé') >= 0 || s.indexOf('payé') >= 0 || s.indexOf('réglé') >= 0;
}

function _donneesJustificatif(code, lignes) {
  try {
    var actives = lignes.filter(function(r) {
      var st = lireStatutInscription(r).toLowerCase();
      return st.indexOf('supprim') < 0 && st.indexOf('attente') < 0;
    });
    if (!actives.length) return {regle: false};
    var regle = actives.every(function(r) { return _estStatutRegle(r[21]); });
    var r0 = actives[0];
    var calc = lireLignesRestantes(SpreadsheetApp.openById(SHEET_ID), code, null);
    var licence = Math.round(((calc.fnsmr || 0) + (calc.fftt || 0)) * 100) / 100;
    var activites = Math.round(((calc.totalNet || 0) - licence) * 100) / 100;
    // Aides : montant le plus élevé trouvé par type (la même chaîne est recopiée sur chaque ligne)
    var aidesMax = {PassJeunes: 0, Atout: 0, PASS: 0, ANCV: 0}, avoir = 0;
    actives.forEach(function(r) {
      var pa = String(r[35] || '');
      Object.keys(aidesMax).forEach(function(t) {
        var m = pa.match(new RegExp('(?:^|\\|)' + t + ':([\\d.]+)'));
        if (m) aidesMax[t] = Math.max(aidesMax[t], parseFloat(m[1]) || 0);
      });
      avoir = Math.max(avoir, Number(r[33]) || 0);
    });
    var aides = Math.round((aidesMax.PassJeunes + aidesMax.Atout + aidesMax.PASS) * 100) / 100;
    var libAides = [];
    if (aidesMax.PassJeunes) libAides.push("Pass'Jeunes");
    if (aidesMax.Atout) libAides.push('Atout Normandie');
    if (aidesMax.PASS) libAides.push("Pass'Sport");
    // Modes de règlement : colonne « Mode paiement » + libellé du statut (« ✅ Règlement validé — Chèque »)
    var modes = {};
    actives.forEach(function(r) {
      var m = String(r[32] || '').toLowerCase() + ' ' + String(r[21] || '').toLowerCase();
      if (m.indexOf('espèce') >= 0 || m.indexOf('espece') >= 0) modes.especes = true;
      if (m.indexOf('chèque') >= 0 || m.indexOf('cheque') >= 0) modes.cheques = true;
      if (m.indexOf('helloasso') >= 0 || m.indexOf('carte') >= 0) modes.carte = true;
    });
    if (aidesMax.ANCV || avoir) modes.coupons = true;
    return {
      regle: regle,
      nom: String(r0[2] || '').trim(), prenom: String(r0[3] || '').trim(),
      responsable: lireResponsable(r0),
      adresse: String(r0[7] || '').trim(), cp: String(r0[9] || '').trim(), ville: String(r0[10] || '').trim(),
      tel: String(r0[13] || '').trim(), portable: String(r0[14] || '').trim(),
      email: String(r0[15] || '').trim(),
      activites: activites, licence: licence, aides: aides, libAides: libAides.join(', '),
      total: Math.max(0, Math.round((activites + licence - aides) * 100) / 100),
      modes: Object.keys(modes)
    };
  } catch(e) {
    Logger.log('_donneesJustificatif KO : ' + e);
    return {regle: false};
  }
}

// Pièces reçues pour l'espace famille (contrôle de présence dans le Drive, même logique que l'admin)
function piecesFamilleGAS(codeBrut, emailBrut) {
  var lecture = lireDossierFamille(codeBrut, emailBrut);
  if (lecture.status !== 'ok') return lecture;
  return controlerPiecesDossier(lecture.dossier.code);
}

// ── Contrôle automatique des pièces d'un dossier (console admin, rappel pièces manquantes) ──
// QS santé : attestation attendue si « Attestation OK » / « Non rempli » / « Non signé » ;
// certificat médical attendu si une réponse OUI (« Certificat requis » / « Certificat transmis ») ;
// règlement intérieur : un par dossier. Présence vérifiée dans les dossiers Drive (nom de fichier = code).
var DOSSIERS_QS = ['2-QS Santé Adhérents', '21-QS Santé corrigés'];
var DOSSIERS_RI = ['3-Règlements intérieurs', '31-RI corrigé'];
var DOSSIERS_CERTIF = ['1-Certificats médicaux'];

// Identifiants des dossiers Drive mémorisés 6 h : évite une recherche par nom à chaque contrôle
function _idsDossiersDrive(nd) {
  var cache = CacheService.getScriptCache(), clef = 'dossier_ids3_' + nomDossierDrive(nd);
  var enCache = cache.get(clef);
  if (enCache !== null) return enCache ? enCache.split(',') : [];
  // En test : dossiers « TEST - … » ET dossiers sans préfixe (fichiers déposés avant le mode test)
  var noms = MODE_TEST ? [nomDossierDrive(nd), nd] : [nd];
  var ids = [];
  noms.forEach(function(n) {
    var it = DriveApp.getFoldersByName(n);
    while (it.hasNext()) { var id = it.next().getId(); if (ids.indexOf(id) < 0) ids.push(id); }
  });
  if (ids.length) cache.put(clef, ids.join(','), 21600);
  return ids;
}

// Noms des fichiers d'un dossier Drive (majuscules), gardés 60 s.
// On parcourt la liste plutôt que d'utiliser searchFiles('title contains …') : la recherche Drive
// découpe les noms en mots et ne trouve pas « FRI-8N98 » dans « FRI-8N98_Gabie-…_QS-Sante.pdf ».
function _nomsFichiersDossier(id) {
  var cache = CacheService.getScriptCache(), clef = 'noms_fichiers_' + id;
  var enCache = cache.get(clef);
  if (enCache !== null) { try { return JSON.parse(enCache); } catch(e) {} }
  var noms = [], it = DriveApp.getFolderById(id).getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (!f.isTrashed()) noms.push(f.getName().toUpperCase());
  }
  try { var json = JSON.stringify(noms); if (json.length < 95000) cache.put(clef, json, 60); } catch(e) {}
  return noms;
}

function _fichiersDuDossier(nomsDossiers, code) {
  var noms = [], c = String(code || '').toUpperCase();
  nomsDossiers.forEach(function(nd) {
    try {
      _idsDossiersDrive(nd).forEach(function(id) {
        _nomsFichiersDossier(id).forEach(function(n) { if (n.indexOf(c) >= 0) noms.push(n); });
      });
    } catch(e) { Logger.log('Recherche Drive ' + nd + ' KO : ' + e); }
  });
  return noms;
}

// Recherche globale par nom (index Drive) : complète le parcours des dossiers connus
function _fichiersPartoutDansDrive(code) {
  var noms = [];
  try {
    var it = DriveApp.searchFiles('title contains "' + code + '" and trashed = false');
    while (it.hasNext() && noms.length < 50) noms.push(it.next().getName().toUpperCase());
  } catch(e) { Logger.log('Recherche globale Drive KO : ' + e); }
  return noms;
}

// À lancer depuis l'éditeur Apps Script (remplacer le code) : indique où sont les pièces d'un dossier
function diagnostiquerPiecesDossier(code) {
  code = String(code || 'FRI-8N98').toUpperCase();
  Logger.log('Mode test : ' + MODE_TEST);
  [['QS', DOSSIERS_QS], ['RI', DOSSIERS_RI], ['CERTIF', DOSSIERS_CERTIF]].forEach(function(g) {
    g[1].forEach(function(nd) {
      CacheService.getScriptCache().remove('dossier_ids3_' + nomDossierDrive(nd));
      var ids = _idsDossiersDrive(nd);
      Logger.log(g[0] + ' — dossier « ' + nd + ' » : ' + ids.length + ' dossier(s) trouvé(s)');
      ids.forEach(function(id) {
        var f = DriveApp.getFolderById(id);
        CacheService.getScriptCache().remove('noms_fichiers_' + id);
        var noms = _nomsFichiersDossier(id);
        Logger.log('   « ' + f.getName() + ' » (' + id + ') : ' + noms.length + ' fichier(s), dont pour ' + code + ' : '
          + JSON.stringify(noms.filter(function(n) { return n.indexOf(code) >= 0; })));
      });
    });
  });
  var it = DriveApp.searchFiles('title contains "' + code + '" and trashed = false');
  while (it.hasNext()) {
    var fi = it.next(), parents = [], p = fi.getParents();
    while (p.hasNext()) parents.push(p.next().getName());
    Logger.log('Recherche globale : ' + fi.getName() + ' → dans « ' + parents.join(', ') + ' »');
  }
  Logger.log('Résultat contrôle : ' + JSON.stringify(controlerPiecesDossier(code)));
}

// ── Suivi admin (coches « Reçu / Attente » des documents et règlements) ──
// Onglet « Suivi admin » : une ligne par dossier et par élément.
var SHEET_SUIVI_ADMIN = 'Suivi admin';
function _feuilleSuiviAdmin() {
  var ss = SpreadsheetApp.openById(SHEET_ID), sh = ss.getSheetByName(SHEET_SUIVI_ADMIN);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SUIVI_ADMIN);
    sh.getRange(1, 1, 1, 5).setValues([['Dossier', 'Élément', 'Statut', 'Modifié le', 'Par']])
      .setFontWeight('bold').setBackground('#1a2e22').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function lireSuiviAdmin(codeSeul) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_SUIVI_ADMIN);
  var suivi = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function(r) {
      var code = String(r[0] || '').trim();
      if (!code || (codeSeul && code !== codeSeul)) return;
      if (!suivi[code]) suivi[code] = {};
      suivi[code][String(r[1])] = String(r[2]);
    });
  }
  return {status:'ok', suivi: suivi};
}

function majSuiviAdminGAS(changements, session) {
  if (!changements || typeof changements !== 'object') return {status:'error', message:'Aucun changement'};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return {status:'error', message:'Serveur occupé, réessayez'};
  try {
    var sh = _feuilleSuiviAdmin(), index = {};
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function(r, i) { index[String(r[0]) + '|' + String(r[1])] = i + 2; });
    }
    var maintenant = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm'), par = session ? session.user : '', nouvelles = [], nb = 0;
    Object.keys(changements).forEach(function(code) {
      if (!/^FRI-[A-Z0-9]{4}$/.test(code)) return;
      var items = changements[code] || {};
      Object.keys(items).forEach(function(item) {
        var statut = String(items[item] || '').substring(0, 30), it = String(item).substring(0, 120), clef = code + '|' + it;
        if (index[clef]) sh.getRange(index[clef], 3, 1, 3).setValues([[statut, maintenant, par]]);
        else nouvelles.push([code, it, statut, maintenant, par]);
        nb++;
      });
    });
    if (nouvelles.length) sh.getRange(sh.getLastRow() + 1, 1, nouvelles.length, 5).setValues(nouvelles);
    return {status:'ok', modifies: nb};
  } finally { lock.releaseLock(); }
}

function controlerPiecesDossier(code) {
  if (!/^FRI-[A-Z0-9]{4}$/.test(code)) return {status:'error', message:'Code invalide'};
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_INSCRIPTIONS);
  if (!sh || sh.getLastRow() < 2) return {status:'ok', pieces: []};
  var nbCol = Math.min(Math.max(sh.getLastColumn(), 40), 45);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, nbCol).getValues();
  var membres = {}, ordre = [];
  data.forEach(function(r) {
    if (String(r[19] || '').trim() !== code) return;
    var st = lireStatutInscription(r).toLowerCase();
    if (st.indexOf('supprim') >= 0) return;
    var k = (String(r[3] || '').trim() + ' ' + String(r[2] || '').trim()).trim();
    if (!membres[k]) { membres[k] = { prenom: String(r[3] || '').trim(), nom: String(r[2] || '').trim(), qs: '' }; ordre.push(k); }
    var qs = String(r[34] || '').trim();
    if (qs && (!membres[k].qs || /certificat/i.test(qs))) membres[k].qs = qs;   // une mention certificat l'emporte
  });
  if (!ordre.length) return {status:'ok', pieces: []};
  var fQS = _fichiersDuDossier(DOSSIERS_QS, code), fRI = _fichiersDuDossier(DOSSIERS_RI, code), fCM = _fichiersDuDossier(DOSSIERS_CERTIF, code);
  // Filet de sécurité : fichiers portant le code n'importe où dans le Drive (dossier renommé, déplacé…)
  _fichiersPartoutDansDrive(code).forEach(function(n) {
    if (/CERTIF/.test(n)) { if (fCM.indexOf(n) < 0) fCM.push(n); }
    else if (/(^|[^A-Z])QS([^A-Z]|$)|SANTE|SANTÉ/.test(n)) { if (fQS.indexOf(n) < 0) fQS.push(n); }
    else if (/^RI[_-]|REGLEMENT|RÈGLEMENT/.test(n)) { if (fRI.indexOf(n) < 0) fRI.push(n); }
  });
  var norm = function(v) { return String(v || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z]/g, ''); };
  var contientMembre = function(liste, m) {
    var p = norm(m.prenom), n = norm(m.nom);
    return liste.some(function(nomF) { var t = norm(nomF); return t.indexOf(p) >= 0 && t.indexOf(n) >= 0; });
  };
  var pieces = [];
  ordre.forEach(function(k) {
    var m = membres[k], qs = m.qs.toLowerCase();
    if (qs.indexOf('pas de qs') >= 0) return;
    if (qs.indexOf('certificat') >= 0) {
      pieces.push({ type: 'CERTIF', membre: k, label: 'Certificat médical — ' + k, recu: contientMembre(fCM, m) });
    } else {
      pieces.push({ type: 'QS', membre: k, label: 'Attestation de santé (QS) — ' + k, recu: contientMembre(fQS, m) });
    }
  });
  pieces.push({ type: 'RI', membre: '', label: 'Règlement intérieur signé', recu: fRI.length > 0 });
  return {status:'ok', pieces: pieces};
}

// Certificat reçu → la feuille l'indique pour ce membre (colonne QS santé)
function marquerCertificatTransmis(code, nom, prenom) {
  try {
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_INSCRIPTIONS);
    if (!sh || sh.getLastRow() < 2) return;
    var norm = function(v) { return String(v || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z]/g, ''); };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 35).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][19] || '').trim() !== code) continue;
      if (norm(data[i][2]) !== norm(nom) || norm(data[i][3]) !== norm(prenom)) continue;
      sh.getRange(i + 2, 35).setValue('Certificat transmis').setFontColor('#1b5e20');
    }
  } catch(e) { Logger.log('marquerCertificatTransmis KO : ' + e); }
}

// ── Demandes de modification : onglet dédié, une ligne par ajout / retrait ──
var SHEET_DEMANDES = 'Demandes modification';
var COLS_DEMANDES = ['ID', 'Date', 'Dossier', 'Responsable', 'Email', 'Type', 'Membre',
                     'ID activité', 'Activité', 'Commentaire', 'Statut', 'Traité le', 'Traité par', 'Réponse'];

function getOrCreateDemandesSheet(ss) {
  var sh = ss.getSheetByName(SHEET_DEMANDES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_DEMANDES);
    sh.getRange(1, 1, 1, COLS_DEMANDES.length).setValues([COLS_DEMANDES])
      .setFontWeight('bold').setBackground('#1a2e22').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Demande d'ajout / de retrait d'activités : enregistrée pour la console admin,
// email à l'admin + accusé de réception à la famille
function demanderModificationGAS(p) {
  var lecture = lireDossierFamille(p.code, p.email);
  if (lecture.status !== 'ok') return lecture;
  var d = lecture.dossier;
  if (_compteurDepasse('dm_' + d.code, 5, 3600)) return {status:'error', message:'Trop de demandes pour ce dossier, réessayez plus tard.'};

  // Retraits : uniquement des activités réellement présentes dans le dossier
  var presentes = {};
  (d.membres || []).forEach(function(m) {
    (m.activites || []).forEach(function(a) { presentes[(m.prenom + ' ' + m.nom).trim().toUpperCase() + '|' + a.id] = {membre: (m.prenom + ' ' + m.nom).trim(), act: a}; });
  });
  var items = [];
  (Array.isArray(p.retraits) ? p.retraits : []).slice(0, 20).forEach(function(r) {
    var k = _texteCourt(r && r.membre, 120).toUpperCase() + '|' + _texteCourt(r && r.actId, 60);
    if (presentes[k]) items.push({type: 'Retrait', membre: presentes[k].membre, actId: presentes[k].act.id, actNom: presentes[k].act.nom});
  });
  (Array.isArray(p.ajouts) ? p.ajouts : []).slice(0, 20).forEach(function(a) {
    var actId = _texteCourt(a && a.actId, 60);
    if (!/^[a-z0-9-]+$/i.test(actId)) return;
    items.push({type: 'Ajout', membre: _texteCourt(a.membre, 120) || 'Nouveau membre', actId: actId, actNom: _texteCourt(a.actNom, 200)});
  });
  var commentaire = _texteCourt(p.commentaire, 2000);
  if (!items.length && !commentaire) return {status:'error', message:'Indiquez au moins une modification.'};
  if (!items.length) items.push({type: 'Commentaire', membre: '', actId: '', actNom: ''});

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = getOrCreateDemandesSheet(ss);
  var maintenant = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
  var base = 'DM-' + Utilities.formatDate(new Date(), 'Europe/Paris', 'yyMMddHHmmss');
  var lignes = items.map(function(it, i) {
    return [base + '-' + (i + 1), maintenant, d.code, d.responsable, d.email, it.type, it.membre,
            it.actId, it.actNom, commentaire, 'En attente', '', '', ''];
  });
  sh.getRange(sh.getLastRow() + 1, 1, lignes.length, COLS_DEMANDES.length).setValues(lignes);

  var resume = items.filter(function(it) { return it.type !== 'Commentaire'; })
    .map(function(it) { return (it.type === 'Ajout' ? '+ AJOUT : ' : '- RETRAIT : ') + it.membre + ' — ' + it.actNom; }).join('\n');
  var corps = 'Demande de modification — dossier ' + d.code + '\n'
    + 'Responsable : ' + d.responsable + ' <' + d.email + '>\n\n'
    + (resume ? resume + '\n\n' : '')
    + (commentaire ? 'Commentaire :\n' + commentaire + '\n\n' : '');
  envoyerEmail(EMAIL_ADMIN, '[FRI] Demande de modification ' + d.code, corps
    + '→ À valider dans la console admin, rubrique « Demandes de modification ».', {name: NOM_ASSO, replyTo: d.email});
  envoyerEmail(d.email, 'Foyer Rural — demande de modification reçue (' + d.code + ')',
    'Bonjour,\n\nNous avons bien reçu votre demande de modification pour le dossier ' + d.code + '.\n'
    + 'Elle sera traitée par l\'équipe du Foyer Rural, qui reviendra vers vous si un complément ou un règlement est nécessaire.\n\n'
    + corps + 'Cordialement,\n' + NOM_ASSO, {name: NOM_ASSO});
  return {status:'ok'};
}

// Console admin : demandes en attente
function lireDemandesModification() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_DEMANDES);
  if (!sh || sh.getLastRow() < 2) return {status:'ok', demandes: []};
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_DEMANDES.length).getValues();
  var demandes = [];
  data.forEach(function(r) {
    if (String(r[10]) !== 'En attente') return;
    demandes.push({id: String(r[0]), date: r[1] instanceof Date ? Utilities.formatDate(r[1], 'Europe/Paris', 'dd/MM/yyyy HH:mm') : String(r[1]),
      code: String(r[2]), responsable: String(r[3]), email: String(r[4]), type: String(r[5]), membre: String(r[6]),
      actId: String(r[7]), actNom: String(r[8]), commentaire: String(r[9])});
  });
  return {status:'ok', demandes: demandes};
}

// Console admin : marquer une demande validée (après application) ou refusée (email à la famille)
function traiterDemandeModificationGAS(p, session) {
  var id = String(p.id || ''), statut = p.statut === 'Refusée' ? 'Refusée' : 'Validée';
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_DEMANDES);
  if (!sh || sh.getLastRow() < 2) return {status:'error', message:'Demande introuvable'};
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_DEMANDES.length).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== id) continue;
    if (String(data[i][10]) !== 'En attente') return {status:'ok', deja: true};
    var reponse = _texteCourt(p.reponse, 1000);
    sh.getRange(i + 2, 11, 1, 4).setValues([[statut, Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm'),
      session ? session.user : '', reponse]]);
    if (statut === 'Refusée' && _emailValide(data[i][4])) {
      envoyerEmail(String(data[i][4]), 'Foyer Rural — votre demande de modification (' + data[i][2] + ')',
        'Bonjour,\n\nVotre demande concernant le dossier ' + data[i][2] + ' n\'a pas pu être acceptée :\n'
        + data[i][5] + ' — ' + data[i][6] + ' — ' + data[i][8] + '\n\n'
        + (reponse ? 'Motif : ' + reponse + '\n\n' : '')
        + 'N\'hésitez pas à nous contacter pour en parler.\n\nCordialement,\n' + NOM_ASSO, {name: NOM_ASSO});
    }
    return {status:'ok'};
  }
  return {status:'error', message:'Demande introuvable'};
}

// Message libre envoyé depuis le site
function envoyerMessageContactGAS(p) {
  if (String(p.site_web || '')) return {status:'ok'}; // champ piège rempli par les robots : on ignore
  var nom = _texteCourt(p.nom, 120), email = _texteCourt(p.email, 200).toLowerCase();
  var code = _texteCourt(p.code, 12).toUpperCase(), sujet = _texteCourt(p.sujet, 150), message = _texteCourt(p.message, 3000);
  if (!nom || !_emailValide(email)) return {status:'error', message:'Indiquez votre nom et une adresse email valide.'};
  if (message.length < 5) return {status:'error', message:'Votre message est vide.'};
  if (_compteurDepasse('msg_' + email, 3, 3600) || _compteurDepasse('msg_global', 60, 3600))
    return {status:'error', message:'Trop de messages envoyés, réessayez plus tard.'};
  var corps = 'Message reçu depuis le site d\'inscription\n\n'
    + 'De : ' + nom + ' <' + email + '>\n'
    + (code ? 'Dossier : ' + code + '\n' : '')
    + 'Sujet : ' + (sujet || '(sans sujet)') + '\n\n' + message;
  envoyerEmail(EMAIL_ADMIN, '[FRI] Message : ' + (sujet || nom), corps, {name: NOM_ASSO, replyTo: email});
  envoyerEmail(email, 'Foyer Rural — nous avons bien reçu votre message',
    'Bonjour ' + nom + ',\n\nNous avons bien reçu votre message et vous répondrons dès que possible.\n\n---\n' + message
    + '\n\nCordialement,\n' + NOM_ASSO, {name: NOM_ASSO});
  return {status:'ok'};
}

// Réponse renvoyée au site après addRegistration : « ok » seulement si les lignes sont écrites
// (ou déjà présentes pour cette famille), sinon une erreur que le site sait traiter.
function reponseAddRegistration(result) {
  result = result || {};
  if (result.codePris) return { status: 'error', code: 'CODE_PRIS', message: result.error };
  if (result.error)    return { status: 'error', code: 'OCCUPE', message: result.error };
  return { status: 'ok', inserted: result.inserted || 0, doublon: !!result.doublon };
}

function addRegistration(rows, paymentStatus, payload) {
  if(!rows||rows.length===0)return{inserted:0};
  rows = nettoyerTexteSaisi(rows);

  // ── Verrou anti-concurrence pour les inscriptions simultanées ──
  var lock = LockService.getScriptLock();
  // Forum des associations : beaucoup d'inscriptions simultanées → attente jusqu'à 30 s ;
  // en cas d'échec une vraie erreur est renvoyée (le site réessaie) au lieu d'un « ok » sans écriture.
  if (!lock.tryLock(30000)) {
    Logger.log('addRegistration: verrou indisponible après 30 s — code ' + (rows[0] && rows[0].code_dossier));
    return { inserted: 0, error: 'Serveur occupé, nouvel essai en cours…' };
  }

  try {
  const ss=SpreadsheetApp.openById(SHEET_ID);

  // ── Protection anti-doublon : vérifier si le code_dossier existe déjà ──
  // Peut arriver si confirmPayment() appelle sendToGoogleSheets une 2e fois
  var codeDossier = rows[0].code_dossier || '';
  var ajoutDossier = !!(payload && payload.ajoutDossier);
  // Envoi déjà traité (nouvel essai après une coupure réseau) → ne pas écrire deux fois
  var idEnvoi = payload && payload.idEnvoi ? 'envoi_' + String(payload.idEnvoi).replace(/[^\w-]/g, '').substring(0, 60) : '';
  if (idEnvoi) {
    try { if (CacheService.getScriptCache().get(idEnvoi)) return { inserted: 0, doublon: true }; } catch(eC) {}
  }
  var membresDejaDansDossier = {};
  // Anti-doublon : ignorer les codes de test et vérifier seulement les vrais codes FRI-XXXX
  if (codeDossier && /^FRI-[A-Z0-9]{4}$/.test(codeDossier) && codeDossier !== 'FRI-TEST') {
    try {
      var sheetCheck = ss.getSheetByName(SHEET_INSCRIPTIONS);
      if (sheetCheck && sheetCheck.getLastRow() > 1) {
        var colCodes = sheetCheck.getRange(2, 20, sheetCheck.getLastRow() - 1, 1).getValues();
        for (var dc = 0; dc < colCodes.length; dc++) {
          if (String(colCodes[dc][0] || '').trim() === codeDossier) {
            // Même code : vrai doublon (même famille, nouvel envoi) ou numéro déjà attribué à une autre famille ?
            var emailExistant = String(sheetCheck.getRange(dc + 2, 16).getValue() || '').trim().toLowerCase();
            var emailNouveau  = String(rows[0].email1 || '').trim().toLowerCase();
            if (emailExistant && emailNouveau && emailExistant !== emailNouveau) {
              Logger.log('⚠️ addRegistration — code ' + codeDossier + ' déjà pris par une autre famille');
              return { inserted: 0, codePris: true, error: 'Numéro de dossier déjà utilisé' };
            }
            if (ajoutDossier) break; // ajout d'activités à un dossier existant de la même famille
            Logger.log('⚠️ addRegistration — doublon détecté pour ' + codeDossier + ' → insertion ignorée');
            return { inserted: 0, doublon: true };
          }
        }
      }
    } catch(eCheck) { Logger.log('Vérif doublon KO (non bloquant) : ' + eCheck); }
    // Ajout à un dossier existant : membres déjà présents (leur adhésion FNSMR est déjà comptée)
    if (ajoutDossier) {
      try {
        var shM = ss.getSheetByName(SHEET_INSCRIPTIONS);
        if (shM && shM.getLastRow() > 1) {
          shM.getRange(2, 1, shM.getLastRow() - 1, 20).getValues().forEach(function(rm) {
            if (String(rm[19] || '').trim() !== codeDossier) return;
            membresDejaDansDossier[(String(rm[3] || '') + ' ' + String(rm[2] || '')).trim().toUpperCase()] = true;
          });
        }
      } catch(eM) { Logger.log('Lecture membres existants KO : ' + eM); }
    }
  }

  var modePaiement=rows[0].mode_paiement||'helloasso';
  var modeLabelMap={helloasso:'HelloAsso',cheque:'Cheque',cheque3:'Chèques échelonnés',especes:'Especes',ancv:'Coupon sport ANCV',aide:'Soldé par aides'};
  var modeLabel=modeLabelMap[modePaiement]||modePaiement;
  // statut col22 : calculé par ligne selon liste d'attente ou non
  // On utilise r.statut_inscription du frontend ('Liste attente' ou '') car stGas pas encore calculé
  function getStatutPaiement(r) {
    var frontSt = String(r.statut_inscription||'').toLowerCase();
    var isLA = frontSt.indexOf('attente') >= 0 || frontSt.indexOf('liste') >= 0;
    return isLA ? '— Non concerné' : '⏳ En attente de règlement — '+modeLabel;
  }
  var statut = '⏳ En attente de règlement — '+modeLabel; // fallback pour majOngletsActivites
  let sheet=ss.getSheetByName(SHEET_INSCRIPTIONS);
  if(!sheet){sheet=ss.insertSheet(SHEET_INSCRIPTIONS);ecrireEnTeteInscriptions(sheet);}
  else if(sheet.getLastRow()===0){ecrireEnTeteInscriptions(sheet);}
  else{var firstCell=sheet.getRange(1,1).getValue();if(firstCell&&String(firstCell).indexOf('Num')<0&&String(firstCell).length>0){sheet.insertRowBefore(1);ecrireEnTeteInscriptions(sheet);}}
  // S'assurer que le Sheet a 41 colonnes (structure v8.8)
  if(sheet.getMaxColumns() < 41){
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 41 - sheet.getMaxColumns());
    Logger.log('Colonnes ajoutées pour atteindre 41');
  }

  var NON_REMISABLE_IDS = ['PINGL1945/ME1945','PINGM1930/ME21','PINGJ1530',
    'PINGL18/J17','PINGME1730','PINGL1745ME1830','JAZME1015','JAZME1115','JAZME1315','JAZME1415','JAZME1515',
    'JAZME1615','JAZME1715','JAZJ1745','MNOME/S10','COUNV1830','COUNV1930','COUNV2030','COUNV2130'];
  var commune0 = rows[0].commune || '';
  function estEligRemise(pid) {
    return estEligibleRemise(pid, commune0);
  }
  const newRows=rows.map(function(r){
    var commentaire=(r.activite||'').replace(/\n/g,' — ')+(r.jour?' | '+r.jour:'')+(r.heure?' '+r.heure:'')+(r.lieu?' ('+r.lieu+')':'');
    return [
      /* 1  A  */ '',
      /* 2  B  */ '',
      /* 3  C  */ r.membre_nom||'',
      /* 4  D  */ r.membre_prenom||'',
      /* 5  E  */ r.ddn||'',
      /* 6  F  */ '',
      /* 7  G  */ '',
      /* 8  H  */ r.adresse||'',
      /* 9  I  */ '',
      /* 10 J  */ r.cp||'',
      /* 11 K  */ r.ville||'',
      /* 12 L  */ '',
      /* 13 M  */ 'FR',
      /* 14 N  */ '',
      /* 15 O  */ r.tel1||'',
      /* 16 P  */ r.email1||'',
      /* 17 Q  */ commentaire,
      /* 18 R  */ r.tel2||'',
      /* 19 S  */ r.email2||'',
      /* 20 T  */ r.code_dossier||'',
      /* 21 U  */ r.date||'',
      /* 22 V  */ getStatutPaiement(r),
      /* 23 W  */ (r.activite||'').replace(/\n/g,' — '),
      /* 24 X  */ r.jour||'',
      /* 25 Y  */ r.heure||'',
      /* 26 Z  */ r.lieu||'',
      /* 27 AA */ r.animateur||'',
      /* 28 AB */ Number(r.tarif_brut)||Number(r.tarif)||0,
      /* 29 AC */ (String(r.statut_inscription||'').toLowerCase().indexOf('attente') >= 0)
        ? 0 // liste d'attente = non éligible remise
        : estEligRemise(getPlacesId(r.activite_id||'')) ? 1 : 0,
      /* 30 AD */ 0,
      /* 31 AE */ (function(){
        // Ajout à un dossier existant : adhésion déjà réglée pour ce membre
        if (membresDejaDansDossier[((r.membre_prenom||'') + ' ' + (r.membre_nom||'')).trim().toUpperCase()]) return 0;
        // FNSMR : 15€ SAUF si activité en attente ET le membre a d'autres lignes payantes
        var stInscrit = String(r.statut_inscription||'');
        if (stInscrit.toLowerCase().indexOf('attente') < 0) return 15; // payante → FNSMR dû
        // En attente → vérifier si le membre a d'autres lignes payantes dans rows
        var memKey = (r.membre_prenom||'') + ' ' + (r.membre_nom||'');
        var autresPayantes = rows.filter(function(other){
          var otherKey = (other.membre_prenom||'') + ' ' + (other.membre_nom||'');
          return otherKey === memKey
            && String(other.statut_inscription||'').toLowerCase().indexOf('attente') < 0;
        });
        return autresPayantes.length > 0 ? 0 : 15; // 0 si déjà facturé, 15 si seul en attente
      })(),
      /* 32 AF */ 0,
      /* 33 AG */ r.mode_paiement||'helloasso',
      /* 34 AH */ Number(r.avoir_montant)||0, // montant numérique
      /* 35 AI */ r.qs_sante||'',
      /* 36 AJ */ (function(){
        var parts = [];
        var pa = r.pass_aide||'';
        // Label aide (Pass'jeunes, Atout...) sans les montants redondants
        if(pa && pa.indexOf('PassJeunes:') < 0 && pa.indexOf('Atout:') < 0) parts.push(pa);
        // Montants structurés pour parsing côté GAS
        if(Number(r.pass_jeunes_montant)>0) parts.push('PassJeunes:'+Number(r.pass_jeunes_montant).toFixed(2)+':'+(r.pass_jeunes_type||'1ere'));
        if(Number(r.atout_montant)>0)       parts.push('Atout:'+Number(r.atout_montant).toFixed(2)+(r.atout_code?':'+r.atout_code:''));
        if(Number(r.ancv_montant)>0)        parts.push('ANCV:'+Number(r.ancv_montant).toFixed(2));
        if(Number(r.pass_sport_montant)>0)  parts.push('PASS:'+Number(r.pass_sport_montant).toFixed(2));
        // Stocker le code avoir pour utiliserAvoirGAS lors de la validation
        if(r.avoir_code && Number(r.avoir_montant)>0) parts.push('AVOIR:'+r.avoir_code+':'+Number(r.avoir_montant).toFixed(2));
        return parts.join('|');
      })(),
      /* 37 AK */ r.sexe||'',
      /* 38 AL */ getPlacesId(r.activite_id||''),
      /* 39 AM */ (r.responsable_prenom||'')+' '+(r.responsable_nom||''),
      /* 40 AN */ '',
      /* 41 AO */ Number(r.fftt_price)||0 // lireFFTT au moment de l\'inscription
    ];
  });
  var startRow=sheet.getLastRow()+1;
  // S'assurer que le Sheet a au moins 41 colonnes
  if(sheet.getMaxColumns() < 41){
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 41 - sheet.getMaxColumns());
    Logger.log('Colonnes ajoutées → ' + sheet.getMaxColumns());
  }
  Logger.log('addReg — newRows:'+newRows.length+' cols:'+sheet.getMaxColumns()+' r0.activite_id:'+(rows[0]?rows[0].activite_id:'?'));
  // Vérifier que chaque row a bien 41 éléments
  var rowLens = newRows.map(function(r){ return r.length; });
  Logger.log('addReg — longueurs rows: ' + JSON.stringify(rowLens));
  try {
    sheet.getRange(startRow,1,newRows.length,41).setValues(newRows);
    Logger.log('addReg step1: setValues OK — startRow:'+startRow);
  } catch(esv) {
    Logger.log('addReg setValues ERREUR: '+esv.toString());
    return {inserted:0, error: esv.toString()};
  }
  for(var i=0;i<newRows.length;i++){
    var rn=startRow+i;
    var bgR=rn%2===0?'#f0f7f3':'#ffffff';
    sheet.getRange(rn,1,1,41).setBackground(bgR);
    sheet.getRange(rn,20).setBackground('#e8f4fd').setFontColor('#1565c0').setFontWeight('bold');
    // col 29 AC : couleur éligible (valeur déjà dans newRow)
    var eligVal = Number(sheet.getRange(rn, 29).getValue());
    sheet.getRange(rn, 29)
      .setBackground(eligVal === 1 ? '#e8f5e9' : '#fce4ec')
      .setFontColor(eligVal === 1 ? '#1b5e20' : '#b71c1c').setFontWeight('bold');
    // col 30 AD : formule tarif net
    sheet.getRange(rn, 30)
      .setFormula(formuleAD(rn))
      .setBackground('#d8f3dc').setFontColor('#1b5e20');
    // col 32 AF : formule total famille (SUMIF des AD du même dossier)
    sheet.getRange(rn, 32)
      .setValue(0) // AF calculé par calcTotalFamille() après inscription
      .setBackground('#e8f4fd').setFontColor('#1565c0').setFontWeight('bold');
    // col 28 AB : mise en évidence si éligible (AC=1)
    if(Number(rows[i].remise||0)>0||estEligRemise(getPlacesId(rows[i].activite_id||''))){sheet.getRange(rn,28).setBackground('#d8f3dc').setFontColor('#1b5e20').setFontWeight('bold');}
  }
  Logger.log("addReg step2: boucle for terminée");
  sheet.autoResizeColumns(1,41);
  // Calculer et écrire le total famille AF après toutes les formules AD
  SpreadsheetApp.flush();
  Logger.log("addReg step3: avant calcTotalFamille");
  try { calcTotalFamille(ss, rows[0].code_dossier || ''); } catch(etf) { Logger.log('calcTotalFamille KO: '+etf); }
  var statutMap={};
  try{statutMap=incrementPlaces(ss,rows);}catch(e){Logger.log('Places KO: '+e);}

  var actCounts={};
  rows.forEach(function(r){var pid=getPlacesId(r.activite_id||'');actCounts[pid]=(actCounts[pid]||0);});
  rows.forEach(function(r,idx){
    var pid=getPlacesId(r.activite_id||'');
    var jIdx=actCounts[pid]||0;actCounts[pid]=jIdx+1;
    // Priorité : si le frontend a déjà marqué 'Liste attente', conserver ce statut
    var frontendStatut = String(r.statut_inscription||'').toLowerCase();
    var statutMapVal   = String(statutMap[pid+'_'+jIdx]||'');
    var isAttenteLA    = frontendStatut.indexOf('attente') >= 0
                      || statutMapVal.toLowerCase().indexOf('attente') >= 0;
    // Statut inscription (col 40 AN) : libellés clairs pour l'admin
    var stGas = isAttenteLA
              ? '⏳ En attente de place'
              : '⏳ En cours de validation';
    r.stGas = stGas; // stocker pour majOngletListeAttente
    var rn=startRow+idx;
    // Nouvelles lignes → col 40 AN (index 39) = Statut inscription
    sheet.getRange(rn,40).setValue(stGas);
    if(isAttenteLA){
      sheet.getRange(rn,1,1,41).setBackground('#b2ebf2');
      sheet.getRange(rn,40).setFontColor('#006064').setFontWeight('bold');
      sheet.getRange(rn,28).setValue(0); // tarif brut = 0 pour liste attente
      sheet.getRange(rn,29).setValue(0); // éligible = 0
    } else {
      sheet.getRange(rn,40).setFontColor('#e65100').setFontWeight('bold'); // orange = en attente
    }
    r.statut_inscription=stGas;
  });

  // Créer les onglets LA-{placesId} APRÈS que statutMap est complet
  try {
    var rowsParPlacesId = {};
    rows.forEach(function(r) {
      var pid = getPlacesId(r.activite_id||'');
      // Créer onglet LA UNIQUEMENT si l'activité est vraiment en liste d'attente de place
      // (pas '⏳ En cours de validation' qui est le statut normal d'une nouvelle inscription)
      var stGasR = String(r.stGas||'');
      var isVraiAttente = stGasR.indexOf('En attente de place') >= 0;
      if (isVraiAttente) {
        if (!rowsParPlacesId[pid]) rowsParPlacesId[pid] = [];
        rowsParPlacesId[pid].push(r);
      }
    });
    Object.keys(rowsParPlacesId).forEach(function(pid) {
      majOngletListeAttente(ss, pid, rowsParPlacesId[pid]);
    });
  } catch(ela) { Logger.log('Onglet LA KO : ' + ela); }

  Logger.log("addReg step4: avant majRecapitulatif");
  try{majRecapitulatif(ss,rows,statut);}catch(e){Logger.log('Recap KO: '+e);}
  Logger.log("addReg step5: avant majOngletsActivites");
  try{majOngletsActivites(ss,rows,statut);}catch(e){Logger.log('Onglets KO: '+e);}
  var cheques=null;
  if(payload&&payload.cheques&&payload.cheques.length>0)cheques=payload.cheques;
  if(!cheques&&rows[0]&&rows[0].cheques)cheques=rows[0].cheques;
  if(cheques&&cheques.length>0){try{majOngletsChequesWrite(ss,rows,cheques);}catch(e){Logger.log('Cheques KO: '+e);}}
  // Alimenter les onglets paiements et aides des l'inscription (statut = En attente de validation)
  // L'admin validera chaque element depuis la console admin
  try {
    var r0ins = rows[0];
    var modeIns = String(r0ins.mode_paiement || 'cheque').toLowerCase();
    var codeIns = r0ins.code_dossier || '';
    // Paiement principal
    if (modeIns === 'especes') {
      try { majOngletEspeces(ss, rows); } catch(eE) { Logger.log('Especes inscription KO: '+eE); }
    } else if (modeIns === 'helloasso' || modeIns === 'helloasso3x' || modeIns === 'ha' || String(modeIns).indexOf('hello') >= 0) {
      try { majOngletHelloAsso(ss, rows, 0, null); } catch(eHA) { Logger.log('HA inscription KO: '+eHA); }
    }
    // Aides selon pass_aide
    var passAideStr = String(r0ins.pass_aide || '');
    var ancvM2 = passAideStr.match(/(?:^|\|)ANCV:([\d.]+)/);
    if (ancvM2 && parseFloat(ancvM2[1]) > 0) { try { majOngletAideANCv(ss, rows); } catch(eA) { Logger.log('ANCV inscription KO: '+eA); } }
    var psM2 = passAideStr.match(/(?:^|\|)PASS:([\d.]+)/);
    if (psM2 && parseFloat(psM2[1]) > 0) { try { majOngletAidePassSport(ss, rows); } catch(eA) { Logger.log('PassSport inscription KO: '+eA); } }
    var pjM2 = passAideStr.match(/PassJeunes:([\d.]+)/);
    if (pjM2 && parseFloat(pjM2[1]) > 0) { try { majOngletAidePassJeunes(ss, rows); } catch(eA) { Logger.log('PassJeunes inscription KO: '+eA); } }
    var atM2 = passAideStr.match(/Atout:([\d.]+)/);
    if (atM2 && parseFloat(atM2[1]) > 0) { try { majOngletAideAtout(ss, rows); } catch(eA) { Logger.log('Atout inscription KO: '+eA); } }
    // FFTT
    var ffttTotal = rows.reduce(function(s, r) { return s + (Number(r.fftt_price)||0); }, 0);
    if (ffttTotal > 0) {
}
    // Avoir utilise
    var avoirIns = parseFloat(r0ins.avoir_montant) || 0;
    if (avoirIns > 0) { try { ecrireAvoirUtilise(ss, rows, avoirIns); } catch(eAv) { Logger.log('Avoir inscription KO: '+eAv); } }
    Logger.log('Onglets paiements/aides alimentes inscription: ' + codeIns + ' mode=' + modeIns);
  } catch(eIns) { Logger.log('Alimenter onglets inscription KO: ' + eIns); }
  Logger.log('Email adhérent — email:' + (rows[0].email1||'VIDE') + ' modeLabel:' + modeLabel + ' nbRows:' + rows.length);
  Logger.log("addReg step6: avant emails — email1:"+rows[0].email1);
  try{envoyerEmailAdherent(rows[0].email1,rows,false,modeLabel,null);}catch(e){Logger.log('Email adherent KO: '+e.toString());}
  // Email admin à l'inscription supprimé (quota Gmail) — notification visible dans la console admin
  try{verifierSeuilCentDossiers(ss);}catch(e){Logger.log('Seuil100 KO: '+e);}
  // ── Journal des tentatives (PropertiesService) ──
  try {
    var props = PropertiesService.getScriptProperties();
    var journalRaw = props.getProperty('fri_journal_saves') || '[]';
    var journal = JSON.parse(journalRaw);
    var entree = {
      ts:   Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm:ss'),
      code: codeDossier,
      inserted: newRows.length,
      status: newRows.length > 0 ? 'ok' : 'doublon'
    };
    journal.unshift(entree); // Plus récent en premier
    if (journal.length > 200) journal = journal.slice(0, 200); // Garder 200 entrées max
    props.setProperty('fri_journal_saves', JSON.stringify(journal));
  } catch(eJournal) { Logger.log('Journal KO: ' + eJournal); }

  invaliderCachePlaces(); // Invalider le cache places après inscription
  Logger.log("addReg step7: TERMINÉ — inserted:"+newRows.length);
  if (idEnvoi && newRows.length) { try { CacheService.getScriptCache().put(idEnvoi, '1', 21600); } catch(eC2) {} }
  return{inserted:newRows.length};
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// GET DOSSIERS — v8.1 (inchangé)
// ============================================================

// ── Helpers lecture colonnes — compatibles ancienne et nouvelle structure ──
// Structure confirmée par log 27/05/2026 :
// ANCIENNE (données existantes, décalées de -2) :
//   index 35 = ID Activité, index 36 = Responsable(*), index 37 = Statut inscription, index 38 = FFTT
// NOUVELLE (après insertion colonnes AC/AD) :
//   index 37 = ID Activité, index 38 = Responsable, index 39 = Statut inscription, index 40 = FFTT
// Détection : si index 37 contient un placesId valide → nouvelle structure
//             si index 38 = nom avec espace ET index 37 = statut → ancienne structure
// (*) Responsable est toujours en index 38 (AM) selon le log → ne pas utiliser fallback index 36

// Détection ancienne/nouvelle structure :
// Nouvelle ligne (après corrigerEnTetes) : AL[37] = placesId valide (ex: "PILV14")
// Ancienne ligne (avant corrigerEnTetes) : AL[37] = "Inscrit" ou statut texte
function isNewStructure(row) {
  var al = String(row[37] || '').trim();
  // Un placesId valide : pas d'espace, >3 chars, pas un mot de statut
  var STATUTS = ['Inscrit','Supprimée','En attente','Liste attente','En cours de validation','En attente de place'];
  if (al.indexOf(' ') >= 0) return false;
  if (al.length <= 3) return false;
  for (var s = 0; s < STATUTS.length; s++) {
    if (al.toLowerCase().indexOf(STATUTS[s].toLowerCase().substring(0,6)) >= 0) return false;
  }
  return true;
}
// Nouvelle ligne : AL(37)=ID, AM(38)=Resp, AN(39)=Statut, AO(40)=FFTT
// Ancienne ligne : AJ(35)=ID, AK(36)=Resp, AL(37)=Statut, AM(38)=FFTT
// Fonctions lire* — voir définitions complètes ci-dessous
function lireActiviteId(row) {
  // Ancienne : index 35 (AJ) | Nouvelle : index 37 (AL)
  // Pour les lignes en liste d'attente : AL(37) contient le statut → lire AJ(35)
  var al37 = String(row[37]||'').trim();
  var aj35 = String(row[35]||'').trim();
  // Si AL(37) ressemble à un statut (contient "attente" ou "Inscrit") → ancienne structure ou LA
  var STATUTS = ['inscrit','attente','supprim','cours','validation'];
  var al37IsStatut = STATUTS.some(function(s){ return al37.toLowerCase().indexOf(s) >= 0; });
  if (al37IsStatut) return aj35; // Ligne en attente : actId est en AJ(35)
  return isNewStructure(row) ? al37 : aj35;
}

function lireResponsable(row) {
  // AM (index 38) dans les deux structures selon le log
  return String(row[38]||'').trim();
}

function lireStatutInscription(row) {
  // Ancienne : index 37 (AL) = "Inscrit" | Nouvelle : index 39 (AN)
  return isNewStructure(row) ? String(row[39]||'').trim() : String(row[37]||'').trim();
}

function lireSexe(row) {
  // Ancienne : index 34 (AI) | Nouvelle : index 36 (AK)
  return isNewStructure(row) ? String(row[36]||'').trim() : String(row[34]||'').trim();
}

// ── Console Trésorier : lecture 100% serveur ──────────────────────────────
// Remplace les lectures gviz publiques faites depuis le navigateur (qui
// nécessitaient que le classeur entier soit partagé "Toute personne avec le
// lien" — exposant ainsi les données personnelles de l'onglet Inscriptions
// à quiconque connaît l'URL). Toutes les données nécessaires à la Console
// Trésorier sont lues ici côté serveur via SpreadsheetApp (fonctionne que
// le classeur soit public ou privé) et renvoyées en un seul appel.
// Alias d'ID activité connus (mêmes créneaux réels enregistrés sous des ID
// différents selon l'onglet/l'ancienneté de la donnée) → fusionnés sous un
// seul ID canonique. Doit rester identique à ACT_ID_ALIASES côté Index.html.
var ACT_ID_ALIASES_GAS = {
  'PINGME1830': 'PINGL1745ME1830'
};
function canonActIdGAS(id) {
  var t = String(id || '').trim();
  return ACT_ID_ALIASES_GAS[t] || t;
}

function getStatsTresorierGAS() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  function lireOngletVals(nom) {
    var sh = ss.getSheetByName(nom);
    if (!sh || sh.getLastRow() < 2) return [];
    var nbCols = Math.max(sh.getLastColumn(), 1);
    return sh.getRange(2, 1, sh.getLastRow() - 1, nbCols).getValues();
  }
  function num(v) { var n = Number(v); return (v != null && !isNaN(n)) ? n : 0; }
  function str(v) { return v != null ? String(v) : ''; }
  // Somme d'une colonne en sautant les 3 premières lignes (entêtes/totaux),
  // comme sommeColSkip3() côté client — utilisé pour les onglets paiement.
  function sommeColSkip3(rows, idx) {
    var s = 0;
    for (var i = 3; i < rows.length; i++) {
      var v = num(rows[i][idx]);
      if (v > 0 && v < 1e10) s += v;
    }
    return s;
  }
  function sommeCol(rows, idx) {
    var s = 0;
    for (var i = 0; i < rows.length; i++) {
      var v = num(rows[i][idx]);
      if (v > 1e10) v = 0;
      s += v;
    }
    return s;
  }

  var placesRows = lireOngletVals(SHEET_PLACES);

  // ── Places : effectifs inscrits / coût animateur par activité ──────────
  var placesInscrits  = {};
  var placesAnimateur = {};
  var placesSheetNames = {};
  placesRows.forEach(function(r) {
    var id    = canonActIdGAS(str(r[0]).trim());   // col A
    var idAct = canonActIdGAS(str(r[9]).trim());   // col J
    var cout  = num(r[15]);                        // col P = coût animateur
    var ins   = num(r[33]);                        // col AH = nombre d'inscrits pour l'activité (col J)
    var key   = idAct || id;
    if (key) {
      placesInscrits[key]  = (placesInscrits[key]  || 0) + ins;
      placesAnimateur[key] = (placesAnimateur[key] || 0) + cout;
      placesSheetNames[key] = key;
      if (id && id !== key) {
        placesInscrits[id]   = placesInscrits[key];
        placesAnimateur[id]  = placesAnimateur[key];
        placesSheetNames[id] = key;
      }
    }
  });

  // ── Inscriptions : dossiers/adhérents/activités/FNSMR ───────────────────
  // (Calculé AVANT le passage Isno/HC ci-dessous, pour pouvoir restreindre
  // ce dernier aux seuls onglets réellement concernés par des inscriptions.)
  var codesDossiers = {}, codesAdherents = {}, activitesCounts = {};
  var fnsmrTotal = 0;
  var inscRows = lireOngletVals(SHEET_INSCRIPTIONS);

  // ── Forfait Danse en ligne Country (2-3 séances, 190€/220€) ─────────────
  // Diagnostic (v8.94) : pour un membre inscrit à plusieurs créneaux Country
  // dans le cadre du forfait, le prix forfaitaire entier est imputé en col
  // AB (tarif brut) sur UN SEUL des créneaux, les autres affichant 0€ — ce
  // qui faisait chuter/gonfler artificiellement le "coût moyen" de chaque
  // créneau Country pris isolément. On répartit donc équitablement le total
  // du forfait entre tous les créneaux Country réellement suivis par ce
  // membre, uniquement pour ces 4 créneaux (les autres activités d'un même
  // dossier, ex: Danse moderne dans le même dossier, ne sont pas touchées).
  var COUNTRY_SHEET_IDS = ['COUNV1830', 'COUNV1930', 'COUNV2030', 'COUNV2130'];
  var countryParMembre = {};
  inscRows.forEach(function(row) {
    var code   = str(row[19]).trim();
    var statut = str(row[39]).toLowerCase();
    if (!code || statut.indexOf('supprim') >= 0) return;
    var actId = canonActIdGAS(str(row[37]).trim() || str(row[35]).trim() || str(row[22]).trim());
    if (COUNTRY_SHEET_IDS.indexOf(actId) < 0) return;
    var memKey = code + '_' + str(row[2]) + str(row[3]);
    if (!countryParMembre[memKey]) countryParMembre[memKey] = { total: 0, count: 0 };
    countryParMembre[memKey].total += num(row[27]);
    countryParMembre[memKey].count++;
  });

  inscRows.forEach(function(row) {
    var code   = str(row[19]).trim();
    var statut = str(row[39]).toLowerCase();
    if (!code || statut.indexOf('supprim') >= 0) return;
    codesDossiers[code] = true;
    var memKey = code + '_' + str(row[2]) + str(row[3]);
    codesAdherents[memKey] = true;
    var actNom = str(row[22]).trim();
    var actId  = canonActIdGAS(str(row[37]).trim() || str(row[35]).trim() || actNom);
    var tarif  = num(row[27]);
    if (COUNTRY_SHEET_IDS.indexOf(actId) >= 0) {
      var cp = countryParMembre[memKey];
      if (cp && cp.count > 0) tarif = Math.round((cp.total / cp.count) * 100) / 100;
    }
    if (actNom) {
      if (!activitesCounts[actId]) activitesCounts[actId] = { nom: actNom, inscrits: 0, tarifTotal: 0 };
      activitesCounts[actId].inscrits++;
      activitesCounts[actId].tarifTotal += tarif;
    }
    fnsmrTotal += num(row[30]);
  });

  // ── Isno/HC : lus depuis l'onglet propre à chaque activité (O1:R2) ─────
  // Lire un onglet par activité coûte cher en appels réseau (≈200-300ms
  // chacun) — avec ~60-90 onglets définis dans "Places", ce seul passage
  // peut à lui seul approcher les 25s de timeout côté navigateur. On lit
  // TOUS les onglets définis dans "Places" (comme avant — un filtrage par
  // "activité active" a été essayé puis abandonné : les clés Places et
  // Inscriptions ne coïncident pas toujours pour un même créneau — ex. TT
  // adultes compétiteurs, Méditation 20 séances — et le filtrage faisait
  // disparaître leur %Isno). On tente d'abord un batchGet via l'API Sheets
  // avancée (1 seul appel réseau pour tous les onglets) quand ce service
  // est activé sur le projet, avant de retomber sur la lecture onglet par
  // onglet si l'API avancée n'est pas disponible.
  var sidsToRead = Object.keys(placesSheetNames);

  var isnoCounts = {};
  function parseIsnoVals(sid, vals) {
    var isno = 0, hc = 0;
    for (var ri = 0; ri < Math.min(vals.length, 2); ri++) {
      var row = vals[ri];
      if (!row) continue;
      var isPing = sid.toLowerCase().indexOf('ping') >= 0;
      var v1 = num(row[1]), v2 = num(row[2]), v3 = num(row[3]);
      if (!isNaN(v1) && v1 > 0) {
        isno = isPing ? v2 : v1;
        hc   = isPing ? v3 : v2;
        break;
      }
    }
    if (isno > 0 || hc > 0) isnoCounts[sid] = { isno: isno, hc: hc };
  }

  var sidsRestants = sidsToRead.slice();
  try {
    // Sheets.Spreadsheets.Values.batchGet — service avancé "Google Sheets
    // API" (Extensions > Services dans l'éditeur). S'il n'est pas activé,
    // l'accès à "Sheets" lève une erreur, capturée ci-dessous.
    if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && sidsToRead.length > 0) {
      var ranges = sidsToRead.map(function(sid) { return "'" + sid.replace(/'/g, "''") + "'!O1:R2"; });
      var resp = Sheets.Spreadsheets.Values.batchGet(SHEET_ID, { ranges: ranges });
      (resp.valueRanges || []).forEach(function(vr, i) {
        parseIsnoVals(sidsToRead[i], vr.values || []);
      });
      sidsRestants = [];
    }
  } catch (eBatch) {
    Logger.log('getStatsTresorierGAS: batchGet Isno indisponible (' + eBatch + ') — repli onglet par onglet');
  }
  sidsRestants.forEach(function(sid) {
    try {
      var sh = ss.getSheetByName(sid);
      if (!sh) return;
      parseIsnoVals(sid, sh.getRange(1, 15, 2, 4).getValues()); // O1:R2
    } catch (eIsno) { /* onglet absent ou illisible — ignorer */ }
  });

  // ── Recettes par mode de paiement ───────────────────────────────────────
  var rowsAvoirs = lireOngletVals(SHEET_AVOIRS_UTILISES);
  var avoirsUtilises = 0;
  for (var i = 1; i < rowsAvoirs.length; i++) {
    var v = num(rowsAvoirs[i][4]);
    if (v > 0 && v < 1e10) avoirsUtilises += v;
  }
  var recettes = {
    ha1x:    sommeColSkip3(lireOngletVals(SHEET_HELLOASSO),   4),
    cheq1:   sommeColSkip3(lireOngletVals(SHEET_CHEQUE_1),    5),
    cheq2:   sommeColSkip3(lireOngletVals(SHEET_CHEQUE_2),    5),
    cheq3:   sommeColSkip3(lireOngletVals(SHEET_CHEQUE_3),    5),
    especes: sommeColSkip3(lireOngletVals(SHEET_ESPECES),     4),
    ancv:    sommeColSkip3(lireOngletVals(SHEET_AIDE_ANCV),   4),
    atout:   sommeColSkip3(lireOngletVals(SHEET_AIDE_ATOUT),  4),
    passJ:   sommeColSkip3(lireOngletVals(SHEET_AIDE_PASS_J), 4),
    passS:   sommeColSkip3(lireOngletVals(SHEET_AIDE_PASS_S), 4),
    avoirs:  avoirsUtilises,
    avoirsGeneres:  sommeCol(lireOngletVals(SHEET_AVOIRS), 4),
    remboursements: sommeCol(lireOngletVals(SHEET_REMBOURSEMENTS), 4)
  };
  Logger.log('getStatsTresorierGAS: ' + sidsToRead.length + ' onglets activité scannés pour Isno/HC (sur '
    + Object.keys(placesSheetNames).length + ' définis dans Places)');
  recettes.totalCheques = recettes.cheq1 + recettes.cheq2 + recettes.cheq3;
  recettes.totalAides   = recettes.ancv + recettes.atout + recettes.passJ + recettes.passS;
  recettes.total        = recettes.ha1x + recettes.totalCheques + recettes.especes + recettes.totalAides;

  return {
    nbDossiers:      Object.keys(codesDossiers).length,
    nbAdherents:     Object.keys(codesAdherents).length,
    fnsmrTotal:      Math.round(fnsmrTotal * 100) / 100,
    recettes:        recettes,
    activitesCounts: activitesCounts,
    placesInscrits:  placesInscrits,
    placesAnimateur: placesAnimateur,
    placesSheetNames: placesSheetNames,
    isnoCounts:      isnoCounts
  };
}

function getDossiersSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (!sheet || sheet.getLastRow() < 2) return { dossiers: [] };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
  var dossiers = [];
  data.forEach(function(row) {
    var code = String(row[19] || '').trim();
    if (!code || !code.match(/^FRI-[A-Z0-9]{4}$/)) return;
    // Log diagnostic colonnes clés

    var respFull = lireResponsable(row);
    var respParts = respFull.split(' ');

    dossiers.push({
      code: code, date: formaterDateHeure(row[20]), statut: String(row[21]||''),
      nom_membre: String(row[2]||''), prenom_membre: String(row[3]||''),
      ddn: formaterDdn(row[4]), sexe: String(row[36]||''),
      adresse: String(row[7]||''), cp: String(row[9]||''), ville: String(row[10]||''),
      tel: String(row[14]||''), email: String(row[15]||''),
      tel2: String(row[17]||''), email2: String(row[18]||''),
      nom_responsable:    respParts.length > 1 ? respParts.slice(1).join(' ') : String(row[2]||''),
      prenom_responsable: respParts[0] || String(row[3]||''),
      activite: String(row[22]||''),
      activite_id: (function() {
        // Pour les lignes en attente, chercher l'actId dans toutes les colonnes candidates
        var statut = lireStatutInscription(row);
        if (statut.toLowerCase().indexOf('attente') >= 0) {
          var candidates = [String(row[35]||'').trim(), String(row[37]||'').trim(), String(row[36]||'').trim()];
          var STATUTS = ['inscrit','attente','supprim','cours','validation'];
          for (var ci=0; ci<candidates.length; ci++) {
            var c = candidates[ci];
            if (c && !STATUTS.some(function(s){ return c.toLowerCase().indexOf(s)>=0; })) return c;
          }
        }
        return lireActiviteId(row);
      })(),
      jour: String(row[23]||''), heure: String(row[24]||''),
      lieu: String(row[25]||''), animateur: String(row[26]||''),
      tarif_brut: Number(row[27]||0), tarif: Number(row[29]||0), total: Number(row[31]||0),
      mode_paiement: String(row[32]||''), avoir: String(row[33]||''),
      qs_sante: String(row[34]||''), pass_aide: String(row[35]||''),
      statut_inscription: lireStatutInscription(row),
      fftt_price: 0
    });
  });
  Logger.log('getDossiersSheet v8.3: ' + dossiers.length + ' lignes retournées');
  return { dossiers: dossiers };
}

// ============================================================
// VALIDER PAIEMENT — v8.3 : emailRows APRÈS recalcul remise
// ============================================================

function validerPaiementSheet(rows, code, modeForce, montantHA, payerInfo) {
  Logger.log('=== validerPaiementSheet START — code: ' + code + ' ===');
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  var updated   = 0;
  var modeLabel = 'Chèque';

  if (!sheet) { Logger.log('❌ Onglet Inscriptions introuvable'); return { updated: 0 }; }

  if (sheet.getLastRow() > 1 && code) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();

    // ── Passe 1 : marquer payé + déterminer modeLabel ──
    // Règle : une ligne en liste d'attente reste '⏳ En attente de place'
    //          tant qu'elle n'a pas été basculée par l'admin dans l'onglet activité
    for (var j = 0; j < data.length; j++) {
      var cellCode = String(data[j][19] || '').trim();
      if (cellCode !== code) continue;
      var modePaiement = String(data[j][32] || 'cheque');
      var modeLabelMap = {helloasso:'HelloAsso',cheque:'Chèque',cheque3:'Chèques échelonnés',especes:'Espèces',ancv:'Coupon ANCV',aide:'Soldé par aides'};
      modeLabel = modeLabelMap[modePaiement] || modePaiement;
      // Lire le statut d'inscription (col 40 AN = statut_inscription)
      var statutLigne = String(data[j][39] || '').toLowerCase();
      // 'En attente de place' = vraie liste d'attente (pas encore basculée)
      // 'En cours de validation' = inscription normale → ne pas bloquer
      var isAttentePlace = statutLigne.indexOf('en attente de place') >= 0
                        || statutLigne.indexOf('liste attente') >= 0
                        || statutLigne === 'liste attente';
      Logger.log('valider L'+(j+2)+' code:'+cellCode+' statut:"'+statutLigne+'" isAttentePlace:'+isAttentePlace);
      if (isAttentePlace) {
        // Ligne en liste d'attente : ne pas toucher col 22 (reste '— Non concerné')
        // Col AN (40) : statut inscription reste 'En attente de place'
        sheet.getRange(j + 2, 40)
          .setValue('⏳ En attente de place')
          .setBackground('#b2ebf2').setFontColor('#006064').setFontWeight('bold');
      } else {
        // Ligne inscrite : paiement validé
        sheet.getRange(j + 2, 22)
          .setValue('✅ Règlement validé — ' + modeLabel)
          .setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
        // Col AN (40) : statut inscription = Inscrit
        sheet.getRange(j + 2, 40)
          .setValue('✅ Inscrit')
          .setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
        updated++;
      }
    }

    // ── Passe 2 : recalcul remise 15% → met à jour col 28 dans le Sheet ──
    try { recalculerRemiseDossier(ss, code); } catch(re) { Logger.log('Recalcul remise KO: '+re); }
    SpreadsheetApp.flush();

    // ── Passe 3 : relire les données APRÈS recalcul pour construire emailRows ──
    var dataApr = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
    var emailRows = [];

    for (var k = 0; k < dataApr.length; k++) {
      if (String(dataApr[k][19] || '').trim() !== code) continue;

      var modePaiem2  = String(dataApr[k][32] || 'cheque'); // col 33 AG = Mode paiement
      var placesId    = String(dataApr[k][37] || '').trim(); // col 38 AL = ID Activité
      // Col 28 AB = Tarif brut, col 29 AC = Remise %, col 30 AD = Tarif net
      var tarifBrut  = Number(dataApr[k][27] || 0); // col 28 AB = Tarif brut EUR
      var remisePct  = Number(dataApr[k][28] || 0); // col 29 AC = Remise %
      var tarifFinal = Number(dataApr[k][29] || 0); // col 30 AD = Tarif net
      if (!tarifFinal && tarifBrut) tarifFinal = Math.round(tarifBrut*(1-remisePct/100)*100)/100;
      // Col 34 AH : peut contenir "Avoir : 20.00 €" ou "Avoir suppression : X€"
      // Parser le montant pour les avoirs de règlement
      // Col AH (index 33) = avoir montant numérique
      var avoirMontant = parseFloat(dataApr[k][33]) || 0;
      // Col AJ (index 35) : "PassJeunes:30.00:1ere|Atout:30.00:1234|ANCV:30.00|PASS:50.00"
      var passAideStr = String(dataApr[k][35] || '');
      var ancvMontant = 0;
      var passSportMontant = 0;
      var passJeunesMontant = 0;
      var atoutMontantV = 0;
      var ancvM  = passAideStr.match(/(?:^|\|)ANCV:([\d.]+)/);
      var passM  = passAideStr.match(/(?:^|\|)PASS:([\d.]+)/);
      var pjM    = passAideStr.match(/PassJeunes:([\d.]+)/);
      var atoutM = passAideStr.match(/Atout:([\d.]+)/);
      if (ancvM)  ancvMontant       = parseFloat(ancvM[1])  || 0;
      if (passM)  passSportMontant  = parseFloat(passM[1])  || 0;
      if (pjM)    passJeunesMontant = parseFloat(pjM[1])    || 0;
      if (atoutM) atoutMontantV     = parseFloat(atoutM[1]) || 0;
      var respFull    = String(dataApr[k][38] || ''); // col 39 AM = Responsable
      var respParts   = respFull.split(' ');

      emailRows.push({
        code_dossier:       String(dataApr[k][19] || ''),
        date:               (function(d){ if(!d) return ''; if(d instanceof Date) return Utilities.formatDate(d,'Europe/Paris','dd/MM/yyyy'); var s=String(d); return s.indexOf('GMT')>=0?Utilities.formatDate(new Date(s),'Europe/Paris','dd/MM/yyyy'):s; })(dataApr[k][20]),
        responsable_nom:    respParts.length > 1 ? respParts.slice(1).join(' ') : String(dataApr[k][2] || ''),
        responsable_prenom: respParts[0] || String(dataApr[k][3] || ''),
        adresse:            String(dataApr[k][7]  || ''),
        cp:                 String(dataApr[k][9]  || ''),
        ville:              String(dataApr[k][10] || ''),
        tel1:               String(dataApr[k][14] || ''),
        email1:             String(dataApr[k][15] || ''),
        membre_nom:         String(dataApr[k][2]  || ''),
        membre_prenom:      String(dataApr[k][3]  || ''),
        ddn:                String(dataApr[k][4]  || ''),
        sexe:               String(dataApr[k][36] || ''), // col 37 AK = Sexe
        activite:           String(dataApr[k][22] || ''),
        activite_id:        placesId,
        jour:               String(dataApr[k][23] || ''),
        heure:              String(dataApr[k][24] || ''),
        lieu:               String(dataApr[k][25] || ''),
        animateur:          String(dataApr[k][26] || ''),
        tarif:              tarifFinal,   // tarif après remise
        tarif_brut:         tarifBrut,   // tarif brut (depuis note)
        remise:             remisePct,
        fnsmr:              15,
        total_famille:      Number(dataApr[k][31] || 0), // col 32 AF = Total famille
        mode_paiement:      modePaiem2,
        avoir_montant:      avoirMontant,
        ancv_montant:       ancvMontant,
        pass_sport_montant: passSportMontant,
        solde_net:          Math.max(0, Math.round((Number(dataApr[k][31]||0) - avoirMontant - ancvMontant - passSportMontant - passJeunesMontant - atoutMontantV) * 100) / 100),
        pass_aide:          String(dataApr[k][35] || ''), // col 36 AJ = Pass/Aide
        qs_sante:           String(dataApr[k][34] || ''), // col 35 AI = QS Santé
        // pass_sport_montant et ancv_montant sont déjà renseignés plus haut — NE PAS écraser à 0
        fftt_price:         lireFFTT(dataApr[k]), // col 41 AO ou 39 selon structure
        statut_inscription: lireStatutInscription(dataApr[k]) // col 40 AN = statut inscription
      });
    }

    // MAJ Récapitulatif
    var recap = ss.getSheetByName(SHEET_RECAPITULATIF);
    if (recap && recap.getLastRow() > 1 && code) {
      var rdata = recap.getRange(2, 1, recap.getLastRow() - 1, 14).getValues();
      for (var r = 0; r < rdata.length; r++) {
        var found = rdata[r].some(function(c) { return String(c || '').trim() === code; });
        if (!found) continue;
        recap.getRange(r + 2, 3).setValue('✅ Payé — ' + modeLabel).setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
      }
    }

    // MAJ onglets activités
    // Construire un set des paires code+activite qui sont en liste d'attente
    var attenteSet = {};
    for (var ja = 0; ja < data.length; ja++) {
      if (String(data[ja][19] || '').trim() !== code) continue;
      var stLa = String(data[ja][39] || '').toLowerCase();
      if (stLa.indexOf('attente') >= 0) {
        // Col 1 = nom activite dans onglet inscription (col B = index 1)
        var actNomWait = String(data[ja][1] || '').trim();
        attenteSet[actNomWait] = true;
      }
    }
    Logger.log('Activites en attente pour ' + code + ': ' + JSON.stringify(Object.keys(attenteSet)));
    var protectedSheets = [SHEET_INSCRIPTIONS,SHEET_RECAPITULATIF,SHEET_PLACES,SHEET_CHEQUE_1,SHEET_CHEQUE_2,SHEET_CHEQUE_3,SHEET_AVOIRS,SHEET_AVOIRS_UTILISES,SHEET_AIDE_ANCV,SHEET_AIDE_ATOUT,SHEET_AIDE_PASS_J,SHEET_AIDE_PASS_S,SHEET_ESPECES,SHEET_HELLOASSO];
    ss.getSheets().forEach(function(actSheet) {
      var sn = actSheet.getName();
      if (protectedSheets.indexOf(sn) >= 0) return;
      // Exclure les onglets LA-* : leur statut est gere par basculerListeAttente
      if (sn.indexOf('LA-') === 0) return;
      var lastRow = actSheet.getLastRow();
      if (lastRow < 3) return;
      var aData = actSheet.getRange(3, 1, lastRow - 2, 3).getValues();
      for (var kk = 0; kk < aData.length; kk++) {
        if (String(aData[kk][0] || '').trim() !== code) continue;
        // Verifier si cette activite est en liste d'attente dans l'onglet Inscriptions
        var sheetName = actSheet.getName();
        // Verifier si CETTE LIGNE PRECISE est en attente dans Inscriptions
        // En lisant l'activite_id de la ligne (col B = aData[kk][1] = nomActivite)
        // et en cherchant la correspondance exacte dans Inscriptions
        var nomActInSheet = String(aData[kk][1] || '').trim().toLowerCase();
        var isLigneEnAttente = false;
        for (var jb = 0; jb < data.length; jb++) {
          if (String(data[jb][19] || '').trim() !== code) continue;
          var stLb = String(data[jb][39] || '').toLowerCase();
          if (stLb.indexOf('attente') < 0) continue; // ligne pas en attente
          // Verifier que cette ligne Inscriptions correspond a l'onglet actuel
          // Col B de l'onglet = nom activite / Col 2 Inscriptions = nom activite (index 1)
          var nomActInscr = String(data[jb][1] || '').trim().toLowerCase();
          // Matcher par nom d'activite OU par ID onglet
          var actIdInscr = String(data[jb][6] || '').trim(); // col G = activite_id
          if ((nomActInSheet && nomActInscr && (nomActInSheet === nomActInscr || nomActInscr.indexOf(nomActInSheet) >= 0 || nomActInSheet.indexOf(nomActInscr) >= 0))
            || (actIdInscr && sheetName.indexOf(actIdInscr) >= 0)) {
            isLigneEnAttente = true; break;
          }
        }
        if (isLigneEnAttente) {
          // Activite en attente : laisser le statut inchange ou mettre 'En attente'
          var curVal = String(actSheet.getRange(kk + 3, 3).getValue() || '');
          if (curVal.toLowerCase().indexOf('pay') < 0) {
            actSheet.getRange(kk + 3, 3)
              .setValue('✅ Règlement reçu — En attente de place')
              .setBackground('#b2ebf2').setFontColor('#006064').setFontWeight('bold');
          }
        } else {
          actSheet.getRange(kk + 3, 3).setValue('✅ Payé — ' + modeLabel).setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
          var nomAct = actSheet.getName();
          if (nomAct.indexOf('PING') >= 0) {
            var colO = actSheet.getRange(kk + 3, 15).getValue();
            if (colO && Number(colO) > 0) actSheet.getRange(kk + 3, 15).setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
          }
        }
      }
    });

    Logger.log('emailRows: ' + emailRows.length + ' — updated: ' + updated);

    // Alimenter l'onglet HA pour TOUT paiement HelloAsso (webhook ou manuel)
    if (modeForce === 'helloasso') {
      try { majOngletHelloAssoWebhook(ss, code, payerInfo || {}, montantHA); majTotalHelloAsso(ss); } catch(e) { Logger.log('HelloAsso onglet KO: ' + e); }
    }
    if (emailRows.length > 0) {
      Logger.log('Destinataire : ' + emailRows[0].email1);
      // Injecter les cheques depuis les onglets Cheques 1/2/3 dans emailRows[0]
      if (emailRows.length > 0 && (modeLabel.toLowerCase().indexOf('ch') >= 0)) {
        var chequesEmail = [];
        [SHEET_CHEQUE_1, SHEET_CHEQUE_2, SHEET_CHEQUE_3].forEach(function(nomCh, idx) {
          var shCh = ss.getSheetByName(nomCh);
          if (!shCh || shCh.getLastRow() < 4) return;
          var dCh = shCh.getRange(4, 1, shCh.getLastRow()-3, 7).getValues();
          for (var ci=0; ci<dCh.length; ci++) {
            if (String(dCh[ci][0]||'').trim() !== code) continue;
            chequesEmail.push({
              numero: idx+1,
              numCheque: String(dCh[ci][4]||''),
              banque:    String(dCh[ci][3]||''),
              montant:   Number(dCh[ci][5]||0)
            });
            break;
          }
        });
        if (chequesEmail.length > 0) emailRows[0].cheques = chequesEmail;
        Logger.log('Cheques email injectes: ' + chequesEmail.length + ' pour ' + code);
      }

    // Mettre a jour le statut de tous les elements en 'Valide admin'
      // (les lignes ont ete ecrites a l'inscription avec statut 'En attente')
      try { validerTousElementsPaiement(ss, code, modeLabel); } catch(eVT) { Logger.log('validerTousElementsPaiement KO: ' + eVT); }
      // Fallback : ecrire si inscription directe sans addRegistration
      try { majOngletEspeces(ss, emailRows); majTotalEspeces(ss); } catch(e) { Logger.log('Especes fallback KO: ' + e); }
      try { majOngletAideANCv(ss, emailRows); }      catch(e){ Logger.log('ANCV fallback KO: '+e); }
      try { majOngletAidePassSport(ss, emailRows); } catch(e){ Logger.log('PassSport fallback KO: '+e); }
      try { majOngletAidePassJeunes(ss, emailRows); }catch(e){ Logger.log('PassJeunes fallback KO: '+e); }
      try { majOngletAideAtout(ss, emailRows); }     catch(e){ Logger.log('Atout fallback KO: '+e); }
      // FFTT : deja ecrit a l'inscription, validerTousElementsPaiement met a jour le statut
      // Onglet "Avoirs utilisés" + déduction dans onglet "Avoirs générés"
      try {
        var fCalc2 = _calcFinancier(emailRows);
        if (fCalc2.deducAvoir > 0) {
          ecrireAvoirUtilise(ss, emailRows, fCalc2.deducAvoir);
          // Déduire du solde de l'avoir dans l'onglet "Avoirs générés"
          var r0em = emailRows[0];
          // Parser AVOIR:CODE:MONTANT depuis pass_aide si avoir_code absent
          var codeAvUsed = String(r0em.avoir_code || r0em.code_avoir || '').trim().toUpperCase();
          if (!codeAvUsed) {
            var paStr2 = String(r0em.pass_aide||'');
            var avoirMatch = paStr2.match(/AVOIR:([A-Z0-9]+):([\.\d]+)/);
            if (avoirMatch) codeAvUsed = avoirMatch[1].trim().toUpperCase();
          }
          if (codeAvUsed) {
            utiliserAvoirGAS(codeAvUsed, fCalc2.deducAvoir);
          } else {
            // Fallback : chercher le code avoir dans l'onglet Avoirs generes
            // Priorite : code avoir dans pass_aide (format AVOIR:AVXXXX:20.00)
            var paStrFb = String(r0em.pass_aide||'');
            var avoirMatchFb = paStrFb.match(/AVOIR:([A-Z0-9]+):/);
            var codeAvoirCherche = avoirMatchFb ? avoirMatchFb[1].trim().toUpperCase() : '';
            Logger.log('Fallback avoir cherche code: ' + codeAvoirCherche + ' dans pass_aide: ' + paStrFb);
            if (codeAvoirCherche) {
              // Code avoir connu : utiliser directement
              utiliserAvoirGAS(codeAvoirCherche, fCalc2.deducAvoir);
            } else {
              // Pas de code avoir dans pass_aide : avoir non deduit
              Logger.log('Code avoir introuvable - avoir non deduit dans Avoirs generes');
            }
          }
        }
      } catch(e){ Logger.log('AvoirsUtilises KO: '+e); }
      if (modeForce === 'helloasso') { try { majOngletHelloAsso(ss, emailRows, montantHA, payerInfo); majTotalHelloAsso(ss); } catch(e) { Logger.log('HelloAsso validation KO: ' + e); } }
      var pdfBlob = null, semBlob = null;
      try { pdfBlob = genererFacturePDF(emailRows, modeLabel); } catch(e) { Logger.log('Facture KO : ' + e); }
      // Semainier PDF supprimé
      try { envoyerEmailAdherent(emailRows[0].email1, emailRows, true, modeLabel, pdfBlob); } catch(e) { Logger.log('Email adhérent KO : ' + e); }
      // Email admin à la validation supprimé (quota Gmail) — notification via console admin
    }
  }

  Logger.log('Paiement validé — code: ' + code + ' — lignes MAJ: ' + updated);
  return { updated: updated };
}

// ============================================================
// AJOUTER UNE ACTIVITÉ — v8.1 (inchangé)
// ============================================================

// ══════════════════════════════════════════════════════════════════════════════
// HELPER CENTRAL : recalculer, écrire avoir, envoyer email après modification
// Utilisé par ajouterActiviteDossierSheet ET supprimer*Sheet
// params : { ss, code, actNomClean, typeModif ('ajout'|'suppression'),
//            estRegle, totalPayeAvant, emailAdherent, responsable }
// ══════════════════════════════════════════════════════════════════════════════
// Frais de dossier retenus (non remboursables) lors de la suppression d'une activité d'un dossier réglé
var FRAIS_DOSSIER_SUPPRESSION = 10;

function appliquerModificationDossier(params) {
  var ss           = params.ss;
  var code         = params.code;
  var actNomClean  = params.actNomClean;
  var typeModif    = params.typeModif;   // 'ajout' ou 'suppression'
  var estRegle     = params.estRegle;
  var totalPayeAvant = params.totalPayeAvant || 0; // total réglé avant modif (0 si non réglé)
  var emailAdherent  = params.emailAdherent || '';
  var responsable    = params.responsable   || '';
  var commentaireAdmin = String(params.commentaireAdmin || '').trim();
  var montantModifie   = params.montantModifie === true;
  var membreNomLog      = params.membreNomLog || '';
  var membrePrenomLog   = params.membrePrenomLog || '';
  var dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');

  // ── 1. Recalculer remise ──
  var recalcResult = null;
  try { recalcResult = recalculerRemiseDossier(ss, code); } catch(re) { Logger.log('Recalcul remise KO: ' + re); }
  SpreadsheetApp.flush();

  // ── 2. Lire les lignes du dossier après recalcul ──
  var infos = lireLignesRestantes(ss, code, null);
  var totalApres   = infos.totalNet;
  var aRemise      = infos.aRemise;
  var lignes       = infos.rows;
  var fnsmrDetail  = infos.fnsmrDetail || {};
  var ffttDetail   = infos.ffttDetail  || {};

  Logger.log('appliquerModificationDossier — typeModif:' + typeModif
    + ' estRegle:' + estRegle
    + ' totalPayeAvant:' + totalPayeAvant
    + ' totalApres:' + totalApres
    + ' aRemise:' + aRemise);

  // ── 3. Calcul avoir ou supplément (seulement si réglé) ──
  var avoir = 0;
  var supplement = 0;
  var adhesionNonRemboursee = 0, fraisDossier = 0;
  if (estRegle && totalPayeAvant > 0) {
    var diff = Math.round((totalPayeAvant - totalApres) * 100) / 100;
    if (typeModif === 'suppression' && diff > 0) {
      // Suppression sur dossier réglé : l'adhésion FNSMR déjà réglée n'est jamais remboursée
      // (cas d'un membre qui n'a plus d'activité) et des frais de dossier sont retenus.
      adhesionNonRemboursee = Math.max(0, (Number(params.fnsmrPaye) || 0) - (Number(infos.fnsmr) || 0));
      fraisDossier = FRAIS_DOSSIER_SUPPRESSION;
      diff = Math.round((diff - adhesionNonRemboursee - fraisDossier) * 100) / 100;
      if (diff < 0) diff = 0;
    }
    if (diff > 0)      { avoir      = diff; }  // remboursement
    else if (diff < 0) { supplement = -diff; } // complément à payer
  }

  // ── 4. Écrire l'avoir dans l'onglet Avoirs générés ──
  var codeAvoirGenere = '';
  if (avoir > 0) {
    try {
      var avoirSheet = getOrCreateAvoirSheet(ss);
      codeAvoirGenere = genererCodeAvoir(ss);
      var nextRow = Math.max(avoirSheet.getLastRow() + 1, 2);
      var bg = nextRow % 2 === 0 ? '#fff8e1' : '#ffffff';
      // 9 colonnes : N°Dossier, Nom, Prénom, Activité, Montant, Date, CodeAvoir, SoldeRestant, Statut
      avoirSheet.getRange(nextRow, 1, 1, 9).setValues([[
        code,
        responsable.split(' ').slice(1).join(' '),
        responsable.split(' ')[0] || '',
        actNomClean,
        avoir,
        dateJour,
        codeAvoirGenere,
        avoir,          // Solde restant = montant initial
        'Disponible'
      ]]).setBackground(bg);
      avoirSheet.getRange(nextRow, 5).setFontColor('#e65100').setFontWeight('bold').setNumberFormat('#,##0.00 €');
      avoirSheet.getRange(nextRow, 7).setFontColor('#1a2e22').setFontWeight('bold').setBackground('#d8f3dc');
      // Col H : solde restant = montant avoir initial (pas encore utilisé)
      avoirSheet.getRange(nextRow, 8).setFontColor('#e65100').setFontWeight('bold').setNumberFormat('#,##0.00 €');
      // S'assurer que le solde initial est bien égal au montant avoir
      if (!avoirSheet.getRange(nextRow, 8).getValue()) {
        avoirSheet.getRange(nextRow, 8).setValue(avoirSheet.getRange(nextRow, 5).getValue());
      }
      avoirSheet.getRange(nextRow, 9).setFontColor('#2d6a4f').setFontWeight('bold').setBackground('#d8f3dc');
      avoirSheet.autoResizeColumns(1, 9);
      Logger.log('✅ Avoir écrit : ' + avoir + ' € — code : ' + codeAvoirGenere);
    } catch(ea) { Logger.log('Onglet Avoirs KO : ' + ea.toString()); }
  }

  // ── 5. Construire et envoyer l'email ──
  try {
    // Lignes activités HTML + texte
    var lignesHtml = lignes.map(function(r, i) {
      var t = (aRemise && r.elig) ? r.tarifNet : r.tarifBrut;
      var remHtml = (aRemise && r.elig) ? '<br><small style="color:#2d6a4f">remise 15% → ' + t.toFixed(2) + ' €</small>' : '';
      var isNew   = (typeModif === 'ajout' && r.activite.toLowerCase().indexOf(actNomClean.toLowerCase().substring(0, 12)) >= 0);
      var bg = isNew ? '#e3f2fd' : (i % 2 === 0 ? '#f0f7f3' : '#ffffff');
      var newBadge = isNew ? ' <span style="font-size:10px;background:#1565c0;color:white;border-radius:3px;padding:1px 5px">NOUVEAU</span>' : '';
      return '<tr style="background:' + bg + '">'
        + '<td style="padding:8px 12px">' + r.membre + '</td>'
        + '<td style="padding:8px 12px">' + r.activite + newBadge + '</td>'
        + '<td style="padding:8px 12px;white-space:nowrap">' + r.jour + ' ' + r.heure + '</td>'
        + '<td style="padding:8px 12px;text-align:right">' + r.tarifBrut.toFixed(2) + ' €' + remHtml + '</td>'
        + '</tr>';
    }).join('');

    var fnsmrHtml = Object.keys(fnsmrDetail).map(function(m) {
      return '<tr style="background:#e8f5e9">'
        + '<td style="padding:8px 12px;color:#1b5e20">' + m + '</td>'
        + '<td style="padding:8px 12px;color:#1b5e20" colspan="2">Adhésion FNSMR</td>'
        + '<td style="padding:8px 12px;text-align:right;color:#1b5e20;font-weight:bold">15.00 €</td>'
        + '</tr>';
    }).join('');

    var ffttHtml = Object.keys(ffttDetail).filter(function(m){ return Number(ffttDetail[m]) > 0; }).map(function(m) {
      return '<tr style="background:#e3f2fd">'
        + '<td style="padding:8px 12px;color:#1565c0">' + m + '</td>'
        + '<td style="padding:8px 12px;color:#1565c0" colspan="2">Licence FFTT</td>'
        + '<td style="padding:8px 12px;text-align:right;color:#1565c0;font-weight:bold">' + Number(ffttDetail[m]).toFixed(2) + ' €</td>'
        + '</tr>';
    }).join('');

    var remiseNoteHtml = aRemise
      ? '<div style="margin:10px 0;padding:8px 14px;background:#d8f3dc;border-radius:4px;color:#1b5e20">✅ Remise famille 15% appliquée (≥ 3 activités éligibles).</div>'
      : '';

    var modifBadgeHtml = typeModif === 'ajout'
      ? '<div style="margin:10px 0;padding:8px 14px;background:#e3f2fd;border-left:4px solid #1565c0;border-radius:4px;color:#1565c0">➕ Activité ajoutée : <strong>' + actNomClean + '</strong></div>'
      : '<div style="margin:10px 0;padding:8px 14px;background:#fff3e0;border-left:4px solid #e65100;border-radius:4px;color:#e65100">➖ Activité supprimée : <strong>' + actNomClean + '</strong></div>';

    // Montant à régler : complément d'un dossier déjà réglé, ou total d'un dossier non réglé
    var montantARegler = estRegle ? supplement : (typeModif === 'ajout' ? totalApres : 0);
    var lienPaiement = '';
    if (montantARegler > 0 && emailAdherent) {
      var partsResp = String(responsable || '').trim().split(' ');
      lienPaiement = creerLienPaiementEmail(code, montantARegler, emailAdherent,
        partsResp[0] || '', partsResp.slice(1).join(' ') || '',
        estRegle ? 'Complément dossier' : 'Règlement dossier');
    }
    var boutonPaiementHtml = lienPaiement
      ? helloassoBoutonHtml(lienPaiement, montantARegler.toFixed(2) + ' €',
          'Paiement sécurisé par carte bancaire — ou par chèque à l\'ordre du FRI / lors des permanences')
      : '';

    // Retenues sur l'avoir (suppression sur dossier réglé)
    var detailRetenuesHtml = (fraisDossier > 0 || adhesionNonRemboursee > 0)
      ? '<div style="font-size:12px;color:#555;margin-top:8px;border-top:1px dashed #e0b080;padding-top:6px">'
        + 'Calcul : montant supprimé ' + (Math.round((totalPayeAvant - totalApres) * 100) / 100).toFixed(2) + ' €'
        + (adhesionNonRemboursee > 0 ? ' − adhésion FNSMR non remboursable ' + adhesionNonRemboursee.toFixed(2) + ' €' : '')
        + (fraisDossier > 0 ? ' − frais de dossier ' + fraisDossier.toFixed(2) + ' € (non remboursables)' : '')
        + ' = <strong>' + avoir.toFixed(2) + ' €</strong></div>'
      : '';
    var regleHtml = '';
    if (estRegle) {
      if (typeModif === 'suppression' && avoir <= 0 && (fraisDossier > 0 || adhesionNonRemboursee > 0)) {
        regleHtml = '<div style="margin:16px 0;padding:14px 18px;background:#fff3e0;border-left:4px solid #e65100;border-radius:6px">'
          + '<div style="font-weight:700;color:#e65100;margin-bottom:6px">Aucun avoir</div>' + detailRetenuesHtml + '</div>';
      } else if (avoir > 0) {
        regleHtml = '<div style="margin:16px 0;padding:14px 18px;background:#fff3e0;border-left:4px solid #e65100;border-radius:6px">'
          + '<div style="font-weight:700;color:#e65100;margin-bottom:6px">💳 Avoir enregistré : ' + avoir.toFixed(2) + ' €</div>'
          + (codeAvoirGenere ? '<div style="background:#d8f3dc;border:2px solid #52b788;border-radius:8px;padding:10px 14px;margin:8px 0;text-align:center">'
            + '<div style="font-size:11px;font-weight:700;color:#2d6a4f;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px">Votre code avoir</div>'
            + '<div style="font-family:monospace;font-size:26px;font-weight:900;color:#1a2e22;letter-spacing:6px">' + codeAvoirGenere + '</div>'
            + '<div style="font-size:11px;color:#555;margin-top:4px">Conservez ce code — il vous sera demandé lors de votre prochaine inscription</div>'
            + '</div>' : '')
          + '<div style="font-size:13px;color:#555">Ce montant sera déduit de votre prochain règlement sur présentation de ce code.</div>'
          + detailRetenuesHtml
          + '</div>';
      } else if (supplement > 0) {
        regleHtml = '<div style="margin:16px 0;padding:14px 18px;background:#fff8e1;border-left:4px solid #f9a825;border-radius:6px">'
          + '<div style="font-weight:700;color:#f57f17;margin-bottom:4px">⚠️ Supplément à régler : ' + supplement.toFixed(2) + ' €</div>'
          + '<div style="font-size:13px;color:#555;margin-bottom:4px">Déjà réglé : ' + totalPayeAvant.toFixed(2) + ' € — nouveau total du dossier : ' + totalApres.toFixed(2) + ' €'
          + ' (l\'adhésion déjà réglée n\'est pas réclamée à nouveau).</div>'
          + '<div style="font-size:13px;color:#555">Merci de régler ce complément via HelloAsso ou par chèque à l\'ordre du FRI.</div>'
          + '</div>' + boutonPaiementHtml;
      } else {
        regleHtml = '<div style="margin:12px 0;padding:10px 14px;background:#f0f7f3;border-radius:4px;font-size:13px;color:#2d6a4f">✅ Aucune différence de règlement.</div>';
      }
    } else if (montantARegler > 0) {
      // Dossier pas encore réglé : on réclame la totalité du dossier
      regleHtml = '<div style="margin:16px 0;padding:14px 18px;background:#fff8e1;border-left:4px solid #f9a825;border-radius:6px">'
        + '<div style="font-weight:700;color:#f57f17;margin-bottom:4px">⚠️ Votre dossier n\'est pas encore réglé — montant total à régler : ' + montantARegler.toFixed(2) + ' €</div>'
        + '<div style="font-size:13px;color:#555">Via HelloAsso ci-dessous, par chèque à l\'ordre du FRI ou lors des permanences.</div>'
        + '</div>' + boutonPaiementHtml;
    }

    var bodyHtml = '<div style="font-family:Arial,sans-serif;max-width:620px;margin:auto">'
      + '<div style="background:#2d6a4f;color:white;padding:18px 24px;border-radius:8px 8px 0 0">'
      + '<h2 style="margin:0">Foyer Rural d\'Isneauville</h2>'
      + '<p style="margin:4px 0;opacity:.85">Modification dossier — N°' + code + '</p></div>'
      + '<div style="padding:20px 24px;background:#f9fafb;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px">'
      + '<p>Bonjour <strong>' + responsable + '</strong>,</p>'
      + modifBadgeHtml + remiseNoteHtml
      + (lignes.length > 0
          ? '<table width="100%" style="border-collapse:collapse;margin:12px 0">'
            + '<thead><tr style="background:#2d6a4f;color:white">'
            + '<th style="padding:8px 12px;text-align:left">Membre</th>'
            + '<th style="padding:8px 12px;text-align:left">Activité</th>'
            + '<th style="padding:8px 12px;text-align:left">Horaire</th>'
            + '<th style="padding:8px 12px;text-align:right">Tarif brut</th>'
            + '</tr></thead><tbody>'
            + lignesHtml + fnsmrHtml + ffttHtml
            + '</tbody></table>'
            + '<p style="text-align:right;font-weight:bold;font-size:16px;margin:4px 0">Total dossier : ' + totalApres.toFixed(2) + ' €</p>'
          : '<p><em>Il ne reste plus d\'activité dans votre dossier.</em></p>')
      + regleHtml
      + blocCommentaireAdminHtml(commentaireAdmin)
      + '<hr style="border:none;border-top:1px solid #e0e0e0;margin:16px 0">'
      + '<small style="color:#888">Contact : frisneauville@orange.fr | 02.35.59.01.01</small></div></div>';

    // Corps texte
    var lignesTexte = lignes.map(function(r) {
      var t = (aRemise && r.elig) ? Math.round(r.tarifBrut * 0.85 * 100) / 100 : r.tarifBrut;
      var remLabel = (aRemise && r.elig) ? ' (remise 15% → ' + t.toFixed(2) + ' €)' : '';
      return '  • ' + r.membre + ' — ' + r.activite + ' : ' + r.tarifBrut.toFixed(2) + ' €' + remLabel;
    });
    Object.keys(fnsmrDetail).forEach(function(m) { lignesTexte.push('  • ' + m + ' — Adhésion FNSMR : 15.00 €'); });
    Object.keys(ffttDetail).forEach(function(m)  { if (Number(ffttDetail[m]) > 0) lignesTexte.push('  • ' + m + ' — Licence FFTT : ' + Number(ffttDetail[m]).toFixed(2) + ' €'); });

    var regleTexte = estRegle
      ? (avoir > 0 ? '\n\nAvoir enregistré : ' + avoir.toFixed(2) + ' €'
          + (fraisDossier > 0 ? ' (frais de dossier de ' + fraisDossier.toFixed(2) + ' € retenus'
             + (adhesionNonRemboursee > 0 ? ', adhésion FNSMR non remboursable' : '') + ')' : '')
        : supplement > 0 ? '\n\nSupplément à régler : ' + supplement.toFixed(2) + ' €'
        : '\n\nAucune différence de règlement.')
      : (montantARegler > 0 ? '\n\nVotre dossier n\'est pas encore réglé — montant total à régler : ' + montantARegler.toFixed(2) + ' €' : '');
    if (estRegle && supplement > 0) regleTexte += ' (déjà réglé : ' + totalPayeAvant.toFixed(2) + ' €, nouveau total : ' + totalApres.toFixed(2) + ' €)';
    if (lienPaiement) regleTexte += '\n\nRégler ' + montantARegler.toFixed(2) + ' € par HelloAsso : ' + lienPaiement;

    var bodyTexte = 'Bonjour ' + responsable + ',\n\n'
      + (typeModif === 'ajout' ? 'Activité ajoutée' : 'Activité supprimée') + ' : "' + actNomClean + '" — Dossier N°' + code + '\n\n'
      + 'Détail complet du dossier :\n' + lignesTexte.join('\n')
      + '\n\nTotal dossier : ' + totalApres.toFixed(2) + ' €'
      + (aRemise ? '\n\nRemise famille 15% appliquée (≥ 3 activités éligibles).' : '')
      + regleTexte
      + (commentaireAdmin ? '\n\nNote de l\'équipe FRI : ' + commentaireAdmin : '')
      + '\n\nContact : frisneauville@orange.fr | 02.35.59.01.01\nCordialement,\nLe Foyer Rural d\'Isneauville';

    // Traçabilité : n'écrit dans "Commentaires admin" que si un commentaire a
    // été saisi ou si le montant a été modifié manuellement par l'admin.
    logCommentaireAdmin(ss, (typeModif === 'ajout' ? 'Ajout activité' : 'Suppression activité'),
      code, membreNomLog, membrePrenomLog, actNomClean,
      (avoir > 0 ? avoir : (supplement > 0 ? supplement : totalApres)),
      commentaireAdmin, montantModifie);

    if (emailAdherent && emailAdherent.indexOf('@') > 0) {
      var objet = '[FRI] ' + (typeModif === 'ajout' ? 'Ajout activité' : 'Modification dossier')
        + ' — N°' + code + ' — ' + actNomClean;
      var opts = { name: NOM_ASSO, replyTo: EMAIL_ADMIN, htmlBody: bodyHtml };
      envoyerEmail(emailAdherent, objet, bodyTexte, opts);
      Logger.log('✅ Email modif → adhérent : ' + emailAdherent);
    }

    // Email admin
    var adminSuffix = estRegle
      ? (avoir > 0 ? ' | Avoir : ' + avoir.toFixed(2) + ' €'
        : supplement > 0 ? ' | Supplément : ' + supplement.toFixed(2) + ' €'
        : ' | Pas de différence')
      : ' (non réglé)';
    envoyerEmail(EMAIL_ADMIN,
      '[FRI Admin] ' + (typeModif === 'ajout' ? 'Ajout' : 'Suppression') + ' — N°' + code + adminSuffix,
      (typeModif === 'ajout' ? '➕' : '➖') + ' ' + actNomClean + ' | ' + code
        + '\nRemise famille : ' + (aRemise ? 'OUI' : 'NON')
        + '\nDétail :\n' + lignesTexte.join('\n')
        + '\nTotal dossier : ' + totalApres.toFixed(2) + ' €',
      { name: 'FRI Admin' });
    Logger.log('✅ Email modif → admin');
  } catch(emailErr) { Logger.log('Email modif KO : ' + emailErr.toString()); }

  // Mettre à jour col I (Total EUR) du Récapitulatif
  try {
    var recapSheet = ss.getSheetByName(SHEET_RECAPITULATIF);
    if (recapSheet && recapSheet.getLastRow() > 1) {
      var recapData = recapSheet.getRange(2, 1, recapSheet.getLastRow()-1, 9).getValues();
      for (var rr=0; rr<recapData.length; rr++) {
        var hasCode = recapData[rr].some(function(c){ return String(c||'').trim()===code; });
        if (!hasCode) continue;
        recapSheet.getRange(rr+2, 9).setValue(totalApres)
          .setFontColor('#1565c0').setFontWeight('bold');
        Logger.log('✅ Recap col I mise à jour: '+code+' total='+totalApres);
        break;
      }
    }
  } catch(eR) { Logger.log('Recap MAJ KO: '+eR); }

  return { totalApres: totalApres, avoir: avoir, supplement: supplement, aRemise: aRemise };
}

function ajouterActiviteDossierSheet(payload) {
  var ss            = SpreadsheetApp.openById(SHEET_ID);
  var code          = payload.code          || '';
  var actNom        = payload.actNom        || '';
  var actId         = payload.actId         || '';
  var placesId      = payload.placesId      || getPlacesId(actId);
  var membreNom     = payload.membreNom     || '';
  var membrePrenom  = payload.membrePrenom  || '';
  var membreDdn     = formaterDdn(payload.membreDdn || '');
  var membreSexe    = payload.membreSexe    || '';
  var jour          = payload.jour          || '';
  var heure         = payload.heure         || '';
  var lieu          = payload.lieu          || '';
  var animateur     = payload.animateur     || '';
  var tarif         = parseFloat(payload.tarif) || 0;
  var estRegle      = payload.estRegle === true || payload.estRegle === 'true';
  var statutInscrit = payload.statutInscription || '';
  var modePaiement  = String(payload.modePaiement || 'cheque');
  var dateJour      = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');

  var emailAdherent = '', responsable = '', adresse = '', cp = '', ville = '', tel1 = '', modePaiement = 'cheque';
  var totalActifExistant = 0;
  var membresExistants   = {};
  var estRegleServeur    = false;

  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][19] || '').trim() !== code) continue;
      if (!emailAdherent) {
        emailAdherent = String(data[i][15] || '');
        adresse       = String(data[i][7]  || '');
        cp            = String(data[i][9]  || '');
        ville         = String(data[i][10] || '');
        tel1          = String(data[i][14] || '');
        responsable   = String(data[i][38] || '');
        modePaiement  = String(data[i][32] || 'cheque');
      }
      var statLigne = String(data[i][21] || '');
      if (statLigne.indexOf('supprimée') >= 0) continue;
      // Dossier réglé ? même règle que modifierActivite (statut de paiement « validé » / « payé »)
      var statLower = statLigne.toLowerCase();
      if (statLower.indexOf('valid\u00e9') >= 0 || statLower.indexOf('pay\u00e9') >= 0) estRegleServeur = true;
      var isWaiting = String(data[i][39] || '').toLowerCase().indexOf('attente') >= 0;
      if (!isWaiting) totalActifExistant += Number(data[i][27] || 0);
      var mbKey = (String(data[i][3] || '') + ' ' + String(data[i][2] || '')).trim();
      membresExistants[mbKey] = true;
    }
  }

  // Montant déjà dû / réglé AVANT l'ajout : activités (remise comprise) + adhésions FNSMR + FFTT.
  // (Avant, l'adhésion déjà réglée n'était pas comptée : l'email réclamait 15 € de trop.)
  var totalDossierAvant = 0;
  try { totalDossierAvant = lireLignesRestantes(ss, code, null).totalNet || 0; } catch(eAv) { Logger.log('Total avant ajout KO : ' + eAv); }

  var nouveauMembreKey  = (membrePrenom + ' ' + membreNom).trim();
  var estNouveauMembre  = !membresExistants[nouveauMembreKey];
  var fnsmrNouvelleLigne = estNouveauMembre ? 15 : 0;
  var tarifNouvelleLigne = (statutInscrit.indexOf('attente') >= 0) ? 0 : tarif;
  var nouveauTotal = Math.round((totalActifExistant + tarifNouvelleLigne + fnsmrNouvelleLigne) * 100) / 100;

  Logger.log('ajouterActiviteDossierSheet — code:' + code + ' act:' + actNom
    + ' tarif:' + tarif + ' totalExistant:' + totalActifExistant
    + ' fnsmr:' + fnsmrNouvelleLigne + ' nouveauTotal:' + nouveauTotal);
  // Traçabilité debug — vérifie que ce que l'admin a saisi manuellement arrive
  // bien jusqu'ici depuis le payload envoyé par le frontend.
  Logger.log('ajouterActiviteDossierSheet — [DEBUG commentaire/montant] '
    + 'commentaireAdmin reçu:"' + String(payload.commentaireAdmin || '') + '" '
    + 'montantModifie reçu:' + payload.montantModifie + ' (typeof ' + typeof payload.montantModifie + ')');

  var modeLabel = modePaiement === 'helloasso' ? 'HelloAsso' : (modePaiement.charAt(0).toUpperCase() + modePaiement.slice(1));
  var statut = estRegle ? '✅ Payé — ' + modeLabel : '⏳ En cours de validation — ' + modeLabel;

  var newRow = [
    // Col 1-11 : identite membre
    '', '', membreNom, membrePrenom, membreDdn, '', '', adresse, '', cp, ville,
    // Col 12-19
    '', 'FR', '', tel1, emailAdherent, actNom, '', '',
    // Col 20-22 : dossier
    code, dateJour, statut,
    // Col 23-27 : activite
    actNom, jour, heure, lieu, animateur,
    // Col 28-32 : tarifs
    tarif, 0, tarif, fnsmrNouvelleLigne, nouveauTotal,
    // Col 33-39 : mode, avoir, QS, pass, sexe, idAct, responsable
    modePaiement, '', '', '', membreSexe, placesId, responsable,
    // Col 40 AN : statut inscription, Col 41 AO : licence FFTT
    statutInscrit || 'Inscrit', 0
  ]; // 41 colonnes
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, 1, 41).setValues([newRow]);
  // Fond bleu clair (#bbdefb) pour toute activité ajoutée par admin (réglée ou non)
  // Fond vert clair (#d8f3dc) si dossier réglé et paiement validé
  // Fond cyan (#b2ebf2) si en attente de place
  var bgColor = statutInscrit.indexOf('attente') >= 0 ? '#b2ebf2'
              : estRegle ? '#d8f3dc'
              : '#bbdefb'; // bleu clair = ajout admin avant règlement
  sheet.getRange(startRow, 1, 1, 41).setBackground(bgColor);
  sheet.getRange(startRow, 20).setBackground('#e8f4fd').setFontColor('#1565c0').setFontWeight('bold');
  // ── Nouvelles colonnes AC et AD pour l'activité ajoutée ──
  // (estEligible/placesIdCible n'existent que dans supprimerActiviteDossierSheet —
  //  ici il faut estEligibleRemise(pid, commune) avec les variables locales placesId/ville)
  var eligAjout = estEligibleRemise(placesId, communeFromVille(ville));
  // col 29 AC : Éligible remise
  sheet.getRange(startRow, 29).setValue(eligAjout ? 1 : 0)
    .setBackground(eligAjout ? '#e8f5e9' : '#fce4ec')
    .setFontColor(eligAjout ? '#1b5e20' : '#b71c1c').setFontWeight('bold');
  // col 30 AD : Tarif net — formule Excel
  sheet.getRange(startRow, 30)
    .setFormula(formuleAD(startRow))
    .setBackground('#d8f3dc').setFontColor('#1b5e20');
  // col 32 AF : Total famille — formule Excel (même dossier)
  // NB : les autres lignes du dossier ont déjà cette formule — elle se recalcule auto
  // On l'écrit sur la nouvelle ligne pour cohérence
  sheet.getRange(startRow, 32)
    .setValue(0) // AF calculé par calcTotalFamille() après
    .setBackground('#e8f4fd').setFontColor('#1565c0').setFontWeight('bold');
  // col 28 AB : tarif brut déjà dans newRow[27]

  if (sheet.getLastRow() > 1) {
    var data2 = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
    for (var j = 0; j < data2.length; j++) {
      if (String(data2[j][19] || '').trim() === code) sheet.getRange(j + 2, 32).setValue(nouveauTotal);
    }
  }
  SpreadsheetApp.flush();

  // ── Si liste d'attente : ajouter à l'onglet LA (format canonique, cf. majOngletListeAttente) et envoyer email ──
  if (statutInscrit.indexOf('attente') >= 0) {
    // Réutiliser majOngletListeAttente pour garder EXACTEMENT le même format
    // (N° Dossier, Date, Statut, Nom, Prénom, Tél, Email, Activité, Jour, Heure)
    // que celui créé à l'inscription initiale — un format différent (ex: colonnes
    // Date naiss./Sexe/Responsable/Ville) décale les données dans l'onglet LA existant.
    try {
      majOngletListeAttente(ss, placesId, [{
        code_dossier:   code,
        membre_nom:     membreNom,
        membre_prenom:  membrePrenom,
        tel1:           tel1,
        email1:         emailAdherent,
        activite:       actNom,
        jour:           jour,
        heure:          heure
      }]);
      Logger.log('✅ Onglet LA mis à jour (format canonique) pour ' + code);
    } catch(eLA) { Logger.log('LA tab KO : ' + eLA); }

    // Envoyer email à l'adhérent : inscription en liste d'attente
    try {
      if (emailAdherent) {
        var subjLA = '[FRI] Dossier ' + code + ' - Inscription en liste d\'attente';
        var bodyLA = 'Bonjour ' + membrePrenom + ' ' + membreNom + ',\n\n'
          + 'Votre inscription à l\'activité suivante a été enregistrée en liste d\'attente :\n\n'
          + '  Activité : ' + actNom + '\n'
          + '  Horaire  : ' + jour + ' ' + heure + '\n'
          + '  Lieu     : ' + lieu + '\n\n'
          + 'Nous vous contacterons dès qu\'une place se libère.\n\n'
          + 'Foyer Rural d\'Isneauville\nfrisneauville@orange.fr | 02.35.59.01.01';
        envoyerEmail(emailAdherent, subjLA, bodyLA, {
          name: NOM_ASSO, replyTo: EMAIL_ADMIN
        });
        Logger.log('✅ Email LA envoyé : ' + code + ' → ' + emailAdherent);
      }
    } catch(eMailLA) { Logger.log('Email LA KO : ' + eMailLA); }
  }

  // N'écrire dans l'onglet {placesId} QUE si l'adhérent est inscrit (pas liste attente)
  if (statutInscrit.indexOf('attente') < 0) {
    var rowAct = [code, dateJour, statut, membreNom, membrePrenom, membreDdn, membreSexe, responsable, tel1, emailAdherent, ville, '', modePaiement, ''];
    var actSheet = ss.getSheetByName(placesId);
    if (!actSheet) {
      actSheet = ss.insertSheet(placesId);
      var couleur2 = getCouleurActivite(actId);
      var hdr = ['N° Dossier','Date','Statut','Nom','Prenom','Date naiss.','Sexe','Responsable','Telephone','Email','Ville','QS Sante','Paiement','Pass Aide'];
      actSheet.getRange(1, 1, 1, hdr.length).setValues([hdr]).setBackground(couleur2).setFontColor('#ffffff').setFontWeight('bold');
      actSheet.getRange(2, 1, 1, hdr.length).setValue(actNom+' | '+animateur+' | '+jour+' '+heure+' | '+lieu).setFontStyle('italic').setFontSize(9).setFontColor('#555').setBackground('#f5f5f5');
      actSheet.setFrozenRows(2);
    }
    var nextActRow = Math.max(actSheet.getLastRow() + 1, 3);
    actSheet.getRange(nextActRow, 1, 1, rowAct.length).setValues([rowAct]).setBackground(nextActRow % 2 === 0 ? '#f0f7f3' : '#ffffff');
    if (estRegle) actSheet.getRange(nextActRow, 3).setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
  }

  if (statutInscrit.indexOf('attente') < 0) {
    try {
      var placesSheet = ss.getSheetByName(SHEET_PLACES);
      if (placesSheet && placesSheet.getLastRow() > 1) {
        var pData = placesSheet.getRange(2, 1, placesSheet.getLastRow() - 1, 5).getValues();
        for (var p = 0; p < pData.length; p++) {
          if (String(pData[p][0]).trim() !== placesId) continue;
          var inscrits2 = (parseInt(pData[p][3]) || 0) + 1;
          var cap2      = parseInt(pData[p][2]) || 20;
          placesSheet.getRange(p + 2, 4).setValue(inscrits2);
          placesSheet.getRange(p + 2, 5).setValue(Math.max(0, cap2 - inscrits2));
          break;
        }
      }
    } catch(ep) { Logger.log('Places KO : ' + ep); }
  }

  // ── Recalcul remise + email + avoir via helper central ──
  var actNomClean2 = actNom.replace(/\n/g, ' — ');
  var modeLabel2 = modePaiement==='helloasso'?'HelloAsso':modePaiement==='especes'?'Espèces':'Chèque';
  // Ecrire dans l'onglet paiement si dossier regle et activite non en attente
  if (estRegle && statutInscrit.indexOf('attente') < 0) {
    try {
      var dA = Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy');
      if (modePaiement==='especes') {
        var shEA=getOrCreateEspecesSheet(ss); var nrEA=Math.max(shEA.getLastRow()+1,4);
        shEA.getRange(nrEA,1,1,7).setValues([[code,membreNom,membrePrenom,emailAdherent,tarifNouvelleLigne,dA,'✅ Validé — Admin']]);
        shEA.getRange(nrEA,5).setFontColor('#4e342e').setFontWeight('bold');
        shEA.getRange(nrEA,7).setFontColor('#1b5e20').setFontWeight('bold');
        majTotalEspeces(ss);
      } else if (modePaiement==='cheque') {
        var shCA=getOrCreateChequeSheet(ss,SHEET_CHEQUE_1); var nrCA=Math.max(shCA.getLastRow()+1,4);
        shCA.getRange(nrCA,1,1,7).setValues([[code,membreNom,membrePrenom,'','',tarifNouvelleLigne,'✅ Validé — Admin']]);
        shCA.getRange(nrCA,6).setFontColor('#1b5e20').setFontWeight('bold');
        shCA.getRange(nrCA,7).setFontColor('#1b5e20').setFontWeight('bold');
        majTotalCheque(ss,SHEET_CHEQUE_1);
      } else if (modePaiement==='helloasso') {
        var shHA2=getOrCreateHelloAssoSheet(ss); var nrHA2=Math.max(shHA2.getLastRow()+1,4);
        shHA2.getRange(nrHA2,1,1,9).setValues([[code,membreNom,membrePrenom,emailAdherent,tarifNouvelleLigne,tarifNouvelleLigne,'HelloAsso',dA,'✅ Validé — Admin']]);
        shHA2.getRange(nrHA2,6).setFontColor('#1b5e20').setFontWeight('bold');
        majTotalHelloAsso(ss);
      }
      Logger.log('✅ Paiement ajout: '+code+' '+modePaiement+' '+tarifNouvelleLigne+'EUR');
    } catch(ePay){Logger.log('Paiement ajout KO: '+ePay);}
  }
  var modifResult = appliquerModificationDossier({
    ss: ss, code: code, actNomClean: actNomClean2,
    // Email : le statut réel du dossier (Sheet) décide. Dossier réglé → seul le complément
    // est réclamé ; dossier non réglé → le total du dossier est réclamé.
    typeModif: 'ajout', estRegle: estRegleServeur,
    totalPayeAvant: estRegleServeur ? totalDossierAvant : 0,
    emailAdherent: emailAdherent, responsable: responsable,
    modePaiement: modeLabel2,
    commentaireAdmin: payload.commentaireAdmin || '',
    montantModifie:   payload.montantModifie === true || payload.montantModifie === 'true',
    membreNomLog: membreNom, membrePrenomLog: membrePrenom
  });
  // Retourner le nouveau total depuis le Sheet (après recalcul remise)
  var nouveauTotalReel = 0;
  try {
    SpreadsheetApp.flush();
    var sheetT = ss.getSheetByName(SHEET_INSCRIPTIONS);
    if (sheetT && sheetT.getLastRow() > 1) {
      var dataT = sheetT.getRange(2, 1, sheetT.getLastRow() - 1, 32).getValues();
      for (var ti = 0; ti < dataT.length; ti++) {
        if (String(dataT[ti][19] || '').trim() === code) {
          var v = Number(dataT[ti][31] || 0);
          if (v > 0) { nouveauTotalReel = v; break; }
        }
      }
    }
  } catch(et) { Logger.log('Lecture total ajout KO : ' + et); }
  return { inserted: 1, nouveauTotal: nouveauTotalReel };
}

// ============================================================
// SUPPRIMER UNE ACTIVITÉ (DOSSIER RÉGLÉ) — v8.8
// Avoir = Total_payé - Total_dû_après_suppression (sans remise si < 3 élig.)
// + recalcul remise sur lignes restantes + email cohérent Sheet
// ============================================================
function supprimerActiviteDossierSheet(code, actNom, actId, membreNom, avoirMontant, placesIdParam, commentaireAdmin, montantModifie) {
  Logger.log('supprimerActiviteDossierSheet v8.8 — code:'+code+' actId:'+actId);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var placesIdCible = placesIdParam || getPlacesId(actId || '');
  var deleted = 0;
  var actNomClean = actNom.replace(/\n/g, ' — ');
  var dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  var emailAdherent = '', responsable = '';

  var NON_REMISABLE = ['PINGL1945/ME1945','PINGM1930/ME21','PINGJ1530',
    'PINGL18/J17','PINGME1730','PINGL1745ME1830','JAZME1015','JAZME1115','JAZME1315','JAZME1415','JAZME1515',
    'JAZME1615','JAZME1715','JAZJ1745','MNOME/S10','COUNV1830','COUNV1930','COUNV2030','COUNV2130'];
  function estEligible(pid) {
    if (!pid) return false;
    if (NON_REMISABLE.indexOf(pid) >= 0) return false;
    if (pid.indexOf('PING') >= 0) return false;
    return true;
  }

  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);

  // ── Passe 1 : lire toutes les lignes actives AVANT suppression ──
  var totalPayeActuel = 0;
  var lignesActives = [];
  var membresUniques = {}, ffttVus = {};
  var fnsmrTotal = 0, ffttTotal = 0;
  var seenPidsAvant = {}, nbEligiblesAvant = 0;

  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][19] || '').trim() !== code) continue;
      var st38 = String(data[i][39] || '').toLowerCase();
      var st22 = String(data[i][21] || '').toLowerCase();
      if (st38.indexOf('supprimée') >= 0 || st22.indexOf('supprimée') >= 0) continue;
      if (st38.indexOf('attente') >= 0) continue;

      var pid = String(data[i][37] || '').trim();
      var tarifActuel = Number(data[i][27] || 0);
      var note = sheet.getRange(i + 2, 28).getNote() || '';
      var brutMatch = note.match(/Tarif brut\s*:\s*([\d.]+)/);
      var tarifBrut = brutMatch ? parseFloat(brutMatch[1]) : tarifActuel;
      var memKey = String(data[i][3] || '') + ' ' + String(data[i][2] || '');
      var ffttVal = 0; // FFTT supprimé

      if (!membresUniques[memKey]) { membresUniques[memKey] = true; fnsmrTotal += 15; }
      if (ffttVal > 0 && !ffttVus[memKey]) { ffttVus[memKey] = true; ffttTotal += ffttVal; }

      // Détecter l'email et le responsable
      if (!emailAdherent) emailAdherent = String(data[i][15] || '');
      if (!responsable)   responsable   = String(data[i][38] || '');

      var rowActNorm = String(data[i][22] || '').replace(/\n/g, ' — ').toLowerCase();
      var actNomNorm = actNomClean.toLowerCase();
      var isCible = (pid === placesIdCible)
        || (actNomNorm.length >= 8 && rowActNorm.indexOf(actNomNorm.substring(0, 15)) >= 0);

      lignesActives.push({ idx: i, pid: pid, tarifActuel: tarifActuel, tarifBrut: tarifBrut, isCible: isCible,
        activite: String(data[i][22] || '').replace(/\n/g, ' — '),
        membre: String(data[i][3] || '') + ' ' + String(data[i][2] || ''),
        jour: String(data[i][23] || ''), heure: String(data[i][24] || ''), elig: estEligible(pid) });

      totalPayeActuel += tarifActuel;
      if (estEligible(pid) && !seenPidsAvant[pid]) { seenPidsAvant[pid] = true; nbEligiblesAvant++; }
    }
  }

  var totalPayeAvecFNSMR = Math.round((totalPayeActuel + fnsmrTotal + ffttTotal) * 100) / 100;

  // ── Calcul total dû APRÈS suppression ──
  var lignesRestantes = lignesActives.filter(function(l) { return !l.isCible; });
  var seenPidsApres = {}, nbEligiblesApres = 0;
  lignesRestantes.forEach(function(l) {
    if (estEligible(l.pid) && !seenPidsApres[l.pid]) { seenPidsApres[l.pid] = true; nbEligiblesApres++; }
  });
  var aRemiseApres = nbEligiblesApres >= 3;

  // Recalculer FNSMR et FFTT sur les membres restants uniquement
  var membresApres = {}, ffttApres = {}, fnsmrApres = 0, ffttApresTotal = 0;
  lignesRestantes.forEach(function(l) {
    var mk = l.membre;
    if (!membresApres[mk]) { membresApres[mk] = true; fnsmrApres += 15; }
  });
  // FFTT : recalculer depuis les lignes restantes
  if (sheet && sheet.getLastRow() > 1) {
    var dataFF2 = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
    var ffttVus2 = {};
    for (var fi2 = 0; fi2 < dataFF2.length; fi2++) {
      if (String(dataFF2[fi2][19] || '').trim() !== code) continue;
      if (String(dataFF2[fi2][39] || '').toLowerCase().indexOf('supprimée') >= 0) continue;
      var pid2 = String(dataFF2[fi2][37] || '').trim();
      if (pid2 === placesIdCible) continue; // exclure la ligne supprimée
      var memK2 = String(dataFF2[fi2][3] || '') + ' ' + String(dataFF2[fi2][2] || '');
      var ffV2 = Number(dataFF2[fi2][40] || 0);
      if (ffV2 > 0 && !ffttVus2[memK2]) { ffttVus2[memK2] = true; ffttApresTotal += ffV2; }
    }
  }

  var totalDuApres = 0;
  lignesRestantes.forEach(function(l) {
    totalDuApres += (aRemiseApres && estEligible(l.pid))
      ? Math.round(l.tarifBrut * 0.85 * 100) / 100 : l.tarifBrut;
  });
  totalDuApres = Math.round((totalDuApres + fnsmrApres + ffttApresTotal) * 100) / 100;

  var avoir = Math.max(0, Math.round((totalPayeAvecFNSMR - totalDuApres) * 100) / 100);

  Logger.log('supprimerActiviteDossierSheet v8.8'
    + ' | nbEligAvant:' + nbEligiblesAvant + ' | nbEligApres:' + nbEligiblesApres
    + ' | aRemiseApres:' + aRemiseApres
    + ' | totalPayé:' + totalPayeAvecFNSMR + ' | totalDûApres:' + totalDuApres + ' | avoir:' + avoir);

  // ── Passe 2 : marquer la ligne supprimée en orange ──
  var ligneSupprimeeRow = 0;
  if (sheet && sheet.getLastRow() > 1) {
    var data2 = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
    for (var j = 0; j < data2.length; j++) {
      if (String(data2[j][19] || '').trim() !== code) continue;
      var pid2 = lireActiviteId(data2[j]);
      var act2 = String(data2[j][22] || '').replace(/\n/g, ' — ').toLowerCase();
      var isCible2 = (pid2 === placesIdCible)
        || (actNomClean.toLowerCase().length >= 8 && act2.indexOf(actNomClean.toLowerCase().substring(0, 15)) >= 0);
      if (!isCible2) continue;
      sheet.getRange(j + 2, 1, 1, 41).setBackground('#ffe0b2');
      sheet.getRange(j + 2, 22).setValue('🗑 Activité supprimée').setFontColor('#e65100').setFontWeight('bold');
      sheet.getRange(j + 2, 40).setValue('Supprimée').setFontColor('#bf360c').setFontWeight('bold');
      sheet.getRange(j + 2, 30).setValue(0).setFontColor('#e65100'); // col 30 AD = 0 (ligne supprimée)
      ligneSupprimeeRow = j + 2;
      deleted++;
      break;
    }
  }

  // ── Passe 3 : supprimer des onglets activités ──
  var protectedNames = [SHEET_INSCRIPTIONS,SHEET_RECAPITULATIF,SHEET_PLACES,SHEET_CHEQUE_1,SHEET_CHEQUE_2,
    SHEET_CHEQUE_3,SHEET_AVOIRS,SHEET_AIDE_ANCV,SHEET_AIDE_ATOUT,SHEET_AIDE_PASS_J,SHEET_AIDE_PASS_S,SHEET_ESPECES];
  ss.getSheets().forEach(function(s) {
    var nom = s.getName();
    if (protectedNames.indexOf(nom) >= 0) return;
    if (nom !== placesIdCible && nom !== actId) return;
    if (s.getLastRow() < 3) return;
    var aData = s.getRange(3, 1, s.getLastRow() - 2, 1).getValues();
    for (var k = aData.length - 1; k >= 0; k--) {
      if (String(aData[k][0] || '').trim() === code) { s.deleteRow(k + 3); break; }
    }
  });

  // ── Passe 4 : décrémenter les places ──
  try {
    var placesSheet = ss.getSheetByName(SHEET_PLACES);
    if (placesSheet && placesSheet.getLastRow() > 1) {
      var pData = placesSheet.getRange(2, 1, placesSheet.getLastRow() - 1, 5).getValues();
      for (var p = 0; p < pData.length; p++) {
        var pId = String(pData[p][0]).trim();
        if (pId === placesIdCible || pId === actId) {
          var inscrits = Math.max(0, (parseInt(pData[p][3]) || 0) - 1);
          var capacity = parseInt(pData[p][2]) || 20;
          placesSheet.getRange(p + 2, 4).setValue(inscrits);
          placesSheet.getRange(p + 2, 5).setValue(capacity - inscrits);
          break;
        }
      }
    }
  } catch(ep) { Logger.log('Places update KO: ' + ep.toString()); }

  // ── Passses 5-6-7 : recalcul + avoir + email via helper central ──
  var modifResultSuppr = appliquerModificationDossier({
    ss: ss, code: code, actNomClean: actNomClean,
    typeModif: 'suppression', estRegle: true,
    totalPayeAvant: totalPayeAvecFNSMR,
    fnsmrPaye: fnsmrTotal,
    emailAdherent: emailAdherent, responsable: responsable,
    commentaireAdmin: commentaireAdmin || '',
    montantModifie: montantModifie === true || montantModifie === 'true',
    membreNomLog: membreNom || ''
  });
  var avoir = modifResultSuppr.avoir;
  if (ligneSupprimeeRow) {
    sheet.getRange(ligneSupprimeeRow, 34).setValue(avoir > 0
      ? 'Avoir suppression : ' + avoir.toFixed(2) + ' € (frais de dossier ' + FRAIS_DOSSIER_SUPPRESSION + ' € retenus)'
      : 'Suppression — aucun avoir (frais de dossier ' + FRAIS_DOSSIER_SUPPRESSION + ' € retenus)');
  }

  Logger.log('✅ supprimerActiviteDossierSheet v8.8 terminé — deleted:' + deleted + ' avoir:' + avoir + '€');
  return { deleted: deleted, avoir: avoir, nouveauTotal: modifResultSuppr.totalApres };
}


// ============================================================
// SUPPRIMER ACTIVITÉ NON RÉGLÉE — v8.8
// Après suppression : recalcul remise + email cohérent Sheet
// ============================================================
function supprimerActiviteNonRegleSheet(code, actNom, actId, membreNom, placesIdParam, nouveauTotal, commentaireAdmin, montantModifie) {
  Logger.log('supprimerActiviteNonRegleSheet v8.8 — code:' + code + ' act:' + actNom);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var deleted = 0;
  var placesIdCible = placesIdParam || getPlacesId(actId || '');
  var dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  var actNomClean = actNom.replace(/\n/g, ' — ');
  var emailAdherent = '', responsable = '';

  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      var rowCode  = String(data[i][19] || '').trim();
      // ID Activité : col 38 AL (index 37) en v8.8 41 colonnes
      // ou col 36 AJ (index 35) en ancienne structure 39 colonnes
      // Chercher dans les deux et prendre celui qui ressemble à un placesId
      var rowPid = lireActiviteId(data[i]);
      var rowNorm  = String(data[i][22] || '').replace(/\n/g, ' — ').toLowerCase();
      var actNorm  = actNomClean.toLowerCase();
      if (rowCode !== code) continue;
      var matchPid  = (rowPid === placesIdCible) || (rowPid37 === placesIdCible) || (rowPid35 === placesIdCible);
      var matchNom1 = (actNorm.length >= 8 && rowNorm.indexOf(actNorm.substring(0, 15)) >= 0);
      var matchNom2 = (actNorm.length >= 8 && actNorm.indexOf(rowNorm.substring(0, 15)) >= 0);
      var match = matchPid || matchNom1 || matchNom2;
      Logger.log('supprimerNonRegle row'+i+' pid:"'+rowPid+'" vs "'+placesIdCible+'" match:'+match);
      if (!match) continue;
      if (!emailAdherent) emailAdherent = String(data[i][15] || '');
      if (!responsable)   responsable   = String(data[i][38] || '');
      sheet.deleteRow(i + 2);
      deleted++;
      Logger.log('✅ Ligne supprimée (non réglé) : ' + rowCode + ' / ' + rowNorm);
      break;
    }
    SpreadsheetApp.flush();
  }

  // ── Recalcul remise via helper central (avant suppression onglets) ──
  // (appliquerModificationDossier sera appelé après suppression onglets/places)

  // ── Supprimer des onglets activités ──
  var protectedNames = [SHEET_INSCRIPTIONS,SHEET_RECAPITULATIF,SHEET_PLACES,SHEET_CHEQUE_1,SHEET_CHEQUE_2,
    SHEET_CHEQUE_3,SHEET_AVOIRS,SHEET_AIDE_ANCV,SHEET_AIDE_ATOUT,SHEET_AIDE_PASS_J,SHEET_AIDE_PASS_S,SHEET_ESPECES];
  ss.getSheets().forEach(function(s) {
    var nom = s.getName();
    if (protectedNames.indexOf(nom) >= 0) return;
    if (nom !== placesIdCible && nom !== actId) return;
    if (s.getLastRow() < 3) return;
    var aData = s.getRange(3, 1, s.getLastRow() - 2, 1).getValues();
    for (var k = aData.length - 1; k >= 0; k--) {
      if (String(aData[k][0] || '').trim() === code) { s.deleteRow(k + 3); break; }
    }
  });

  // ── Décrémenter les places ──
  try {
    var placesSheet = ss.getSheetByName(SHEET_PLACES);
    if (placesSheet && placesSheet.getLastRow() > 1) {
      var pData = placesSheet.getRange(2, 1, placesSheet.getLastRow() - 1, 5).getValues();
      for (var p = 0; p < pData.length; p++) {
        var pId = String(pData[p][0]).trim();
        if (pId === placesIdCible || pId === actId) {
          var inscrits = Math.max(0, (parseInt(pData[p][3]) || 0) - 1);
          var capacity = parseInt(pData[p][2]) || 20;
          placesSheet.getRange(p + 2, 4).setValue(inscrits);
          placesSheet.getRange(p + 2, 5).setValue(capacity - inscrits);
          break;
        }
      }
    }
  } catch(ep) { Logger.log('Places update KO: ' + ep.toString()); }

  // ── Recalcul remise + email via helper central ──
  var modifResultNR = appliquerModificationDossier({
    ss: ss, code: code, actNomClean: actNomClean,
    typeModif: 'suppression', estRegle: false,
    totalPayeAvant: 0,
    emailAdherent: emailAdherent, responsable: responsable,
    commentaireAdmin: commentaireAdmin || '',
    montantModifie: montantModifie === true || montantModifie === 'true',
    membreNomLog: membreNom || ''
  });

  Logger.log('✅ supprimerActiviteNonRegleSheet v8.8 terminé — deleted:' + deleted);
  return { deleted: deleted, nouveauTotal: modifResultNR.totalApres };
}


// ============================================================
// RECALCULER LA REMISE 15% — inchangé v8.2
// ============================================================
function recalculerRemiseDossier(ss, code) {
  Logger.log('recalculerRemiseDossier — code: ' + code);
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (!sheet || sheet.getLastRow() < 2) return;
  var NON_REMISABLE = ['PINGL1945/ME1945','PINGM1930/ME21','PINGJ1530',
    'PINGL18/J17','PINGME1730','PINGL1745ME1830','JAZME1015','JAZME1115','JAZME1315','JAZME1415','JAZME1515',
    'JAZME1615','JAZME1715','JAZJ1745','MNOME/S10','COUNV1830','COUNV1930','COUNV2030','COUNV2130'];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
  var communeRec = communeFromVille((data||[]).length > 0 ? ((data||[]).find(function(r){ return String(r[19]||'').trim()===code; })||{})[10] : '');
  function eligibleRemise(pid) { return estEligibleRemise(pid, communeRec); }
  var lignesActives = [], seenPids = {}, nbEligibles = 0;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][19] || '').trim() !== code) continue;
    var statut38 = String(data[i][39] || '').toLowerCase();
    var statut22 = String(data[i][21] || '').toLowerCase();
    if (statut38.indexOf('supprimée') >= 0 || statut22.indexOf('supprimée') >= 0) continue;
    if (statut38.indexOf('attente') >= 0) continue;
    var pid = String(data[i][37] || '').trim();
    var tarifActuel = Number(data[i][27] || 0);
    lignesActives.push({ idx: i, pid: pid, tarifActuel: tarifActuel });
    if (eligibleRemise(pid) && !seenPids[pid]) { seenPids[pid] = true; nbEligibles++; }
  }
  var aRemise = nbEligibles >= 3;
  Logger.log('recalculerRemiseDossier — '+lignesActives.length+' lignes actives, '+nbEligibles+' éligibles, remise: '+aRemise);
  if (lignesActives.length === 0) return;
  // ── Lecture des tarifs bruts ──
  // Stratégie : lire TOUTES les notes en une seule passe après flush
  // Si la note est absente → le tarif actuel est déjà le brut (pas encore remisé)
  // JAMAIS appliquer la remise sur un tarif déjà remisé
  SpreadsheetApp.flush();
  var tarifsBruts = {};
  var eligibles   = {};
  if (lignesActives.length > 0) {
    // Lire col 28 (tarif brut) et col 29 (éligible) en batch
    var brutVals  = sheet.getRange(2, 28, sheet.getLastRow() - 1, 2).getValues(); // [AB, AC]
    lignesActives.forEach(function(l) {
      var brutVal  = brutVals[l.idx] ? Number(brutVals[l.idx][0]) : 0;
      var eligVal  = brutVals[l.idx] ? Number(brutVals[l.idx][1] || 0) : 0;
      // Tarif brut : utiliser col 28 directement
      var brut = brutVal > 0 ? brutVal : l.tarifActuel;
      // Si col 28 = 0 (ligne ancienne sans tarif brut dédié), écrire le brut
      if (brutVal <= 0 && l.tarifActuel > 0) {
        sheet.getRange(l.idx + 2, 28).setValue(l.tarifActuel);
        brut = l.tarifActuel;
      }
      // Éligible : col 29 (Oui/Non) — fallback sur estEligible(pid)
      var elig = eligVal === 1 || (eligVal !== 1 && eligVal !== 0 && estEligible(l.pid));
      if (eligVal !== 1 && eligVal !== 0) {
        sheet.getRange(l.idx + 2, 29).setValue(elig ? 1 : 0);
      }
      tarifsBruts[l.idx] = brut;
      eligibles[l.idx]   = elig;
    });
    SpreadsheetApp.flush();
  }
  // ── GAS écrit uniquement col 28 (brut) et col 29 (éligible) ──
  // col 30 AD = formule Excel → Sheets calcule le tarif net automatiquement
  // col 32 AF = formule SUMIF → Sheets calcule le total famille automatiquement
  lignesActives.forEach(function(l) {
    var brut    = tarifsBruts[l.idx];
    var estElig = eligibles[l.idx] !== undefined ? eligibles[l.idx] : eligibleRemise(l.pid);
    var rn      = l.idx + 2;
    // col 28 AB : tarif brut (valeur fixe, ne change jamais)
    sheet.getRange(rn, 28).setValue(brut);
    // col 29 AC : éligible (écrire si absent)
    var acVal = data[l.idx][28]; if (acVal !== 1 && acVal !== 0) {
      sheet.getRange(rn, 29).setValue(estElig ? 1 : 0)
        .setBackground(estElig ? '#e8f5e9' : '#fce4ec')
        .setFontColor(estElig ? '#1b5e20' : '#b71c1c').setFontWeight('bold');
    }
    // col 30 AD : réécrire la formule si cellule vide
    if (!sheet.getRange(rn, 30).getFormula()) {
      sheet.getRange(rn, 30)
        .setFormula(formuleAD(rn))
        .setBackground('#d8f3dc').setFontColor('#1b5e20');
    }
    // col 32 AF : réécrire SUMIF si cellule vide
    if (!sheet.getRange(rn, 32).getFormula()) {
      sheet.getRange(rn, 32)
        .setValue(0) // AF calculé par calcTotalFamille() après inscription
        .setBackground('#e8f4fd').setFontColor('#1565c0').setFontWeight('bold');
    }
  });
  SpreadsheetApp.flush();
  // Calculer et écrire le total famille (gère nbElig, FNSMR, FFTT)
  var totalFamille = calcTotalFamille(ss, code);
  Logger.log('✅ recalculerRemiseDossier — remise:'+aRemise+' nbEligibles:'+nbEligibles+' totalFamille:'+totalFamille);
  return { aRemise: aRemise, nbEligibles: nbEligibles, totalFamille: totalFamille };
}

// ============================================================
// SUPPRIMER UN DOSSIER PAR CODE
// ============================================================
function supprimerDossierSheet(code) {
  invaliderCachePlaces();
  Logger.log('supprimerDossierSheet v8.8 — code: ' + code);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var deleted = 0;
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);

  // ── Passe 1 : collecter les placesIds des activités du dossier AVANT suppression ──
  // Nécessaire pour décrémenter les places disponibles dans l'onglet Places
  var placesIdsASupprimer = {}; // { placesId: count }
  if (sheet && sheet.getLastRow() > 1) {
    var dataLecture = sheet.getRange(2, 1, sheet.getLastRow() - 1, 41).getValues();
    for (var r = 0; r < dataLecture.length; r++) {
      if (String(dataLecture[r][19] || '').trim() !== code) continue;
      var st = String(dataLecture[r][39] || '').toLowerCase();
      if (st.indexOf('supprimée') >= 0 || st.indexOf('attente') >= 0) continue;
      var pid = String(dataLecture[r][37] || '').trim();
      if (pid) placesIdsASupprimer[pid] = (placesIdsASupprimer[pid] || 0) + 1;
    }
    Logger.log('Places à décrémenter : ' + JSON.stringify(placesIdsASupprimer));
  }

  // ── Passe 2 : supprimer les lignes Inscriptions ──
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 40).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][19] || '').trim() === code) { sheet.deleteRow(i + 2); deleted++; }
    }
  }

  // ── Passe 3 : supprimer du Récapitulatif ──
  var recap = ss.getSheetByName(SHEET_RECAPITULATIF);
  if (recap && recap.getLastRow() > 1) {
    var rdata = recap.getRange(2, 1, recap.getLastRow() - 1, 14).getValues();
    for (var j = rdata.length - 1; j >= 0; j--) {
      var found = rdata[j].some(function(c) { return String(c || '').trim() === code; });
      if (found) recap.deleteRow(j + 2);
    }
  }

  // ── Passe 4 : supprimer des onglets activités ──
  var protectedSheets = [SHEET_INSCRIPTIONS, SHEET_RECAPITULATIF, SHEET_PLACES,
    SHEET_CHEQUE_1, SHEET_CHEQUE_2, SHEET_CHEQUE_3, SHEET_AVOIRS, SHEET_AVOIRS_UTILISES,
    SHEET_AIDE_ANCV, SHEET_AIDE_ATOUT, SHEET_AIDE_PASS_J, SHEET_AIDE_PASS_S, SHEET_ESPECES];
  ss.getSheets().forEach(function(actSheet) {
    if (protectedSheets.indexOf(actSheet.getName()) >= 0) return;
    if (actSheet.getLastRow() < 3) return;
    var aData = actSheet.getRange(3, 1, actSheet.getLastRow() - 2, 1).getValues();
    for (var k = aData.length - 1; k >= 0; k--) {
      if (String(aData[k][0] || '').trim() === code) actSheet.deleteRow(k + 3);
    }
  });

  // ── Passe 5 : supprimer des feuilles chèques ──
  [SHEET_CHEQUE_1, SHEET_CHEQUE_2, SHEET_CHEQUE_3].forEach(function(nom) {
    var ch = ss.getSheetByName(nom);
    if (!ch || ch.getLastRow() < 2) return;
    var cData = ch.getRange(2, 1, ch.getLastRow() - 1, 1).getValues();
    for (var l = cData.length - 1; l >= 0; l--) {
      if (String(cData[l][0] || '').trim() === code) ch.deleteRow(l + 2);
    }
  });

  // ── Passe 6 : supprimer des Avoirs ──
  var avoirSh = ss.getSheetByName(SHEET_AVOIRS);
  if (avoirSh && avoirSh.getLastRow() > 1) {
    var avData = avoirSh.getRange(2, 1, avoirSh.getLastRow() - 1, 1).getValues();
    for (var av = avData.length - 1; av >= 0; av--) {
      if (String(avData[av][0] || '').trim() === code) avoirSh.deleteRow(av + 2);
    }
  }

  // ── Passe 7b : supprimer des onglets Aides (ANCV, Atout, Pass Jeunes, Pass Sport) ──
  [SHEET_AIDE_ANCV, SHEET_AIDE_ATOUT, SHEET_AIDE_PASS_J, SHEET_AIDE_PASS_S].forEach(function(nom) {
    var sh = ss.getSheetByName(nom);
    if (!sh || sh.getLastRow() < 2) return;
    var lr = sh.getLastRow();
    var d  = sh.getRange(2, 1, lr - 1, 1).getValues();
    for (var i = d.length - 1; i >= 0; i--) {
      if (String(d[i][0] || '').trim() === code) {
        sh.deleteRow(i + 2);
        Logger.log('✅ Ligne supprimée dans ' + nom + ' — code: ' + code);
      }
    }
  });

  // ── Passe 7c : supprimer des Espèces ──
  var especesSh = ss.getSheetByName(SHEET_ESPECES);
  if (especesSh && especesSh.getLastRow() > 1) {
    var espData = especesSh.getRange(2, 1, especesSh.getLastRow() - 1, 1).getValues();
    for (var es = espData.length - 1; es >= 0; es--) {
      if (String(espData[es][0] || '').trim() === code) {
        especesSh.deleteRow(es + 2);
        Logger.log('✅ Ligne supprimée dans Espèces — code: ' + code);
      }
    }
  }

  // ── Passe 7d : supprimer de HelloAsso ──
  var haSh = ss.getSheetByName(SHEET_HELLOASSO);
  if (haSh && haSh.getLastRow() > 1) {
    var haData = haSh.getRange(2, 1, haSh.getLastRow() - 1, 1).getValues();
    for (var ha = haData.length - 1; ha >= 0; ha--) {
      if (String(haData[ha][0] || '').trim() === code) {
        haSh.deleteRow(ha + 2);
        Logger.log('✅ Ligne supprimée dans HelloAsso — code: ' + code);
      }
    }
  }

  // ── Passe 7e : supprimer des Licences FFTT ──
  var ffttSh = ss.getSheetByName('Licences FFTT');
  if (ffttSh && ffttSh.getLastRow() > 1) {
    var ffttData = ffttSh.getRange(2, 1, ffttSh.getLastRow() - 1, 1).getValues();
    for (var ff = ffttData.length - 1; ff >= 0; ff--) {
      if (String(ffttData[ff][0] || '').trim() === code) {
        ffttSh.deleteRow(ff + 2);
        Logger.log('✅ Ligne supprimée dans FFTT — code: ' + code);
      }
    }
  }

  // ── Passe 7f : recalculer les totaux des onglets Chèques ──
  try {
    [SHEET_CHEQUE_1, SHEET_CHEQUE_2, SHEET_CHEQUE_3].forEach(function(nom) {
      majTotalCheque(ss, nom);
    });
    Logger.log('✅ Totaux chèques recalculés après suppression dossier ' + code);
  } catch(eTot) {
    Logger.log('❌ Recalcul totaux chèques KO : ' + eTot.toString());
  }

  // ── Passe 7 : décrémenter les places disponibles ──
  try {
    var placesSheet = ss.getSheetByName(SHEET_PLACES);
    if (placesSheet && placesSheet.getLastRow() > 1 && Object.keys(placesIdsASupprimer).length > 0) {
      var pData = placesSheet.getRange(2, 1, placesSheet.getLastRow() - 1, 5).getValues();
      for (var p = 0; p < pData.length; p++) {
        var pId = String(pData[p][0] || '').trim();
        if (!pId || !placesIdsASupprimer[pId]) continue;
        var nbSuppr   = placesIdsASupprimer[pId];
        var inscrits  = Math.max(0, (parseInt(pData[p][3]) || 0) - nbSuppr);
        var capacity  = parseInt(pData[p][2]) || 20;
        var dispos    = Math.max(0, capacity - inscrits);
        placesSheet.getRange(p + 2, 4).setValue(inscrits);
        placesSheet.getRange(p + 2, 5).setValue(dispos);
        Logger.log('✅ Places MAJ — ' + pId + ' : inscrits=' + inscrits + ' dispos=' + dispos);
      }
    }
  } catch(ep) {
    Logger.log('❌ Mise à jour Places KO : ' + ep.toString());
  }

  Logger.log('✅ Dossier ' + code + ' supprimé — ' + deleted + ' ligne(s) Inscriptions');
  return { deleted: deleted };
}

// ============================================================
// EMAIL ADHÉRENT — v8.3 : total recalculé depuis emailRows
// ============================================================

// ══════════════════════════════════════════════════════════════
// RÈGLEMENT INTÉRIEUR — Sauvegarde PDF sur Drive
// ══════════════════════════════════════════════════════════════
function sauvegarderReglementDrive(code, nom, prenom, email, dateSignature, signatureBase64, dossierOverride) {
  var payload = { dossierOverride: dossierOverride || '' };
  var nomDossierRI = String(payload && payload.dossierOverride ? payload.dossierOverride : '3-Règlements intérieurs');
  var dossier = creerDossierSecurise(nomDossierRI);
  var dateStr = dateSignature || Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  var nomFichier = 'RI_' + code + '_' + (nom || 'INCONNU').replace(/[^A-Z0-9]/gi, '_') + '_' + Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyyMMdd') + '.pdf';

  // Créer un document Google Doc temporaire
  var doc = DocumentApp.create('RI_temp_' + code);
  var body = doc.getBody();

  body.setMarginTop(18).setMarginBottom(18).setMarginLeft(36).setMarginRight(36);

  // En-tête
  var titre = body.appendParagraph('RÈGLEMENT INTÉRIEUR');
  titre.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titre.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titre.editAsText().setFontSize(13).setBold(true).setForegroundColor('#1a2e22');

  body.appendParagraph("Foyer Rural d\'Isneauville")
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setFontSize(10).setForegroundColor('#2d6a4f');

  body.appendParagraph('');

  // Infos adhérent
  var infoStyle = {};
  body.appendParagraph('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━').editAsText().setFontSize(10).setForegroundColor('#cccccc');
  body.appendParagraph('Adhérent : ' + prenom + ' ' + nom).editAsText().setFontSize(9).setBold(true);
  body.appendParagraph('Email : ' + email).editAsText().setFontSize(9);
  body.appendParagraph('Dossier : ' + code).editAsText().setFontSize(9);
  body.appendParagraph('Date de signature : ' + dateStr).editAsText().setFontSize(9);
  body.appendParagraph('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━').editAsText().setFontSize(10).setForegroundColor('#cccccc');
  body.appendParagraph('');

  // Contenu du règlement
  var sections = [
    ["STATUTS", "Par l'apposition de sa signature, l'adherent reconnait avoir pris connaissance des statuts et declare y adherer sans reserve. L'adhesion donne le droit d'expression et de vote lors des assemblees generales (adherents depuis plus de 3 mois et ages de plus de 16 ans)."],
    ["ACTIVITES", "Les activites et horaires ne sont pas contractuels. Un certificat medical est exige pour la pratique du Tennis de table en competition (FFTT) et pour les adherents avec reponse OUI au questionnaire de sante."],
    ["INSCRIPTIONS", "L'inscription n'est validee qu'apres versement de la cotisation. Tout remboursement entraine un forfait de 30 EUR de frais de secretariat, principalement sous forme d'avoir. Un essai est conseille avant engagement."],
    ["COTISATION", "3 activites ou plus dans la meme famille : reduction de 15% (hors forfait Country, Marche Nordique, droit adhesion, assurance). En cas de crise sanitaire, la cotisation ne donne pas droit a remboursement automatique."],
    ["ACCOMPAGNEMENT DES MINEURS", "Les mineurs doivent etre accompagnes d'un adulte jusqu'a la porte et recuperes a la fin du cours. Les parents ne sont pas autorises a assister aux cours."],
    ["ASSIDUITE", "Arriver a l'heure (5 min avant idealement). En cas d'absence, prevenir l'animateur ou frisneauville@orange.fr. L'assiduite est primordiale pour la progression du groupe."],
    ["TENUE", "Tenue adaptee obligatoire. Baskets propres exigees dans le complexe sportif. Serviette de protection obligatoire sur tapis. Pour le country : chaussures propres (pas de chaussures exterieures)."],
    ["COMPORTEMENT", "Comportement correct et respectueux envers tous (eleves, animateurs, locaux). Tout manquement peut entrainer une exclusion temporaire ou definitive par le CA."],
    ["SPECTACLES", "Spectacles en fin d'annee. Presence obligatoire aux repetitions generales pour y participer."],
    ["DROIT A L'IMAGE", "L'adherent autorise le FRI a publier des photographies pour information ou publicite de l'activite uniquement. Supports : site FRI, mailing, reseaux sociaux, YouTube, journal municipal."],
    ["RGPD", "Le FRI se conforme au RGPD. Les donnees personnelles ne sont ni partagees ni vendues. Droit de retrait garanti."]
  ];

  sections.forEach(function(s) {
    var p = body.appendParagraph(s[0]);
    p.editAsText().setFontSize(9).setBold(true).setForegroundColor('#1a2e22');
    var p2 = body.appendParagraph(s[1]);
    p2.editAsText().setFontSize(8).setForegroundColor('#333333');
    body.appendParagraph('');
  });

  // Signature
  body.appendParagraph('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━').editAsText().setFontSize(10).setForegroundColor('#cccccc');
  body.appendParagraph('ACCEPTATION ET SIGNATURE ÉLECTRONIQUE').editAsText().setFontSize(10).setBold(true).setForegroundColor('#1a2e22');
  body.appendParagraph('');
  body.appendParagraph("Je soussigné(e) " + prenom + " " + nom + " déclare avoir lu et accepté le règlement intérieur du Foyer Rural d\'Isneauville.").editAsText().setFontSize(9);
  body.appendParagraph('Signé électroniquement le : ' + dateStr).editAsText().setFontSize(9).setBold(true);
  body.appendParagraph('');

  // Image de signature si fournie
  if (signatureBase64 && signatureBase64.indexOf('data:image/png;base64,') === 0) {
    try {
      var base64Data = signatureBase64.replace('data:image/png;base64,', '');
      var imgBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', 'signature.png');
      body.appendImage(imgBlob).setWidth(160).setHeight(50);
    } catch(imgErr) {
      Logger.log('Image signature KO : ' + imgErr);
      body.appendParagraph('[Signature électronique enregistrée]').editAsText().setItalic(true).setForegroundColor('#666666');
    }
  }

  doc.saveAndClose();

  // Convertir en PDF
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(nomFichier);

  // Sauvegarder dans le dossier Drive
  var fichier = dossier.createFile(pdfBlob);
  securiserFichier(fichier);

  // Supprimer le Doc temporaire
  docFile.setTrashed(true);

  Logger.log('✅ 3-Règlements intérieurs sauvegardé : ' + nomFichier);
  return { fichierNom: nomFichier, fichierId: fichier.getId() };
}


// ══════════════════════════════════════════════════════════════
// ADMIN — Basculer liste d'attente → inscrit
// ══════════════════════════════════════════════════════════════
// ============================================================
// EMAIL ADHÉRENT — Bascule liste d'attente
// Récapitulatif COMPLET du dossier + mise en évidence de la nouvelle activité à régler
// ============================================================
function envoyerEmailBasculeListe(email, rows, actNomBasculee, tarifBascule, isPaidDossier, actIdBasculee, commentaireAdmin) {
  commentaireAdmin = String(commentaireAdmin || '').trim();
  if (!email || email.indexOf('@') < 0) return;
  var r0 = rows[0];
  var prenom = r0.responsable_prenom || '';
  var nom    = r0.responsable_nom    || '';
  var code   = r0.code_dossier || '';

  // ── Calcul financier global ──
  var f = _calcFinancier(rows);

  // ── Identifier la nouvelle activité basculée ──
  // On prend la PREMIÈRE ligne marquée basculée (évite d'écraser avec un autre membre)
  var rowBasculee = null;
  rows.forEach(function(r) {
    if (!rowBasculee
        && (r._estBasculee || lireActiviteId_fromRow(r) === actIdBasculee)
        && r.statut_inscription.toLowerCase().indexOf('attente') < 0) {
      rowBasculee = r;
    }
  });
  var tarifARegler = tarifBascule > 0 ? tarifBascule : (rowBasculee ? (parseFloat(rowBasculee.tarif_brut)||parseFloat(rowBasculee.tarif)||0) : 0);

  // ── Séparer activités : déjà payées vs basculée (à régler) vs encore en attente ──
  var parMembre = {};
  rows.forEach(function(r) {
    var k = (r.membre_prenom||'') + ' ' + (r.membre_nom||'');
    if (!parMembre[k]) parMembre[k] = [];
    parMembre[k].push(r);
  });

  var lignesHtml = '';
  Object.keys(parMembre).forEach(function(membre) {
    lignesHtml += '<tr><td colspan="5" style="background:#2d6a4f;color:white;padding:7px 10px;font-weight:bold;font-size:13px">👤 ' + membre + '</td></tr>';
    parMembre[membre].forEach(function(r) {
      var statutR = String(r.statut_inscription||'').toLowerCase();
      var isAttente  = statutR.indexOf('attente') >= 0;
      var isBasculee = r._estBasculee || lireActiviteId_fromRow(r) === actIdBasculee;
      var tarif = parseFloat(r.tarif_brut)||parseFloat(r.tarif)||0;
      var actNom = (r.activite||'').replace(/\n/g,' — ');
      var joHe   = (r.jour||'') + (r.heure ? ' · ' + r.heure : '');

      var bg, statutBadge, tarifCell;
      if (isBasculee && !isAttente) {
        // Nouvelle activité basculée — à régler
        bg = '#fff8e1';
        statutBadge = '<span style="background:#e65100;color:white;font-size:10px;font-weight:bold;padding:2px 7px;border-radius:10px;margin-left:6px">A REGLER</span>';
        tarifCell = '<strong style="color:#e65100">' + tarifARegler.toFixed(2) + ' €</strong>';
      } else if (isAttente) {
        // Encore en liste d'attente
        bg = '#f5f5f5';
        statutBadge = '<span style="background:#9e9e9e;color:white;font-size:10px;padding:2px 7px;border-radius:10px;margin-left:6px">LISTE ATTENTE</span>';
        tarifCell = '<span style="color:#9e9e9e">—</span>';
      } else {
        // Activité déjà payée
        bg = '#f0f7f3';
        statutBadge = '<span style="background:#2d6a4f;color:white;font-size:10px;padding:2px 7px;border-radius:10px;margin-left:6px">REGLE</span>';
        tarifCell = '<span style="color:#2d6a4f">' + tarif.toFixed(2) + ' €</span>';
      }
      lignesHtml += '<tr style="background:' + bg + '">'
        + '<td style="padding:7px 10px;border-bottom:1px solid #eee">' + actNom + statutBadge + '</td>'
        + '<td style="padding:7px 10px;border-bottom:1px solid #eee;color:#555">' + joHe + '</td>'
        + '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right">' + tarifCell + '</td>'
        + '</tr>';
    });
    // FFTT
    if (f.ffttMembres[membre]) {
      lignesHtml += '<tr style="background:#fff3e0"><td style="padding:6px 10px;font-size:12px;color:#e65100" colspan="2">🏓 Licence FFTT — ' + membre + '</td>'
        + '<td style="padding:6px 10px;text-align:right;font-weight:bold;color:#e65100">' + f.ffttMembres[membre].toFixed(2) + ' €</td></tr>';
    }
  });

  // ── Bloc "Ce qui reste à régler" — tenant compte des paiements déjà effectués ──
  // Calcul : totalNonPaye = somme des lignes NON payées (activite basculee + autres non réglées)
  //          totalDejaRegle = somme des lignes déjà payées
  var totalDejaRegle = 0;
  var totalNonPaye   = 0;
  var lignesBilanHtml = '';

  rows.forEach(function(r) {
    var statutPaie = String(r.statut_paiement || '');
    var statutInsc = String(r.statut_inscription || '').toLowerCase();
    var isAttente  = statutInsc.indexOf('attente') >= 0;
    if (isAttente) return; // lignes attente ignorées du bilan financier
    var tarif = parseFloat(r.tarif_brut) || parseFloat(r.tarif) || 0;
    var actNomLigne = (r.activite || '').replace(/\n/g,' — ');
    var isPaye = statutPaie.indexOf('✅') >= 0 || statutPaie.toLowerCase().indexOf('pay') >= 0;
    if (isPaye) {
      totalDejaRegle += tarif;
      lignesBilanHtml += '<tr style="background:#f0f7f3">'
        + '<td style="padding:5px 8px;font-size:12px">' + actNomLigne + '</td>'
        + '<td style="text-align:right;color:#2d6a4f;padding:5px 8px;font-size:12px">✅ ' + tarif.toFixed(2) + ' €</td></tr>';
    } else {
      totalNonPaye += tarif;
      lignesBilanHtml += '<tr style="background:#fff8e1">'
        + '<td style="padding:5px 8px;font-size:12px">' + actNomLigne + '</td>'
        + '<td style="text-align:right;color:#e65100;padding:5px 8px;font-size:12px">⏳ ' + tarif.toFixed(2) + ' €</td></tr>';
    }
  });
  // Ajouter ligne FFTT si présente
  if (f.totalFftt > 0) {
    var ffttPaye = isPaidDossier; // heuristique : si dossier payé, FFTT aussi
    if (ffttPaye) {
      totalDejaRegle += f.totalFftt;
    } else {
      totalNonPaye += f.totalFftt;
      lignesBilanHtml += '<tr style="background:#fff8e1">'
        + '<td style="padding:5px 8px;font-size:12px">🏓 Licence(s) FFTT</td>'
        + '<td style="text-align:right;color:#e65100;padding:5px 8px;font-size:12px">⏳ ' + f.totalFftt.toFixed(2) + ' €</td></tr>';
    }
  }

  var soldeNet = Math.max(0, totalNonPaye - (f.totalDeductions || 0));
  var totalARegler = soldeNet;

  var blocRegler = soldeNet > 0
    ? '<div style="background:#fff3e0;border:2px solid #e65100;border-radius:10px;padding:14px 18px;margin:16px 0">'
      + '<div style="font-weight:bold;font-size:15px;color:#e65100;margin-bottom:8px">💳 Solde à régler pour votre dossier</div>'
      + '<table style="width:100%;border-collapse:collapse;font-size:13px">'
      + lignesBilanHtml
      + (totalDejaRegle > 0
          ? '<tr style="border-top:1px solid #ddd;background:#f0f7f3"><td style="padding:5px 8px;font-size:12px;color:#2d6a4f">Déjà réglé</td>'
            + '<td style="text-align:right;color:#2d6a4f;padding:5px 8px;font-size:12px">− ' + totalDejaRegle.toFixed(2) + ' €</td></tr>'
          : '')
      + ((f.totalDeductions || 0) > 0
          ? '<tr style="background:#f0f7f3"><td style="padding:5px 8px;font-size:12px;color:#2d6a4f">Déductions / réductions</td>'
            + '<td style="text-align:right;color:#2d6a4f;padding:5px 8px;font-size:12px">− ' + (f.totalDeductions||0).toFixed(2) + ' €</td></tr>'
          : '')
      + '<tr style="border-top:2px solid #e65100"><td style="padding:6px 8px;font-weight:bold;font-size:14px;color:#e65100">SOLDE À RÉGLER</td>'
      + '<td style="text-align:right;font-weight:bold;font-size:16px;color:#e65100;padding:6px 8px">' + soldeNet.toFixed(2) + ' €</td></tr>'
      + '</table>'
      + '<div style="margin-top:12px;font-size:12px;color:#777;line-height:1.7">'
      + '<strong>Comment régler ?</strong><br>'
      + '💳 En ligne via HelloAsso · 📅 Aux permanences (mardi 16h30–18h30) · ✉️ Chèque à l\'ordre du FRI'
      + '</div>'
      + helloassoBoutonHtml(null, soldeNet > 0 ? soldeNet.toFixed(2) + ' €' : null)
      + '</div>'
    : '<div style="background:#d8f3dc;border:2px solid #52b788;border-radius:10px;padding:14px 18px;margin:16px 0">'
      + '<div style="font-weight:bold;font-size:15px;color:#2d6a4f;margin-bottom:4px">✅ Dossier entièrement réglé</div>'
      + '<div style="font-size:13px;color:#555">L\'activité <strong>' + actNomBasculee.replace(/\n/g,' — ') + '</strong> est maintenant confirmée. Votre dossier est soldé.</div>'
      + '</div>';

  // ── QS Santé ──
  var santeHtml = '', vuQs = {};
  rows.forEach(function(r) {
    var k = (r.membre_prenom||'') + ' ' + (r.membre_nom||'');
    if (!vuQs[k]) {
      vuQs[k] = true;
      var qsLabel2, qsColor2;
      if (r.qs_sante === 'Attestation OK') {
        qsLabel2 = '✅ QS signé en ligne'; qsColor2 = '#2d6a4f';
      } else if (r.qs_sante === 'Certificat requis') {
        qsLabel2 = '⚠️ Certificat médical requis'; qsColor2 = '#c0392b';
      } else if (r.qs_sante && r.qs_sante.indexOf('non sportive') >= 0) {
        qsLabel2 = '✅ Pas de QS requis (activité non sportive)'; qsColor2 = '#2d6a4f';
      } else if (r.qs_sante === 'Non rempli' || !r.qs_sante) {
        qsLabel2 = '⚠️ QS non complété'; qsColor2 = '#e65100';
      } else { qsLabel2 = r.qs_sante; qsColor2 = '#555'; }
      if (qsLabel2) santeHtml += '<li style="margin-bottom:4px"><strong>' + k + '</strong> — <span style="color:' + qsColor2 + '">' + qsLabel2 + '</span></li>';
    }
  });

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f0f4f2;margin:0;padding:20px">'
    + '<div style="max-width:640px;margin:0 auto;background:white;border-radius:12px;overflow:hidden">'
    // En-tête
    + '<div style="background:#1a2e22;padding:20px 24px;text-align:center">'
    + '<img src="cid:logo_fri" alt="FRI" style="width:66px;height:auto;margin-bottom:10px;border-radius:8px;display:block;margin-left:auto;margin-right:auto">'
    + '<h1 style="color:white;margin:0;font-size:19px">Place disponible — Inscription confirmée !</h1>'
    + '<p style="color:#52b788;margin:5px 0 0;font-size:13px">' + NOM_ASSO + ' — Saison 2026/2027</p></div>'
    + '<div style="background:#fff8e1;padding:10px 20px;text-align:center;font-weight:bold;color:#856404;border-bottom:2px solid #e8c84a">'
    + '🎉 Une place s\'est libérée pour <strong>' + actNomBasculee.replace(/\n/g,' — ') + '</strong></div>'
    + '<div style="padding:20px 24px">'
    + '<p style="color:#333">Bonjour <strong>' + prenom + ' ' + nom + '</strong>,</p>'
    + '<p style="color:#555;line-height:1.7">Suite à un désistement, une place s\'est libérée dans l\'activité <strong>' + actNomBasculee.replace(/\n/g,' — ') + '</strong>. Votre inscription est maintenant <strong>confirmée</strong>.</p>'
    // Numéro dossier
    + '<div style="background:#d8f3dc;border:2px solid #52b788;border-radius:10px;padding:10px 18px;text-align:center;margin:14px 0">'
    + '<div style="font-size:11px;font-weight:700;color:#2d6a4f;text-transform:uppercase;letter-spacing:1px;margin-bottom:3px">Numéro de dossier</div>'
    + '<div style="font-family:monospace;font-size:26px;font-weight:900;color:#1a2e22;letter-spacing:5px">' + code + '</div>'
    + '</div>'
    // Tableau récapitulatif global
    + '<h3 style="color:#1a2e22;font-size:14px;margin:16px 0 6px;padding-bottom:3px;border-bottom:2px solid #d8f3dc">Récapitulatif complet de vos inscriptions</h3>'
    + '<table style="width:100%;border-collapse:collapse;font-size:13px">'
    + '<thead><tr style="background:#1a2e22;color:white">'
    + '<th style="padding:7px 10px;text-align:left">Activité</th>'
    + '<th style="padding:7px 10px;text-align:left">Jour / Heure</th>'
    + '<th style="padding:7px 10px;text-align:right">Tarif</th>'
    + '</tr></thead><tbody>' + lignesHtml + '</tbody></table>'
    // FNSMR déjà payé
    + '<div style="background:#f5f5f5;border-radius:6px;padding:8px 12px;margin:8px 0;font-size:12px;color:#555">'
    + 'Adhésion FNSMR (' + f.nbMembres + ' pers. × 15 €) — déjà réglée lors de l\'inscription initiale</div>'
    // Bloc à régler
    + blocRegler
    + blocCommentaireAdminHtml(commentaireAdmin)
    // Documents
    + (santeHtml ? '<h3 style="color:#1a2e22;font-size:14px;margin:16px 0 6px">Documents</h3><ul style="color:#555;line-height:1.9;margin-top:4px">' + santeHtml + '</ul>' : '')
    + ffttBloc
    + '<div style="background:#e8f4fd;border-left:4px solid #2980b9;border-radius:6px;padding:11px;margin-top:14px;font-size:13px;line-height:1.7">'
    + '<strong>Permanences</strong> : tous les <strong>mardis de 16h30 à 18h30</strong> <em>(période scolaire)</em><br>'
    + 'Salle des fêtes — Place A. Cramilly, 76230 Isneauville</div>'
    + '<p style="margin-top:12px;font-size:12px;color:#555">Contact : <a href="mailto:frisneauville@orange.fr" style="color:#2d6a4f">frisneauville@orange.fr</a> — <a href="tel:0235590101" style="color:#2d6a4f">02.35.59.01.01</a></p>'
    + '</div>'
    + (function() {
        var hasTT = rows.some(function(r){ var aid = String(r.activite_id||''); return aid.indexOf('PING') >= 0 || aid.indexOf('tt-') === 0; });
        var hasCult = rows.some(function(r){ var a=String(r.activite_id||''); return ['couture-','peinture-','theatre-','guitare-'].some(function(p){ return a.indexOf(p)===0; }); });
        var out = '';
        if (hasTT) {
          out += '<div style="background:#e8eaf6;border:2px solid #1a237e;border-radius:8px;padding:14px 18px;margin:12px 0;">'
            + '<strong style="color:#1a237e;">🏓 Tennis de Table — Document sanitaire obligatoire</strong>'
            + '<p style="font-size:13px;color:#333;margin:8px 0;">Vous devez fournir un des documents suivants :</p>'
            + '<ul style="font-size:13px;color:#333;margin:0 0 10px;padding-left:18px;line-height:1.8;">'
            + '<li>QS Santé FNSMR — à signer en ligne</li>'
            + '<li>Parcours Personnel de Santé (PPS FFTT)</li>'
            + '<li>Certificat médical de non contre-indication (- 3 ans)</li>'
            + '</ul>'
            + '<a href="https://script.google.com/macros/s/AKfycbx-Y6io0i42BbjalFcG45--tq5-k9I-AU5kcQ7QHhI1zgh-X2baR3dU7TKWD1X8KQwI/exec" style="display:inline-block;background:#1a237e;color:white;border-radius:6px;padding:8px 16px;font-size:13px;font-weight:700;text-decoration:none;">📋 Remplir le PPS FFTT en ligne</a>'
            + '</div>';
        }
        if (hasCult) {
          out += '<div style="background:#e8f5e9;border:2px solid #2d6a4f;border-radius:8px;padding:12px 18px;margin:10px 0;">'
            + '<strong style="color:#2d6a4f;">✅ Activité culturelle — Pas de QS requis</strong>'
            + '<p style="font-size:13px;color:#2d6a4f;margin:6px 0 0;">Les activités culturelles (peinture, guitare, théâtre, couture) ne nécessitent pas de questionnaire de santé sportive.</p>'
            + '</div>';
        }
        return out;
      })()
    + '<div style="background:#f5f5f5;padding:10px;text-align:center;font-size:11px;color:#aaa">' + NOM_ASSO + ' — www.frisneauville.fr</div>'
    + '</div></body></html>';

  var sujet = 'Place disponible - Inscription confirmee FRI - N. ' + code + ' - ' + prenom + ' ' + nom;
  var logoB = getLogoBlob();
  var opts = { htmlBody: html, charset: 'UTF-8', name: NOM_ASSO, replyTo: EMAIL_ADMIN, charset: 'UTF-8', inlineImages: logoB ? {logo_fri: logoB} : {} };
  var bodyTxt = 'Bonjour ' + prenom + ' ' + nom + ','
    + '\n\nUne place s\'est liberee pour : ' + actNomBasculee.replace(/\n/g,' - ')
    + '\nDossier : ' + code
    + (totalARegler > 0
        ? '\nMontant a regler : ' + totalARegler.toFixed(2) + ' EUR'
        : '\nAucun supplement - adhesion deja reglee.')
    + (commentaireAdmin ? '\n\nNote de l\'equipe FRI : ' + commentaireAdmin : '')
    + '\n\nContact : frisneauville@orange.fr | 02.35.59.01.01';
  envoyerEmail(email, sujet, bodyTxt, opts);
  Logger.log('Email bascule adherent envoye : ' + email);
}

// ── Helper pour lire l'activite_id depuis un row objet ──
function lireActiviteId_fromRow(r) {
  return String(r.activite_id || r._actId || '').trim();
}

// ============================================================
// EMAIL ADMIN — Bascule liste d'attente
// ============================================================
function envoyerEmailAdminBascule(emailAdmin, rows, actNomBasculee, tarifBascule, isPaidDossier, actIdBasculee) {
  if (!emailAdmin || emailAdmin.indexOf('@') < 0) return;
  var r0 = rows[0];
  var responsable = (r0.responsable_prenom||'') + ' ' + (r0.responsable_nom||'');
  var code = r0.code_dossier || '';
  var f = _calcFinancier(rows);

  var lignesAdmin = rows.map(function(r, idx) {
    var statutR   = String(r.statut_inscription||'').toLowerCase();
    var isAttente = statutR.indexOf('attente') >= 0;
    var isBasc    = r._estBasculee || lireActiviteId_fromRow(r) === actIdBasculee;
    var tarif     = parseFloat(r.tarif_brut)||parseFloat(r.tarif)||0;
    var bg = isBasc && !isAttente ? '#fff8e1' : (isAttente ? '#f5f5f5' : (idx%2===0?'#f0f7f3':'#ffffff'));
    var statutLabel = isBasc && !isAttente ? '<strong style="color:#e65100">A REGLER</strong>' : (isAttente ? '<span style="color:#9e9e9e">Attente</span>' : '<span style="color:#2d6a4f">Regle</span>');
    var tarifLabel  = isAttente ? '—' : tarif.toFixed(2) + ' €';
    return '<tr style="background:' + bg + '">'
      + '<td style="padding:6px 10px">' + (r.membre_prenom||'') + ' ' + (r.membre_nom||'') + '</td>'
      + '<td style="padding:6px 10px">' + (r.activite||'').replace(/\n/g,' — ') + '</td>'
      + '<td style="padding:6px 10px">' + (r.jour||'') + '</td>'
      + '<td style="padding:6px 10px;text-align:right">' + tarifLabel + '</td>'
      + '<td style="padding:6px 10px;text-align:center">' + statutLabel + '</td>'
      + '</tr>';
  }).join('');

  var tarifARegler = isPaidDossier ? 0 : (tarifBascule > 0 ? tarifBascule : 0);

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f0f4f2;margin:0;padding:16px">'
    + '<div style="max-width:620px;margin:0 auto;background:white;border-radius:12px;overflow:hidden">'
    + '<div style="background:#1a2e22;padding:16px 22px;text-align:center">'
    + '<img src="cid:logo_fri" alt="FRI" style="width:52px;height:auto;border-radius:6px;display:block;margin:0 auto 8px">'
    + '<h1 style="color:#52b788;margin:0;font-size:16px">[ADMIN] Bascule liste d\'attente</h1>'
    + '<p style="color:rgba(255,255,255,.5);margin:3px 0 0;font-size:12px">Dossier ' + code + '</p></div>'
    + '<div style="background:#fff8e1;padding:9px 18px;font-weight:bold;color:#856404">'
    + 'Activite basculee : ' + actNomBasculee.replace(/\n/g,' — ') + (isPaidDossier ? ' — Dossier deja regle' : ' — A REGLER : ' + tarifARegler.toFixed(2) + ' EUR') + '</div>'
    + '<div style="padding:16px 20px">'
    + '<table style="width:100%;font-size:13px;border-collapse:collapse;margin-bottom:12px">'
    + '<tr><td style="padding:4px 0;color:#888;width:38%">N. dossier</td><td style="font-weight:bold;font-family:monospace;font-size:14px;color:#1565c0;letter-spacing:3px">' + code + '</td></tr>'
    + '<tr><td style="padding:4px 0;color:#888">Responsable</td><td style="font-weight:bold">' + responsable + '</td></tr>'
    + '<tr><td style="padding:4px 0;color:#888">Email</td><td>' + (r0.email1||'—') + '</td></tr>'
    + '<tr><td style="padding:4px 0;color:#888">Tel.</td><td>' + (r0.tel1||'—') + '</td></tr>'
    + '</table>'
    + '<table style="width:100%;font-size:13px;border-collapse:collapse">'
    + '<thead><tr style="background:#d8f3dc"><th style="padding:7px 10px;text-align:left">Membre</th><th style="text-align:left;padding:7px 10px">Activite</th><th style="text-align:left;padding:7px 10px">Jour</th><th style="text-align:right;padding:7px 10px">Tarif</th><th style="text-align:center;padding:7px 10px">Statut</th></tr></thead>'
    + '<tbody>' + lignesAdmin + '</tbody></table>'
    + '<div style="background:#f5f5f5;border-radius:6px;padding:8px 12px;margin:8px 0;font-size:12px;color:#555">'
    + 'FNSMR (' + f.nbMembres + ' pers. x 15 €) = ' + f.totalFnsmr.toFixed(2) + ' € — deja regle</div>'
    + (tarifARegler > 0
      ? '<div style="background:#fff3e0;border:1px solid #e65100;border-radius:8px;padding:10px 14px;margin-top:10px;font-weight:bold;color:#e65100;font-size:14px">Montant a regler : ' + tarifARegler.toFixed(2) + ' EUR</div>'
      : '<div style="background:#d8f3dc;border:1px solid #52b788;border-radius:8px;padding:10px 14px;margin-top:10px;font-weight:bold;color:#2d6a4f;font-size:13px">Aucun supplement — dossier deja regle</div>')
    + '</div></div></body></html>';

  var sujet = '[FRI] Bascule liste attente N. ' + code + ' - ' + responsable + (tarifARegler > 0 ? ' - A REGLER ' + tarifARegler.toFixed(2) + ' EUR' : ' - Deja regle');
  var logoBA = getLogoBlob();
  envoyerEmail(emailAdmin, sujet, 'Bascule liste attente : ' + code + ' - ' + actNomBasculee.replace(/\n/g,' - '), { htmlBody: html, charset: 'UTF-8', name: 'Site FRI Inscriptions', charset: 'UTF-8', inlineImages: logoBA ? {logo_fri: logoBA} : {} });
  Logger.log('Email admin bascule envoye : ' + emailAdmin);
}

function basculerListeAttenteGAS(code, actId, actNom, prenomMembre, nomMembre, tarifBrut, commentaireAdmin, montantModifie) {
  commentaireAdmin = String(commentaireAdmin || '').trim();
  montantModifie = montantModifie === true || montantModifie === 'true';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (!sheet) return {status:'error', message:'Onglet Inscriptions introuvable'};

  // Lire 42 colonnes pour avoir toutes les données
  var data = sheet.getRange(2, 1, Math.max(sheet.getLastRow()-1,1), 42).getValues();
  var rowsUpdated = 0;
  var emailRows = [];
  var isPaidDossier = false;

  // Logger toutes les lignes du dossier pour diagnostic
  Logger.log('=== BASCULE '+code+' actId=['+actId+'] nom=['+nomMembre+'] prenom=['+prenomMembre+'] ===');
  for (var di=0; di<data.length; di++) {
    if (String(data[di][19]||'').trim() !== code) continue;
    Logger.log('L'+(di+2)+' AJ35=['+String(data[di][35]||'')+'] AL37=['+String(data[di][37]||'')+'] AN39=['+String(data[di][39]||'')+'] statut=['+lireStatutInscription(data[di])+'] actId=['+lireActiviteId(data[di])+']');
  }

  // ── Étape 1 : mettre à jour les lignes dans le Sheet ──────────
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowCode   = String(row[19] || '').trim(); // col T = code_dossier
    var rowNom    = String(row[2]  || '').trim(); // col C = Nom
    var rowPrenom = String(row[3]  || '').trim(); // col D = Prénom
    var rowStatut = lireStatutInscription(row);   // AN(39) ou AL(37)

    if (rowCode !== code) continue;
    if (rowStatut.toLowerCase().indexOf('attente') < 0) continue;

    if (!actId) return {status:'error',message:'actId manquant'};
    if (!nomMembre || !prenomMembre) return {status:'error',message:'nom membre manquant'};

    // Matching nom + prenom
    var matchMem = rowNom.toLowerCase() === nomMembre.toLowerCase()
                && rowPrenom.toLowerCase() === prenomMembre.toLowerCase();
    if (!matchMem) continue;

    // Pour les lignes en attente, chercher l'actId dans TOUTES les colonnes candidates
    // AJ(35), AL(37), col W(22)=nom activité → tenter plusieurs colonnes
    var candidates = [
      String(row[35]||'').trim(), // AJ
      String(row[37]||'').trim(), // AL
      String(row[36]||'').trim(), // AK
      String(row[22]||'').trim(), // W = nom activité (parfois contient l'ID)
    ];
    var matchAct = candidates.some(function(c){ return c === actId; });

    Logger.log('Bascule L'+(i+2)+': matchMem='+matchMem+' matchAct='+matchAct
      +' candidates='+JSON.stringify(candidates)+' actId=['+actId+']');

    if (!matchAct) continue;

    // ── Capturer la ligne matchée pour l'écriture dans l'onglet (hors boucle) ──
    var matchedRow = row.slice(); // copie de la ligne matchée
    var matchedRowCode = rowCode;
    var matchedRowNom = rowNom;
    var matchedRowPrenom = rowPrenom;

    var sheetRow = i + 2; // +1 header, +1 base-1
    // Mettre à jour statut → 'Inscrit' (col AN = index 39 = col 40)
    // Mettre à jour le statut inscription selon la structure de la ligne
    var colStatut = isNewStructure(row) ? 40 : 38; // AN(40) nouvelle / AL(38) ancienne
    sheet.getRange(sheetRow, colStatut).setValue('✅ Inscrit').setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
    // Mettre a jour col 22 (statut paiement) selon que le dossier est deja regle ou non
    var modeP2 = String(row[32]||'').toLowerCase();
    var modeLabelB = {helloasso:'HelloAsso',cheque:'Chèque',cheque3:'Chèques',especes:'Espèces'};
    var modeDisplayB = modeLabelB[modeP2] || modeP2;
    // Verifier que d'autres lignes du dossier sont deja en statut Paye
    var autresPayees = false;
    for (var jj2 = 0; jj2 < data.length; jj2++) {
      if (String(data[jj2][19]||'').trim() !== code) continue;
      var st22 = String(data[jj2][21]||'').toLowerCase();
      if (st22.indexOf('pay') >= 0) { autresPayees = true; break; }
    }
    // Ligne basculee : toujours en attente de reglement (activite non payee)
    sheet.getRange(sheetRow, 22)
      .setValue('⏳ En attente de règlement — '+modeDisplayB)
      .setBackground('#fff8e1').setFontColor('#e65100').setFontWeight('bold');
    // Mettre à jour tarif brut (col AB = col 28)
    if (tarifBrut > 0) {
      sheet.getRange(sheetRow, 28).setValue(tarifBrut);
      // Recalculer tarif net col AD = AB*(1-0.15*AC)
      var ac = parseFloat(row[28])||0; // col AC = éligible remise
      sheet.getRange(sheetRow, 30).setValue(tarifBrut * (1 - 0.15 * ac));
    }
    // Recalculer le total famille : somme de toutes les activités du dossier + FNSMR
    // Le FNSMR (15€) est payé une seule fois par membre — vérifier si déjà payé
    var totalFamille = 0;
    var fnsmrDejaCompte = {};
    for (var jj = 0; jj < data.length; jj++) {
      var rj = data[jj];
      if (String(rj[19]||'').trim() !== code) continue;
      var statutJ = lireStatutInscription(rj);
      if (statutJ.toLowerCase().indexOf('supprim') >= 0) continue;
      if (statutJ.toLowerCase().indexOf('attente de place') >= 0 && jj !== i) continue;
      var tarifJ = parseFloat(rj[27]||0); // col AB = tarif brut
      var memKeyJ = (String(rj[2]||'')+'_'+String(rj[3]||'')).toLowerCase();
      totalFamille += tarifJ;
      if (!fnsmrDejaCompte[memKeyJ]) { totalFamille += 15; fnsmrDejaCompte[memKeyJ] = true; }
    }
    // Ajouter la nouvelle activité basculée (son tarif n'est pas encore compté)
    if (tarifBrut > 0) { totalFamille += tarifBrut; }
    // Remise 15% famille si 3+ activités éligibles
    var actsEligibles = 0;
    for (var jj3 = 0; jj3 < data.length; jj3++) {
      var rj3 = data[jj3];
      if (String(rj3[19]||'').trim() !== code) continue;
      var stj3 = lireStatutInscription(rj3);
      if (stj3.toLowerCase().indexOf('supprim') >= 0) continue;
      if (stj3.toLowerCase().indexOf('attente de place') >= 0 && jj3 !== i) continue;
      var actIdJ3 = lireActiviteId(rj3);
      if (!actIdJ3.startsWith('tt-')) actsEligibles++;
    }
    if (tarifBrut > 0 && !actId.startsWith('tt-')) actsEligibles++;
    var remiseFamille = actsEligibles >= 3 ? Math.round(totalFamille * 0.15 * 100) / 100 : 0;
    var totalNet = Math.round((totalFamille - remiseFamille) * 100) / 100;
    // Mettre à jour col AF (32) pour toutes les lignes du dossier
    for (var jj4 = 0; jj4 < data.length; jj4++) {
      if (String(data[jj4][19]||'').trim() !== code) continue;
      sheet.getRange(jj4 + 2, 32).setValue(totalNet);
    }
    Logger.log('Bascule recalcul total: '+totalFamille+' remise:'+remiseFamille+' net:'+totalNet+' actsEligibles:'+actsEligibles);
    Logger.log('Bascule L'+sheetRow+' statut col'+colStatut+'→Inscrit tarif:'+tarifBrut);
    rowsUpdated++;
    // Construire emailRows pour l'email
    emailRows.push({
      code_dossier:       matchedRowCode,
      responsable_nom:    String(matchedRow[38]||'').trim().split(' ').slice(1).join(' ') || matchedRowNom,
      responsable_prenom: String(matchedRow[38]||'').trim().split(' ')[0] || matchedRowPrenom,
      email1:             String(matchedRow[15]||''),
      tel1:               String(matchedRow[14]||''),
      adresse:            String(matchedRow[7]||'')+', '+String(matchedRow[9]||'')+' '+String(matchedRow[10]||''),
      commune:            String(matchedRow[10]||'').toUpperCase().indexOf('ISNEAUVILLE')>=0?'isno':'',
      membre_nom:         matchedRowNom,
      membre_prenom:      matchedRowPrenom,
      activite:           actNom || String(matchedRow[22]||''),
      activite_id:        actId,
      jour:               String(matchedRow[23]||''),
      heure:              String(matchedRow[24]||''),
      lieu:               String(matchedRow[25]||''),
      tarif_brut:         tarifBrut > 0 ? tarifBrut : Number(matchedRow[27]||0),
      tarif:              tarifBrut > 0 ? tarifBrut : Number(matchedRow[27]||0),
      statut_inscription: 'Inscrit',
      mode_paiement:      String(matchedRow[32]||'') || 'À régler',
      total_famille:      String(totalNet || tarifBrut || Number(matchedRow[27]||0)),
      qs_sante:           String(matchedRow[34]||''),
      avoir_montant:      String(matchedRow[33]||''),
      pass_sport_montant: '0',
      ancv_montant:       '0',
      pass_aide:          String(matchedRow[35]||''),
      fnsmr:              0,
      remise:             0,
      note_tarif:         ''
    });
    // Vérifier si le dossier est déjà réglé
    var modeP = String(matchedRow[32]||'').toLowerCase();
    if (modeP && modeP !== 'pending') isPaidDossier = true;
    break; // ← Sortir dès le premier match trouvé
  }

  if (rowsUpdated === 0) return {status:'error', message:'Ligne liste attente introuvable pour ' + code + ' / ' + actId};
  Logger.log('✅ Basculé ' + rowsUpdated + ' ligne(s) pour ' + code + ' — ' + actNom);

  // ── Mise à jour onglet Places : incrementer Inscrits, decrementer Places restantes ──
  try {
    var placesSheet = getOrCreatePlacesSheet(ss);
    var placesId2 = getPlacesId(actId) || actId;
    var plData = placesSheet.getRange(2,1,Math.max(placesSheet.getLastRow()-1,1),5).getValues();
    for (var pi=0; pi<plData.length; pi++) {
      var pid = String(plData[pi][0]||'').trim();
      if (pid !== placesId2 && pid !== actId) continue;
      var cap  = parseInt(plData[pi][2])||20;
      var insc = parseInt(plData[pi][3])||0;
      var newInsc = insc + rowsUpdated;
      var newDisp = Math.max(0, cap - newInsc);
      placesSheet.getRange(pi+2, 4).setValue(newInsc);
      placesSheet.getRange(pi+2, 5).setValue(newDisp)
        .setBackground(newDisp === 0 ? '#ffcdd2' : newDisp <= 2 ? '#fff9c4' : '#d8f3dc')
        .setFontColor(newDisp === 0 ? '#c62828' : '#2d6a4f').setFontWeight('bold');
      Logger.log('✅ Places bascule: ' + pid + ' inscrits=' + newInsc + ' dispo=' + newDisp);
      break;
    }
  } catch(ePlaces) { Logger.log('Places bascule KO: ' + ePlaces); }

  // ── Étape 2 : écrire dans l'onglet {placesId} ──
  try {
    var actSheetBascule = ss.getSheetByName(actId);
    if (!actSheetBascule) {
      actSheetBascule = ss.insertSheet(actId);
      var couleurB = getCouleurActivite(actId);
      var hdrB = ['N° Dossier','Date','Statut','Nom','Prénom','Date naiss.','Sexe','Responsable','Téléphone','Email','Ville','QS Santé','Paiement','Pass Aide'];
      actSheetBascule.getRange(1,1,1,hdrB.length).setValues([hdrB]).setBackground(couleurB).setFontColor('#ffffff').setFontWeight('bold');
      actSheetBascule.getRange(2,1,1,hdrB.length).setValue(actNom||actId).setFontStyle('italic').setFontSize(9).setFontColor('#555').setBackground('#f5f5f5');
      actSheetBascule.setFrozenRows(2);
    }
    var dateJourB = Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy à HH:mm');
    var statutActB = '⏳ En attente de règlement — activité confirmée';
    var bgActB = '#fff8e1';
    // Chercher si la ligne existe déjà (inscrite depuis LA lors de l'inscription initiale)
    var existingRowB = -1;
    if (actSheetBascule.getLastRow() >= 3) {
      var existDataB = actSheetBascule.getRange(3, 1, actSheetBascule.getLastRow()-2, 1).getValues();
      for (var ei=0; ei<existDataB.length; ei++) {
        if (String(existDataB[ei][0]||'').trim() === matchedRowCode) { existingRowB = ei+3; break; }
      }
    }
    if (existingRowB > 0) {
      // Mettre à jour le statut de la ligne existante
      actSheetBascule.getRange(existingRowB, 3).setValue(statutActB)
        .setBackground(bgActB).setFontColor('#e65100').setFontWeight('bold');
      Logger.log('✅ Mise à jour onglet '+actId+' L'+existingRowB+' pour '+matchedRowCode);
    } else {
      // Créer une nouvelle ligne
      var rowActB = [matchedRowCode, dateJourB, statutActB, matchedRowNom, matchedRowPrenom,
        (function(d){ return d instanceof Date ? Utilities.formatDate(d,'Europe/Paris','dd/MM/yyyy') : String(d||''); })(matchedRow[4]),
        lireSexe(matchedRow), String(matchedRow[38]||''),
        String(matchedRow[14]||''), String(matchedRow[15]||''), String(matchedRow[10]||''),
        String(matchedRow[34]||''), String(matchedRow[32]||''), String(matchedRow[35]||'')];
      var nextActRowB = Math.max(actSheetBascule.getLastRow()+1, 3);
      actSheetBascule.getRange(nextActRowB,1,1,rowActB.length).setValues([rowActB]).setBackground(bgActB);
      Logger.log('✅ Inscrit dans onglet '+actId+' pour '+matchedRowCode+' statut:'+statutActB);
    }
  } catch(eActSheet){ Logger.log('Onglet activité bascule KO: '+eActSheet); }

  // ── Étape 3 : mettre à jour l'onglet LA ──
  try {
    var nomLA = 'LA-' + actId.replace(/[\/\\:?*[\]]/g, '-').substring(0, 28);
    var sheetLA = ss.getSheetByName(nomLA);
    if (sheetLA && sheetLA.getLastRow() > 1) {
      var laData = sheetLA.getRange(2, 1, sheetLA.getLastRow()-1, 10).getValues();
      for (var j = 0; j < laData.length; j++) {
        var laCode = String(laData[j][0]||'').trim();
        var laNom  = String(laData[j][3]||'').trim();
        if (laCode === code && laNom.toUpperCase() === nomMembre.toUpperCase()) {
          // Mettre à jour le statut dans l'onglet LA selon paiement
          // Onglet LA : conserver la ligne avec statut 'Inscrit' (la place est confirmee)
          // Le reglement de l'activite basculee se fait separement
          sheetLA.getRange(j+2, 3).setValue('Inscrit — ⏳ Règlement à valider').setFontColor('#1565c0').setFontWeight('bold');
          sheetLA.getRange(j+2, 1, 1, 10).setBackground('#e3f2fd');
          Logger.log('✅ Onglet LA mis à jour pour ' + code);
          break;
        }
      }
    }
  } catch(eLA) { Logger.log('Onglet LA update KO: ' + eLA); }

  // ── Étape 3 : lire TOUTES les lignes du dossier après basculement ──
  SpreadsheetApp.flush();
  var allEmailRows = [];
  try {
    var dataAll = sheet.getRange(2, 1, Math.max(sheet.getLastRow()-1,1), 41).getValues();
    var r0Bascule = emailRows.length > 0 ? emailRows[0] : null;
    var respFullB = r0Bascule ? (r0Bascule.responsable_prenom + ' ' + r0Bascule.responsable_nom) : '';

    for (var bi = 0; bi < dataAll.length; bi++) {
      var bRow = dataAll[bi];
      if (String(bRow[19]||'').trim() !== code) continue;
      var bStatut = lireStatutInscription(bRow).toLowerCase();
      if (bStatut.indexOf('supprim') >= 0) continue; // exclure les lignes supprimées

      var bRespFull = String(bRow[38]||'').trim();
      var bRespParts = bRespFull ? bRespFull.split(' ') : [];
      var bMode = String(bRow[32]||'');
      var modeLabelMap2 = {helloasso:'HelloAsso',cheque:'Cheque',cheque3:'Cheques 3x',especes:'Especes'};
      var passAideB = String(bRow[35]||'');
      var ancvB=0,passB=0,pjB=0,atoutB=0,avoirB=parseFloat(bRow[33])||0;
      var ancvMB=passAideB.match(/(?:^|\|)ANCV:([\d.]+)/),passM2B=passAideB.match(/(?:^|\|)PASS:([\d.]+)/);
      var pjMB=passAideB.match(/PassJeunes:([\d.]+)/),atMB=passAideB.match(/Atout:([\d.]+)/);
      if(ancvMB) ancvB=parseFloat(ancvMB[1])||0;
      if(passM2B) passB=parseFloat(passM2B[1])||0;
      if(pjMB)  pjB=parseFloat(pjMB[1])||0;
      if(atMB)  atoutB=parseFloat(atMB[1])||0;

      allEmailRows.push({
        code_dossier:       code,
        date:               (function(d){ if(!d)return''; if(d instanceof Date)return Utilities.formatDate(d,'Europe/Paris','dd/MM/yyyy'); var s=String(d); return s.indexOf('GMT')>=0?Utilities.formatDate(new Date(s),'Europe/Paris','dd/MM/yyyy'):s; })(bRow[20]),
        responsable_nom:    bRespParts.length > 1 ? bRespParts.slice(1).join(' ') : String(bRow[2]||''),
        responsable_prenom: bRespParts[0] || String(bRow[3]||''),
        email1:             String(bRow[15]||''),
        tel1:               String(bRow[14]||''),
        adresse:            String(bRow[7]||''),
        cp:                 String(bRow[9]||''),
        ville:              String(bRow[10]||''),
        membre_nom:         String(bRow[2]||''),
        membre_prenom:      String(bRow[3]||''),
        ddn:                String(bRow[4]||''),
        sexe:               String(bRow[36]||''),
        activite:           String(bRow[22]||''),
        activite_id:        lireActiviteId(bRow),
        jour:               String(bRow[23]||''),
        heure:              String(bRow[24]||''),
        lieu:               String(bRow[25]||''),
        animateur:          String(bRow[26]||''),
        tarif_brut:         Number(bRow[27]||0),
        tarif:              Number(bRow[27]||0),
        fnsmr:              Number(bRow[30]||0),
        total_famille:      Number(bRow[31]||0),
        mode_paiement:      bMode,
        statut_paiement:    String(bRow[21]||''),   // col 22 = statut paiement (⏳ En attente / ✅ Payé...)
        avoir_montant:      avoirB,
        ancv_montant:       ancvB,
        pass_sport_montant: passB,
        pass_aide:          passAideB,
        qs_sante:           String(bRow[34]||''),
        statut_inscription: lireStatutInscription(bRow),
        note_tarif:         '',
        fftt_price:         0, // FFTT supprimé
        _estBasculee:       (lireActiviteId(bRow) === actId
                             && String(bRow[2]||'').trim().toLowerCase() === nomMembre.toLowerCase()
                             && String(bRow[3]||'').trim().toLowerCase() === prenomMembre.toLowerCase()) // marquer l'activité basculée (bon membre uniquement)
      });
    }
  } catch(eRead) { Logger.log('Lecture allEmailRows bascule KO: ' + eRead); }

  // ── Étape 4 : envoyer emails avec récapitulatif complet ──────
  var rowsToSend = allEmailRows.length > 0 ? allEmailRows : emailRows;
  if (rowsToSend.length > 0) {
    try {
      var r0Send = rowsToSend[0];
      var emailDest = r0Send.email1 || (emailRows.length > 0 ? emailRows[0].email1 : '');
      var modeLabel = isPaidDossier ? 'Deja regle' : 'helloasso';
      var modeLabelFr = isPaidDossier ? 'Deja regle' : 'HelloAsso';
      envoyerEmailBasculeListe(emailDest, rowsToSend, actNom, tarifBrut, isPaidDossier, actId, commentaireAdmin);
      envoyerEmailAdminBascule(EMAIL_ADMIN, rowsToSend, actNom, tarifBrut, isPaidDossier, actId);
      Logger.log('✅ Emails bascule complets envoyés pour ' + code);
    } catch(eEmail) { Logger.log('Email bascule KO: ' + eEmail); }
  }
  logCommentaireAdmin(ss, 'Bascule liste attente', code, nomMembre, prenomMembre, actNom, tarifBrut, commentaireAdmin, montantModifie);

  // Recalculer col AF (total famille) + Recap col I apres bascule
  try {
    var ss2=SpreadsheetApp.openById(SHEET_ID);
    var newTotal=calcTotalFamille(ss2,code);
    var recapSh=ss2.getSheetByName(SHEET_RECAPITULATIF);
    if (recapSh && recapSh.getLastRow()>1) {
      var rD=recapSh.getRange(2,1,recapSh.getLastRow()-1,9).getValues();
      for (var ri=0;ri<rD.length;ri++) {
        if (!rD[ri].some(function(c){return String(c||'').trim()===code;})) continue;
        recapSh.getRange(ri+2,9).setValue(newTotal).setFontColor('#1565c0').setFontWeight('bold');
        break;
      }
    }
    Logger.log('✅ Total famille bascule: '+code+'='+newTotal);
  } catch(eTot){Logger.log('calcTotalFamille bascule KO: '+eTot);}

  return {
    status: 'ok',
    rowsUpdated: rowsUpdated,
    message: rowsUpdated + ' ligne(s) basculée(s) — emails envoyés'
  };
}

// ============================================================
// HELPER PARTAGÉ — calcul financier depuis rows
// Retourne l'objet { parMembre, nbMembres, totalFnsmr, ffttMembres, totalFFTT,
//   aRemise, nbEligibles, totalBrut, totalRemise, totalNet, totalActivites,
//   deducPassSport, deducAncv, deducAvoir, deducPassJeunes, deducAtout,
//   totalDeductions, solde, cheques }
// ============================================================
function _calcFinancier(rows) {
  // Activites non remisables pour TOUS : uniquement la Marche Nordique
  // Jazz et Country (seances normales) sont remisables
  // Country en FORFAIT est exclu via note_tarif (!note dans le calcul)
  var NON_REM_TOUS = ['MNOME/S10'];
  // Activites TT (PING*) : remisables uniquement pour habitants d'Isneauville
  var _ville0 = rows[0] ? (rows[0].ville || '') : '';
  var _commune0 = rows[0] ? (rows[0].commune || '') : '';
  var commune_calc = (_commune0 === 'isno' || _commune0 === 'Isneauville')
    ? 'Isneauville'
    : communeFromVille(_ville0);
  var isIsneauville = (commune_calc === 'Isneauville');
  function elig(pid) {
    if (!pid) return false;
    // Jazz, Marche Nordique, Country : jamais remisables
    if (NON_REM_TOUS.indexOf(pid) >= 0) return false;
    // TT (PING*) : remisable seulement pour Isneauville
    if (pid.indexOf('PING') >= 0) return isIsneauville;
    return true;
  }

  var parMembre = {}, ffttMembres = {}, seenPids = {}, nbElig = 0;
  rows.forEach(function(r){
    var k = (r.membre_prenom||'') + ' ' + (r.membre_nom||'');
    if (!parMembre[k]) parMembre[k] = [];
    parMembre[k].push(r);
    var pid = r.activite_id || '';
    var isWait = String(r.statut_inscription||'').toLowerCase().indexOf('attente') >= 0;
    if (!isWait && !seenPids[pid]) { seenPids[pid]=true; if(elig(pid)) nbElig++; }
    var ffttP = parseFloat(r.fftt_price)||0;
    if (ffttP>0 && !ffttMembres[k]) ffttMembres[k] = ffttP;
  });

  // Adhésion déjà réglée (ajout d'activités à un dossier existant) : pas de nouvelle FNSMR
  var nbMembres  = Object.keys(parMembre).filter(function(k) {
    return parMembre[k].some(function(r) { return !r.adhesion_deja_reglee; });
  }).length;
  var totalFnsmr = nbMembres * 15;
  // Activités éligibles déjà présentes dans le dossier : comptent pour la remise famille
  nbElig += Number((rows[0] && rows[0].nb_elig_existants) || 0);
  var totalFFTT  = Object.keys(ffttMembres).reduce(function(s,k){ return s+ffttMembres[k]; }, 0);
  var aRemise    = nbElig >= 3;
  Logger.log('_calcFinancier: ville=['+_ville0+'] commune=['+commune_calc+'] nbElig='+nbElig+' aRemise='+aRemise
    +' pids='+JSON.stringify(Object.keys(seenPids))
    +' tarifs='+JSON.stringify(rows.map(function(r){return {id:r.activite_id,b:r.tarif_brut,st:r.statut_inscription};})));

  var totalBrut=0, totalRemise=0, totalNet=0;
  rows.forEach(function(r){
    var b = parseFloat(r.tarif_brut)||parseFloat(r.tarif)||0;
    var isWait = String(r.statut_inscription||'').toLowerCase().indexOf('attente') >= 0;
    var note   = String(r.note_tarif||'');
    var e = !isWait && !note && aRemise && elig(r.activite_id||'');
    var rem = e ? Math.round(b*0.15*100)/100 : 0;
    totalBrut   += b;
    totalRemise += rem;
    totalNet    += Math.round((b-rem)*100)/100;
  });
  totalBrut   = Math.round(totalBrut*100)/100;
  totalRemise = Math.round(totalRemise*100)/100;
  totalNet    = Math.round(totalNet*100)/100;
  var totalActivites = Math.round((totalNet + totalFnsmr + totalFFTT)*100)/100;

  // Chercher les aides sur toutes les lignes (pas seulement r0)
  // car l'ordre des lignes peut varier selon l'activité
  var r0 = rows[0];
  var deducPassSport = 0, deducAncv = 0, deducAvoir = 0;
  var deducPassJeunes = 0, deducAtout = 0;
  var _paStrBest = '';
  rows.forEach(function(r) {
    var ps = parseFloat(r.pass_sport_montant)||0;
    var an = parseFloat(r.ancv_montant)||0;
    var av = parseFloat(r.avoir_montant)||0;
    var pj = parseFloat(r.pass_jeunes_montant)||0;
    var at = parseFloat(r.atout_montant)||0;
    if (ps > deducPassSport) deducPassSport = ps;
    if (an > deducAncv)     deducAncv     = an;
    if (av > deducAvoir)    deducAvoir    = av;
    if (pj > deducPassJeunes) deducPassJeunes = pj;
    if (at > deducAtout)    deducAtout    = at;
    var pa = String(r.pass_aide||'');
    if (pa.length > _paStrBest.length) _paStrBest = pa; // garder la plus complète
  });
  // Fallback depuis pass_aide si champs directs absents
  // pass_aide format : "PassJeunes:30.00:1ere|Atout:30.00:CODE|ANCV:30.00|PASS:50.00"
  var paStr = _paStrBest || String(r0.pass_aide||'');
  var ancvM2 = paStr.match(/(?:^|\|)ANCV:([\d.]+)/);
  var passM2 = paStr.match(/(?:^|\|)PASS:([\d.]+)/);
  var pjM    = paStr.match(/PassJeunes:([\d.]+)/);
  var atM    = paStr.match(/Atout:([\d.]+)/);
  if (!deducPassSport && passM2)   deducPassSport  = parseFloat(passM2[1])||0;
  if (!deducAncv && ancvM2)        deducAncv       = parseFloat(ancvM2[1])||0;
  if (!deducPassJeunes && pjM)     deducPassJeunes = parseFloat(pjM[1])||0;
  if (!deducAtout && atM)          deducAtout      = parseFloat(atM[1])||0;
  var totalDeductions = Math.round((deducPassSport+deducAncv+deducAvoir+deducPassJeunes+deducAtout)*100)/100;
  var solde           = Math.max(0, Math.round((totalActivites-totalDeductions)*100)/100);

  // cheques
  var cheques = [];
  try {
    var src = (r0.cheques&&r0.cheques.length) ? r0.cheques : null;
    if (src) cheques = src;
  } catch(e){}

  return {
    parMembre:parMembre, nbMembres:nbMembres, totalFnsmr:totalFnsmr,
    ffttMembres:ffttMembres, totalFFTT:totalFFTT,
    aRemise:aRemise, nbEligibles:nbElig,
    totalBrut:totalBrut, totalRemise:totalRemise, totalNet:totalNet,
    totalActivites:totalActivites,
    deducPassSport:deducPassSport, deducAncv:deducAncv, deducAvoir:deducAvoir,
    deducPassJeunes:deducPassJeunes, deducAtout:deducAtout,
    totalDeductions:totalDeductions, solde:solde,
    cheques:cheques,
    rows: rows // pour commune_calc dans _htmlPartie1Activites
  };
}

// ============================================================
// HELPER PARTAGÉ — HTML Partie 1 : tableau activités par membre
// ============================================================
function _htmlPartie1Activites(f, isAdminView) {
  // f = résultat de _calcFinancier(rows)
  var html = '';

  // En-tête du tableau
  html += '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:8px">';
  html += '<thead><tr style="background:#1a2e22;color:white">';
  html += '<th style="padding:8px 10px;text-align:left">Activité</th>';
  html += '<th style="padding:8px 10px;text-align:left">Jour / Heure</th>';
  if (f.aRemise) {
    html += '<th style="padding:8px 10px;text-align:right">Tarif brut</th>';
    html += '<th style="padding:8px 10px;text-align:right">Remise -15%</th>';
    html += '<th style="padding:8px 10px;text-align:right">Tarif net</th>';
  } else {
    html += '<th style="padding:8px 10px;text-align:right">Tarif</th>';
  }
  html += '</tr></thead><tbody>';

  Object.keys(f.parMembre).forEach(function(membre) {
    // Ligne d'en-tête membre
    var colspanMembre = f.aRemise ? 5 : 3;
    html += '<tr><td colspan="' + colspanMembre + '" style="background:#2d6a4f;color:white;padding:7px 10px;font-weight:bold;font-size:13px">👤 ' + membre + '</td></tr>';

    f.parMembre[membre].forEach(function(r) {
      var b = parseFloat(r.tarif_brut)||parseFloat(r.tarif)||0;
      var isWait = String(r.statut_inscription||'').toLowerCase().indexOf('attente') >= 0;
      var note   = String(r.note_tarif||'');
      var _commune1 = communeFromVille(f.rows ? (f.rows[0]||{}).ville||'' : '');
      // note_tarif = descriptif (ex: '2 seances Lu+Me') ou forfait ('Forfait country N seances')
      var isForfait = note && note.toLowerCase().indexOf('forfait') >= 0;
      var e      = !isWait && !isForfait && f.aRemise && (function(pid){
        if (!pid) return false;
        if (pid === 'MNOME/S10') return false;
        if (pid.indexOf('PING') >= 0) return _commune1 === 'Isneauville';
        return true;
      })(r.activite_id||'');
      var rem   = e ? Math.round(b*0.15*100)/100 : 0;
      var net   = isWait ? 0 : Math.round((b-rem)*100)/100;
      var actNom = (r.activite||'').replace(/\n/g,' — ');
      var joHe   = (r.jour||'') + (r.heure ? ' · ' + r.heure : '');
      var rowBg  = isWait ? '#fff3cd' : (e ? '#d8f3dc' : '#ffffff');

      if (isWait) {
        // Liste d'attente
        html += '<tr style="background:' + rowBg + '">';
        html += '<td style="padding:7px 10px;border-bottom:1px solid #eee"><em style="color:#e65100">⏳ ' + actNom + '</em></td>';
        html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;color:#e65100">' + joHe + '</td>';
        if (f.aRemise) {
          html += '<td colspan="3" style="padding:7px 10px;border-bottom:1px solid #eee;text-align:center;color:#e65100;font-style:italic">En liste d\'attente — règlement non demandé</td>';
        } else {
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:center;color:#e65100;font-style:italic">En liste d\'attente</td>';
        }
        html += '</tr>';
      } else if (isForfait) {
        // Forfait country : pas de remise
        html += '<tr style="background:#f0f7f3">';
        html += '<td style="padding:7px 10px;border-bottom:1px solid #eee">' + actNom + ' <span style="font-size:11px;color:#2d6a4f;font-style:italic">(forfait)</span></td>';
        html += '<td style="padding:7px 10px;border-bottom:1px solid #eee">' + joHe + '</td>';
        if (f.aRemise) {
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right">' + (b>0?b.toFixed(2)+' €':'—') + '</td>';
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;color:#1b5e20;font-style:italic">' + note + '</td>';
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;font-weight:bold">' + (b>0?b.toFixed(2)+' €':'Inclus') + '</td>';
        } else {
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;font-weight:bold">' + (b>0?b.toFixed(2)+' €':'Inclus') + '</td>';
        }
        html += '</tr>';
      } else {
        // Activité standard
        html += '<tr style="background:' + rowBg + '">';
        html += '<td style="padding:7px 10px;border-bottom:1px solid #eee">' + actNom + '</td>';
        html += '<td style="padding:7px 10px;border-bottom:1px solid #eee">' + joHe + '</td>';
        if (f.aRemise) {
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right">' + b.toFixed(2) + ' €</td>';
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;color:' + (e?'#2d6a4f':'#999') + '">' + (e?'- '+rem.toFixed(2)+' €':'—') + '</td>';
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;font-weight:bold;color:' + (e?'#1b5e20':'#222') + '">' + net.toFixed(2) + ' €</td>';
        } else {
          html += '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right;font-weight:bold">' + b.toFixed(2) + ' €</td>';
        }
        html += '</tr>';
      }
    });

    // Ligne FFTT membre si applicable
    if (f.ffttMembres[membre]) {
      var colspanFftt = f.aRemise ? 4 : 2;
      html += '<tr style="background:#fff3e0">';
      html += '<td style="padding:6px 10px;font-size:12px;color:#e65100" colspan="2">🏓 Licence FFTT — ' + membre + '</td>';
      if (f.aRemise) {
        html += '<td style="padding:6px 10px;text-align:right" colspan="2"></td>';
      }
      html += '<td style="padding:6px 10px;text-align:right;font-weight:bold;color:#e65100">' + f.ffttMembres[membre].toFixed(2) + ' €</td>';
      html += '</tr>';
    }
  });

  // Lignes récap totaux
  var colspan0 = f.aRemise ? 3 : 1;
  if (f.aRemise && f.totalRemise > 0) {
    html += '<tr style="background:#e8f5e9"><td colspan="' + colspan0 + '" style="padding:7px 10px;font-size:12px;color:#2d6a4f">Sous-total activités (brut)</td>'
      + '<td style="padding:7px 10px;text-align:right;font-size:12px">' + f.totalBrut.toFixed(2) + ' €</td>'
      + '<td style="padding:7px 10px;text-align:right;font-weight:bold;font-size:12px;color:#2d6a4f"></td></tr>';
    html += '<tr style="background:#d8f3dc"><td colspan="' + colspan0 + '" style="padding:7px 10px;font-size:12px;color:#2d6a4f">🎉 Remise famille −15% (' + f.nbEligibles + ' activité(s) éligible(s))</td>'
      + '<td style="padding:7px 10px;text-align:right;font-weight:bold;color:#2d6a4f">− ' + f.totalRemise.toFixed(2) + ' €</td>'
      + '<td style="padding:7px 10px;text-align:right;font-weight:bold;color:#1b5e20">' + f.totalNet.toFixed(2) + ' €</td></tr>';
  }
  // FNSMR
  html += '<tr style="background:#f5f5f5"><td colspan="' + (f.aRemise?4:2) + '" style="padding:7px 10px;color:#555;font-size:12px">Adhésion FNSMR (' + f.nbMembres + ' pers. × 15,00 €)</td>'
    + '<td style="padding:7px 10px;text-align:right;font-size:12px">' + f.totalFnsmr.toFixed(2) + ' €</td></tr>';
  // Total général
  html += '<tr style="background:#1a2e22"><td colspan="' + (f.aRemise?4:2) + '" style="padding:9px 10px;font-weight:bold;font-size:14px;color:white">TOTAL GÉNÉRAL</td>'
    + '<td style="padding:9px 10px;text-align:right;font-weight:bold;font-size:16px;color:#52b788">' + f.totalActivites.toFixed(2) + ' €</td></tr>';

  html += '</tbody></table>';
  if (f.aRemise) {
    html += '<div style="background:#d8f3dc;border:1px solid #52b788;border-radius:6px;padding:9px 12px;margin-top:6px;font-size:12px;color:#1b5e20">'
      + '🎉 <strong>Remise famille −15% appliquée</strong> — économie de <strong>' + f.totalRemise.toFixed(2) + ' €</strong> sur ' + f.nbEligibles + ' activité(s)</div>';
  }
  return html;
}

// ============================================================
// HELPER PARTAGÉ — HTML Partie 2 : aides / pass / avoirs
// isPaid=true → "Vérifié par le FRI ✅"  |  false → "À valider par le FRI ⏳"
// ============================================================
function _htmlPartie2Aides(f, isPaid) {
  var hasAides = f.deducPassSport>0 || f.deducAncv>0 || f.deducAvoir>0
              || f.deducPassJeunes>0 || f.deducAtout>0;
  if (!hasAides && f.totalDeductions===0) return '';

  var verif = isPaid
    ? '<span style="background:#d8f3dc;color:#1b5e20;font-size:11px;font-weight:bold;padding:2px 8px;border-radius:10px;margin-left:8px">✅ Vérifié par le FRI</span>'
    : '<span style="background:#fff8e1;color:#856404;font-size:11px;font-weight:bold;padding:2px 8px;border-radius:10px;margin-left:8px">⏳ À valider par le FRI</span>';

  var html = '<div style="border:2px solid ' + (isPaid?'#52b788':'#e8c84a') + ';border-radius:10px;overflow:hidden;margin-top:14px">';
  html += '<div style="background:' + (isPaid?'#2d6a4f':'#856404') + ';padding:10px 14px;display:flex;align-items:center">';
  html += '<span style="color:white;font-weight:bold;font-size:14px">🎫 Aides / Pass / Avoirs</span>' + verif + '</div>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:13px">';

  function ligneAide(label, montant, couleur, note) {
    return '<tr style="border-bottom:1px solid #eee">'
      + '<td style="padding:8px 14px;color:#333">' + label + (note?'<br><span style="font-size:11px;color:#888;font-style:italic">'+note+'</span>':'') + '</td>'
      + '<td style="padding:8px 14px;text-align:right;font-weight:bold;color:' + couleur + '">− ' + montant.toFixed(2) + ' €</td>'
      + '</tr>';
  }

  if (f.deducPassSport>0)   html += ligneAide('Pass\'sport État', f.deducPassSport, '#1565c0', 'Déduction sur présentation du coupon en permanence');
  if (f.deducAncv>0)        html += ligneAide('Coupon ANCV', f.deducAncv, '#1565c0', 'Déduction sur remise du coupon en permanence');
  if (f.deducAvoir>0)       html += ligneAide('Avoir dossier FRI', f.deducAvoir, '#1565c0', 'Avoir enregistré sur votre dossier');
  if (f.deducPassJeunes>0)  html += ligneAide('Pass\'jeunes 76 / Handipass\'sport', f.deducPassJeunes, '#6a1b9a', 'Déduction sur présentation du pass en permanence');
  if (f.deducAtout>0)       html += ligneAide('Atout Normandie', f.deducAtout, '#e65100', 'Déduction sur présentation de l\'attestation en permanence');

  // Ligne solde
  html += '<tr style="background:' + (isPaid?'#d8f3dc':'#fff8e1') + '">'
    + '<td style="padding:10px 14px;font-weight:bold;font-size:14px;color:#1a2e22">Total déductions</td>'
    + '<td style="padding:10px 14px;text-align:right;font-weight:bold;color:#1a2e22">− ' + f.totalDeductions.toFixed(2) + ' €</td></tr>';
  html += '<tr style="background:' + (isPaid?'#1a2e22':'#fff3cd') + '">'
    + '<td style="padding:10px 14px;font-weight:bold;font-size:15px;color:' + (isPaid?'white':'#856404') + '">💰 Solde à régler</td>'
    + '<td style="padding:10px 14px;text-align:right;font-weight:bold;font-size:18px;color:' + (isPaid?'#52b788':'#856404') + '">' + f.solde.toFixed(2) + ' €</td></tr>';
  html += '</table></div>';
  return html;
}

// ============================================================
// HELPER PARTAGÉ — HTML Partie 3 : mode de paiement
// ============================================================
function _htmlPartie3Paiement(f, rows, modeLabel, isPaid) {
  var r0 = rows[0];
  var modePaiement = r0.mode_paiement || 'helloasso';
  var montantRegle = f.totalDeductions>0 ? f.solde : f.totalActivites;

  var couleurMode = isPaid ? '#1a2e22' : '#856404';
  var bgMode      = isPaid ? '#d8f3dc' : '#fff8e1';
  var borderMode  = isPaid ? '#52b788' : '#e8c84a';
  var statut = isPaid
    ? '<span style="background:#d8f3dc;color:#1b5e20;font-size:11px;font-weight:bold;padding:2px 8px;border-radius:10px">✅ Paiement confirmé</span>'
    : '<span style="background:#fff8e1;color:#856404;font-size:11px;font-weight:bold;padding:2px 8px;border-radius:10px">⏳ En attente</span>';

  var html = '<div style="border:2px solid ' + borderMode + ';border-radius:10px;overflow:hidden;margin-top:14px">';
  html += '<div style="background:' + couleurMode + ';padding:10px 14px;display:flex;align-items:center;justify-content:space-between">';
  html += '<span style="color:white;font-weight:bold;font-size:14px">💳 Mode de paiement</span>' + statut + '</div>';
  html += '<div style="padding:12px 14px">';

  // HelloAsso
  if (modePaiement === 'helloasso') {
    html += '<div style="display:flex;align-items:center;gap:10px;padding:8px 0">';
    html += '<span style="font-size:22px">💳</span>';
    html += '<div><div style="font-weight:bold;color:#1a2e22">HelloAsso (paiement en ligne)</div>';
    html += '<div style="font-size:12px;color:#555">Paiement sécurisé en ligne</div></div>';
    html += '<div style="margin-left:auto;font-weight:bold;font-size:16px;color:#1a2e22">' + montantRegle.toFixed(2) + ' €</div>';
    html += '</div>';
  }
  // Chèque unique
  else if (modePaiement === 'cheque') {
    html += '<div style="display:flex;align-items:center;gap:10px;padding:8px 0">';
    html += '<span style="font-size:22px">📝</span>';
    html += '<div style="flex:1"><div style="font-weight:bold;color:#1a2e22">Chèque à l\'ordre du FRI</div>';
    if (f.cheques.length > 0) {
      var ch = f.cheques[0];
      html += '<div style="font-size:12px;color:#555;margin-top:2px">';
      if (ch.numCheque) html += 'N° ' + ch.numCheque;
      if (ch.banque)    html += (ch.numCheque?' — ':'') + ch.banque;
      html += '</div>';
    }
    html += '</div>';
    html += '<div style="margin-left:auto;font-weight:bold;font-size:16px;color:#1a2e22">' + montantRegle.toFixed(2) + ' €</div>';
    html += '</div>';
  }
  // 3 chèques
  else if (modePaiement === 'cheque3') {
    html += '<div style="font-weight:bold;color:#1a2e22;margin-bottom:8px">📝 Paiement en 3 chèques à l\'ordre du FRI</div>';
    var totalCheques = 0;
    (f.cheques.length>0 ? f.cheques : [{numero:1},{numero:2},{numero:3}]).forEach(function(ch, i) {
      var mt = parseFloat(ch.montant)||0;
      totalCheques += mt;
      html += '<div style="display:flex;gap:8px;padding:5px 0;border-bottom:1px solid #eee;font-size:13px">';
      html += '<span style="color:#555;min-width:80px">Chèque n°' + (ch.numero||i+1) + '</span>';
      if (ch.numCheque) html += '<span style="color:#333">' + ch.numCheque + '</span>';
      if (ch.banque)    html += '<span style="color:#888;font-size:12px"> — ' + ch.banque + '</span>';
      html += '<span style="margin-left:auto;font-weight:bold">' + (mt>0?mt.toFixed(2)+' €':'—') + '</span>';
      html += '</div>';
    });
    html += '<div style="display:flex;justify-content:space-between;padding-top:8px;font-weight:bold;color:#1a2e22">';
    html += '<span>Total</span><span>' + (totalCheques>0?totalCheques.toFixed(2):montantRegle.toFixed(2)) + ' €</span></div>';
  }
  // Espèces
  else if (modePaiement === 'especes') {
    html += '<div style="display:flex;align-items:center;gap:10px;padding:8px 0">';
    html += '<span style="font-size:22px">💵</span>';
    html += '<div style="flex:1"><div style="font-weight:bold;color:#1a2e22">Règlement en espèces</div>';
    html += '<div style="font-size:12px;color:#555">À remettre en permanence (mardi 16h30–18h30)</div></div>';
    html += '<div style="margin-left:auto;font-weight:bold;font-size:16px;color:#1a2e22">' + montantRegle.toFixed(2) + ' €</div>';
    html += '</div>';
  }
  // Autre
  else {
    html += '<div style="padding:8px 0;font-weight:bold;color:#1a2e22">' + modeLabel + ' — ' + montantRegle.toFixed(2) + ' €</div>';
  }

  html += '</div></div>';
  return html;
}

// ============================================================
// EMAIL ADHÉRENT — v8.82 restructuré 3 parties
// ============================================================
function envoyerEmailAdherent(email, rows, isPaid, modeLabel, pdfBlob) {
  var semainierBlob = null;
  if (isPaid===undefined) isPaid=true;
  if (!email||email.indexOf('@')<0) { Logger.log('envoyerEmailAdherent — email invalide: "'+email+'"'); return; }

  var r0     = rows[0];
  var prenom = r0.responsable_prenom||'';
  var nom    = r0.responsable_nom||'';
  var code   = r0.code_dossier||'';
  var f      = _calcFinancier(rows);
  if (!modeLabel) {
    var mlMap={helloasso:'HelloAsso 💳',cheque:'Chèque 📝',cheque3:'Paiement 3 chèques 📝',especes:'Espèces 💵'};
    modeLabel = mlMap[r0.mode_paiement||'helloasso'] || (r0.mode_paiement||'HelloAsso');
  }

  // Alerte liste d'attente
  var attenteRows = rows.filter(function(r){ return String(r.statut_inscription||'').toLowerCase().indexOf('attente')>=0; });
  var htmlAttente = '';
  if (attenteRows.length>0) {
    htmlAttente = '<div style="background:#e0f7fa;border:2px solid #00838f;border-radius:10px;padding:14px;margin:12px 0">'
      + '<div style="font-size:14px;font-weight:bold;color:#006064;margin-bottom:8px">⏳ Inscription(s) sur liste d\'attente</div>'
      + '<ul style="margin:0;padding-left:18px;color:#006064;font-size:13px;line-height:1.9">'
      + attenteRows.map(function(r){ return '<li><strong>'+(r.activite||'').replace(/\n/g,' — ')+'</strong> — '+(r.jour||'')+(r.heure?' '+r.heure:'')+'</li>'; }).join('')
      + '</ul><div style="background:#b2ebf2;border-radius:6px;padding:10px 12px;margin-top:10px;font-size:12px;color:#006064;line-height:1.6">'
      + '✅ Vous serez informé(e) dès qu\'une place se libère.<br>⚠️ Le règlement ne sera demandé que si une place se libère.</div></div>';
  }

  // QS Santé
  var santeHtml = '', vuQs = {};
  rows.forEach(function(r){
    var k = (r.membre_prenom||'')+' '+(r.membre_nom||'');
    if (!vuQs[k]) {
      vuQs[k]=true;
      var qsLabel, qsColor, qsStyle;
      if (r.qs_sante === 'Attestation OK') {
        // QS completé et signé en ligne
        qsLabel = '✅ Questionnaire santé signé — déjà envoyé en ligne';
        qsColor = '#2d6a4f';
        qsStyle = 'color:'+qsColor+';font-style:italic';
      } else if (r.qs_sante === 'Certificat requis') {
        // Certificat médical nécessaire
        qsLabel = '⚠️ Certificat médical à apporter obligatoirement';
        qsColor = '#c0392b';
        qsStyle = 'color:'+qsColor+';font-weight:bold';
      } else if (r.qs_sante && r.qs_sante.indexOf('non sportive') >= 0) {
        qsLabel = '✅ Pas de QS requis pour cette activité';
        qsColor = '#2d6a4f';
        qsStyle = 'color:'+qsColor;
      } else if (r.qs_sante === 'Non rempli' || !r.qs_sante) {
        qsLabel = '⚠️ Questionnaire santé non complété — à apporter signé à la première séance';
        qsColor = '#e65100';
        qsStyle = 'color:'+qsColor+';font-weight:bold';
      } else { qsLabel = r.qs_sante; qsColor = '#555'; qsStyle = 'color:'+qsColor; }
      if (qsLabel) santeHtml += '<li style="margin-bottom:5px"><strong>'+k+'</strong> — <span style="'+qsStyle+'">'+qsLabel+'</span></li>';
    }
  });

  // Bloc FFTT si des activités TT sont présentes
  var hasTTemail = rows.some(function(r){ var aid = String(r.activite_id||''); return aid.indexOf('PING') >= 0 || aid.indexOf('tt-') === 0; });
  var ffttBloc = hasTTemail
    ? '<div style="background:#e8eaf6;border:2px solid #1a237e;border-radius:8px;padding:14px 18px;margin:12px 0;">'
      + '<strong style="color:#1a237e;font-size:14px;">🏓 Tennis de Table — Documents obligatoires</strong>'
      + '<p style="font-size:13px;color:#333;margin:8px 0;line-height:1.6;">Pour participer aux compétitions, vous devez également :</p>'
      + '<ul style="font-size:13px;color:#333;margin:0 0 10px;padding-left:18px;line-height:1.8;">'
      + '<li><strong>S’inscrire à la FFTT</strong> (Fédération Française de Tennis de Table) — obligatoire pour la compétition</li>'
      + '<li><strong>Régler la licence FFTT</strong> — directement auprès du club (tarif indicatif : 37–70 € selon catégorie)</li>'
      + '<li>Fournir un <strong>QS Santé FFTT</strong>, un <strong>Parcours Personnel de Santé (PPS FFTT)</strong> ou un <strong>certificat médical</strong> (- 3 ans)</li>'
      + '</ul>'
      + '<a href="https://script.google.com/macros/s/AKfycbx-Y6io0i42BbjalFcG45--tq5-k9I-AU5kcQ7QHhI1zgh-X2baR3dU7TKWD1X8KQwI/exec"'
      + ' style="display:inline-block;background:#1a237e;color:white;border-radius:6px;padding:8px 16px;font-size:13px;font-weight:700;text-decoration:none;">'
      + '📋 Remplir le PPS FFTT en ligne</a>'
      + '</div>'
    : '';

  var bandeau = isPaid
    ? '<div style="background:#d8f3dc;padding:11px 20px;text-align:center;font-weight:bold;color:#2d6a4f">✅ Paiement confirmé — '+modeLabel+' — '+(f.totalDeductions>0?f.solde:f.totalActivites).toFixed(2)+' €</div>'
    : '<div style="background:#fff8e1;padding:11px 20px;text-align:center;font-weight:bold;color:#856404;border-bottom:2px solid #e8c84a">⏳ En attente de règlement — '+modeLabel+' — '+f.totalActivites.toFixed(2)+' €</div>';

  var mentionFacture = (isPaid&&pdfBlob)
    ? '<div style="background:#e8f4fd;border:1.5px solid #2980b9;border-radius:8px;padding:12px 16px;margin:12px 0;font-size:14px">📄 <strong>Votre reçu de paiement est joint à cet email</strong> (fichier PDF).</div>' : '';

  var paraIntro = isPaid
    ? '<p style="color:#333;line-height:1.7">Votre inscription pour la saison <strong>2026/2027</strong> est <strong>validée</strong>.</p>'
    : '<p style="color:#333;line-height:1.7">Le Foyer Rural d\'Isneauville a bien reçu votre demande d\'inscription.<br>'
      + '<strong>Votre inscription sera confirmée dès réception de votre règlement.</strong></p>';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f0f4f2;margin:0;padding:20px">'
    + '<div style="max-width:640px;margin:0 auto;background:white;border-radius:12px;overflow:hidden">'
    // En-tête
    + '<div style="background:#1a2e22;padding:20px 24px;text-align:center">'
    + '<img src="cid:logo_fri" alt="FRI" style="width:66px;height:auto;margin-bottom:10px;border-radius:8px;display:block;margin-left:auto;margin-right:auto">'
    + '<h1 style="color:white;margin:0;font-size:20px">' + (isPaid?'🎉 Inscription validée !':'📋 Demande d\'inscription enregistrée') + '</h1>'
    + '<p style="color:#52b788;margin:5px 0 0;font-size:13px">' + NOM_ASSO + ' — Saison 2026/2027</p></div>'
    + bandeau
    + '<div style="padding:22px 24px">'
    + '<p style="color:#333">Bonjour <strong>' + prenom + ' ' + nom + '</strong>,</p>'
    + paraIntro
    + htmlAttente
    + mentionFacture
    // N° dossier
    + '<div style="background:#d8f3dc;border:2px solid #52b788;border-radius:10px;padding:12px 18px;text-align:center;margin:14px 0">'
    + '<div style="font-size:11px;font-weight:700;color:#2d6a4f;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px">Numéro de dossier</div>'
    + '<div style="font-family:monospace;font-size:28px;font-weight:900;color:#1a2e22;letter-spacing:5px">' + code + '</div>'
    + '<div style="font-size:11px;color:#2d6a4f;margin-top:4px">Conservez ce numéro pour tout suivi de votre dossier</div></div>'

    // PARTIE 1 — Activités
    + '<h3 style="color:#1a2e22;font-size:15px;margin:18px 0 6px;padding-bottom:4px;border-bottom:2px solid #d8f3dc">1. Activités inscrites</h3>'
    + _htmlPartie1Activites(f, false)

    // PARTIE 2 — Aides / Pass / Avoirs
    + (f.totalDeductions>0
        ? '<h3 style="color:#1a2e22;font-size:15px;margin:18px 0 6px;padding-bottom:4px;border-bottom:2px solid #d8f3dc">2. Aides, pass et avoirs</h3>'
          + _htmlPartie2Aides(f, isPaid)
        : '')

    // PARTIE 3 — Mode de paiement
    + '<h3 style="color:#1a2e22;font-size:15px;margin:18px 0 6px;padding-bottom:4px;border-bottom:2px solid #d8f3dc">' + (f.totalDeductions>0?'3.':'2.') + ' Mode de règlement</h3>'
    + _htmlPartie3Paiement(f, rows, modeLabel, isPaid)

    // Santé + documents
    // Verifier si au moins un membre a un QS non envoye (certif requis ou non rempli)
    + (santeHtml ? (function(sh){
        var hasCertif = sh.indexOf('⚠️') >= 0;
        var titre = hasCertif
          ? '<div style="background:#c62828;border-radius:8px;padding:10px 14px;margin:16px 0">'
            + '<p style="color:white;font-weight:bold;font-size:14px;margin:0 0 4px">⚠️ Documents à apporter impérativement</p>'
            + '<p style="color:rgba(255,255,255,0.85);font-size:12px;margin:0">Sans ces documents, votre inscription ne pourra pas être finalisée.</p>'
            + '</div>'
          : '<h3 style="color:#1a2e22;font-size:15px;margin:18px 0 6px;padding-bottom:4px;border-bottom:2px solid #d8f3dc">📋 Documents</h3>';
        return titre + '<ul style="color:#555;line-height:2;margin-top:8px">' + sh + '</ul>';
      })(santeHtml) : '')
    + ffttBloc

    + '<div style="background:#e8f4fd;border-left:4px solid #2980b9;border-radius:6px;padding:12px;margin-top:16px;font-size:13px;line-height:1.7">'
    + '<strong>Permanences</strong> : tous les <strong>mardis de 16h30 à 18h30</strong> <em>(période scolaire)</em><br>'
    + 'Salle des fêtes — Place A. Cramilly, 76230 Isneauville</div>'
    + '<p style="margin-top:14px;font-size:13px;color:#555">Contact : <a href="mailto:frisneauville@orange.fr" style="color:#2d6a4f">frisneauville@orange.fr</a> — <a href="tel:0235590101" style="color:#2d6a4f">02.35.59.01.01</a></p>'
    + '</div>'
    + '<div style="background:#f5f5f5;padding:10px;text-align:center;font-size:11px;color:#aaa">' + NOM_ASSO + ' — www.frisneauville.fr</div>'
    + '</div></body></html>';

  var sujet = isPaid
    ? 'Inscription validee FRI 2026/2027 - N° ' + code + ' - ' + prenom + ' ' + nom
    : 'Dossier enregistre FRI 2026/2027 - N° ' + code + ' - En attente reglement';
  var logoB = getLogoBlob();
  var opts = {htmlBody:html, charset:'UTF-8', name:NOM_ASSO, replyTo:EMAIL_ADMIN, charset:'UTF-8', inlineImages:logoB?{logo_fri:logoB}:{}};
  if (isPaid&&(pdfBlob||semainierBlob)) {
    var att=[]; if(pdfBlob)att.push(pdfBlob); if(semainierBlob)att.push(semainierBlob);
    opts.attachments = att;
  }
  envoyerEmail(email, sujet,
    'Bonjour '+prenom+' '+nom+',\n\n'+(isPaid?'Votre inscription est validée.':'Votre demande est enregistrée.')
    +'\nTotal : '+f.totalActivites.toFixed(2)+' €'+(f.totalDeductions>0?'\nSolde : '+f.solde.toFixed(2)+' €':'')
    +'\n\nContact : frisneauville@orange.fr | 02.35.59.01.01', opts);
  Logger.log('Email adhérent envoyé : '+email+(pdfBlob?' + PDF joint':''));
}

// ============================================================
// EMAIL ADMIN — v8.82 restructuré 3 parties
// ============================================================
function envoyerEmailAdmin(emailAdmin, rows, isPaid, modeLabel) {
  if (isPaid===undefined) isPaid=true;
  if (!emailAdmin||emailAdmin.indexOf('@')<0) return;

  var r0           = rows[0];
  var responsable  = (r0.responsable_prenom||'')+' '+(r0.responsable_nom||'');
  var code         = r0.code_dossier||'N/A';
  var certifReq    = rows.some(function(r){ return r.qs_sante==='Certificat requis'; });
  var f            = _calcFinancier(rows);
  if (!modeLabel) {
    var mlMap2={helloasso:'HelloAsso 💳',cheque:'Chèque 📝',cheque3:'Paiement 3 chèques 📝',especes:'Espèces 💵'};
    modeLabel = mlMap2[r0.mode_paiement||'helloasso'] || (r0.mode_paiement||'HelloAsso');
  }

  var alerteSante = certifReq
    ? '<div style="background:#ffeeed;border-left:5px solid #c0392b;padding:10px 14px;margin:10px 0;color:#c0392b;font-weight:bold;font-size:13px">⚠️ CERTIFICAT MÉDICAL REQUIS pour ce dossier</div>'
    : '<div style="background:#d8f3dc;border-left:5px solid #2d6a4f;padding:10px 14px;margin:10px 0;color:#2d6a4f;font-size:13px">✅ Attestation santé signée — toutes réponses négatives</div>';

  var bandeau = '<div style="background:' + (isPaid?'#d8f3dc':'#fff8e1') + ';padding:10px 16px;font-weight:bold;color:'
    + (isPaid?'#2d6a4f':'#856404') + '">' + (isPaid?'✅ ':'⏳ ') + modeLabel + ' — ' + f.totalActivites.toFixed(2) + ' €'
    + (f.totalDeductions>0?' (solde : '+f.solde.toFixed(2)+' €)':'') + '</div>';

  var infoCard = '<table style="width:100%;font-size:13px;border-collapse:collapse;margin-bottom:14px">'
    + '<tr><td style="padding:5px 0;color:#888;width:38%">N° dossier</td>'
    + '<td style="font-weight:bold;font-family:monospace;font-size:15px;color:#1565c0;letter-spacing:3px">' + code + '</td></tr>'
    + '<tr><td style="padding:5px 0;color:#888">Responsable</td><td style="font-weight:bold">' + responsable + '</td></tr>'
    + '<tr><td style="padding:5px 0;color:#888">Email</td><td>' + (r0.email1||'—') + '</td></tr>'
    + '<tr><td style="padding:5px 0;color:#888">Téléphone</td><td>' + (r0.tel1||'—') + '</td></tr>'
    + '<tr><td style="padding:5px 0;color:#888">Date inscription</td><td>' + (r0.date||'—') + '</td></tr>'
    + '</table>';

  // Détail texte pour fallback
  var detailText = rows.map(function(r){
    var isWait = String(r.statut_inscription||'').toLowerCase().indexOf('attente')>=0;
    return '- '+(r.membre_prenom||'')+' '+(r.membre_nom||'')+' : '+(r.activite||'').replace(/\n/g,' — ')
      +' ('+r.jour+' '+r.heure+') — '+(isWait?'Liste d\'attente':((parseFloat(r.tarif_brut)||parseFloat(r.tarif)||0)+' €'));
  }).join('\n');

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f0f4f2;margin:0;padding:16px">'
    + '<div style="max-width:620px;margin:0 auto;background:white;border-radius:12px;overflow:hidden">'
    // En-tête
    + '<div style="background:#1a2e22;padding:16px 22px;text-align:center">'
    + '<img src="cid:logo_fri" alt="FRI" style="width:52px;height:auto;border-radius:6px;display:block;margin:0 auto 8px">'
    + '<h1 style="color:#52b788;margin:0;font-size:17px">[ADMIN] ' + (isPaid?'Inscription validée':'Nouvelle demande en attente') + '</h1>'
    + '<p style="color:rgba(255,255,255,.5);margin:3px 0 0;font-size:12px">' + NOM_ASSO + ' — Dossier ' + code + '</p></div>'
    + bandeau
    + '<div style="padding:18px 20px">'
    + infoCard
    + alerteSante

    // PARTIE 1 — Activités
    + '<h3 style="color:#1a2e22;font-size:14px;margin:14px 0 5px;padding-bottom:3px;border-bottom:2px solid #d8f3dc">1. Activités</h3>'
    + _htmlPartie1Activites(f, true)

    // PARTIE 2 — Aides
    + (f.totalDeductions>0
        ? '<h3 style="color:#1a2e22;font-size:14px;margin:14px 0 5px;padding-bottom:3px;border-bottom:2px solid #d8f3dc">2. Aides / Pass / Avoirs</h3>'
          + _htmlPartie2Aides(f, isPaid)
        : '')

    // PARTIE 3 — Mode de paiement
    + '<h3 style="color:#1a2e22;font-size:14px;margin:14px 0 5px;padding-bottom:3px;border-bottom:2px solid #d8f3dc">'+(f.totalDeductions>0?'3.':'2.')+'  Mode de paiement</h3>'
    + _htmlPartie3Paiement(f, rows, modeLabel, isPaid)

    + '</div></div></body></html>';

  var logoBA = getLogoBlob();
  var sujetAdmin = '[FRI] N.' + code + ' - ' + responsable + ' - ' + f.totalActivites.toFixed(2) + ' EUR'
    + (f.totalDeductions>0?' (solde '+f.solde.toFixed(2)+' EUR)':'') + ' - ' + modeLabel;
  envoyerEmail(emailAdmin, sujetAdmin, detailText,
    {htmlBody:html, charset:'UTF-8', name:'Site FRI Inscriptions', charset:'UTF-8', inlineImages:logoBA?{logo_fri:logoBA}:{}});
  Logger.log('Email admin envoyé : '+emailAdmin+' — total: '+f.totalActivites+' solde: '+f.solde);
}

// ============================================================
// VÉRIFIER SEUIL 100 DOSSIERS
// ============================================================
function verifierSeuilCentDossiers(ss) {
  var sheet=ss.getSheetByName(SHEET_INSCRIPTIONS);if(!sheet||sheet.getLastRow()<2)return;
  var data=sheet.getRange(2,20,sheet.getLastRow()-1,1).getValues();var codes={};
  data.forEach(function(r){var c=String(r[0]||'').trim();if(c.match(/^FRI-[A-Z0-9]{4}$/))codes[c]=true;});
  var nb=Object.keys(codes).length;if(nb>0&&nb%100===0)envoyerExportCent(ss,nb);
}
function envoyerExportCent(ss,nb){
  try{var date=Utilities.formatDate(new Date(),'Europe/Paris','yyyy-MM-dd');var nom='Inscriptions_FRI_2026-2027_'+nb+'dossiers_'+date+'.xlsx';// Copie spreadsheet sans UrlFetchApp
    var ssCap=SpreadsheetApp.openById(SHEET_ID);var copieCap=ssCap.copy(nom.replace('.xlsx',''));
    var fCap=DriveApp.getFileById(copieCap.getId());
    envoyerEmail(EMAIL_ADMIN,'[FRI] '+nb+' dossiers - export inscriptions','Le cap des '+nb+' dossiers a été atteint le '+date+'.\nLien : '+copieCap.getUrl(),{name:NOM_ASSO});fCap.setTrashed(true);}catch(e){Logger.log('Export 100 dossiers KO : '+e);}
}

// ============================================================
// EN-TÊTE FEUILLE INSCRIPTIONS
// ============================================================
function ecrireEnTeteInscriptions(sheet) {
  var h=[
    'Numéro de licence','Civilité','Nom','Prénom','Date de naissance',
    'Appartement - Etage','Batiment - Résidence','N° et nom de voie',
    'Lieu-dit ou boîte postale','Code postal','Ville','Cedex','Code pays',
    'Tél. fixe','Tél. portable','Adresse e-mail','Commentaire','Tél. 2','Email 2',
    'N° Dossier','Date inscription','Statut paiement','Activité','Jour','Heure',
    'Lieu','Animateur',
    'Tarif brut EUR',      // col 28 AB
    'Éligible remise',     // col 29 AC — Oui/Non
    'Tarif net EUR',       // col 30 AD — formule =SI(nbElig>=3 ET éligible, AB*0.85, AB)
    'FNSMR EUR',           // col 31 AE
    'Total famille EUR',   // col 32 AF ← PIVOT
    'Mode paiement',       // col 33 AG
    'Avoir',               // col 34 AH
    'QS Santé',            // col 35 AI
    'Pass / Aide',         // col 36 AJ
    'Sexe',                // col 37 AK
    'ID Activité',         // col 38 AL
    'Responsable',         // col 39 AM
    'Statut inscription',  // col 40 AN
    'Licence FFTT (€)'     // col 41 AO
  ];
  var r=sheet.getRange(1,1,1,h.length);r.setValues([h]);
  r.setBackground('#2d6a4f').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(1,1,1,19).setBackground('#1b5e20');
  sheet.getRange(1,20).setBackground('#1565c0').setFontColor('#ffffff');
  sheet.getRange(1,29).setBackground('#e8f5e9').setFontColor('#1b5e20').setFontWeight('bold'); // Éligible
  sheet.getRange(1,30).setBackground('#d8f3dc').setFontColor('#1b5e20').setFontWeight('bold'); // Tarif net
  sheet.getRange(1,34).setBackground('#e65100').setFontColor('#ffffff');                        // Avoir
  sheet.getRange(1,40).setBackground('#4a148c').setFontColor('#ffffff');                        // Statut inscription
  sheet.getRange(1,41).setBackground('#bf360c').setFontColor('#ffffff');                        // FFTT
  sheet.setFrozenRows(1);sheet.autoResizeColumns(1,h.length);
}

// ============================================================
// GESTION DES PLACES
// ============================================================
function getOrCreatePlacesSheet(ss) {
  var sheet=ss.getSheetByName(SHEET_PLACES);
  if(!sheet){sheet=ss.insertSheet(SHEET_PLACES);var h=['ID Activite','Nom activite','Capacite max','Inscrits','Places restantes'];sheet.getRange(1,1,1,h.length).setValues([h]).setBackground('#c9a84c').setFontColor('#1a2e22').setFontWeight('bold');sheet.setFrozenRows(1);initAllActivities(sheet);}
  return sheet;
}

// ============================================================
// HELPER : lire les lignes actives restantes d'un dossier
// Retourne { rows, totalNet, totalBrut, aRemise, nbEligibles }
// ─ ss         : SpreadsheetApp instance
// ─ code       : code dossier
// ─ excluePid  : placesId à exclure (déjà supprimé ou à supprimer)
// ============================================================
function lireLignesRestantes(ss, code, excluePid) {
  // Utilise la même logique que calcTotalFamille pour cohérence Sheet/email/admin
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (!sheet || sheet.getLastRow() < 2) return {
    rows:[], totalNet:0, totalBrut:0, aRemise:false, nbEligibles:0,
    fnsmr:0, fftt:0, fnsmrDetail:{}, ffttDetail:{}
  };
  SpreadsheetApp.flush();
  var nbCols = Math.min(sheet.getLastColumn(), 41);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, nbCols).getValues();

  var rows = [], nbElig = 0, sumAB = 0, sumAD = 0;
  var membresVus = {}, ffttVus = {}, fnsmrTotal = 0, ffttTotal = 0;
  var fnsmrDetail = {}, ffttDetail = {};

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][19] || '').trim() !== code) continue;
    // col 40 AN = statut inscription (index 39)
    var statut = lireStatutInscription(data[i]).toLowerCase();
    if (statut.indexOf('supprim') >= 0) continue;
    if (statut.indexOf('attente') >= 0) continue;

    var pid     = lireActiviteId(data[i]);
    if (excluePid && pid === excluePid) continue;

    var ab      = Number(data[i][27] !== undefined ? data[i][27] : 0) || 0;  // col 28 AB = tarif brut
    var ac      = Number(data[i][28] !== undefined ? data[i][28] : 0) || 0;  // col 29 AC = éligible (1/0)
    var ad      = Math.round(ab * (1 - 0.15 * ac) * 100) / 100;              // calculer comme calcTotalFamille
    var ffttVal = 0; // FFTT supprimé // col 41 AO ou 39 selon structure

    var memKey = String(data[i][3] || '') + ' ' + String(data[i][2] || '');

    if (ac === 1) nbElig++;
    sumAB += ab;
    sumAD += ad;

    if (!membresVus[memKey]) {
      membresVus[memKey] = true;
      fnsmrTotal += 15;
      fnsmrDetail[memKey] = 15;
    }
    if (ffttVal > 0 && !ffttVus[memKey]) {
      ffttVus[memKey] = true;
      ffttTotal += ffttVal;
      ffttDetail[memKey] = ffttVal;
    }

    rows.push({
      pid:       pid,
      tarifBrut: ab,
      tarifNet:  ad,
      elig:      ac === 1,
      activite:  String(data[i][22] || '').replace(/\n/g, ' — '),
      membre:    String(data[i][3] || '') + ' ' + String(data[i][2] || ''),
      jour:      String(data[i][23] || ''),
      heure:     String(data[i][24] || '')
    });
  }

  var aRemise   = nbElig >= 3;
  // Même logique que calcTotalFamille
  var totalActif = aRemise ? sumAD : sumAB;
  var totalNet   = Math.round((totalActif + fnsmrTotal + ffttTotal) * 100) / 100;
  var totalBrut  = Math.round((sumAB      + fnsmrTotal + ffttTotal) * 100) / 100;

  Logger.log('lireLignesRestantes — code:'+code+' nbElig:'+nbElig+' aRemise:'+aRemise
    +' sumAD:'+sumAD+' sumAB:'+sumAB+' fnsmr:'+fnsmrTotal+' fftt:'+ffttTotal+' total:'+totalNet);

  return {
    rows: rows, totalNet: totalNet, totalBrut: totalBrut,
    aRemise: aRemise, nbEligibles: nbElig,
    fnsmr: fnsmrTotal, fftt: ffttTotal,
    fnsmrDetail: fnsmrDetail, ffttDetail: ffttDetail
  };
}

// ============================================================
// HELPER : envoyer email suppression activité
// Utilisé par supprimerActiviteDossierSheet ET supprimerActiviteNonRegleSheet
// ============================================================
function envoyerEmailSuppressionActivite(params) {
  var code         = params.code;
  var actNomClean  = params.actNomClean;
  var emailAdherent= params.emailAdherent;
  var responsable  = params.responsable;
  var estRegle     = params.estRegle;
  var avoir        = params.avoir || 0;
  var lignes       = params.lignesRestantes || [];
  var totalApres   = params.totalApres || 0;
  var aRemiseApres = params.aRemiseApres || false;
  var fnsmr        = params.fnsmr || 0;   // total FNSMR (15€ × nb membres)
  var fftt         = params.fftt  || 0;   // total licences FFTT
  var fnsmrDetail  = params.fnsmrDetail || {}; // { "Prénom NOM": 15 }
  var ffttDetail   = params.ffttDetail  || {}; // { "Prénom NOM": 70 }
  var dateJour     = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');

  // ── Corps texte ──
  var lignesTexte = lignes.map(function(r) {
    var tarif = (aRemiseApres && r.elig) ? Math.round(r.tarifBrut * 0.85 * 100) / 100 : r.tarifBrut;
    var remiseLabel = (aRemiseApres && r.elig) ? ' (remise 15% → ' + tarif.toFixed(2) + ' €)' : '';
    return '  • ' + r.membre + ' — ' + r.activite + ' : ' + r.tarifBrut.toFixed(2) + ' €' + remiseLabel;
  });
  // Ajouter FNSMR par membre
  Object.keys(fnsmrDetail).forEach(function(mem) {
    lignesTexte.push('  • ' + mem + ' — Adhésion FNSMR : ' + Number(fnsmrDetail[mem]).toFixed(2) + ' €');
  });
  // Ajouter FFTT par membre
  Object.keys(ffttDetail).forEach(function(mem) {
    if (Number(ffttDetail[mem]) > 0)
      lignesTexte.push('  • ' + mem + ' — Licence FFTT : ' + Number(ffttDetail[mem]).toFixed(2) + ' €');
  });

  var remiseNote = aRemiseApres
    ? '\n\nVotre dossier bénéficie toujours de la remise famille 15% (≥3 activités éligibles).'
    : (lignes.length >= 2 ? '\n\nAttention : votre dossier ne bénéficie plus de la remise famille 15% (moins de 3 activités éligibles). Les tarifs sont revenus au tarif plein.' : '');

  var avoirTexte = estRegle
    ? (avoir > 0
        ? '\n\nUn avoir de ' + avoir.toFixed(2) + ' € a été enregistré. Il sera déduit de votre prochain règlement ou remboursé sur demande.'
        : '\n\nAucun avoir à enregistrer.')
    : '';

  var bodyTexte = 'Bonjour ' + responsable + ',\n\n'
    + 'L\'activité "' + actNomClean + '" a été retirée de votre dossier N°' + code + (estRegle ? ' (dossier réglé)' : '') + '.\n\n'
    + (lignes.length > 0
        ? 'Détail de votre dossier :\n' + lignesTexte.join('\n') + '\n\nNouveau montant total : ' + totalApres.toFixed(2) + ' €'
        : 'Il ne reste plus d\'activité dans votre dossier.')
    + remiseNote + avoirTexte
    + '\n\nContact : frisneauville@orange.fr | 02.35.59.01.01\nCordialement,\nLe Foyer Rural d\'Isneauville';

  // ── Lignes activités HTML ──
  var lignesHtml = lignes.map(function(r, i) {
    var tarif = (aRemiseApres && r.elig) ? Math.round(r.tarifBrut * 0.85 * 100) / 100 : r.tarifBrut;
    var remiseHtml = (aRemiseApres && r.elig)
      ? '<br><small style="color:#2d6a4f">Remise 15% → ' + tarif.toFixed(2) + ' €</small>' : '';
    var bg = i % 2 === 0 ? '#f0f7f3' : '#ffffff';
    return '<tr style="background:' + bg + '">'
      + '<td style="padding:8px 12px">' + r.membre + '</td>'
      + '<td style="padding:8px 12px">' + r.activite + '</td>'
      + '<td style="padding:8px 12px;white-space:nowrap">' + r.jour + ' ' + r.heure + '</td>'
      + '<td style="padding:8px 12px;text-align:right">' + r.tarifBrut.toFixed(2) + ' €' + remiseHtml + '</td>'
      + '</tr>';
  }).join('');

  // ── Lignes FNSMR HTML (une ligne par membre) ──
  var fnsmrHtml = Object.keys(fnsmrDetail).map(function(mem) {
    return '<tr style="background:#e8f5e9">'
      + '<td style="padding:8px 12px;color:#1b5e20">' + mem + '</td>'
      + '<td style="padding:8px 12px;color:#1b5e20">Adhésion FNSMR</td>'
      + '<td style="padding:8px 12px;white-space:nowrap;color:#1b5e20">—</td>'
      + '<td style="padding:8px 12px;text-align:right;color:#1b5e20;font-weight:bold">' + Number(fnsmrDetail[mem]).toFixed(2) + ' €</td>'
      + '</tr>';
  }).join('');

  // ── Lignes FFTT HTML (une ligne par membre si > 0) ──
  var ffttHtml = Object.keys(ffttDetail).filter(function(mem){ return Number(ffttDetail[mem]) > 0; }).map(function(mem) {
    return '<tr style="background:#e3f2fd">'
      + '<td style="padding:8px 12px;color:#1565c0">' + mem + '</td>'
      + '<td style="padding:8px 12px;color:#1565c0">Licence FFTT</td>'
      + '<td style="padding:8px 12px;white-space:nowrap;color:#1565c0">—</td>'
      + '<td style="padding:8px 12px;text-align:right;color:#1565c0;font-weight:bold">' + Number(ffttDetail[mem]).toFixed(2) + ' €</td>'
      + '</tr>';
  }).join('');

  var remiseNoteHtml = aRemiseApres
    ? '<div style="margin:10px 0;padding:8px 14px;background:#d8f3dc;border-radius:4px;color:#1b5e20">✅ Remise famille 15% toujours active.</div>'
    : (lignes.length >= 2
        ? '<div style="margin:10px 0;padding:8px 14px;background:#fff8e1;border-radius:4px;color:#e65100">⚠️ La remise famille 15% ne s\'applique plus. Les tarifs sont revenus au tarif plein.</div>'
        : '');

  var avoirHtml = '';
  if (estRegle) {
    avoirHtml = avoir > 0
      ? '<div style="margin:16px 0;padding:12px 16px;background:#fff3e0;border-left:4px solid #e65100;border-radius:4px">'
        + '💳 <strong>Avoir enregistré : ' + avoir.toFixed(2) + ' €</strong><br>'
        + '<small>Sera déduit de votre prochain règlement ou remboursé sur demande.</small></div>'
      : '<div style="margin:12px 0;padding:10px 14px;background:#f5f5f5;border-radius:4px">Aucun avoir à enregistrer.</div>';
  }

  var tableHtml = lignes.length > 0
    ? '<table width="100%" style="border-collapse:collapse;margin:12px 0">'
      + '<thead><tr style="background:#2d6a4f;color:white">'
      + '<th style="padding:8px 12px;text-align:left">Membre</th>'
      + '<th style="padding:8px 12px;text-align:left">Activité</th>'
      + '<th style="padding:8px 12px;text-align:left">Horaire</th>'
      + '<th style="padding:8px 12px;text-align:right">Tarif</th>'
      + '</tr></thead><tbody>'
      + lignesHtml + fnsmrHtml + ffttHtml
      + '</tbody></table>'
      + '<p style="text-align:right;font-weight:bold;font-size:16px;margin-top:4px">Total : ' + totalApres.toFixed(2) + ' €</p>'
    : '<p><em>Il ne reste plus d\'activité dans votre dossier.</em></p>';

  var bodyHtml = '<div style="font-family:Arial,sans-serif;max-width:620px;margin:auto">'
    + '<div style="background:#2d6a4f;color:white;padding:18px 24px;border-radius:8px 8px 0 0">'
    + '<h2 style="margin:0">Foyer Rural d\'Isneauville</h2>'
    + '<p style="margin:4px 0;opacity:.85">Modification de dossier — N°' + code + '</p></div>'
    + '<div style="padding:20px 24px;background:#f9fafb;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px">'
    + '<p>Bonjour <strong>' + responsable + '</strong>,</p>'
    + '<p>L\'activité <strong>"' + actNomClean + '"</strong> a été retirée de votre dossier'
    + (estRegle ? ' <em>(dossier réglé)</em>' : '') + '.</p>'
    + remiseNoteHtml + tableHtml + avoirHtml
    + '<hr style="border:none;border-top:1px solid #e0e0e0;margin:16px 0">'
    + '<small style="color:#888">Contact : frisneauville@orange.fr | 02.35.59.01.01</small></div></div>';

  // ── Envoi adhérent ──
  try {
    if (emailAdherent && emailAdherent.indexOf('@') > 0) {
      envoyerEmail(emailAdherent,
        '[FRI] Modification dossier N.' + code + ' — ' + actNomClean + ' supprimée',
        bodyTexte,
        { name: 'Foyer Rural d\'Isneauville', replyTo: EMAIL_ADMIN, htmlBody: bodyHtml });
      Logger.log('✅ Email suppression → adhérent : ' + emailAdherent);
    }
  } catch(e) { Logger.log('❌ Email adhérent KO : ' + e.toString()); }

  // ── Envoi admin ──
  try {
    var adminLines = lignes.map(function(r) {
      var t = (aRemiseApres && r.elig) ? Math.round(r.tarifBrut * 0.85 * 100)/100 : r.tarifBrut;
      return '  • ' + r.membre + ' — ' + r.activite + ' : ' + t.toFixed(2) + ' €';
    });
    Object.keys(fnsmrDetail).forEach(function(m){ adminLines.push('  • ' + m + ' — FNSMR : ' + Number(fnsmrDetail[m]).toFixed(2) + ' €'); });
    Object.keys(ffttDetail).forEach(function(m){ if(Number(ffttDetail[m])>0) adminLines.push('  • ' + m + ' — FFTT : ' + Number(ffttDetail[m]).toFixed(2) + ' €'); });
    envoyerEmail(EMAIL_ADMIN,
      '[FRI Admin] Suppression ' + (estRegle ? 'réglée' : 'non réglée') + ' — N°' + code + ' — ' + actNomClean,
      '🗑 ' + actNomClean + ' supprimée\nDossier : ' + code + (estRegle ? ' (RÉGLÉ)' : ' (non réglé)')
        + '\nDate : ' + dateJour
        + (estRegle && avoir > 0 ? '\n💳 Avoir : ' + avoir.toFixed(2) + ' €' : '')
        + '\nRemise famille après : ' + (aRemiseApres ? 'OUI' : 'NON')
        + '\nDétail :\n' + (adminLines.join('\n') || '  (aucune activité restante)')
        + '\nTotal dossier : ' + totalApres.toFixed(2) + ' €',
      { name: 'FRI Admin' });
    Logger.log('✅ Email suppression → admin');
  } catch(e) { Logger.log('❌ Email admin KO : ' + e.toString()); }
}

function getAllActivityIds(){
  var NOMS = {
    'FITL1730': 'Body Sculpt — Lundi 17h30',
    'FITL1930': 'Circuit training adapté — Lundi 19h30',
    'FITL2030': 'Zumba — Lundi 20h30',
    'FITM10': 'Body Sculpt — Mardi 10h00',
    'FITV10': 'Circuit training adapté — Vendredi 10h00',
    'FITV11': 'Stretching — Vendredi 11h00',
    'FITV1830': 'Body Sculpt — Vendredi 18h30',
    'FITS0945': 'Zumba — Samedi 09h45',
    'FITS1045': 'Body Sculpt — Samedi 10h45',
    'FITS1145': 'Stretching — Samedi 11h45',
    'YOGAL18': 'Yoga — Lundi 18h00',
    'MEDL1930': 'Méditation — Lundi 19h30',
    'YOGAM11': 'Yoga — Mardi 11h00',
    'YOGAJ1615': 'Stretching Yoga — Jeudi 16h15',
    'YOGAJ1835': 'Yoga — Jeudi 18h35',
    'YOGAJ2005': 'Yoga — Jeudi 20h05',
    'PILL11': 'Pilates — Lundi 11h00',
    'PILL12': 'Pilates — Lundi 12h00',
    'PILL14': 'Pilates — Lundi 14h00',
    'PILL1830': 'Pilates Flow — Lundi 18h30',
    'PILM1830': 'Pilates Flow — Mardi 18h30',
    'PILME10': 'Pilates — Mercredi 10h00',
    'PILME11': 'Pilates Séniors — Mercredi 11h00',
    'PILJ11': 'Pilates — Jeudi 11h00',
    'PILJ12': 'Pilates — Jeudi 12h00',
    'PILJ1515': 'Pilates — Jeudi 15h15',
    'PILV14': 'Pilates — Vendredi 14h00',
    'PINGL1945/ME1945': 'Tennis de table adultes compétiteurs (Lu+Me)',
    'PINGM1930/ME21': 'Tennis de table libre (Ma+Me)',
    'PINGJ1530': 'Ping bien-être / Pong santé — Jeudi 15h30',
    'SOPHM11': 'Sophrologie — Mardi 11h00',
    'SOPHME1830': 'Sophrologie — Mercredi 18h30',
    'MNOME/S10': 'Marche Nordique — Mercredi/Samedi 10h00',
    'APAL15': 'Sport Santé / APA — Lundi 15h00',
    'APAJ14': 'Sport Santé / APA — Jeudi 14h00',
    'APAV1730': 'Sport Santé / APA — Vendredi 17h30',
    'YOGAM1430': 'Yoga doux / chaise — Mardi 14h30',
    'PEINL1545': 'Atelier Peinture — Lundi 15h45',
    'PEINL1845': 'Atelier Peinture — Lundi 18h45',
    'GUITJ19': 'Atelier Guitare / Chant — Jeudi 19h00',
    'THEL2030': 'Théâtre Adultes création — Lundi 20h30',
    'THEJ19': 'Théâtre Adultes loisirs — Jeudi 19h00',
    'COUTJ0930': 'Couture — Jeudi 09h30',
    'COUTJ13': 'Couture — Jeudi 13h00',
    'COUTJ1800': 'Couture — Jeudi 18h00',
    'COUNV1830': 'Country Initiation — Vendredi 18h30',
    'COUNV1930': 'Country Débutants — Vendredi 19h30',
    'COUNV2030': 'Country Novices — Vendredi 20h30',
    'COUNV2130': 'Country Intermédiaires — Vendredi 21h30',
    'JAZJ1845': 'Danse Moderne Avancés 1 — Jeudi 18h45',
    'JAZJ2015': 'Danse Moderne Avancés 2 — Jeudi 20h15',
    'CLASSL1845': 'Danse Classique Inter Ados/Adultes — Lundi 18h45',
    'CLASSL21': 'Danse Classique Pointes — Lundi 20h00',
    'CLASSJ1215': 'Danse Classique Barre à terre — Jeudi 12h15',
    'CLASSJ1315': 'Danse Classique Débutants/Inters — Jeudi 13h15',
    'CLASSJ1430': 'Danse Classique Pointes — Jeudi 14h30',
    'PINGL18/J17': 'Tennis de table Jeunes Compétiteurs (Lu+Je)',
    'PINGME1730': 'Tennis de table Jeunes Débutants — Mercredi 17h15',
    'PINGL1745ME1830': 'Tennis de table Jeunes Confirmés (Lu+Me)',
    'GYMME10': 'Baby Gym Parents — Mercredi 10h00',
    'GYMME1045': 'Gym Parent-Enfant — 3-5 ans — Mercredi 10h45',
    'GYMME1130': 'Gym — 3-5 ans — Mercredi 11h30',
    'GYMME1215': 'Gym — Inter/Avancés 1 — Mercredi 12h15',
    'GYMME1315': 'Gym Débutants — 6-8 ans (1) — Mercredi 13h15',
    'GYMME1415': 'Gym Débutants — 6-8 ans (2) — Mercredi 14h15',
    'GYMME1515': 'Gym — Inter/Avancés 2 — Mercredi 15h15',
    'JAZME1015': 'Danse Moderne CP-CE1 — Mercredi 10h15',
    'JAZME1115': 'Danse Moderne CE2-CM1 — Mercredi 11h15',
    'JAZME1315': 'Danse Moderne CM1-CM2 — Mercredi 13h15',
    'JAZME1415': 'Danse Moderne 6e-5e — Mercredi 14h15',
    'JAZME1515': 'Danse Moderne Déb.4/Inter.1 — Mercredi 15h15',
    'JAZME1615': 'Danse Moderne Inter 2&3 — Mercredi 16h15',
    'JAZME1715': 'Danse Moderne Inter 4&5 — Mercredi 17h15',
    'JAZJ1745': 'Danse Moderne Éveil Petits — Jeudi 17h45',
    'CLASSL1745': 'Danse Classique 11-13 ans — Lundi 17h45',
    'CLASSME0930': 'Danse Classique 4-5 ans — Mercredi 09h30',
    'CLASSME1015': 'Danse Classique 6-7 ans — Mercredi 10h15',
    'CLASSME11': 'Danse Classique 8-10 ans — Mercredi 11h00',
    'PEINL1715': 'Atelier Peinture Ados — Lundi 17h15',
    'PEINME0930': 'Atelier Peinture Enfants — Mercredi 09h30',
    'PEINME11': 'Atelier Peinture Enfants — Mercredi 11h00',
    'THEV1630': 'Théâtre Primaires CM1/CM2 — Vendredi 16h30',
    'THEV1730': 'Théâtre Collégiens 6e/5e — Vendredi 17h30',
    'THEV1830': 'Théâtre 4e, 3e et Lycéens — Vendredi 18h30',
  };
  var ids = [
    'FITL1730',
    'FITL1930',
    'FITL2030',
    'FITM10',
    'FITV10',
    'FITV11',
    'FITV1830',
    'FITS0945',
    'FITS1045',
    'FITS1145',
    'YOGAL18',
    'MEDL1930',
    'YOGAM11',
    'YOGAJ1615',
    'YOGAJ1835',
    'YOGAJ2005',
    'PILL11',
    'PILL12',
    'PILL14',
    'PILL1830',
    'PILM1830',
    'PILME10',
    'PILME11',
    'PILJ11',
    'PILJ12',
    'PILJ1515',
    'PILV14',
    'PINGL1945/ME1945',
    'PINGM1930/ME21',
    'PINGJ1530',
    'SOPHM11',
    'SOPHME1830',
    'MNOME/S10',
    'APAL15',
    'APAJ14',
    'APAV1730',
    'YOGAM1430',
    'PEINL1545',
    'PEINL1845',
    'GUITJ19',
    'THEL2030',
    'THEJ19',
    'COUTJ0930',
    'COUTJ13',
    'COUTJ1800',
    'COUNV1830',
    'COUNV1930',
    'COUNV2030',
    'COUNV2130',
    'JAZJ1845',
    'JAZJ2015',
    'CLASSL1845',
    'CLASSL21',
    'CLASSJ1215',
    'CLASSJ1315',
    'CLASSJ1430',
    'PINGL18/J17',
    'PINGME1730',
    'PINGL1745ME1830',
    'JAZME1215',
    'JAZME1315',
    'JAZME1415',
    'JAZME1515',
    'JAZME1615',
    'JAZME1715',
    'JAZJ1745',
    'CLASSL1745',
    'CLASSME0930',
    'CLASSME1015',
    'CLASSME11',
    'PEINL1715',
    'PEINME0930',
    'PEINME11',
    'GYMME10',
    'GYMME1045',
    'GYMME1130',
    'GYMME1215',
    'GYMME1315',
    'GYMME1415',
    'GYMME1515',
    'THEV1630',
    'THEV1730',
    'THEV1830'];
  return ids.map(function(id){ return [id, NOMS[id] || id]; });
}

function initAllActivities(sheet){
  // Capacités fixes par activité — source : tableau Places libres saison 2026/2027
  var CAPACITES = {
    'FITL1730':24, 'FITL1930':24, 'FITL2030':0,
    'FITM10':24,
    'FITV10':24, 'FITV11':24, 'FITV1830':40,
    'FITS0945':40, 'FITS1045':40, 'FITS1145':0,
    'YOGAL18':25, 'MEDL1930':8, 'YOGAM11':22,
    'YOGAJ1615':0, 'YOGAJ1835':18, 'YOGAJ2005':18,
    'PILL11':24, 'PILL12':24, 'PILL14':24,
    'PILL1830':24, 'PILM1830':24,
    'PILME10':24, 'PILME11':18,
    'PILJ11':24, 'PILJ12':24, 'PILJ1515':24, 'PILV14':24,
    'PINGL1945/ME1945':24, 'PINGM1930/ME21':24, 'PINGJ1530':10,
    'SOPHM11':8, 'SOPHME1830':10,
    'MNOME/S10':70,
    'APAL15':8, 'APAJ14':8, 'APAV1730':8,
    'YOGAM1430':20,
    'PEINL1545':8, 'PEINL1845':9,
    'GUITJ19':10,
    'THEL2030':12, 'THEJ19':12,
    'COUTJ0930':8, 'COUTJ13':8, 'COUTJ1800':8,
    'COUNV1830':55, 'COUNV1930':55, 'COUNV2030':55, 'COUNV2130':55,
    'JAZJ1845':16, 'JAZJ2015':16,
    'CLASSL1845':0, 'CLASSL21':0,
    'CLASSJ1215':0, 'CLASSJ1315':0, 'CLASSJ1430':0,
    'PINGL18/J17':20, 'PINGME1730':20, 'PINGL1745ME1830':20,
    'JAZME1215':15, 'JAZME1315':15, 'JAZME1415':15,
    'JAZME1515':15, 'JAZME1615':15, 'JAZME1715':15, 'JAZJ1745':15,
    'CLASSL1745':0, 'CLASSME0930':0, 'CLASSME1015':0, 'CLASSME11':0,
    'PEINL1715':8, 'PEINME0930':8, 'PEINME11':8,
    'THEV1630':10, 'THEV1730':10, 'THEV1830':10,
    'GYMME10':12, 'GYMME1045':24, 'GYMME1130':24,
    'GYMME1215':12, 'GYMME1315':16, 'GYMME1415':16, 'GYMME1515':16
  };
  var acts = getAllActivityIds();
  // Filtrer les capacités 0 (activités suspendues) et trier alphabétiquement par ID
  var rows = acts
    .map(function(a){
      var cap = (CAPACITES[a[0]] !== undefined) ? CAPACITES[a[0]] : DEFAULT_CAPACITY;
      return [a[0], a[1], cap, 0, cap];
    })
    .filter(function(r){ return r[2] > 0; })  // supprimer cap=0
    .sort(function(a, b){ return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; }); // tri alpha
  if(rows.length > 0){
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
    for(var i = 2; i <= rows.length + 1; i++){
      sheet.getRange(i, 5)
        .setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
    }
    sheet.autoResizeColumns(1, 5);
  }
}
function getPlacesId(actId){
  var m={'body-sculpt-lu':'FITL1730','zumba-lu':'FITL1930','body-sculpt-ma':'FITM10','circuit-ve':'FITV10','stretching-ve':'FITV11','body-sculpt-cm-ve':'FITV1830','zumba-sa':'FITS0945','body-sculpt-sa':'FITS1045','stretching-sa':'FITS1145','yoga-sm-lu':'YOGAL18','meditation-lu':'MEDL1930','yoga-mb-ma':'YOGAM11','stretching-yoga-je':'YOGAJ1615','yoga-sm-je1':'YOGAJ1835','yoga-sm-je2':'YOGAJ2005','pilates-lu1':'PILL11','pilates-lu2':'PILL12','pilates-lu3':'PILL14','pilates-flow-lu':'PILL1830','pilates-flow-ma':'PILM1830','pilates-me1':'PILME10','pilates-seniors-me':'PILME11','pilates-je1':'PILJ11','pilates-je2':'PILJ12','pilates-je3':'PILJ1515','pilates-ve':'PILV14','tt-comp-entrain':'PINGL1945/ME1945','tt-comp-entrain-me':'PINGL1945/ME1945','tt-comp-sans':'PINGM1930/ME21','tt-comp-sans-me':'PINGM1930/ME21','ping-bienetre':'PINGJ1530','sophro-ma':'SOPHM11','sophro-me':'SOPHME1830','marche-nordic':'MNOME/S10','apa-lu':'APAL15','apa-je':'APAJ14','apa-ve':'APAV1730','yoga-doux':'YOGAM1430','peinture-lu1':'PEINL1545','peinture-lu2':'PEINL1845','guitare-a':'GUITJ19','guitare-j':'GUITJ19','theatre-creation':'THEL2030','theatre-loisirs-a':'THEJ19','couture-je1':'COUTJ0930','couture-je2':'COUTJ13','couture-je3':'COUTJ1800','country-init-a':'COUNV1830','country-debu-a':'COUNV1930','country-novice':'COUNV2030','country-inter':'COUNV2130','country-init-j':'COUNV1830','country-debu-j':'COUNV1930','danse-moderne-av1':'JAZJ1845','danse-moderne-av2':'JAZJ2015','danse-classique-inter-a':'CLASSL1845','danse-classique-pointes-lu':'CLASSL21','danse-classique-barre':'CLASSJ1215','danse-classique-debu-je':'CLASSJ1315','danse-classique-pointes-je':'CLASSJ1430','tt-jeunes-comp':'PINGL18/J17','tt-jeunes-comp-je':'PINGL18/J17','tt-debu-1':'PINGME1730','tt-debu-2':'PINGL1745ME1830','tt-debu-2-me':'PINGL1745ME1830','gym-eveil':'GYMME10','gym-parent-3-5':'GYMME1045','gym-3-5':'GYMME1130','gym-inter-av1':'GYMME1215','gym-debu-6-8-1':'GYMME1315','gym-debu-6-8-2':'GYMME1415','gym-inter-av2':'GYMME1515','danse-moderne-init1':'JAZME1215','danse-moderne-debu12':'JAZME1315','danse-moderne-debu23':'JAZME1415','danse-moderne-debu4':'JAZME1515','danse-moderne-inter23':'JAZME1615','danse-moderne-inter45':'JAZME1715','danse-eveil':'JAZJ1745','danse-classique-11-13':'CLASSL1745','danse-classique-4-5':'CLASSME0930','danse-classique-6-7':'CLASSME1015','danse-classique-8-10':'CLASSME11','peinture-ados':'PEINL1715','peinture-enfants1':'PEINME0930','peinture-enfants2':'PEINME11','theatre-cm1cm2':'THEV1630','theatre-6e5e':'THEV1730','theatre-lycee':'THEV1830'};
  return m[actId]||actId;
}
function incrementPlaces(ss,rows){
  var sheet=getOrCreatePlacesSheet(ss);var lr=sheet.getLastRow();if(lr<2)return{};
  var data=sheet.getRange(2,1,lr-1,5).getValues();var counts={},seenPairs={};
  rows.forEach(function(r){
    if(!r.activite_id)return;
    // Ne pas compter les lignes déjà en liste d'attente côté frontend
    if(String(r.statut_inscription||'').toLowerCase().indexOf('attente')>=0)return;
    var placesId=getPlacesId(r.activite_id);
    if(r.activite_id!==placesId){var pairKey=(r.code_dossier||r.id||'')+'_'+placesId;if(seenPairs[pairKey])return;seenPairs[pairKey]=true;}
    counts[placesId]=(counts[placesId]||0)+1;
  });
  var statutMap={};
  for(var i=0;i<data.length;i++){var id=String(data[i][0]).trim();if(!counts[id])continue;var current=parseInt(data[i][3])||0;var capacity=parseInt(data[i][2])||20;var newTotal=current+counts[id];sheet.getRange(i+2,4).setValue(newTotal);sheet.getRange(i+2,5).setValue(Math.max(0,capacity-newTotal));for(var j=0;j<counts[id];j++){var rang=current+j+1;statutMap[id+'_'+j]=rang<=capacity?'Inscrit':'Liste attente #'+(rang-capacity);}}
  return statutMap;
}

// ============================================================
// ONGLETS ACTIVITÉS
// ============================================================

// ══════════════════════════════════════════════════════════════
// ONGLET LISTE D'ATTENTE : LA-{placesId}
// Une ligne par membre en attente, colonnes : N°Dossier, Date, Nom, Prénom, Tel, Email, Activité
// ══════════════════════════════════════════════════════════════
function majOngletListeAttente(ss, placesId, rows) {
  var nomOnglet = 'LA-' + placesId.replace(/[\/\\:?*\[\]]/g, '-').substring(0, 28);
  var sheet = ss.getSheetByName(nomOnglet);
  if (!sheet) {
    sheet = ss.insertSheet(nomOnglet);
    var h = ['N° Dossier','Date','Statut','Nom','Prénom','Tél','Email','Activité','Jour','Heure'];
    sheet.getRange(1,1,1,h.length).setValues([h])
      .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('✅ Onglet liste attente créé : ' + nomOnglet);
  }
  var dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  rows.forEach(function(r) {
    // Vérifier que la ligne n'existe pas déjà
    var lr = sheet.getLastRow();
    if (lr > 1) {
      var existing = sheet.getRange(2, 1, lr-1, 2).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]||'').trim() === (r.code_dossier||'') &&
            String(existing[i][3]||'').trim() === (r.membre_nom||'')) return;
      }
    }
    var nextRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(nextRow, 1, 1, 10).setValues([[
      r.code_dossier||'', dateJour, 'En attente',
      r.membre_nom||'', r.membre_prenom||'',
      r.tel1||'', r.email1||'',
      (r.activite||'').replace(/\n/g,' — '),
      r.jour||'', r.heure||''
    ]]);
    sheet.getRange(nextRow, 1, 1, 10)
      .setBackground(nextRow % 2 === 0 ? '#fff3e0' : '#ffffff');
    sheet.getRange(nextRow, 3).setFontColor('#e65100').setFontWeight('bold');
    sheet.autoResizeColumns(1, 10);
  });
}

function majOngletsActivites(ss,rows,statut){
  rows.forEach(function(r){
    // Ne pas écrire dans l'onglet activité si liste d'attente de place
    if (String(r.statut_inscription||'').toLowerCase().indexOf('attente') >= 0) return;
    var placesId=getPlacesId(r.activite_id||'');var nomActivite=placesId||(r.activite||'Activite').replace(/[\\/:?*\[\]]/g,'').substring(0,28).trim();
    var sheet=ss.getSheetByName(nomActivite);
    if(!sheet){
      sheet=ss.insertSheet(nomActivite);var couleur=getCouleurActivite(r.activite_id||'');
      var isPingSheet=(placesId&&placesId.indexOf('PING')>=0)||(r.activite_id&&(r.activite_id.indexOf('tt-')===0||r.activite_id.indexOf('ping')===0));
      var h=['N° Dossier','Date','Statut','Nom','Prenom','Date naiss.','Sexe','Responsable','Telephone','Email','Ville','QS Sante','Paiement','Pass Aide'];
      if(isPingSheet)h.push('Licence FFTT (€)');
      var hr=sheet.getRange(1,1,1,h.length);hr.setValues([h]).setBackground(couleur).setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
      if(isPingSheet)sheet.getRange(1,15).setBackground('#e65100');
      sheet.setFrozenRows(1);sheet.getRange('A2').setValue(r.activite+' | '+r.animateur+' | '+r.jour+' '+r.heure+' | Salle: '+(r.lieu||'—'));
      sheet.getRange(2,1,1,h.length).setBackground('#f5f5f5').setFontStyle('italic').setFontSize(9).setFontColor('#555');sheet.setFrozenRows(2);
    }
    var nextRow=Math.max(sheet.getLastRow()+1,3);var bg=nextRow%2===0?'#f0f7f3':'#ffffff';var ffttPrice=parseFloat(r.fftt_price)||0;
    var isPingRow=(placesId&&placesId.indexOf('PING')>=0)||(r.activite_id&&(r.activite_id.indexOf('tt-')===0||r.activite_id.indexOf('ping')===0));
    var rowData=[r.code_dossier||'',r.date||'',statut,r.membre_nom||'',r.membre_prenom||'',r.ddn||'',r.sexe||'',(r.responsable_prenom||'')+' '+(r.responsable_nom||''),r.tel1||'',r.email1||'',r.ville||'',r.qs_sante||'',r.mode_paiement||'',r.pass_aide||''];
    if(isPingRow)rowData.push(ffttPrice>0?ffttPrice:'');
    sheet.getRange(nextRow,1,1,rowData.length).setValues([rowData]).setBackground(bg);
    sheet.getRange(nextRow,1).setFontColor('#1565c0').setFontWeight('bold');
    if(isPingRow&&ffttPrice>0)sheet.getRange(nextRow,15).setBackground('#fff3e0').setFontColor('#e65100').setFontWeight('bold');
    sheet.autoResizeColumns(1,rowData.length);
  });
}
function getCouleurActivite(actId){
  if(actId.indexOf('tt-')===0||actId.indexOf('ping')===0)return'#e65100';
  if(actId.indexOf('yoga')>=0||actId.indexOf('meditation')>=0)return'#1565c0';
  if(actId.indexOf('pilates')>=0)return'#6a1b9a';
  if(actId.indexOf('danse-moderne')>=0)return'#ad1457';
  if(actId.indexOf('danse-classique')>=0)return'#880e4f';
  if(actId.indexOf('gym')>=0||actId.indexOf('baby')>=0)return'#2e7d32';
  if(actId.indexOf('body')>=0||actId.indexOf('zumba')>=0||actId.indexOf('circuit')>=0||actId.indexOf('stretching')>=0)return'#558b2f';
  if(actId.indexOf('theatre')>=0)return'#4527a0';if(actId.indexOf('peinture')>=0)return'#bf360c';if(actId.indexOf('guitare')>=0)return'#4e342e';
  if(actId.indexOf('couture')>=0)return'#00695c';if(actId.indexOf('country')>=0)return'#f57f17';if(actId.indexOf('sophro')>=0)return'#37474f';
  if(actId.indexOf('marche')>=0)return'#1b5e20';if(actId.indexOf('apa')>=0)return'#004d40';
  return'#2d6a4f';
}
function majRecapitulatif(ss,rows,statut){
  var recap=ss.getSheetByName(SHEET_RECAPITULATIF);
  if(!recap){recap=ss.insertSheet(SHEET_RECAPITULATIF);var rh=['Date','ID','Statut','Responsable','Ville','Commune','Nb membres','Nb activites','Total EUR','Paiement','Email','Tel','Sante','Pass Aide'];recap.getRange(1,1,1,rh.length).setValues([rh]).setBackground('#1a2e22').setFontColor('#52b788').setFontWeight('bold');recap.setFrozenRows(1);}
  // Lire le total AF directement depuis le Sheet (source fiable)
  var totalParCode = {};
  var shInscr = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (shInscr && shInscr.getLastRow() > 1) {
    var dataInscr = shInscr.getRange(2, 1, shInscr.getLastRow()-1, 32).getValues();
    dataInscr.forEach(function(row) {
      var code = String(row[19]||'').trim(); // col T = N° dossier
      var af   = Number(row[31]||0);         // col AF (index 31) = Total famille
      if (code && af > 0) totalParCode[code] = af;
    });
  }

  var fam={};
  rows.forEach(function(r){var famKey=r.code_dossier||r.id||'';if(!fam[famKey]){fam[famKey]={date:r.date,id:famKey,statut:statut,responsable:r.responsable_nom+' '+r.responsable_prenom,ville:r.ville,commune:r.commune||'',membres:{},nbActs:0,total:totalParCode[famKey]||r.total_famille||0,paiement:r.mode_paiement||'helloasso',email:r.email1,tel:r.tel1,sante:[],pass:r.pass_aide};}fam[famKey].membres[r.membre_prenom+' '+r.membre_nom]=true;fam[famKey].nbActs++;if(r.qs_sante)fam[famKey].sante.push(r.membre_prenom+': '+r.qs_sante);});
  // Lire les lignes existantes pour éviter les doublons
  var existingData = recap.getLastRow() > 1
    ? recap.getRange(2, 1, recap.getLastRow()-1, 14).getValues()
    : [];
  var existingByCode = {};
  existingData.forEach(function(row, i) {
    var code = String(row[1]||'').trim();
    if (code) existingByCode[code] = i + 2; // numéro de ligne Sheet (base 1)
  });

  Object.keys(fam).forEach(function(k) {
    var f = fam[k];
    var rowData = [f.date, f.id, f.statut, f.responsable, f.ville, f.commune,
      Object.keys(f.membres).length, f.nbActs, f.total, f.paiement,
      f.email, f.tel, f.sante.join(' | '), f.pass];
    if (existingByCode[k]) {
      // Mettre à jour la ligne existante
      recap.getRange(existingByCode[k], 1, 1, 14).setValues([rowData]);
    } else {
      // Ajouter une nouvelle ligne
      recap.appendRow(rowData);
    }
  });
  recap.autoResizeColumns(1, 14);
}
function getOrCreateChequeSheet(ss, nom) {
  var sheet = ss.getSheetByName(nom);
  if (!sheet) {
    sheet = ss.insertSheet(nom);
    var couleurs = { 'Cheques 1': '#1b5e20', 'Cheques 2': '#0d47a1', 'Cheques 3': '#4a148c',
                     'Cheques LA': '#4e342e', 'Especes LA': '#4e342e' };
    var labels   = { 'Cheques 1': 'Chèques uniques / 1er chèque', 'Cheques 2': '2e chèque', 'Cheques 3': '3e chèque',
                     'Cheques LA': 'Chèque — Activité bassculée (liste d\'attente)', 'Especes LA': 'Espèces — Activité basculée (liste d\'attente)' };
    var couleur  = couleurs[nom] || '#1b5e20';
    var label    = labels[nom]   || nom;
    var h = ['N° Dossier', 'Nom', 'Prénom', 'Banque', 'N° Chèque', 'Montant (€)', 'Statut'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]).setBackground(couleur).setFontColor('#ffffff').setFontWeight('bold');
    var chqPastelMap = {'#1b5e20':'#f0fdf4','#0d47a1':'#eff6ff','#4a148c':'#faf5ff'};
    var bgLigne2Chq = chqPastelMap[couleur] || '#f5f5f5';
    sheet.getRange(2, 1, 1, h.length).mergeAcross().setValue(label + ' — Saison 2026/2027').setFontStyle('italic').setFontSize(9).setFontColor('#555555').setBackground(bgLigne2Chq);
    sheet.setFrozenRows(2);
    // Ligne TOTAL (ligne 3)
    sheet.getRange(3, 1, 1, h.length).setBackground('#e8f5e9');
    sheet.getRange(3, 5).setValue('TOTAL CHÈQUES').setFontWeight('bold').setFontColor('#1b5e20');
    sheet.getRange(3, 6).setValue(0).setFontWeight('bold').setFontColor('#1b5e20').setNumberFormat('#,##0.00 €');
    sheet.getRange(3, 7).setValue('Statut').setFontStyle('italic').setFontColor('#888888');
    sheet.setFrozenRows(3);
    sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 140); sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 160); sheet.setColumnWidth(5, 140); sheet.setColumnWidth(6, 130); sheet.setColumnWidth(7, 160);
  }
  return sheet;
}
function getOrCreateAvoirSheet(ss){
  var sheet=ss.getSheetByName(SHEET_AVOIRS);
  if(!sheet){
    sheet=ss.insertSheet(SHEET_AVOIRS);
    var h=['N° Dossier','Nom','Prénom','Activité supprimée','Montant avoir (€)','Date génération','Code avoir','Solde restant (€)','Statut'];
    sheet.getRange(1,1,1,h.length).setValues([h]).setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(1,5).setBackground('#bf360c');
    sheet.getRange(1,7).setBackground('#1a2e22'); // Code avoir — vert foncé
    sheet.getRange(1,8).setBackground('#bf360c');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,110);sheet.setColumnWidth(2,140);sheet.setColumnWidth(3,140);
    sheet.setColumnWidth(4,260);sheet.setColumnWidth(5,130);sheet.setColumnWidth(6,150);
    sheet.setColumnWidth(7,110);sheet.setColumnWidth(8,130);sheet.setColumnWidth(9,110);
  } else {
    // Migration : ajouter les colonnes manquantes si l'onglet existait avec l'ancienne structure
    var nbCols = sheet.getLastColumn();
    if (nbCols < 9) {
      if (nbCols < 7) sheet.getRange(1,7).setValue('Code avoir').setBackground('#1a2e22').setFontColor('#ffffff').setFontWeight('bold');
      if (nbCols < 8) sheet.getRange(1,8).setValue('Solde restant (€)').setBackground('#bf360c').setFontColor('#ffffff').setFontWeight('bold');
      if (nbCols < 9) sheet.getRange(1,9).setValue('Statut').setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setColumnWidth(7,110);sheet.setColumnWidth(8,130);sheet.setColumnWidth(9,110);
    }
  }
  return sheet;
}

// ── Onglet Remboursements : identique à Avoirs générés, sans le code à 6 caractères,
//    avec en plus le mode de remboursement (Chèque / Virement / HelloAsso).
//    Positionné juste après l'onglet Avoirs générés. ──
function getOrCreateRemboursementSheet(ss){
  var sheet=ss.getSheetByName(SHEET_REMBOURSEMENTS);
  if(!sheet){
    // Positionner juste après "Avoirs générés" : insertSheet(nom, indexPos) prend un index
    // 0-based, et getIndex() de l'onglet Avoirs renvoie son rang 1-based → l'utiliser
    // directement comme indexPos place le nouvel onglet juste après.
    var avoirSh = getOrCreateAvoirSheet(ss);
    sheet = ss.insertSheet(SHEET_REMBOURSEMENTS, avoirSh.getIndex());
    var h=['N° Dossier','Nom','Prénom','Motif','Montant remboursement (€)','Date','Mode remboursement','Statut'];
    sheet.getRange(1,1,1,h.length).setValues([h]).setBackground('#6a1b9a').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(1,5).setBackground('#4a148c');
    sheet.getRange(1,7).setBackground('#4a148c');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,110);sheet.setColumnWidth(2,140);sheet.setColumnWidth(3,140);
    sheet.setColumnWidth(4,260);sheet.setColumnWidth(5,170);sheet.setColumnWidth(6,110);
    sheet.setColumnWidth(7,150);sheet.setColumnWidth(8,110);
  }
  return sheet;
}

// ══════════════════════════════════════════════════════════════
// ONGLET "Commentaires admin" — traçabilité des montants modifiés
// et commentaires libres saisis par l'admin lors d'un ajout/suppression
// d'activité ou d'une bascule liste d'attente.
// ══════════════════════════════════════════════════════════════
var SHEET_COMMENTAIRES_ADMIN = 'Commentaires admin';
function getOrCreateCommentairesSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_COMMENTAIRES_ADMIN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COMMENTAIRES_ADMIN);
    var h = ['Date','Action','N° Dossier','Nom','Prénom','Activité','Montant (€)','Commentaire admin'];
    sheet.getRange(1,1,1,h.length).setValues([h])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,140);sheet.setColumnWidth(2,170);sheet.setColumnWidth(3,110);
    sheet.setColumnWidth(4,140);sheet.setColumnWidth(5,140);sheet.setColumnWidth(6,220);
    sheet.setColumnWidth(7,110);sheet.setColumnWidth(8,320);
  }
  return sheet;
}
// N'écrit une ligne que si un commentaire a été saisi OU si le montant a été
// modifié manuellement (montantModifie=true) — pour ne pas polluer l'onglet
// à chaque action routinière sans intervention de l'admin.
function logCommentaireAdmin(ss, action, code, nom, prenom, actNom, montant, commentaire, montantModifie) {
  try {
    commentaire = String(commentaire || '').trim();
    if (!commentaire && !montantModifie) return;
    var sheet = getOrCreateCommentairesSheet(ss);
    var dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
    var ligne = [
      dateJour, action, code || '', nom || '', prenom || '', actNom || '',
      (typeof montant === 'number' ? montant : parseFloat(montant) || 0),
      commentaire || (montantModifie ? '(montant modifié manuellement, sans commentaire)' : '')
    ];
    sheet.getRange(Math.max(sheet.getLastRow() + 1, 2), 1, 1, ligne.length).setValues([ligne]);
  } catch (eLog) {
    Logger.log('⚠️ logCommentaireAdmin KO: ' + eLog);
  }
}
// Bloc HTML à insérer dans un email quand un commentaire admin est présent.
function blocCommentaireAdminHtml(commentaire) {
  commentaire = String(commentaire || '').trim();
  if (!commentaire) return '';
  return '<div style="background:#fff8e1;border-left:4px solid #f9a825;border-radius:6px;'
    + 'padding:12px 16px;margin:16px 0;font-size:13px;color:#5d4a00;line-height:1.6">'
    + '<strong>Note de l\'équipe FRI :</strong><br>' + commentaire.replace(/\n/g,'<br>')
    + '</div>';
}

// Générer un code avoir unique 6 caractères alphanum (ex: AV3K9X)
function genererCodeAvoir(ss) {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // sans I,O,0,1 pour éviter confusions
  var sheet = getOrCreateAvoirSheet(ss);
  var existants = {};
  if (sheet.getLastRow() > 1) {
    var codes = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues();
    codes.forEach(function(r){ if(r[0]) existants[String(r[0]).trim()] = true; });
  }
  var tentatives = 0;
  while (tentatives < 100) {
    var code = 'AV';
    for (var i = 0; i < 4; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
    if (!existants[code]) return code;
    tentatives++;
  }
  return 'AV' + Math.random().toString(36).substring(2,6).toUpperCase();
}

// ── Créer un avoir manuellement depuis la console admin ──────────────────────
// Insère une ligne dans "Avoirs générés" et envoie un email à l'adhérent
function creerAvoirManuelGAS(code, motif, montant, emailDest, nomDest, prenomDest) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var avoirSh = getOrCreateAvoirSheet(ss);
    var codeAvoir = genererCodeAvoir(ss);
    var dateAvoir = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');

    // Ligne : [N°Dossier, Nom, Prénom, Activité supprimée (= motif), Montant, Date, Code avoir, Solde restant, Statut]
    var ligne = [
      code,
      nomDest || '',
      prenomDest || '',
      motif || 'Avoir manuel',
      montant,
      dateAvoir,
      codeAvoir,
      montant,
      'Disponible'
    ];
    avoirSh.getRange(Math.max(avoirSh.getLastRow() + 1, 2), 1, 1, ligne.length)
      .setValues([ligne]);
    // Mise en forme : code avoir en vert foncé, solde en rouge clair
    var newRow = Math.max(avoirSh.getLastRow(), 2);
    avoirSh.getRange(newRow, 7).setBackground('#d8f3dc').setFontColor('#1a2e22').setFontWeight('bold').setFontFamily('monospace');
    avoirSh.getRange(newRow, 8).setBackground('#fff3e0').setFontColor('#e65100').setFontWeight('bold');
    avoirSh.getRange(newRow, 9).setBackground('#d8f3dc').setFontColor('#2d6a4f');
    Logger.log('✅ Avoir manuel créé: ' + codeAvoir + ' — ' + montant + ' € — ' + code);

    // ── Envoi email adhérent ──
    if (emailDest && emailDest.indexOf('@') >= 0) {
      var logoBA = getLogoBlob();
      var sujetEmail = '🎁 Votre avoir FRI — ' + montant.toFixed(2) + ' €';
      var htmlEmail = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
        + '<body style="font-family:Arial,sans-serif;background:#f0f4f2;margin:0;padding:20px">'
        + '<div style="max-width:600px;margin:0 auto;background:white;border-radius:12px;overflow:hidden">'
        + '<div style="background:#1a2e22;padding:20px 24px;text-align:center">'
        + '<img src="cid:logo_fri" alt="FRI" style="width:60px;height:auto;border-radius:8px;display:block;margin:0 auto 10px">'
        + '<div style="font-family:Georgia,serif;font-size:22px;color:white;font-weight:bold">Foyer Rural d\'Isneauville</div>'
        + '<div style="font-size:12px;color:#52b788;letter-spacing:1px;margin-top:4px">SAISON 2026 / 2027</div>'
        + '</div>'
        + '<div style="padding:24px 28px">'
        + '<p style="font-size:15px;color:#1a2e22;margin-bottom:16px">Bonjour <strong>' + prenomDest + ' ' + nomDest + '</strong>,</p>'
        + '<p style="font-size:13px;color:#555;line-height:1.7;margin-bottom:20px">'
        + 'Nous avons le plaisir de vous informer qu\'un <strong>avoir</strong> a été généré pour votre dossier.</p>'
        // Bloc avoir
        + '<div style="background:#d8f3dc;border:2px solid #52b788;border-radius:10px;padding:20px 24px;margin:20px 0;text-align:center">'
        + '<div style="font-size:13px;color:#2d6a4f;margin-bottom:8px">🎁 Montant de l\'avoir</div>'
        + '<div style="font-size:36px;font-weight:900;color:#1a2e22;margin-bottom:8px">' + montant.toFixed(2) + ' €</div>'
        + '<div style="font-size:13px;color:#555;margin-bottom:14px">Motif : <strong>' + motif + '</strong></div>'
        + '<div style="background:#1a2e22;border-radius:8px;padding:12px 20px;display:inline-block;margin:0 auto">'
        + '<div style="font-size:11px;color:#52b788;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px">Votre code avoir</div>'
        + '<div style="font-family:monospace;font-size:28px;font-weight:900;color:white;letter-spacing:6px">' + codeAvoir + '</div>'
        + '</div>'
        + '</div>'
        // Instructions
        + '<div style="background:#f5f5f5;border-radius:8px;padding:14px 18px;margin:16px 0;font-size:13px;color:#555;line-height:1.8">'
        + '<strong style="color:#1a2e22">Comment utiliser votre avoir ?</strong><br>'
        + '➡️ Lors de votre prochaine inscription sur le site FRI,<br>'
        + '&nbsp;&nbsp;&nbsp;saisissez ce code dans le champ <em>"Avoir FRI"</em> à l\'étape Aides &amp; Mode de règlement.'
        + '<br><br>'
        + '<span style="font-size:11px;color:#888">Code dossier : <strong>' + code + '</strong> — Avoir valable jusqu\'à utilisation complète.</span>'
        + '</div>'
        + '<div style="background:#e8f4fd;border-left:4px solid #2980b9;border-radius:6px;padding:11px;margin-top:14px;font-size:13px;line-height:1.7">'
        + '<strong>Permanences</strong> : tous les <strong>mardis de 16h30 à 18h30</strong> <em>(période scolaire)</em><br>'
        + 'Salle des fêtes — Place A. Cramilly, 76230 Isneauville</div>'
        + '<p style="margin-top:16px;font-size:12px;color:#555">Contact : <a href="mailto:frisneauville@orange.fr" style="color:#2d6a4f">frisneauville@orange.fr</a> — <a href="tel:0235590101" style="color:#2d6a4f">02.35.59.01.01</a></p>'
        + '<p style="font-size:12px;color:#aaa;margin-top:8px">Cordialement,<br>L\'équipe du Foyer Rural d\'Isneauville</p>'
        + '</div></div></body></html>';

      envoyerEmail(
        emailDest,
        sujetEmail,
        'Votre avoir FRI (' + montant.toFixed(2) + ' €) — Code : ' + codeAvoir,
        {
          htmlBody: htmlEmail,
          charset: 'UTF-8',
          name: 'FRI Inscriptions',
          inlineImages: logoBA ? { logo_fri: logoBA } : {}
        }
      );
      Logger.log('✅ Email avoir envoyé à ' + emailDest);
    }

    return { status: 'ok', codeAvoir: codeAvoir, montant: montant, message: 'Avoir créé et email envoyé.' };
  } catch(e) {
    Logger.log('❌ creerAvoirManuelGAS: ' + e);
    return { status: 'error', message: e.toString() };
  }
}

// ── Créer un remboursement manuel depuis la console admin ────────────────────
// Insère une ligne dans "Remboursements" et envoie un email à l'adhérent.
// Contrairement à l'avoir, pas de code à 6 caractères (rien à réutiliser plus
// tard) — juste le mode par lequel l'argent est effectivement rendu.
function creerRemboursementManuelGAS(code, motif, montant, modeRemb, emailDest, nomDest, prenomDest) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rembSh = getOrCreateRemboursementSheet(ss);
    var dateRemb = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');
    var modeLabelMapR = {cheque:'Chèque', virement:'Virement', helloasso:'HelloAsso'};
    var modeLabelR = modeLabelMapR[String(modeRemb||'').toLowerCase()] || (modeRemb || 'Non précisé');

    // Ligne : [N°Dossier, Nom, Prénom, Motif, Montant, Date, Mode remboursement, Statut]
    var ligneR = [
      code,
      nomDest || '',
      prenomDest || '',
      motif || 'Remboursement manuel',
      montant,
      dateRemb,
      modeLabelR,
      'À traiter'
    ];
    rembSh.getRange(Math.max(rembSh.getLastRow() + 1, 2), 1, 1, ligneR.length)
      .setValues([ligneR]);
    var newRowR = Math.max(rembSh.getLastRow(), 2);
    rembSh.getRange(newRowR, 7).setBackground('#f3e5f5').setFontColor('#4a148c').setFontWeight('bold');
    rembSh.getRange(newRowR, 8).setBackground('#fff3e0').setFontColor('#e65100').setFontWeight('bold');
    Logger.log('✅ Remboursement manuel créé: ' + modeLabelR + ' — ' + montant + ' € — ' + code);

    // ── Envoi email adhérent ──
    if (emailDest && emailDest.indexOf('@') >= 0) {
      var logoBAR = getLogoBlob();
      var sujetEmailR = '💶 Remboursement FRI — ' + montant.toFixed(2) + ' €';
      var htmlEmailR = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
        + '<body style="font-family:Arial,sans-serif;background:#f0f4f2;margin:0;padding:20px">'
        + '<div style="max-width:600px;margin:0 auto;background:white;border-radius:12px;overflow:hidden">'
        + '<div style="background:#1a2e22;padding:20px 24px;text-align:center">'
        + '<img src="cid:logo_fri" alt="FRI" style="width:60px;height:auto;border-radius:8px;display:block;margin:0 auto 10px">'
        + '<div style="font-family:Georgia,serif;font-size:22px;color:white;font-weight:bold">Foyer Rural d\'Isneauville</div>'
        + '<div style="font-size:12px;color:#52b788;letter-spacing:1px;margin-top:4px">SAISON 2026 / 2027</div>'
        + '</div>'
        + '<div style="padding:24px 28px">'
        + '<p style="font-size:15px;color:#1a2e22;margin-bottom:16px">Bonjour <strong>' + prenomDest + ' ' + nomDest + '</strong>,</p>'
        + '<p style="font-size:13px;color:#555;line-height:1.7;margin-bottom:20px">'
        + 'Nous avons le plaisir de vous informer qu\'un <strong>remboursement</strong> a été enregistré pour votre dossier.</p>'
        // Bloc remboursement
        + '<div style="background:#f3e5f5;border:2px solid #6a1b9a;border-radius:10px;padding:20px 24px;margin:20px 0;text-align:center">'
        + '<div style="font-size:13px;color:#4a148c;margin-bottom:8px">💶 Montant du remboursement</div>'
        + '<div style="font-size:36px;font-weight:900;color:#1a2e22;margin-bottom:8px">' + montant.toFixed(2) + ' €</div>'
        + '<div style="font-size:13px;color:#555;margin-bottom:4px">Motif : <strong>' + motif + '</strong></div>'
        + '<div style="font-size:13px;color:#555">Mode de remboursement : <strong>' + modeLabelR + '</strong></div>'
        + '</div>'
        // Instructions (adaptées au mode)
        + '<div style="background:#f5f5f5;border-radius:8px;padding:14px 18px;margin:16px 0;font-size:13px;color:#555;line-height:1.8">'
        + '<strong style="color:#1a2e22">Comment allez-vous être remboursé ?</strong><br>'
        + (modeLabelR === 'Chèque'
            ? '➡️ Un chèque de ' + montant.toFixed(2) + ' € vous sera remis ou envoyé prochainement.'
            : modeLabelR === 'Virement'
            ? '➡️ Un virement de ' + montant.toFixed(2) + ' € sera effectué sur votre compte prochainement.'
            : '➡️ Le remboursement de ' + montant.toFixed(2) + ' € sera traité via HelloAsso prochainement.')
        + '<br><br>'
        + '<span style="font-size:11px;color:#888">Code dossier : <strong>' + code + '</strong></span>'
        + '</div>'
        + '<div style="background:#e8f4fd;border-left:4px solid #2980b9;border-radius:6px;padding:11px;margin-top:14px;font-size:13px;line-height:1.7">'
        + '<strong>Permanences</strong> : tous les <strong>mardis de 16h30 à 18h30</strong> <em>(période scolaire)</em><br>'
        + 'Salle des fêtes — Place A. Cramilly, 76230 Isneauville</div>'
        + '<p style="margin-top:16px;font-size:12px;color:#555">Contact : <a href="mailto:frisneauville@orange.fr" style="color:#2d6a4f">frisneauville@orange.fr</a> — <a href="tel:0235590101" style="color:#2d6a4f">02.35.59.01.01</a></p>'
        + '<p style="font-size:12px;color:#aaa;margin-top:8px">Cordialement,<br>L\'équipe du Foyer Rural d\'Isneauville</p>'
        + '</div></div></body></html>';

      envoyerEmail(
        emailDest,
        sujetEmailR,
        'Remboursement FRI (' + montant.toFixed(2) + ' €) — Mode : ' + modeLabelR,
        {
          htmlBody: htmlEmailR,
          charset: 'UTF-8',
          name: 'FRI Inscriptions',
          bcc: EMAIL_TRESORIER,
          inlineImages: logoBAR ? { logo_fri: logoBAR } : {}
        }
      );
      Logger.log('✅ Email remboursement envoyé à ' + emailDest + ' (cci trésorier)');
    } else {
      // Pas d'email adhérent : le trésorier doit quand même être notifié
      envoyerEmail(
        EMAIL_TRESORIER,
        '💶 Remboursement FRI créé — ' + montant.toFixed(2) + ' € (' + code + ')',
        'Un remboursement a été enregistré :\n\n'
          + 'Dossier : ' + code + '\n'
          + 'Nom : ' + (prenomDest||'') + ' ' + (nomDest||'') + '\n'
          + 'Motif : ' + (motif||'') + '\n'
          + 'Montant : ' + montant.toFixed(2) + ' €\n'
          + 'Mode de remboursement : ' + modeLabelR + '\n'
          + '(Aucun e-mail adhérent renseigné pour ce dossier.)',
        { name: 'FRI Inscriptions' }
      );
      Logger.log('✅ Email remboursement envoyé au trésorier (pas d\'email adhérent)');
    }

    return { status: 'ok', montant: montant, mode: modeLabelR, message: 'Remboursement créé et email envoyé.' };
  } catch(e) {
    Logger.log('❌ creerRemboursementManuelGAS: ' + e);
    return { status: 'error', message: e.toString() };
  }
}

// Vérifier un code avoir : retourne { valide, nom, prenom, montantDisponible, message }
function verifierCodeAvoirGAS(codeAvoir, montantDemande) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_AVOIRS);
    if (!sheet || sheet.getLastRow() < 2) return { valide: false, message: 'Aucun avoir enregistré.' };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    for (var i = 0; i < data.length; i++) {
      var codeRow = String(data[i][6] || '').trim().toUpperCase();
      if (codeRow !== codeAvoir.trim().toUpperCase()) continue;
      var solde    = parseFloat(data[i][7]) || 0;
      var statut   = String(data[i][8] || '').toLowerCase();
      var nom      = String(data[i][1] || '');
      var prenom   = String(data[i][2] || '');
      var montantOrig = parseFloat(data[i][4]) || 0;
      if (statut === 'épuisé') return { valide: false, message: 'Cet avoir a déjà été entièrement utilisé.' };
      if (solde <= 0) return { valide: false, message: 'Solde de cet avoir épuisé (' + montantOrig.toFixed(2) + ' € initialement).' };
      if (montantDemande > 0 && montantDemande > solde + 0.005) {
        return { valide: false, message: 'Montant demandé (' + montantDemande.toFixed(2) + ' €) supérieur au solde disponible (' + solde.toFixed(2) + ' €).' };
      }
      return { valide: true, nom: nom, prenom: prenom, montantDisponible: solde, codeAvoir: codeRow, ligne: i + 2 };
    }
    return { valide: false, message: 'Code avoir "' + codeAvoir + '" introuvable.' };
  } catch(e) {
    return { valide: false, message: 'Erreur vérification : ' + e.toString() };
  }
}

// Déduire un montant d'un avoir (appelé à la validation du paiement)
function utiliserAvoirGAS(codeAvoir, montantUtilise) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_AVOIRS);
    if (!sheet || sheet.getLastRow() < 2) return { ok: false, message: 'Onglet introuvable' };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    for (var i = 0; i < data.length; i++) {
      var codeRow = String(data[i][6] || '').trim().toUpperCase();
      if (codeRow !== codeAvoir.trim().toUpperCase()) continue;
      var solde = parseFloat(data[i][7]) || 0;
      var nouveauSolde = Math.max(0, Math.round((solde - montantUtilise) * 100) / 100);
      var newStatut = nouveauSolde <= 0 ? 'Épuisé' : 'Partiel';
      // Mettre à jour col 8 (solde) et col 9 (statut)
      sheet.getRange(i + 2, 8).setValue(nouveauSolde).setNumberFormat('#,##0.00 €');
      sheet.getRange(i + 2, 9).setValue(newStatut)
        .setBackground(nouveauSolde <= 0 ? '#fce4ec' : '#fff8e1')
        .setFontColor(nouveauSolde <= 0 ? '#c0392b' : '#856404')
        .setFontWeight('bold');
      Logger.log('✅ Avoir ' + codeAvoir + ' utilisé : ' + montantUtilise + ' € — solde restant : ' + nouveauSolde + ' €');
      return { ok: true, soldePrecedent: solde, nouveauSolde: nouveauSolde };
    }
    return { ok: false, message: 'Code avoir introuvable : ' + codeAvoir };
  } catch(e) {
    return { ok: false, message: e.toString() };
  }
}
function getOrCreateAvoirsUtilisesSheet(ss) {
  if (!ss) return null;
  var sheet = ss.getSheetByName(SHEET_AVOIRS_UTILISES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_AVOIRS_UTILISES);
    var h = ['N° Dossier', 'Nom', 'Prénom', 'Email', 'Montant utilisé (€)', 'Date paiement', 'Statut'];
    sheet.getRange(1, 1, 1, h.length).setValues([h])
      .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(2, 1, 1, h.length).mergeAcross()
      .setValue('Avoirs utilisés au règlement — Saison 2026/2027')
      .setFontStyle('italic').setFontSize(9).setFontColor('#555555').setBackground('#fff3e0');
    sheet.setFrozenRows(2);
    sheet.getRange(3, 1, 1, h.length).setBackground('#ffe0b2');
    sheet.getRange(3, 4).setValue('TOTAL AVOIRS UTILISÉS').setFontWeight('bold').setFontColor('#bf360c');
    sheet.getRange(3, 5).setValue(0).setFontWeight('bold').setFontColor('#bf360c').setNumberFormat('#,##0.00 €');
    sheet.getRange(3, 7).setValue('Statut').setFontStyle('italic').setFontColor('#888888');
    sheet.setFrozenRows(3);
    sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 140); sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 200); sheet.setColumnWidth(5, 140); sheet.setColumnWidth(6, 140); sheet.setColumnWidth(7, 160);
    // Positionner juste après "Avoirs générés"
    try {
      var avoirs = ss.getSheetByName(SHEET_AVOIRS);
      var nbS2 = ss.getNumSheets();
      if (avoirs && avoirs.getIndex() + 1 <= nbS2) {
        ss.setActiveSheet(sheet); ss.moveActiveSheet(avoirs.getIndex() + 1);
      }
    } catch(ep) { Logger.log('Positionnement Avoirs utilisés KO : ' + ep); }
  }
  return sheet;
}

function majTotalAvoirsUtilises(ss) {
  try {
    if (!ss) return;
    var sheet = ss.getSheetByName(SHEET_AVOIRS_UTILISES);
    if (!sheet || sheet.getLastRow() < 4) {
      if (sheet) sheet.getRange(3, 5).setValue(0).setFontWeight('bold').setFontColor('#bf360c').setBackground('#ffe0b2').setNumberFormat('#,##0.00 €');
      return;
    }
    var total = 0;
    var data = sheet.getRange(4, 5, sheet.getLastRow() - 3, 1).getValues();
    data.forEach(function(row){ total += Number(row[0]) || 0; });
    total = Math.round(total * 100) / 100;
    sheet.getRange(3, 5).setValue(total).setFontWeight('bold').setFontColor('#bf360c').setBackground('#ffe0b2').setNumberFormat('#,##0.00 €');
    Logger.log('✅ Total Avoirs utilisés : ' + total + ' €');
  } catch(e) { Logger.log('majTotalAvoirsUtilises KO : ' + e.toString()); }
}

function ecrireAvoirUtilise(ss, rows, avoirMontant) {
  if (!ss || !rows || rows.length === 0 || avoirMontant <= 0) return;
  var r0 = rows[0];
  var code   = r0.code_dossier || '';
  var nom    = r0.responsable_nom    || r0.membre_nom    || '';
  var prenom = r0.responsable_prenom || r0.membre_prenom || '';
  var email  = r0.email1 || '';
  var dateJour = r0.date || Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  try {
    var sheet = getOrCreateAvoirsUtilisesSheet(ss);
    // Anti-doublon : ne pas écrire si le code est déjà présent
    if (sheet.getLastRow() >= 4) {
      var existing = sheet.getRange(4, 1, sheet.getLastRow() - 3, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]).trim() === code) {
          Logger.log('ecrireAvoirUtilise — doublon ignoré : ' + code);
          return;
        }
      }
    }
    var nextRow = Math.max(sheet.getLastRow() + 1, 4);
    var bg = nextRow % 2 === 0 ? '#fff3e0' : '#ffffff';
    sheet.getRange(nextRow, 1, 1, 7).setValues([[code, nom, prenom, email, avoirMontant, dateJour, '⏳ En cours de validation']])
      .setBackground(bg);
    sheet.getRange(nextRow, 5).setFontColor('#bf360c').setFontWeight('bold').setNumberFormat('#,##0.00 €');
    sheet.getRange(nextRow, 7).setFontColor('#856404').setFontWeight('bold');
    sheet.autoResizeColumns(1, 7);
    majTotalAvoirsUtilises(ss);
    // Mettre à jour le solde restant dans l'onglet Avoirs générés
    // Cherche par code avoir (col G) ou par code dossier d'origine (col A)
    var avoirCode = String(r0.avoir_code || r0.code_avoir || '').trim().toUpperCase();
    // Extraire depuis commentaire si format AVOIR:XXXX:montant
    if (!avoirCode && r0.comment) {
      var acMatch = String(r0.comment).match(/AVOIR:([A-Z0-9-]+):/);
      if (acMatch) avoirCode = acMatch[1];
    }
    try { majSoldeAvoirGenere(ss, avoirCode || code, avoirMontant); } catch(eS) { Logger.log('majSoldeAvoirGenere KO: ' + eS); }
    Logger.log('✅ Avoir utilisé écrit : ' + avoirMontant + ' € pour ' + code);
  } catch(e) { Logger.log('ecrireAvoirUtilise KO : ' + e.toString()); }
}

// ── Mettre à jour le solde restant (col H) dans l'onglet Avoirs générés ──
function majSoldeAvoirGenere(ss, emailOuCode, montantUtilise) {
  if (!ss || !emailOuCode || montantUtilise <= 0) return;
  var sheet = ss.getSheetByName(SHEET_AVOIRS);
  if (!sheet || sheet.getLastRow() < 2) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  var updated = false;
  // Structure onglet : A=N°Dossier B=Nom C=Prénom D=Activité supprimée E=Montant avoir
  //                    F=Date génération G=Code avoir H=Solde restant I=Statut
  for (var i = 0; i < data.length; i++) {
    var rowCode  = String(data[i][0] || '').trim(); // col A = N° Dossier d'origine
    var codeAvRow = String(data[i][6] || '').trim().toUpperCase(); // col G = Code avoir
    var searchKey = String(emailOuCode).trim().toUpperCase();
    // Chercher par code avoir (col G) en priorité, puis par N° dossier (col A)
    if (codeAvRow === searchKey || rowCode === searchKey) {
      var montantAvoir = Number(data[i][4] || 0); // col E : montant avoir initial
      var soldeActuel  = (data[i][7] !== '' && data[i][7] !== null && data[i][7] !== undefined)
                       ? Number(data[i][7] || 0) : montantAvoir; // col H : solde restant
      var nouveauSolde  = Math.max(0, Math.round((soldeActuel - montantUtilise) * 100) / 100);
      var statut        = nouveauSolde <= 0 ? '✅ Solde épuisé' : '⏳ Solde restant : ' + nouveauSolde.toFixed(2) + ' €';
      sheet.getRange(i + 2, 8).setValue(nouveauSolde)
        .setFontColor(nouveauSolde <= 0 ? '#555555' : '#bf360c')
        .setFontWeight('bold')
        .setNumberFormat('#,##0.00 €');
      sheet.getRange(i + 2, 9).setValue(statut)
        .setFontColor(nouveauSolde <= 0 ? '#555555' : '#856404')
        .setFontWeight('bold');
      Logger.log('✅ Solde avoir mis à jour — ' + emailOuCode + ' : ' + soldeActuel + ' - ' + montantUtilise + ' = ' + nouveauSolde + ' €');
      updated = true;
      break;
    }
  }
  if (!updated) Logger.log('⚠️ majSoldeAvoirGenere : avoir non trouvé pour ' + emailOuCode);
}

function getOrCreateAideSheet(ss,nomOnglet,labelAide,couleur){
  var sheet=ss.getSheetByName(nomOnglet);if(!sheet){sheet=ss.insertSheet(nomOnglet);var h=['N° Dossier','Nom','Prénom','Email','Montant aide (€)','Date inscription'];sheet.getRange(1,1,1,h.length).setValues([h]).setBackground(couleur).setFontColor('#ffffff').setFontWeight('bold');sheet.getRange(2,1,1,h.length).mergeAcross().setValue('Aide : '+labelAide).setFontStyle('italic').setFontSize(9).setFontColor('#555555').setBackground('#f5f5f5');sheet.setFrozenRows(2);sheet.setColumnWidth(1,110);sheet.setColumnWidth(2,140);sheet.setColumnWidth(3,140);sheet.setColumnWidth(4,200);sheet.setColumnWidth(5,130);sheet.setColumnWidth(6,130);}return sheet;
}

// ── Onglet HelloAsso ──
function getOrCreateHelloAssoSheet(ss) {
  if (!ss) return null;
  var sheet = ss.getSheetByName(SHEET_HELLOASSO);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HELLOASSO);
    var h = ['N° Dossier','Nom','Prénom','Email','Montant dossier (€)','1er versement (€)','Mode paiement','Date 1er versement','Statut','Type campagne HA','Désignation HA','2e versement (€)','3e versement (€)'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]).setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(2, 1, 1, h.length).mergeAcross().setValue('Règlements HelloAsso — Saison 2026/2027').setFontStyle('italic').setFontSize(9).setFontColor('#555555').setBackground('#fbe9e7');
    sheet.setFrozenRows(2);
    // Ligne TOTAL (ligne 3)
    sheet.getRange(3, 1, 1, h.length).setBackground('#fbe9e7');
    sheet.getRange(3, 4).setValue('TOTAL HELLOASSO').setFontWeight('bold').setFontColor('#bf360c');
    sheet.getRange(3, 5).setValue(0).setFontWeight('bold').setFontColor('#bf360c').setNumberFormat('#,##0.00 €');
    sheet.getRange(3, 6).setValue(0).setFontWeight('bold').setFontColor('#bf360c').setNumberFormat('#,##0.00 €');
    sheet.setFrozenRows(3);
    sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 140); sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 200); sheet.setColumnWidth(5, 100); sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 120); sheet.setColumnWidth(8, 130); sheet.setColumnWidth(9, 100);
    sheet.setColumnWidth(10, 130); sheet.setColumnWidth(11, 130); sheet.setColumnWidth(12, 100); sheet.setColumnWidth(13, 100);
    // Positionner après Espèces
    try {
      var especes = ss.getSheetByName(SHEET_ESPECES);
      var nbS3 = ss.getNumSheets();
      if (especes && especes.getIndex() + 1 <= nbS3) {
        ss.setActiveSheet(sheet); ss.moveActiveSheet(especes.getIndex() + 1);
      }
    } catch(ep) { Logger.log('Positionnement HelloAsso KO : ' + ep); }
  }
  return sheet;
}



// ── Helper : mettre à jour la cellule TOTAL d'un onglet ──
// ligne 3, col colTotal (5 pour HA/Espèces, 6 pour Chèques)
function majTotalOnglet(sheet, colTotal) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) { sheet.getRange(3, colTotal).setValue(0); return; }
  var vals = sheet.getRange(4, colTotal, lastRow - 3, 1).getValues();
  var total = vals.reduce(function(sum, row) { return sum + (Number(row[0]) || 0); }, 0);
  total = Math.round(total * 100) / 100;
  sheet.getRange(3, colTotal).setValue(total).setNumberFormat('#,##0.00 €');
}

// majOngletsAides supprimée v8.82 — remplacée par majOngletAideANCv, majOngletAidePassSport, majOngletAidePassJeunes, majOngletAideAtout (avec anti-doublon)

// ── Onglet Espèces ──
function getOrCreateEspecesSheet(ss) {
  if (!ss) { Logger.log('getOrCreateEspecesSheet : ss undefined'); return null; }
  var sheet = ss.getSheetByName(SHEET_ESPECES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ESPECES);
    var h = ['N° Dossier', 'Nom', 'Prénom', 'Email', 'Montant (€)', 'Date inscription', 'Statut'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]).setBackground('#795548').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(2, 1, 1, h.length).mergeAcross().setValue('Règlements en espèces — Saison 2026/2027').setFontStyle('italic').setFontSize(9).setFontColor('#555555').setBackground('#efebe9');
    sheet.setFrozenRows(2);
    sheet.getRange(3, 1, 1, h.length).setBackground('#d7ccc8');
    sheet.getRange(3, 4).setValue('TOTAL ESPÈCES').setFontWeight('bold').setFontColor('#3e2723');
    sheet.getRange(3, 5).setValue(0).setFontWeight('bold').setFontColor('#3e2723').setNumberFormat('#,##0.00 €');
    sheet.setFrozenRows(3);
    sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 140); sheet.setColumnWidth(3, 140);
    sheet.setColumnWidth(4, 200); sheet.setColumnWidth(5, 130); sheet.setColumnWidth(6, 130); sheet.setColumnWidth(7, 160);
    try { var avoirs = ss.getSheetByName(SHEET_AVOIRS); if (avoirs) { var positionCible = avoirs.getIndex(); ss.setActiveSheet(sheet); ss.moveActiveSheet(positionCible); } } catch(ep) { Logger.log('Positionnement Espèces KO : ' + ep); }
  }
  return sheet;
}
function majTotalEspeces(ss) {
  try{if(!ss)return;var sheet=ss.getSheetByName(SHEET_ESPECES);if(!sheet||sheet.getLastRow()<4)return;var total=0;var data=sheet.getRange(4,5,sheet.getLastRow()-3,1).getValues();data.forEach(function(row){total+=Number(row[0])||0;});total=Math.round(total*100)/100;sheet.getRange(3,5).setValue(total).setFontWeight('bold').setFontColor('#3e2723').setBackground('#d7ccc8').setNumberFormat('#,##0.00 €');Logger.log('✅ Total Espèces : '+total+' €');}catch(e){Logger.log('majTotalEspeces KO : '+e.toString());}
}

// Calcule le montant réellement réglé = total famille - aides/avoirs/ANCV
function calculerSoldeReel(r0, montantForce) {
  if (montantForce && montantForce > 0) return montantForce; // montant réel HA webhook
  // Utiliser solde_net précalculé si disponible (validerPaiementSheet)
  if (r0.solde_net !== undefined && r0.solde_net >= 0) return r0.solde_net;
  var total          = Number(r0.total_famille)            || 0;
  var deducAvoir     = parseFloat(r0.avoir_montant)        || 0;
  var deducPass      = parseFloat(r0.pass_sport_montant)   || 0;
  var deducAncv      = parseFloat(r0.ancv_montant)         || 0;
  var deducPJ        = parseFloat(r0.pass_jeunes_montant)  || 0;
  var deducAtout     = parseFloat(r0.atout_montant)        || 0;
  // Parser aussi depuis pass_aide si champs directs non disponibles
  var passAideStr = String(r0.pass_aide || '');
  if (!deducPJ) {
    var pjParse = passAideStr.match(/PassJeunes:([\d.]+)/);
    if (pjParse) deducPJ = parseFloat(pjParse[1]) || 0;
  }
  if (!deducAtout) {
    var atoutParse = passAideStr.match(/Atout:([\d.]+)/);
    if (atoutParse) deducAtout = parseFloat(atoutParse[1]) || 0;
  }
  return Math.max(0, Math.round((total - deducAvoir - deducPass - deducAncv - deducPJ - deducAtout) * 100) / 100);
}


// ══════════════════════════════════════════════════════════════
// ONGLETS AIDES — ANCV, Pass Jeunes, Atout Normandie, Pass Sport
// ══════════════════════════════════════════════════════════════

// ── Helpers génériques ────────────────────────────────────────

function getOrCreateAideSheet(ss, nomOnglet, titre, couleurFond, couleurTexte, labelTotal) {
  var sheet = ss.getSheetByName(nomOnglet);
  if (!sheet) {
    sheet = ss.insertSheet(nomOnglet);
    var h = ['N° Dossier','Nom','Prénom','Email','Montant aide (€)','Type / Détail','Date inscription','Statut'];
    sheet.getRange(1,1,1,h.length).setValues([h])
      .setBackground(couleurFond).setFontColor(couleurTexte).setFontWeight('bold');
    // Calculer une couleur pastel pour la ligne 2 (fond dilue)
    var pastelMap = {
      '#1565c0':'#dbeafe', '#1b5e20':'#dcfce7', '#6a1b9a':'#f3e8ff',
      '#e65100':'#ffedd5', '#795548':'#f5f0ed', '#0d47a1':'#dbeafe',
      '#4a148c':'#f3e8ff', '#e65100':'#ffedd5'
    };
    var bgLigne2 = pastelMap[couleurFond] || '#f5f5f5';
    sheet.getRange(2,1,1,h.length).mergeAcross()
      .setValue(titre + ' — Saison 2026/2027')
      .setFontStyle('italic').setFontSize(9).setFontColor('#555555')
      .setBackground(bgLigne2);
    sheet.getRange(3,1,1,h.length).setBackground('#eeeeee');
    sheet.getRange(3,5).setValue(labelTotal).setFontWeight('bold').setFontColor(couleurFond);
    sheet.getRange(3,6).setValue(0).setFontWeight('bold').setFontColor(couleurFond)
      .setNumberFormat('#,##0.00 €');
    sheet.setFrozenRows(3);
    sheet.setColumnWidth(1,110); sheet.setColumnWidth(2,140); sheet.setColumnWidth(3,140);
    sheet.setColumnWidth(4,200); sheet.setColumnWidth(5,130);
    sheet.setColumnWidth(6,180); sheet.setColumnWidth(7,150); sheet.setColumnWidth(8,160);
  }
  return sheet;
}

function majTotalAide(ss, nomOnglet, couleurFond) {
  try {
    if (!ss) return;
    var sheet = ss.getSheetByName(nomOnglet);
    if (!sheet || sheet.getLastRow() < 4) return;
    var total = 0;
    var data = sheet.getRange(4, 5, sheet.getLastRow()-3, 1).getValues();
    data.forEach(function(row){ total += Number(row[0]) || 0; });
    total = Math.round(total * 100) / 100;
    sheet.getRange(3,6).setValue(total).setFontWeight('bold')
      .setFontColor(couleurFond).setBackground('#eeeeee')
      .setNumberFormat('#,##0.00 €');
    Logger.log('Total ' + nomOnglet + ' : ' + total + ' €');
  } catch(e) { Logger.log('majTotalAide KO (' + nomOnglet + ') : ' + e); }
}

function ecrireAideSheet(ss, nomOnglet, titre, couleurFond, couleurTexte, labelTotal,
                          code, nom, prenom, email, montant, detail, dateJour) {
  try {
    var sh = getOrCreateAideSheet(ss, nomOnglet, titre, couleurFond, couleurTexte, labelTotal);
    // Anti-doublon : si le code dossier est déjà présent, mettre à jour le montant si 0, sinon ignorer
    if (sh.getLastRow() >= 4) {
      var existing = sh.getRange(4, 1, sh.getLastRow() - 3, 5).getValues();
      for (var ii = 0; ii < existing.length; ii++) {
        if (String(existing[ii][0]).trim() === code) {
          var existMontant = Number(existing[ii][4]);
          if (existMontant === 0 && montant > 0) {
            // Mise à jour : remplacer le montant 0 par le vrai montant
            sh.getRange(ii + 4, 5).setValue(montant).setFontColor(couleurFond).setFontWeight('bold').setNumberFormat('#,##0.00 €');
            sh.getRange(ii + 4, 6).setValue(detail);
            majTotalAide(ss, nomOnglet, couleurFond);
            Logger.log('ecrireAideSheet — montant mis à jour pour ' + code + ' (' + nomOnglet + ') : ' + montant);
          } else {
            Logger.log('ecrireAideSheet — doublon ignoré : ' + code + ' (' + nomOnglet + ')');
          }
          return;
        }
      }
    }
    var nextRow = Math.max(sh.getLastRow() + 1, 4);
    var bgPastelMap2 = {
      '#1565c0':'#eff6ff', '#1b5e20':'#f0fdf4', '#6a1b9a':'#faf5ff',
      '#e65100':'#fff7ed', '#795548':'#fdf8f6', '#0d47a1':'#eff6ff',
      '#4a148c':'#faf5ff'
    };
    var bg = nextRow % 2 === 0 ? (bgPastelMap2[couleurFond] || '#f9f9f9') : '#ffffff';
    sh.getRange(nextRow, 1, 1, 8).setValues([[
      code, nom, prenom, email, montant, detail, dateJour, '⏳ En cours de validation'
    ]]).setBackground(bg);
    sh.getRange(nextRow, 5).setFontColor(couleurFond).setFontWeight('bold').setNumberFormat('#,##0.00 €');
    sh.getRange(nextRow, 8).setFontColor('#856404').setFontWeight('bold');
    sh.autoResizeColumns(1, 8);
    majTotalAide(ss, nomOnglet, couleurFond);
  } catch(e) { Logger.log('ecrireAideSheet KO (' + nomOnglet + ') : ' + e); }
}

// ── ANCV ──────────────────────────────────────────────────────
function majOngletAideANCv(ss, rows) {
  if (!ss || !rows || rows.length === 0) return;
  var r0 = rows[0];
  var code = r0.code_dossier || '';
  var ancvMontant = parseFloat(r0.ancv_montant) || 0;
  if (!ancvMontant) { var paA=String(r0.pass_aide||''); var mA=paA.match(/(?:^|\|)ANCV:([\d.]+)/); if(mA) ancvMontant=parseFloat(mA[1])||0; }
  if (ancvMontant <= 0) return;
  // Anti-doublon par code dossier
  var sh = getOrCreateAideSheet(ss, SHEET_AIDE_ANCV, 'Aides Coupon Sport ANCV', '#1565c0', '#ffffff', 'TOTAL ANCV déduit');
  if (sh && sh.getLastRow() >= 4) {
    var existing = sh.getRange(4, 1, sh.getLastRow()-3, 1).getValues();
    for (var i=0; i<existing.length; i++) { if (String(existing[i][0]).trim() === code) { Logger.log('majOngletAideANCv — doublon ignoré : '+code); return; } }
  }
  var nom     = r0.responsable_nom || r0.membre_nom || '';
  var prenom  = r0.responsable_prenom || r0.membre_prenom || '';
  var email   = r0.email1 || '';
  var dateJour = r0.date || Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  ecrireAideSheet(ss, SHEET_AIDE_ANCV,
    'Aides Coupon Sport ANCV', '#1565c0', '#ffffff', 'TOTAL ANCV déduit',
    code, nom, prenom, email, ancvMontant, 'Coupon ANCV', dateJour);
}

// ── Pass Sport État ──────────────────────────────────────────
function majOngletAidePassSport(ss, rows) {
  if (!ss || !rows || rows.length === 0) return;
  var r0 = rows[0];
  var code = r0.code_dossier || '';
  var passSportMontant = parseFloat(r0.pass_sport_montant) || 0;
  if (!passSportMontant) { var paP=String(r0.pass_aide||''); var mP=paP.match(/(?:^|\|)PASS:([\d.]+)/); if(mP) passSportMontant=parseFloat(mP[1])||0; }
  if (passSportMontant <= 0) return;
  // Anti-doublon par code dossier
  var sh = getOrCreateAideSheet(ss, SHEET_AIDE_PASS_S, "Aides Pass'sport État", '#1b5e20', '#ffffff', "TOTAL Pass'sport déduit");
  if (sh && sh.getLastRow() >= 4) {
    var existing = sh.getRange(4, 1, sh.getLastRow()-3, 1).getValues();
    for (var i=0; i<existing.length; i++) { if (String(existing[i][0]).trim() === code) { Logger.log("majOngletAidePassSport — doublon ignoré : "+code); return; } }
  }
  var nom    = r0.responsable_nom || r0.membre_nom || '';
  var prenom = r0.responsable_prenom || r0.membre_prenom || '';
  var email  = r0.email1 || '';
  var dateJour = r0.date || Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  ecrireAideSheet(ss, SHEET_AIDE_PASS_S,
    "Aides Pass'sport État", '#1b5e20', '#ffffff', "TOTAL Pass'sport déduit",
    code, nom, prenom, email, passSportMontant, "Pass'sport État — 70 €", dateJour);
}

// ── Pass Jeunes 76 ───────────────────────────────────────────
function majOngletAidePassJeunes(ss, rows) {
  if (!ss || !rows || rows.length === 0) return;
  var r0 = rows[0];
  var code = r0.code_dossier || '';
  var passAideStr = String(r0.pass_aide || '');
  var pjM = passAideStr.match(/PassJeunes:([\d.]+):(\w+)/);
  if (!pjM) return;
  var pjMontant = parseFloat(pjM[1]) || 0;
  var pjType    = pjM[2] === '1ere' ? '1ère inscription (30 €)' : '2ème inscription (20 €)';
  if (pjMontant <= 0) return;
  // Anti-doublon par code dossier
  var sh = getOrCreateAideSheet(ss, SHEET_AIDE_PASS_J, "Aides Pass'jeunes 76 / Handipass'sport", '#6a1b9a', '#ffffff', "TOTAL Pass'jeunes déduit");
  if (sh && sh.getLastRow() >= 4) {
    var existing = sh.getRange(4, 1, sh.getLastRow()-3, 1).getValues();
    for (var i=0; i<existing.length; i++) { if (String(existing[i][0]).trim() === code) { Logger.log("majOngletAidePassJeunes — doublon ignoré : "+code); return; } }
  }
  var nom    = r0.responsable_nom || r0.membre_nom || '';
  var prenom = r0.responsable_prenom || r0.membre_prenom || '';
  var email  = r0.email1 || '';
  var dateJour = r0.date || Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  ecrireAideSheet(ss, SHEET_AIDE_PASS_J,
    "Aides Pass'jeunes 76 / Handipass'sport", '#6a1b9a', '#ffffff', "TOTAL Pass'jeunes déduit",
    code, nom, prenom, email, pjMontant, "Pass'jeunes 76 — " + pjType, dateJour);
}

// ── Atout Normandie ───────────────────────────────────────────
function majOngletAideAtout(ss, rows) {
  if (!ss || !rows || rows.length === 0) return;
  var r0 = rows[0];
  var code = r0.code_dossier || '';
  var passAideStr = String(r0.pass_aide || '');
  var atoutM = passAideStr.match(/Atout:([\d.]+):?(\d*)/);
  if (!atoutM) return;
  var atoutMontant = parseFloat(atoutM[1]) || 0;
  var atoutCode    = atoutM[2] || '';
  if (atoutMontant <= 0) return;
  // Anti-doublon : ne pas écrire si le code dossier est déjà présent
  var sh = getOrCreateAideSheet(ss, SHEET_AIDE_ATOUT, 'Aides Atout Normandie', '#e65100', '#ffffff', 'TOTAL Atout Normandie déduit');
  if (sh && sh.getLastRow() >= 4) {
    var existing = sh.getRange(4, 1, sh.getLastRow()-3, 1).getValues();
    for (var i=0; i<existing.length; i++) { if (String(existing[i][0]).trim() === code) { Logger.log('majOngletAideAtout — doublon ignoré : '+code); return; } }
  }
  var nom    = r0.responsable_nom || r0.membre_nom || '';
  var prenom = r0.responsable_prenom || r0.membre_prenom || '';
  var email  = r0.email1 || '';
  var dateJour = r0.date || Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  var detail = 'Atout Normandie — 30 €' + (atoutCode ? ' (code: ' + atoutCode + ')' : '');
  ecrireAideSheet(ss, SHEET_AIDE_ATOUT,
    'Aides Atout Normandie', '#e65100', '#ffffff', 'TOTAL Atout Normandie déduit',
    code, nom, prenom, email, atoutMontant, detail, dateJour);
}


// ══════════════════════════════════════════════════════════════
// ONGLET LICENCES FFTT
// ══════════════════════════════════════════════════════════════


// ══════════════════════════════════════════════════════════════
// Marquer tous les elements d'un dossier comme Valides dans leurs onglets respectifs
// Appelee par validerPaiementSheet quand l'admin valide le paiement global
function validerTousElementsPaiement(ss, code, modeLabel) {
  var onglets = Object.keys(COL_STATUT_MAP);
  var count = 0;
  onglets.forEach(function(nomOnglet) {
    var cfg = COL_STATUT_MAP[nomOnglet];
    var sh = ss.getSheetByName(nomOnglet);
    if (!sh || sh.getLastRow() < cfg.debut) return;
    var nbCols = Math.max(sh.getLastColumn(), cfg.col);
    var data = sh.getRange(cfg.debut, 1, sh.getLastRow()-cfg.debut+1, nbCols).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]||'').trim() !== code) continue;
      var stCur = String(data[i][cfg.col-1]||'').toLowerCase();
      if (stCur.indexOf('valid') >= 0) continue; // deja valide
      sh.getRange(i + cfg.debut, cfg.col)
        .setValue('✅ Validé — ' + (modeLabel||'Admin'))
        .setFontColor('#1b5e20').setFontWeight('bold');
      count++;
    }
  });
  Logger.log('validerTousElementsPaiement: ' + count + ' elements valides pour ' + code);
}

// ============================================================
// GESTION ELEMENTS PAIEMENT - Statut par onglet
// ============================================================
var COL_STATUT_MAP = {
  'Espèces':             {col:7, debut:4},
  'HelloAsso':           {col:9, debut:4},
  'Cheques 1':           {col:7, debut:4},
  'Cheques 2':           {col:7, debut:4},
  'Cheques 3':           {col:7, debut:4},
  'Avoirs utilisés':    {col:7, debut:4},
  'Licences FFTT':       {col:8, debut:4},
  'Aides ANCV':          {col:8, debut:4},
  'Aides Pass Sport':    {col:8, debut:4},
  'Aides Pass Jeunes':   {col:8, debut:4},
  'Aides Atout Normandie': {col:8, debut:4}
};

function getElementsPaiementGAS(code) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var resultats = [];
    var onglets = Object.keys(COL_STATUT_MAP);
    onglets.forEach(function(nomOnglet) {
      var cfg = COL_STATUT_MAP[nomOnglet];
      var sh = ss.getSheetByName(nomOnglet);
      if (!sh || sh.getLastRow() < cfg.debut) return;
      var nbCols = Math.max(sh.getLastColumn(), cfg.col);
      var data = sh.getRange(cfg.debut, 1, sh.getLastRow()-cfg.debut+1, nbCols).getValues();
      data.forEach(function(row, i) {
        if (String(row[0]||'').trim() !== code) return;
        // Structure selon l'onglet :
        // Chèques : A=code B=nom C=prénom D=banque E=numCheque F=montant G=statut → montant=row[5]
        // Espèces : A=code B=nom C=prénom D=email  E=montant  F=date    G=statut → montant=row[4]
        var isCheque = nomOnglet.indexOf('Cheque') >= 0 || nomOnglet.indexOf('Chèque') >= 0;
        var montIdx = isCheque ? 5 : 4;
        var montRaw = row[montIdx];
        var montEl = (typeof montRaw === 'number' && montRaw > 0 && montRaw < 1e10) ? montRaw
                   : (typeof montRaw === 'string' && parseFloat(montRaw) > 0) ? parseFloat(montRaw) : 0;
        var statutEl = String(row[cfg.col-1]||'');
        // Compter les occurrences precedentes du meme onglet pour ce code
        var occCount = resultats.filter(function(x){return x.onglet===nomOnglet;}).length;
        var ongletLabel = occCount > 0
          ? nomOnglet + ' (' + (montEl>0?montEl.toFixed(2)+' EUR':'#'+(occCount+1)) + ')'
          : nomOnglet;
        resultats.push({
          onglet:  ongletLabel,
          ongletReel: nomOnglet,
          ligne:   i + cfg.debut,
          nom:     String(row[1]||''),
          prenom:  String(row[2]||''),
          montant: montEl,
          detail:  String(row[5]||row[6]||nomOnglet),
          statut:  statutEl || 'En attente'
        });
      });
    });
    return {status:'ok', elements: resultats};
  } catch(e) { return {status:'error', message:e.toString()}; }
}

function validerElementPaiementGAS(code, onglet, ligneTxt) {
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    // onglet peut etre un label avec montant ex: 'Cheques 1 (200.00 EUR)'
    // Chercher le vrai nom d'onglet dans COL_STATUT_MAP
    var ongletReel = onglet;
    if (!COL_STATUT_MAP[onglet]) {
      var keys = Object.keys(COL_STATUT_MAP);
      for (var ki=0;ki<keys.length;ki++) {
        if (onglet.indexOf(keys[ki])===0) { ongletReel = keys[ki]; break; }
      }
    }
    var cfg = COL_STATUT_MAP[ongletReel];
    if (!cfg) return {status:'error', message:'Onglet inconnu: '+onglet};
    onglet = ongletReel; // utiliser le vrai nom pour getSheetByName
    var sh = ss.getSheetByName(onglet);
    if (!sh) return {status:'error', message:'Onglet introuvable: '+onglet};
    var ligne = parseInt(ligneTxt);
    if (isNaN(ligne) || ligne < cfg.debut) return {status:'error', message:'Ligne invalide'};
    sh.getRange(ligne, cfg.col).setValue('✅ Validé').setFontColor('#1b5e20').setFontWeight('bold');
    Logger.log('validerElementPaiement OK: '+code+' / '+onglet+' / ligne '+ligne);
    // Verifier si tous les elements du dossier sont valides
    var toutValide = true;
    var onglets = Object.keys(COL_STATUT_MAP);
    onglets.forEach(function(nomOnglet) {
      var cfg2 = COL_STATUT_MAP[nomOnglet];
      var sh2 = ss.getSheetByName(nomOnglet);
      if (!sh2 || sh2.getLastRow() < cfg2.debut) return;
      var nbC = Math.max(sh2.getLastColumn(), cfg2.col);
      var d2 = sh2.getRange(cfg2.debut,1,sh2.getLastRow()-cfg2.debut+1,nbC).getValues();
      d2.forEach(function(row) {
        if (String(row[0]||'').trim() !== code) return;
        var st = String(row[cfg2.col-1]||'').toLowerCase();
        if (st.indexOf('valid') < 0) toutValide = false;
      });
    });
    // Si tout valide -> marquer Inscrit dans Inscriptions (sauf lignes en attente de place)
    if (toutValide) {
      var shInscr = ss.getSheetByName(SHEET_INSCRIPTIONS);
      if (shInscr && shInscr.getLastRow() > 1) {
        var inscData = shInscr.getRange(2,1,shInscr.getLastRow()-1,41).getValues();
        for (var ii=0; ii<inscData.length; ii++) {
          if (String(inscData[ii][19]||'').trim() !== code) continue;
          var stLine = String(inscData[ii][39]||'').toLowerCase();
          if (stLine.indexOf('attente') < 0) {
            shInscr.getRange(ii+2,40).setValue('✅ Inscrit').setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
          }
        }
        Logger.log('Tout validé pour '+code+' -> Inscrit');
      }
    }
    return {status:'ok', toutValide:toutValide};
  } catch(e) { return {status:'error', message:e.toString()}; }
}

function validerPaiementBasculeGAS(code, actId, modePaiement, montant, commentaireAdmin, montantModifie) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');
    var shInscr = ss.getSheetByName(SHEET_INSCRIPTIONS);
    // Col indices (0-based): 0=NumLic, 1=Civilite, 2=Nom, 3=Prenom, 4=DDN, 5=Appart,
    // 14=Tel portable, 15=Email, 19=Code dossier, 20=Date inscription
    var nom='',prenom='',email='';
    if (shInscr && shInscr.getLastRow()>1) {
      var d=shInscr.getRange(2,1,shInscr.getLastRow()-1,21).getValues();
      for (var i=0;i<d.length;i++) {
        if (String(d[i][19]||'').trim()!==code) continue;
        nom    = String(d[i][2]||'');   // col C = Nom
        prenom = String(d[i][3]||'');   // col D = Prenom
        email  = String(d[i][15]||'');  // col P = Email
        break;
      }
    }
    var modeLabel = modePaiement==='cheque'?'Chèque':modePaiement==='especes'?'Espèces':modePaiement==='helloasso'?'HelloAsso':modePaiement;

    // Ecrire dans l'onglet paiement selon le mode
    if (modePaiement==='especes') {
      var shE=getOrCreateEspecesSheet(ss);var nrE=Math.max(shE.getLastRow()+1,4);
      shE.getRange(nrE,1,1,7).setValues([[code,nom,prenom,email,montant,dateJour,'✅ Validé — Admin']]);
      shE.getRange(nrE,5).setFontColor('#4e342e').setFontWeight('bold');
      shE.getRange(nrE,7).setFontColor('#1b5e20').setFontWeight('bold');
      majTotalEspeces(ss);
    } else if (modePaiement==='cheque') {
      var shC=getOrCreateChequeSheet(ss,SHEET_CHEQUE_1);var nrC=Math.max(shC.getLastRow()+1,4);
      shC.getRange(nrC,1,1,7).setValues([[code,nom,prenom,'','',montant,'✅ Validé — Admin']]);
      shC.getRange(nrC,6).setFontColor('#1b5e20').setFontWeight('bold');
      shC.getRange(nrC,7).setFontColor('#1b5e20').setFontWeight('bold');
      majTotalCheque(ss,SHEET_CHEQUE_1);
    } else if (modePaiement==='helloasso') {
      var shH=getOrCreateHelloAssoSheet(ss);var nrH=Math.max(shH.getLastRow()+1,4);
      shH.getRange(nrH,1,1,9).setValues([[code,nom,prenom,email,montant,montant,'HelloAsso',dateJour,'✅ Validé — Admin']]);
      shH.getRange(nrH,6).setFontColor('#1b5e20').setFontWeight('bold');
      majTotalHelloAsso(ss);
    }
    // MAJ onglet LA-{actId}
    var placesId=getPlacesId(actId)||actId;
    var nomLA='LA-'+placesId.replace(/[\/\\:?*[\]]/g,'-').substring(0,28);
    var shLA=ss.getSheetByName(nomLA);
    if (shLA && shLA.getLastRow()>1) {
      var laD=shLA.getRange(2,1,shLA.getLastRow()-1,3).getValues();
      for (var j=0;j<laD.length;j++) {
        if (String(laD[j][0]||'').trim()!==code) continue;
        shLA.getRange(j+2,3).setValue('✅ Payé — '+modeLabel).setFontColor('#1b5e20').setFontWeight('bold');
        shLA.getRange(j+2,1,1,10).setBackground('#d8f3dc'); break;
      }
    }

    // MAJ col 22 Inscriptions pour la ligne en attente
    if (shInscr && shInscr.getLastRow()>1) {
      var d2=shInscr.getRange(2,1,shInscr.getLastRow()-1,41).getValues();
      for (var k=0;k<d2.length;k++) {
        if (String(d2[k][19]||'').trim()!==code) continue;
        var stk=String(d2[k][39]||'').toLowerCase();
        if (stk.indexOf('attente')<0) continue;
        // L'ID d'activité peut se trouver dans plusieurs colonnes selon la structure de la ligne
        // (AJ=35, AL=37, AK=36, ou le nom en W=22) — cf. basculerListeAttenteGAS
        var candidatesV = [
          String(d2[k][35]||'').trim(), // AJ
          String(d2[k][37]||'').trim(), // AL
          String(d2[k][36]||'').trim(), // AK
          String(d2[k][22]||'').trim()  // W = nom activité
        ];
        var matchActV = !actId || candidatesV.some(function(c){ return c === actId; });
        if (!matchActV) continue;
        shInscr.getRange(k+2,22).setValue('✅ Règlement validé — '+modeLabel).setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
        break;
      }
    }
    // MAJ onglet ID activite (ex: FITM10) col 3 = statut
    try {
      var actSheet = ss.getSheetByName(placesId) || ss.getSheetByName(actId);
      if (actSheet && actSheet.getLastRow() >= 3) {
        var aData = actSheet.getRange(3, 1, actSheet.getLastRow()-2, 3).getValues();
        for (var m=0; m<aData.length; m++) {
          if (String(aData[m][0]||'').trim() !== code) continue;
          actSheet.getRange(m+3, 3).setValue('✅ Payé — ' + modeLabel)
            .setBackground('#d8f3dc').setFontColor('#2d6a4f').setFontWeight('bold');
          Logger.log('✅ Onglet activite ' + placesId + ' maj pour ' + code);
          break;
        }
      }
    } catch(eAct) { Logger.log('MAJ onglet activite bascule KO: ' + eAct); }

    logCommentaireAdmin(ss, 'Validation règlement bascule', code, nom, prenom, actId, montant,
      commentaireAdmin, montantModifie === true || montantModifie === 'true');
    Logger.log('validerPaiementBasculeGAS OK: '+code+' '+modePaiement+' '+montant+'EUR');
    return {status:'ok'};
  } catch(e) { Logger.log('validerPaiementBasculeGAS KO: '+e.toString()); return {status:'error',message:e.toString()}; }
}

// ============================================================
// LOOKUP LICENCES FFTT 25-26
// ============================================================
// Recherche adherent FRI annee precedente par 4 premiers chiffres du telephone (col L)
// Recherche d'un adhérent de la saison précédente pour pré-remplir le formulaire.
// SÉCURITÉ : action publique → numéro complet (10 chiffres) ET 3 premières lettres du nom
// exigés, et seuls les adhérents correspondant aux deux critères sont renvoyés.
function rechercherAdherentParTel(telPartiel, nomDebut) {
  try {
    var telChiffres = String(telPartiel || '').replace(/[^0-9]/g, '');
    if (telChiffres.length < 9) return {status:'error', message:'Saisissez votre numéro de téléphone complet'};
    var normNom = function(v) { return String(v || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z]/g, ''); };
    var nomRecherche = normNom(nomDebut).substring(0, 3);
    if (nomRecherche.length < 3) return {status:'error', message:'Saisissez les 3 premières lettres de votre nom'};
    var ss = SpreadsheetApp.openById(SHEET_ADHERENTS_2526);
    var sheet = ss.getSheets()[0]; // premier onglet
    var lr = sheet.getLastRow();
    if (lr < 2) return {status:'ok', adherents:[]};
    var data = sheet.getRange(2, 1, lr-1, 20).getValues();
    // Col L = index 11 = telephone
    var telSearch = telChiffres.replace(/^(33|0)(?=\d{9}$)/, '');
    var resultats = [];
    for (var i=0; i<data.length; i++) {
      var row = data[i];
      // Col L (index 11) = telephone principal, col K (index 10) = telephone secondaire
      var tel  = String(row[11]||'').replace(/[^0-9]/g,''); // col L
      var tel2 = String(row[10]||'').replace(/[^0-9]/g,''); // col K
      // Numéro identique (col L ou col K, avec ou sans le 0) ET même début de nom
      var matched = false;
      [tel, tel2].forEach(function(t) {
        if (t && t.replace(/^(33|0)(?=\d{9}$)/, '') === telSearch) matched = true;
      });
      if (matched && normNom(row[1]).indexOf(nomRecherche) !== 0) matched = false;
      if (matched) {
        var telRaw1 = String(row[11]||'').trim(); // col L = Tel (avec 0)
        // Le numero est deja avec le 0 en col L
        var tel1fmt = telRaw1.replace(/[^0-9]/g,''); // garder uniquement les chiffres
        // Formater en XX XX XX XX XX
        if (tel1fmt.replace(/\D/g,'').length === 10) {
          var d = tel1fmt.replace(/\D/g,'');
          tel1fmt = d.replace(/(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/, '$1 $2 $3 $4 $5');
        }
        resultats.push({
          nom:     String(row[1]||'').trim(),   // col B = Nom
          prenom:  String(row[2]||'').trim(),   // col C = Prenom
          ddn:     String(row[3]||'').trim(),   // col D = Date naissance
          adresse: String(row[6]||'').trim(),   // col G = Adresse
          cp:      String(row[7]||'').trim(),   // col H = CP
          ville:   String(row[8]||'').trim(),   // col I = Ville
          email1:  String(row[9]||'').trim(),   // col J = Email
          tel1:    tel1fmt,                     // col L = Tel (avec 0 + format)
          tel2:    '',
          email2:  ''
        });
        if (resultats.length >= 5) break; // max 5 resultats
      }
    }
    Logger.log('rechercherAdherentParTel -> ' + resultats.length + ' resultat(s)');
    return {status:'ok', adherents:resultats};
  } catch(e) {
    Logger.log('rechercherAdherentParTel KO: '+e);
    return {status:'error', message:e.toString()};
  }
}

function getLicenceFFTT(nom, prenom, numLicence) {
  try {
    var ss2526 = SpreadsheetApp.openById(SHEET_LICENCES_2526);
    var sheets  = ss2526.getSheets();
    if (!sheets || sheets.length === 0) return { found: false, message: 'Sheet licences vide' };
    var sheet = sheets[0];
    if (sheet.getLastRow() < 2) return { found: false, message: 'Aucune donnee' };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).trim().toLowerCase(); });
    function col(kws) {
      for (var ki=0; ki<kws.length; ki++)
        for (var hi=0; hi<headers.length; hi++)
          if (headers[hi].indexOf(kws[ki]) >= 0) return hi;
      return -1;
    }
    var cNum  = col(['numero','licence','lic','n°']);
    var cNom  = col(['nom']);
    var cPren = col(['prénom','prenom']);
    var cDdn  = col(['naissance','ddn','birth','né']);
    var cSexe = col(['sexe','sex']);
    var cNatl = col(['nationalité','nationalite']);
    var cVNaiss = col(['ville nais','commune nais']);
    var cCPN   = col(['cp nais','code postal nais']);
    var cPaysN = col(['pays nais']);
    var cNomN  = col(['nom de nais','nom naiss']);
    var cEmail = col(['email','mail']);
    var cTel   = col(['téléphone','telephone','tel']);
    var cCateg  = col(['catégorie','categorie','categ']);
    var cPoints = col(['points','classement','pts']);
    if (cPoints < 0) cPoints = 15; // colonne P = index 15 par defaut
    Logger.log('getLicenceFFTT cols nom='+cNom+' pren='+cPren+' num='+cNum+' points='+cPoints);
    var data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    var nomN  = nom.trim().toLowerCase().replace(/[^a-z]/gi,'');
    var prenN = prenom.trim().toLowerCase().replace(/[^a-z]/gi,'');
    var numN  = String(numLicence||'').trim().replace(/\s+/g,'');
    for (var i=0; i<data.length; i++) {
      var row = data[i];
      var rowNum  = cNum  >= 0 ? String(row[cNum]||'').trim().replace(/\s+/g,'') : '';
      var rowNom  = cNom  >= 0 ? String(row[cNom]||'').trim().toLowerCase().replace(/[^a-z]/gi,'') : '';
      var rowPren = cPren >= 0 ? String(row[cPren]||'').trim().toLowerCase().replace(/[^a-z]/gi,'') : '';
      var match = false;
      if (numN && rowNum && rowNum === numN) { match = true; }
      else if (!numN && nomN && prenN && rowNom === nomN && rowPren === prenN) { match = true; }
      else if (!numN && nomN && prenN && rowNom.indexOf(nomN)===0 && rowPren.indexOf(prenN)===0) { match = true; }
      if (!match) continue;
      function v(c){ return c>=0 ? String(row[c]||'').trim() : ''; }
      var ddnRaw = cDdn >= 0 ? row[cDdn] : '';
      var ddnFmt = '';
      if (ddnRaw) {
        try {
          var d = (ddnRaw instanceof Date) ? ddnRaw : new Date(ddnRaw);
          if (!isNaN(d.getTime())) ddnFmt = Utilities.formatDate(d,'Europe/Paris','dd/MM/yyyy');
          else ddnFmt = String(ddnRaw).trim();
        } catch(ed){ ddnFmt = String(ddnRaw).trim(); }
      }
      var annee = 0;
      try{ annee = parseInt((ddnFmt||'').split('/')[2])||0; }catch(ea){}
      var categAge = annee<=1986?'veteran':annee<=2007?'senior':annee<=2011?'junior':annee<=2013?'cadet':annee<=2015?'minime':annee<=2017?'benjamin':'poussin';
      return {
        found:true,
        numLicence: v(cNum),
        nom:        cNom>=0?String(row[cNom]||'').trim().toUpperCase():'',
        prenom:     v(cPren),
        ddn:        ddnFmt,
        sexe:       v(cSexe),
        nationalite: v(cNatl)||'Française',
        nomNaissance: v(cNomN)||(cNom>=0?String(row[cNom]||'').trim().toUpperCase():''),
        villeNaissance: v(cVNaiss),
        cpNaissance:    v(cCPN),
        paysNaissance:  v(cPaysN)||'France',
        email: v(cEmail),
        tel:   v(cTel),
        categAge: categAge,
        points:   cPoints >= 0 ? String(row[cPoints]||'').trim() : '',
        source: 'Licences FFTT 2025-2026'
      };
    }
    return { found:false, message:'Licencié introuvable : '+nom+' '+prenom+(numN?' (N°'+numN+')':'') };
  } catch(e) {
    Logger.log('getLicenceFFTT KO: '+e.toString());
    return { found:false, message:'Erreur: '+e.toString() };
  }
}

// ============================================================
// UPLOAD CERTIFICAT MEDICAL vers Drive > dossier '1-Certificats médicaux'
// ============================================================
function uploadJustificatifGAS(code, nom, prenom, filename, mimeType, fileBase64) {
  try {
    if (!fileBase64 || fileBase64.length < 10) return { status: 'ok', message: 'Aucun fichier' };
    var decoded = Utilities.base64Decode(fileBase64);
    var blob    = Utilities.newBlob(decoded, mimeType, filename);
    var dossier = obtenirDossierFRI('6-Justificatifs');
    var fichier = dossier.createFile(blob);
    securiserFichier(fichier);
    Logger.log('Justificatif Drive OK : ' + filename + ' (' + decoded.length + ' bytes) — code:' + code);
    return { status: 'ok', message: 'Justificatif enregistré : ' + filename };
  } catch(e) {
    Logger.log('uploadJustificatifGAS KO : ' + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

function uploadCertificatMedicalGAS(code, nom, prenom, filename, mimeType, fileBase64) {
  try {
    if (!fileBase64 || fileBase64.length < 10) return { status: 'ok', message: 'Aucun fichier' };
    var decoded  = Utilities.base64Decode(fileBase64);
    var blob     = Utilities.newBlob(decoded, mimeType, filename);
    var dossier  = obtenirDossierFRI('1-Certificats médicaux');
    var fichier  = dossier.createFile(blob);
    securiserFichier(fichier);
    if (/^FRI-[A-Z0-9]{4}$/.test(code)) marquerCertificatTransmis(code, nom, prenom);
    Logger.log('Certificat medical Drive OK : ' + filename + ' (' + decoded.length + ' bytes)');
    return { status: 'ok', message: 'Certificat enregistre : ' + filename };
  } catch(e) {
    Logger.log('uploadCertificatMedicalGAS KO : ' + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

// ============================================================
// BORDEREAU LICENCE FFTT -- Generation PDF + Drive + Email
// ============================================================
// ============================================================
// RECHERCHE LICENCIÉ FFTT dans le sheet saison 25/26
// Colonnes supposées (à adapter selon le vrai fichier) :
// A=NumLicence B=Nom C=Prénom D=DDN E=Sexe F=NomNaiss
// G=Nationalité H=PaysNaiss I=VilleNaiss J=CPNaiss K=Adresse L=CP M=Ville
// ============================================================
function rechercherLicencieFFTT(nom, prenom, ddn, numLic) {
  try {
    var SHEET_LIC_ID = '1LvuVT9diwtI4asm7ReFZyeAMlgwXJ99E';
    var ss2 = SpreadsheetApp.openById(SHEET_LIC_ID);
    var sheet = ss2.getSheets()[0]; // première feuille
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { trouve: false, message: 'Feuille vide' };

    // Lire les en-têtes pour mapper dynamiquement les colonnes
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function(h){ return String(h).toLowerCase().trim(); });

    // Index des colonnes utiles (chercher par nom d'en-tête)
    function col(patterns) {
      for (var pi = 0; pi < patterns.length; pi++) {
        for (var hi = 0; hi < headers.length; hi++) {
          if (headers[hi].indexOf(patterns[pi]) >= 0) return hi;
        }
      }
      return -1;
    }
    var iNum     = col(['num','licence','lic','n°']);
    var iNom     = col(['nom']);
    var iPrenom  = col(['prenom','prénom']);
    var iDDN     = col(['naissance','ddn','né','ne']);
    var iSexe    = col(['sexe','genre']);
    var iNomNaiss= col(['nom naiss','naissance nom','nom de naiss']);
    var iNatio   = col(['national']);
    var iPaysN   = col(['pays naiss','pays de naiss']);
    var iVilleN  = col(['ville naiss','ville de naiss']);
    var iCPN     = col(['cp naiss','cp de naiss','cp de n']);
    var iAddr    = col(['adresse','adres']);
    var iCP      = col(['cp','code postal','code post']);
    var iVille   = col(['ville','commune']);
    var iEmail   = col(['email','mail','courriel']);
    var iTel     = col(['tel','tél','phone','portable']);
    var iCateg   = col(['categ','catég','age','âge']);

    // Si l'index nom et num sont confondus, affiner
    if (iNom === iNum) iNom = -1;

    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    function norm(s) { return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,''); }
    function matchDate(cellVal, ddnStr) {
      if (!ddnStr) return true; // pas de filtre DDN
      var cellStr = '';
      if (cellVal instanceof Date) cellStr = Utilities.formatDate(cellVal, 'Europe/Paris', 'dd/MM/yyyy');
      else cellStr = String(cellVal||'');
      return cellStr.replace(/[^0-9]/g,'').indexOf(ddnStr.replace(/[^0-9]/g,'')) >= 0;
    }

    var candidates = [];
    for (var ri = 0; ri < data.length; ri++) {
      var row = data[ri];
      var rowNom    = iNom    >= 0 ? String(row[iNom]||'')    : '';
      var rowPrenom = iPrenom >= 0 ? String(row[iPrenom]||'') : '';
      var rowNum    = iNum    >= 0 ? String(row[iNum]||'')    : '';
      var rowDDN    = iDDN    >= 0 ? row[iDDN]                : '';

      // Correspondance par numéro de licence (prioritaire)
      if (numLic && rowNum && norm(rowNum) === norm(numLic)) {
        candidates.unshift(ri); continue;
      }
      // Correspondance nom + prénom
      var matchNom    = nom    && norm(rowNom).indexOf(norm(nom))    >= 0;
      var matchPrenom = prenom && norm(rowPrenom).indexOf(norm(prenom)) >= 0;
      if (matchNom && matchPrenom && matchDate(rowDDN, ddn)) {
        candidates.push(ri);
      } else if (matchNom && matchPrenom && !ddn) {
        candidates.push(ri);
      }
    }

    if (candidates.length === 0) {
      return { trouve: false, message: 'Licencié non trouvé dans la base 2025/26.' };
    }

    // Prendre le premier candidat
    var r = data[candidates[0]];
    function val(idx) { return idx >= 0 ? String(r[idx]||'').trim() : ''; }
    function valDate(idx) {
      if (idx < 0) return '';
      var v = r[idx];
      if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Paris', 'dd/MM/yyyy');
      return String(v||'').trim();
    }

    var result = {
      trouve:       true,
      numLicence:   val(iNum),
      nom:          val(iNom),
      prenom:       val(iPrenom),
      ddn:          valDate(iDDN),
      sexe:         val(iSexe),
      nomNaissance: val(iNomNaiss) || val(iNom),
      nationalite:  val(iNatio)   || 'Française',
      paysNaissance:val(iPaysN)   || 'France',
      villeNaissance:val(iVilleN),
      cpNaissance:  val(iCPN),
      adresse:      val(iAddr),
      cp:           val(iCP),
      ville:        val(iVille),
      email:        val(iEmail),
      tel:          val(iTel),
      categAge:     val(iCateg),
      nbCandidats:  candidates.length
    };
    Logger.log('FFTT trouvé : ' + result.nom + ' ' + result.prenom + ' lic=' + result.numLicence);
    return result;
  } catch(e) {
    Logger.log('rechercherLicencieFFTT KO : ' + e.toString());
    return { trouve: false, message: 'Erreur recherche : ' + e.toString() };
  }
}

function saveBordereauFFTT(code, signatureBase64, extra, rows) {
  try {
    // Fallback code depuis extra si non transmis directement
    if (!code && extra && extra.code) code = String(extra.code);
    Logger.log('saveBordereauFFTT START code=' + code
      + ' sig_len=' + (signatureBase64 ? signatureBase64.length : 0)
      + ' rows=' + (rows ? rows.length : 0)
      + ' extra_keys=' + (extra ? Object.keys(extra).join(',') : 'none'));
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var r0 = rows && rows.length > 0 ? rows[0] : {};
    // Priorite : extra (direct depuis state), puis r0 (depuis rows Sheet)
    var nom         = extra.nom     || r0.responsable_nom    || '';
    var prenom      = extra.prenom  || r0.responsable_prenom || '';
    var email       = extra.email   || r0.email1             || '';
    var tel         = extra.tel     || r0.tel1               || '';
    var ddn         = extra.ddn     || r0.ddn                || '';
    var sexe        = extra.sexe    || r0.sexe               || '';
    var adresse     = extra.adresse || r0.adresse            || '';
    var cp          = extra.cp      || r0.cp                 || '';
    var ville       = extra.ville   || r0.ville              || '';
    var memNom      = extra.memNom    || r0.membre_nom    || nom;
    var memPrenom   = extra.memPrenom || r0.membre_prenom || prenom;
    var numLicence  = extra.numLicence      || '';
    var categAge    = extra.categAge        || '';
    var nomNaiss    = extra.nomNaissance    || nom;
    var nationalite = extra.nationalite     || 'Francaise';
    var paysNaiss   = extra.paysNaissance   || 'France';
    var villeNaiss  = extra.villeNaissance  || '';
    var cpNaiss     = extra.cpNaissance     || '';
    var certifMed   = extra.certifMed       || 'non';
    var assurance   = extra.assurance       || 'non';
    var dateJour    = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');
    // memNom et memPrenom lus ci-dessus depuis extra en priorite
    var ffttMontant = parseFloat(r0.fftt_price) || 0;
    var typeLic     = numLicence ? 'Renouvellement' : 'Nouvelle licence';

    // -- Creer le document PDF ------------------------------------
    // Ajouter timestamp pour unicite si code vide ou identique entre membres
    var tsNow = Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyyMMdd_HHmmss');
    var codeLabel = code || tsNow;
    var docTitle = 'BordereauFFTT_' + codeLabel + '_' + memPrenom + '_' + memNom;
    var doc  = DocumentApp.create(docTitle);
    var body = doc.getBody();
    body.setMarginTop(18).setMarginBottom(14).setMarginLeft(28).setMarginRight(28);

    // En-tete compact
    var t = body.appendParagraph('FOYER RURAL D\'ISNEAUVILLE  --  Bordereau FFTT  --  Saison 2026/2027');
    t.editAsText().setFontSize(9).setBold(true).setForegroundColor('#1b5e20');
    t.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(2);

    var infos = body.appendParagraph('N. Dossier : ' + codeLabel + '   |   Date : ' + dateJour + '   |   Montant : ' + (ffttMontant > 0 ? ffttMontant.toFixed(2) + ' EUR' : 'Inclus'));
    infos.editAsText().setFontSize(8).setBold(true).setForegroundColor('#1a2e22');
    infos.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(4);

    // Un tableau par section, plusieurs lignes = pas de gaps inter-tableaux
    var _curTable = null;
    function section(titre) {
      _curTable = null; // nouvelle section = nouveau tableau
      var p = body.appendParagraph(titre);
      p.editAsText().setFontSize(8).setBold(true).setForegroundColor('#1b5e20');
      p.setSpacingBefore(6).setSpacingAfter(0);
    }
    function champ(label, valeur) {
      if (!_curTable) {
        _curTable = body.appendTable();
        _curTable.setBorderWidth(0);
      }
      var tr = _curTable.appendTableRow(); tr.setMinimumHeight(1);
      var c1 = tr.appendTableCell(label); c1.setWidth(140);
      c1.editAsText().setFontSize(7).setForegroundColor('#666666').setBold(true);
      c1.setBackgroundColor('#f5f5f5');
      c1.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(3).setPaddingRight(3);
      var c2 = tr.appendTableCell(valeur || ''); c2.setWidth(330);
      c2.editAsText().setFontSize(8).setForegroundColor('#1a2e22');
      c2.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(3).setPaddingRight(3);
    }
    var sexeLabel = (sexe==='M'||sexe==='m')?'Masculin':(sexe==='F'||sexe==='f')?'Feminin':sexe;
    section('IDENTITE');
    champ('Nom de naissance', nomNaiss);
    champ('Nom / Prenom', memNom + ' ' + memPrenom);
    champ('DDN  |  Sexe', ddn + '  |  ' + sexeLabel);
    champ('Nationalite  |  Pays naiss.', nationalite + '  |  ' + paysNaiss);
    champ('Ville naiss.  |  CP naiss.', (villeNaiss||'nd') + '  |  ' + (cpNaiss||'nd'));

    section('COORDONNEES');
    champ('Adresse', adresse + (cp?', '+cp:'') + (ville?' '+ville:''));
    champ('Email  |  Tel.', (email||'nd') + '  |  ' + (tel||'nd'));

    var ptsDisplay = extra.points ? String(extra.points) + ' pts' : '';
    section('LICENCE');
    champ('Type  |  N. licence', typeLic + (numLicence?'  |  N.'+numLicence:''));
    champ(ptsDisplay?'Categorie  |  Points':'Categorie', categAge + (ptsDisplay?'  |  '+ptsDisplay:''));
    var certifLabels = {
      'certif':        '[X] 1. Je joins un certificat medical de pratique sportive de moins d\'un an - etabli par le medecin sur papier libre ou formulaire 26-9',
      'veteranAttest': '[X] 2. J\'ai fourni, si je suis veteran, un certificat medical lors de mon precedent changement de categorie sportive. J\'ai pratique sans discontinuite et je joins l\'attestation certifiant que j\'ai repondu NON a toutes les questions de l\'auto-questionnaire medical (formulaire 26-10-1)',
      'attestMaj':     '[X] 3. Majeur de moins de 40 ans : attestation NON a toutes les questions de l\'auto-questionnaire medical (formulaire 26-10-1)',
      'attestMin':     '[X] 4. Mineur : attestation NON a toutes les questions de l\'auto-questionnaire medical (formulaire 26-10-2)',
      'sansCertif':    '[X] 5. Je ne joins pas de certificat ni d\'attestation - licence ne permettant pas la pratique sportive',
      'oui':           '[X] 1. Certificat medical joint',
      'exempt':        '[X] Exempte'
    };
    var certifLabels2 = {
      'certif':        '[ ] 1.  [ ] 2.  [ ] 3.  [ ] 4.  [ ] 5.',
      'veteranAttest': '[X] 1.  [ ] 2.  [ ] 3.  [ ] 4.  [ ] 5.',
      'attestMaj':     '[ ] 1.  [ ] 2.  [X] 3.  [ ] 4.  [ ] 5.',
      'attestMin':     '[ ] 1.  [ ] 2.  [ ] 3.  [X] 4.  [ ] 5.',
      'sansCertif':    '[ ] 1.  [ ] 2.  [ ] 3.  [ ] 4.  [X] 5.'
    };
    var certifLabel = certifLabels[certifMed] || ('[ ] Non renseigne');
    var assLabel1 = (assurance === 'souscrit' || assurance === 'oui')
      ? '[X] 1. Je souhaite souscrire a la garantie facultative de base dommage corporel (0,38 EUR veterans/seniors - 0,10 EUR jeunes - inclus dans le tarif de la licence)'
      : (assurance === 'renonce')
      ? '[X] 2. Je ne souhaite pas souscrire a la garantie dommage corporel. Je demande le remboursement de la prime et renonce a toute indemnite. J\'ai ete informe des risques et pris connaissance des dispositions relatives a l\'assurance.'
      : ('[ ] Non renseigne');
    // Afficher certification et assurance sur des lignes separees
    champ('CERTIFICATION MEDICALE (cocher 1 case)', certifLabel);
    champ('ASSURANCE (cocher 1 case)', assLabel1);

    section('CONSENTEMENTS');
    champ('Optin assoc.  |  Optin part.', (extra.optinAssoc==='1'?'Accepte':'Refuse') + '  |  ' + (extra.optinPart==='1'?'Accepte':'Refuse'));
    champ('Refus honorar.  |  Refus photo', (extra.refusHonor==='1'?'Oui':'Non') + '  |  ' + (extra.refusPhoto==='1'?'Oui':'Non'));

    section('SIGNATURE  --  Lu et approuve le ' + dateJour);
    if (signatureBase64 && signatureBase64.length > 100) {
      try {
        // Nettoyer le prefixe data:image/...;base64, et les espaces
        var sigData = signatureBase64
          .replace(/^data:image\/[^;]+;base64,/, '')
          .replace(/\s/g, '')  // supprimer espaces/sauts de ligne
          .replace(/-/g, '+').replace(/_/g, '/'); // URL-safe base64
        // Padding
        while (sigData.length % 4 !== 0) sigData += '=';
        var sigBytes = Utilities.base64Decode(sigData);
        // Detecter le format (JPEG commence par FF D8, PNG par 89 50)
        var mime = (sigBytes[0] === 0xFF && sigBytes[1] === 0xD8) ? 'image/jpeg' : 'image/png';
        var sigBlob = Utilities.newBlob(sigBytes, mime, 'signature');
        var sigImg  = body.appendImage(sigBlob);
        sigImg.setWidth(120).setHeight(40);
        Logger.log('Signature inseree : ' + sigBytes.length + ' bytes ' + mime);
      } catch(eSig) {
        Logger.log('Signature image KO: ' + eSig + ' | longueur base64: ' + signatureBase64.length);
        body.appendParagraph('(Signature numerique fournie - ' + signatureBase64.length + ' chars)').editAsText().setFontSize(9).setItalic(true);
      }
    } else {
      Logger.log('Signature absente ou trop courte : ' + (signatureBase64 ? signatureBase64.length : 'null'));
      body.appendParagraph('(Signature a apposer en permanence)').editAsText().setFontSize(9).setItalic(true).setForegroundColor('#aaaaaa');
    }
    body.appendParagraph('');
    body.appendHorizontalRule();
    var pied = body.appendParagraph('Bordereau genere par le site FRI -- ' + NOM_ASSO + ' -- Saison 2026/2027');
    pied.editAsText().setFontSize(8).setForegroundColor('#aaaaaa').setItalic(true);
    pied.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    doc.saveAndClose();
    Utilities.sleep(3000); // laisser Drive indexer le Google Doc

    // -- Sauvegarder PDF dans Drive dossier "5-Bordereaux FFTT" -----
    var docId       = doc.getId();
    var pdfName     = 'BordereauFFTT_' + codeLabel + '_' + memNom + '_' + memPrenom + '_' + tsNow + '.pdf';
    var dossierFFTT = obtenirDossierFRI('5-Bordereaux FFTT');
    var docFile     = null;
    try {
      docFile = DriveApp.getFileById(docId);
      Logger.log('docFile MimeType: ' + docFile.getMimeType() + ' name: ' + docFile.getName());
      // Exporter en PDF via l'API Drive export URL
      var exportUrl = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';
      var token = ScriptApp.getOAuthToken();
      var pdfResp = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (pdfResp.getResponseCode() === 200) {
        var pdfBlob = pdfResp.getBlob().setName(pdfName).setContentType('application/pdf');
        dossierFFTT.createFile(pdfBlob);
        Logger.log('Bordereau FFTT PDF OK via export URL : ' + pdfName);
      } else {
        Logger.log('Export PDF HTTP ' + pdfResp.getResponseCode() + ' - fallback getAs');
        var pdfBlob2 = docFile.getAs('application/pdf').setName(pdfName);
        dossierFFTT.createFile(pdfBlob2);
        Logger.log('Bordereau FFTT PDF OK via getAs : ' + pdfName);
      }
    } catch(ePdf) {
      Logger.log('PDF generation KO : ' + ePdf.toString());
    } finally {
      // Toujours supprimer le Google Doc temporaire
      try { if (docFile) docFile.setTrashed(true); } catch(et) {}
      Logger.log('Bordereau FFTT Drive termine : ' + pdfName);
    }

    // -- Onglet Licences FFTT : deja ecrit a l'inscription via addRegistration
    // Ne pas recrire ici pour eviter les doublons (sans code dossier)

    return { status: 'ok', message: 'Bordereau FFTT genere et sauvegarde dans Drive (dossier 5-Bordereaux FFTT).' };
  } catch(e) {
    Logger.log('saveBordereauFFTT ERREUR : ' + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

// BORDEREAU LICENCE FFTT — Génération PDF pré-rempli


function obtenirDossierFRI(nomSousDossier) {
  var dossiers = dossiersDriveParNom(nomSousDossier);
  if (dossiers.hasNext()) return dossiers.next();
  return creerDossierDrive(nomSousDossier);
}

function majTotalFFTT(ss) {
  try {
    if (!ss) return;
    var sheet = ss.getSheetByName('Licences FFTT');
    if (!sheet || sheet.getLastRow() < 4) return;
    var total = 0;
    var data = sheet.getRange(4, 5, sheet.getLastRow()-3, 1).getValues();
    data.forEach(function(row){ total += Number(row[0]) || 0; });
    total = Math.round(total * 100) / 100;
    sheet.getRange(3,6).setValue(total).setFontWeight('bold')
      .setFontColor('#0d47a1').setBackground('#bbdefb')
      .setNumberFormat('#,##0.00 €');
    Logger.log('Total FFTT : ' + total + ' €');
  } catch(e) { Logger.log('majTotalFFTT KO : ' + e); }
}

function majOngletEspeces(ss, rows) {
  if(!ss||!rows||rows.length===0)return;var r0=rows[0];if((r0.mode_paiement||'')!=='especes')return;
  var code=r0.code_dossier||'',nom=r0.responsable_nom||r0.membre_nom||'',prenom=r0.responsable_prenom||r0.membre_prenom||'',email=r0.email1||'';
  var montant = calculerSoldeReel(r0, null);
  // Reformater la date si elle arrive en format JS brut (ex: "Mon Jun 01 2026 00:00:00 GMT+0200")
  var rawDate = r0.date || '';
  var dateJour;
  if (rawDate && rawDate.indexOf('GMT') >= 0) {
    // Date JS brute → reformater
    try { dateJour = Utilities.formatDate(new Date(rawDate), 'Europe/Paris', 'dd/MM/yyyy'); }
    catch(ed) { dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm'); }
  } else {
    dateJour = rawDate || Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  }
  try {
    var sh = getOrCreateEspecesSheet(ss);
    // Guard : ne pas ajouter si le code est deja present
    if (sh.getLastRow() > 1) {
      var existD = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
      for (var gi=0; gi<existD.length; gi++) { if (String(existD[gi][0]).trim() === code) { Logger.log('Especes: '+code+' deja present'); return; } }
    }
    var nextRow = Math.max(sh.getLastRow()+1, 4);
    var bg = nextRow%2===0?'#efebe9':'#ffffff';
    sh.getRange(nextRow,1,1,7).setValues([[code,nom,prenom,email,montant,dateJour,'⏳ En cours de validation']]).setBackground(bg);
    sh.getRange(nextRow,5).setFontColor('#4e342e').setFontWeight('bold');
    sh.getRange(nextRow,7).setFontColor('#856404').setFontWeight('bold');
    sh.autoResizeColumns(1,7);
    majTotalEspeces(ss);
  } catch(e) { Logger.log('majOngletEspeces KO : '+e.toString()); }
}

// ── Onglet HelloAsso ──

// Alimenter l'onglet HA depuis les données du webhook (quand dossier pas encore dans Sheet)
function majOngletHelloAssoWebhook(ss, code, payerInfo, montantHA, haFormType, haFormName) {
  if (!ss || !code) return;
  var sh = getOrCreateHelloAssoSheet(ss);
  var dateJour = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  var nom    = String((payerInfo && payerInfo.lastName)  || '').toUpperCase();
  var prenom = String((payerInfo && payerInfo.firstName) || '');
  var email  = String((payerInfo && payerInfo.email)     || '');

  // Chercher le montant total du dossier dans Inscriptions
  var montantDossier = 0;
  try {
    var shInsc = ss.getSheetByName('Inscriptions');
    if (shInsc && shInsc.getLastRow() > 1) {
      var dataInsc = shInsc.getRange(2, 1, shInsc.getLastRow()-1, 35).getValues();
      for (var di = 0; di < dataInsc.length; di++) {
        if (String(dataInsc[di][19]||'').trim() === code) {
          if (!nom)    nom    = String(dataInsc[di][1]||'').toUpperCase();
          if (!prenom) prenom = String(dataInsc[di][2]||'');
          if (!email)  email  = String(dataInsc[di][15]||'');
          var tf = Number(dataInsc[di][34]||0); // col AI total_famille
          if (tf > montantDossier) montantDossier = tf;
          break;
        }
      }
    }
  } catch(eInsc) { Logger.log('HA recherche montant dossier KO: ' + eInsc); }

  var montantRecu = (montantHA && montantHA > 0) ? montantHA : 0;
  var is3FoisM = String(code||'').indexOf('3x') >= 0; // sera mis à jour depuis le webhook
  var is3Fois = is3FoisM || (montantRecu > 0 && montantDossier > 0 && montantRecu < montantDossier);
  var modePaiem = is3Fois ? 'HelloAsso 3 fois' : 'HelloAsso 1 fois';

  // Chercher si ce code existe déjà → gérer les versements successifs
  var lr = sh.getLastRow();
  var existRow = -1;
  if (lr > 3) {
    var existing = sh.getRange(4, 1, lr-3, 13).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).trim() === code) { existRow = i + 4; break; }
    }
  }

  if (existRow > 0) {
    // Mettre à jour la ligne existante
    if (montantDossier > 0) sh.getRange(existRow, 5).setValue(montantDossier).setFontColor('#bf360c').setFontWeight('bold');
    if (nom)    sh.getRange(existRow, 2).setValue(nom);
    if (prenom) sh.getRange(existRow, 3).setValue(prenom);
    if (email)  sh.getRange(existRow, 4).setValue(email);
    sh.getRange(existRow, 7).setValue(modePaiem);
    sh.getRange(existRow, 9).setValue('⏳ En cours de validation').setFontColor('#856404').setFontWeight('bold');
    if (montantRecu > 0) {
      var v1 = Number(sh.getRange(existRow, 6).getValue() || 0);
      var v2 = Number(sh.getRange(existRow, 12).getValue() || 0);
      if (v1 === 0) {
        // 1er versement
        sh.getRange(existRow, 6).setValue(montantRecu).setFontColor('#1b5e20').setFontWeight('bold').setNumberFormat('#,##0.00 €');
        sh.getRange(existRow, 8).setValue(dateJour);
      } else if (v2 === 0) {
        // 2e versement → col L (12)
        sh.getRange(existRow, 12).setValue(montantRecu).setFontColor('#1b5e20').setFontWeight('bold').setNumberFormat('#,##0.00 €');
        sh.getRange(existRow, 10).setValue(dateJour);
      } else {
        // 3e versement → col M (13)
        sh.getRange(existRow, 13).setValue(montantRecu).setFontColor('#1b5e20').setFontWeight('bold').setNumberFormat('#,##0.00 €');
        sh.getRange(existRow, 11).setValue(dateJour);
      }
    }
    Logger.log('✅ Onglet HA mis à jour pour ' + code + ' versement:' + montantRecu + '€');
  } else {
    // Nouvelle ligne (13 colonnes)
    var nextRow = Math.max(sh.getLastRow() + 1, 4);
    var bg = nextRow % 2 === 0 ? '#fbe9e7' : '#ffffff';
    sh.getRange(nextRow, 1, 1, 13).setValues([[
      code, nom, prenom, email, montantDossier, montantRecu, modePaiem, dateJour,
      '⏳ En cours de validation', String(haFormType||''), String(haFormName||''), 0, 0
    ]]).setBackground(bg);
    sh.getRange(nextRow, 5).setFontColor('#bf360c').setFontWeight('bold').setNumberFormat('#,##0.00 €');
    sh.getRange(nextRow, 6).setFontColor('#1b5e20').setFontWeight('bold').setNumberFormat('#,##0.00 €');
    sh.getRange(nextRow, 9).setFontColor('#856404').setFontWeight('bold');
    Logger.log('✅ Onglet HA créé pour ' + code + ' 1er versement:' + montantRecu + '€ / dossier:' + montantDossier + '€');
  }
  sh.autoResizeColumns(1, 13);
}

function majOngletHelloAsso(ss, rows, montantHAForce, payerInfo) {
  if (!ss || !rows || rows.length === 0) return;
  var r0 = rows[0];
  var modeP0 = String(r0.mode_paiement || '').toLowerCase();
  if (modeP0.indexOf('hello') < 0 && modeP0 !== 'ha') return;
  var code  = r0.code_dossier || '';
  var nom   = r0.responsable_nom    || r0.membre_nom    || '';
  var pren  = r0.responsable_prenom || r0.membre_prenom || '';
  var email = r0.email1 || '';
  if (payerInfo) {
    if (payerInfo.lastName)  nom  = String(payerInfo.lastName).toUpperCase();
    if (payerInfo.firstName) pren = String(payerInfo.firstName);
    if (payerInfo.email)     email = String(payerInfo.email);
  }
  var dateJour = r0.date || Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy à HH:mm');
  try {
    var sh = getOrCreateHelloAssoSheet(ss);
    var existRow = -1;
    if (sh.getLastRow() > 3) {
      var exHA = sh.getRange(4, 1, sh.getLastRow()-3, 1).getValues();
      for (var gi = 0; gi < exHA.length; gi++) {
        if (String(exHA[gi][0]).trim() === code) { existRow = gi + 4; break; }
      }
    }
    var montantDossier = Number(r0.total_famille || r0.total || 0);
    var montantRecu    = (montantHAForce && montantHAForce > 0) ? montantHAForce : 0;
    // Détecter 3 fois depuis mode_paiement OU depuis le montant reçu partiel
    var is3FoisMode = modeP0.indexOf('3x') >= 0 || modeP0.indexOf('3fois') >= 0
                   || modeP0.indexOf('echelonn') >= 0 || modeP0 === 'helloasso3x';
    var is3Fois     = is3FoisMode || (montantRecu > 0 && montantDossier > 0 && montantRecu < montantDossier);
    var modePaiemHA = is3Fois ? 'HelloAsso 3 fois' : 'HelloAsso 1 fois';
    var statut      = montantRecu > 0 ? '⏳ Paiement reçu — validation admin requise' : '⏳ En attente de paiement';

    if (existRow > 0) {
      // Mise à jour — gérer les versements successifs
      if (montantDossier > 0) sh.getRange(existRow, 5).setValue(montantDossier).setFontColor('#bf360c').setFontWeight('bold').setNumberFormat('#,##0.00 €');
      if (montantRecu > 0) {
        var v1 = Number(sh.getRange(existRow, 6).getValue()  || 0);
        var v2 = Number(sh.getRange(existRow, 12).getValue() || 0);
        if (v1 === 0) {
          sh.getRange(existRow, 6).setValue(montantRecu).setFontColor('#1b5e20').setFontWeight('bold').setNumberFormat('#,##0.00 €');
          sh.getRange(existRow, 8).setValue(dateJour);
        } else if (v2 === 0) {
          sh.getRange(existRow, 12).setValue(montantRecu).setFontColor('#1b5e20').setFontWeight('bold').setNumberFormat('#,##0.00 €');
          sh.getRange(existRow, 10).setValue(dateJour);
        } else {
          sh.getRange(existRow, 13).setValue(montantRecu).setFontColor('#1b5e20').setFontWeight('bold').setNumberFormat('#,##0.00 €');
          sh.getRange(existRow, 11).setValue(dateJour);
        }
      }
      sh.getRange(existRow, 7).setValue(modePaiemHA);
      sh.getRange(existRow, 9).setValue(statut).setFontColor('#856404').setFontWeight('bold');
      Logger.log('HA onglet mis à jour : ' + code + ' versement:' + montantRecu + '€');
    } else {
      // Nouvelle ligne (13 colonnes : A-E montant dossier, F 1er versement, G mode, H date1, I statut, J date2, K date3, L 2e versement, M 3e versement)
      var nextRow = Math.max(sh.getLastRow() + 1, 4);
      var bg = nextRow % 2 === 0 ? '#fbe9e7' : '#ffffff';
      sh.getRange(nextRow, 1, 1, 13).setValues([[
        code, nom, pren, email, montantDossier, montantRecu, modePaiemHA, dateJour, statut,
        'Membership', 'Adhésion FRI 2026-2027', is3Fois ? '' : 'N/A', is3Fois ? '' : 'N/A'
      ]]).setBackground(bg);
      sh.getRange(nextRow, 5).setFontColor('#bf360c').setFontWeight('bold').setNumberFormat('#,##0.00 €');
      sh.getRange(nextRow, 6).setFontColor(montantRecu > 0 ? '#1b5e20' : '#999999').setFontWeight('bold').setNumberFormat('#,##0.00 €');
      sh.getRange(nextRow, 12).setNumberFormat('#,##0.00 €');
      sh.getRange(nextRow, 13).setNumberFormat('#,##0.00 €');
      sh.getRange(nextRow, 9).setFontColor('#856404').setFontWeight('bold');
      sh.autoResizeColumns(1, 13);
      Logger.log('HA onglet créé : ' + code + ' 1er versement:' + montantRecu + '€ / dossier:' + montantDossier + '€');
    }
    majTotalHelloAsso(ss);
  } catch(e) { Logger.log('majOngletHelloAsso KO : ' + e.toString()); }
}

function majTotalHelloAsso(ss) {
  try {
    if (!ss) return;
    var sheet = ss.getSheetByName(SHEET_HELLOASSO);
    if (!sheet || sheet.getLastRow() < 4) return;
    var total = 0;
    // Total = somme de col E (montant dossier) pour la ligne de titre, mais on calcule la somme des versements reçus F+L+M
    var data = sheet.getRange(4, 6, sheet.getLastRow() - 3, 8).getValues(); // cols F-M (index 6 à 13)
    data.forEach(function(row) {
      total += Number(row[0]) || 0;  // col F (index 0) = 1er versement
      total += Number(row[6]) || 0;  // col L (index 6 depuis F) = 2e versement
      total += Number(row[7]) || 0;  // col M (index 7 depuis F) = 3e versement
    });
    total = Math.round(total * 100) / 100;
    sheet.getRange(3, 5).setValue(total).setFontWeight('bold').setFontColor('#bf360c').setBackground('#fbe9e7').setNumberFormat('#,##0.00 €');
    Logger.log('✅ Total HelloAsso : ' + total + ' €');
  } catch(e) { Logger.log('majTotalHelloAsso KO : ' + e.toString()); }
}

// ── Totaux chèques ──
function majTotalCheque(ss, nomOnglet) {
  try {
    if (!ss) return;
    var sheet = ss.getSheetByName(nomOnglet);
    if (!sheet || sheet.getLastRow() < 4) return;
    var total = 0;
    var data = sheet.getRange(4, 6, sheet.getLastRow() - 3, 1).getValues();
    data.forEach(function(row) { total += Number(row[0]) || 0; });
    total = Math.round(total * 100) / 100;
    sheet.getRange(3, 6).setValue(total).setFontWeight('bold').setFontColor('#1b5e20').setBackground('#e8f5e9').setNumberFormat('#,##0.00 €');
    Logger.log('✅ Total ' + nomOnglet + ' : ' + total + ' €');
  } catch(e) { Logger.log('majTotalCheque KO (' + nomOnglet + ') : ' + e.toString()); }
}


function majOngletsChequesWrite(ss, rows, cheques) {
  if (!cheques || cheques.length === 0) return;
  var r0 = rows[0];
  var nom = String(r0.responsable_nom||r0.membre_nom||'');
  var prenom = String(r0.responsable_prenom||r0.membre_prenom||'');
  var code = String(r0.code_dossier||'');
  var noms = [SHEET_CHEQUE_1, SHEET_CHEQUE_2, SHEET_CHEQUE_3];
  function ecrire(sh, ch) {
    if (sh.getLastRow() >= 4) {
      var ex = sh.getRange(4,1,sh.getLastRow()-3,1).getValues();
      for (var gi=0;gi<ex.length;gi++) { if (String(ex[gi][0]).trim()===code) return; }
    }
    var nr = Math.max(sh.getLastRow()+1, 4);
    var bg = nr%2===0?'#e8f5e9':'#ffffff';
    sh.getRange(nr,1,1,7).setValues([[code,nom,prenom,String(ch.banque||''),String(ch.numCheque||''),Number(ch.montant||0),'⏳ En cours de validation']]).setBackground(bg);
    sh.getRange(nr,6).setFontColor('#1b5e20').setFontWeight('bold');
    sh.getRange(nr,7).setFontColor('#856404').setFontWeight('bold');
  }
  if (cheques.length===1) { ecrire(getOrCreateChequeSheet(ss,SHEET_CHEQUE_1),cheques[0]); majTotalCheque(ss,SHEET_CHEQUE_1); }
  else { cheques.forEach(function(ch,i){if(i>=3)return;ecrire(getOrCreateChequeSheet(ss,noms[i]),ch);majTotalCheque(ss,noms[i]);}); }
}

// ============================================================
// SAUVEGARDE QUOTIDIENNE
// ============================================================
var DOSSIER_BACKUP_ID='';
function sauvegardeQuotidienne(){
  try{var aujourdhui=Utilities.formatDate(new Date(),'Europe/Paris','yyyy-MM-dd');var heure=Utilities.formatDate(new Date(),'Europe/Paris','HH:mm');var heureFichier=Utilities.formatDate(new Date(),'Europe/Paris','HH-mm');var nomFichier='FRI_Inscriptions_2026-2027_'+aujourdhui+'_'+heureFichier+'.xlsx';// Copie directe du spreadsheet (sans UrlFetchApp)
    var ss2 = SpreadsheetApp.openById(SHEET_ID);
    var copie2 = ss2.copy(nomFichier.replace('.xlsx',''));
    var dossierNomB='FRI_Backup_Inscriptions';
    var dossiersB=dossiersDriveParNom(dossierNomB);
    var dossier=dossiersB.hasNext()?dossiersB.next():creerDossierDrive(dossierNomB);
    var fichier=DriveApp.getFileById(copie2.getId());
    dossier.addFile(fichier); DriveApp.getRootFolder().removeFile(fichier);
    securiserFichier(fichier);var ss=SpreadsheetApp.openById(SHEET_ID);var shInscr=ss.getSheetByName(SHEET_INSCRIPTIONS);var nbLignes=shInscr?Math.max(0,shInscr.getLastRow()-1):0;envoyerEmail(EMAIL_ADMIN,'FRI Isneauville - Sauvegarde automatique du '+aujourdhui,'Sauvegarde effectuée le '+aujourdhui+' à '+heure+'.\nFichier : '+nomFichier+'\nInscriptions : '+nbLignes+' ligne(s)\nLien Drive : '+fichier.getUrl());}catch(err){Logger.log('ERREUR sauvegarde : '+err.toString());try{envoyerEmail(EMAIL_ADMIN,'ALERTE FRI - Erreur sauvegarde',err.toString());}catch(e2){}}
}
function installerDeclencheurSauvegarde(){ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='sauvegardeQuotidienne')ScriptApp.deleteTrigger(t);});ScriptApp.newTrigger('sauvegardeQuotidienne').timeBased().everyDays(1).atHour(20).create();Logger.log('Déclencheur sauvegarde installé — quotidien à 20h');}

// ============================================================
// EXPORT CSV
// ============================================================
function rowToCsv(row){return row.map(function(cell){var val=(cell===null||cell===undefined)?'':String(cell);if(val.indexOf(';')>=0||val.indexOf('"')>=0||val.indexOf('\n')>=0)val='"'+val.replace(/"/g,'""')+'"';return val;}).join(';');}
function sheetToCsvString(sheet){var lastRow=sheet.getLastRow(),lastCol=sheet.getLastColumn();if(lastRow<1||lastCol<1)return'';var data=sheet.getRange(1,1,lastRow,lastCol).getValues();return'\uFEFF'+data.map(function(row){return rowToCsv(row);}).join('\r\n');}
function exporterCSVParEmail(){var ss=SpreadsheetApp.openById(SHEET_ID);var sheet=ss.getSheetByName(SHEET_INSCRIPTIONS);var dest='fri.inscri@gmail.com';var today=new Date().toLocaleDateString('fr-FR');if(!sheet||sheet.getLastRow()<2){envoyerEmail(dest,'[FRI] Export CSV - Aucune donnee','Aucune inscription au '+today+'.');return;}var blob=Utilities.newBlob(sheetToCsvString(sheet),'text/csv; charset=utf-8','inscriptions_FRI_'+today.replace(/\//g,'-')+'.csv');var nbLignes=sheet.getLastRow()-1;envoyerEmail(dest,'[FRI] Export CSV - '+today+' ('+nbLignes+' inscriptions)','Fichier CSV en pièce jointe.',{attachments:[blob],name:'Site FRI Inscriptions'});}

// ============================================================
// FONCTIONS DE TEST
// ============================================================
function testConnexion(){var ss=SpreadsheetApp.openById(SHEET_ID);Logger.log('Connecté : '+ss.getName());}
function testEmailConfirmationAvecFacture(){
  Logger.log('=== TEST EMAIL + FACTURE v8.3 ===');
  var rowTest=faireRowTest();
  var pdf=genererFacturePDF([rowTest],'Chèque');
  Logger.log(pdf?'✅ PDF OK : '+pdf.getBytes().length+' octets':'❌ PDF KO');
  try{envoyerEmailAdherent(EMAIL_ADMIN,[rowTest],true,'Chèque',pdf);Logger.log('✅ Email envoyé');}catch(e){Logger.log('❌ Email KO : '+e.toString());}
  Logger.log('=== FIN TEST ===');
}
function testEmailAvecRemise(){
  Logger.log('=== TEST EMAIL + REMISE 15% (3 activités) ===');
  var rows=[
    {date:'22/05/2026',code_dossier:'FRI-TEST',responsable_nom:'DUPONT',responsable_prenom:'Marie',adresse:'12 rue des Lilas',cp:'76230',ville:'ISNEAUVILLE',tel1:'06 12 34 56 78',email1:EMAIL_ADMIN,membre_nom:'DUPONT',membre_prenom:'Léa',ddn:'2012-03-15',sexe:'F',activite:'Danse Moderne Inter 2&3',activite_id:'JAZME1615',jour:'Mercredi',heure:'16h15/17h15',lieu:'MB',animateur:'F. CHARROIS',tarif:122.75,tarif_brut:144.42,remise:15,fnsmr:15,total_famille:560,mode_paiement:'cheque',qs_sante:'Attestation OK',pass_aide:'',pass_sport_montant:0,ancv_montant:0,avoir_montant:0,fftt_price:0},
    {date:'22/05/2026',code_dossier:'FRI-TEST',responsable_nom:'DUPONT',responsable_prenom:'Marie',adresse:'12 rue des Lilas',cp:'76230',ville:'ISNEAUVILLE',tel1:'06 12 34 56 78',email1:EMAIL_ADMIN,membre_nom:'DUPONT',membre_prenom:'Léa',ddn:'2012-03-15',sexe:'F',activite:'Pilates Jeudi 1',activite_id:'PILJ11',jour:'Jeudi',heure:'11h/12h',lieu:'PSF',animateur:'M. MARTIN',tarif:122.75,tarif_brut:144.42,remise:15,fnsmr:0,total_famille:560,mode_paiement:'cheque',qs_sante:'Attestation OK',pass_aide:'',pass_sport_montant:0,ancv_montant:0,avoir_montant:0,fftt_price:0},
    {date:'22/05/2026',code_dossier:'FRI-TEST',responsable_nom:'DUPONT',responsable_prenom:'Marie',adresse:'12 rue des Lilas',cp:'76230',ville:'ISNEAUVILLE',tel1:'06 12 34 56 78',email1:EMAIL_ADMIN,membre_nom:'DUPONT',membre_prenom:'Léa',ddn:'2012-03-15',sexe:'F',activite:'Yoga Lundi Soir',activite_id:'YOGAL18',jour:'Lundi',heure:'18h/19h',lieu:'SM',animateur:'S. CLAIRE',tarif:122.75,tarif_brut:144.42,remise:15,fnsmr:0,total_famille:560,mode_paiement:'cheque',qs_sante:'Attestation OK',pass_aide:'',pass_sport_montant:0,ancv_montant:0,avoir_montant:0,fftt_price:0}
  ];
  var pdf=genererFacturePDF(rows,'Chèque');
  Logger.log(pdf?'✅ PDF OK : '+pdf.getBytes().length+' octets':'❌ PDF KO');
  try{envoyerEmailAdherent(EMAIL_ADMIN,rows,true,'Chèque',pdf);Logger.log('✅ Email envoyé');}catch(e){Logger.log('❌ : '+e);}
  Logger.log('=== FIN TEST ===');
}
function faireRowTest(){
  return{date:new Date().toLocaleDateString('fr-FR'),id:String(Date.now()),responsable_nom:'DUPONT',responsable_prenom:'Marie',adresse:'12 rue des Lilas',cp:'76230',ville:'ISNEAUVILLE',commune:'Isneauville',tel1:'06 12 34 56 78',email1:EMAIL_ADMIN,membre_nom:'DUPONT',membre_prenom:'Léa',ddn:'2012-03-15',sexe:'F',activite:'Danse Moderne Inter 2&3',activite_id:'JAZME1615',jour:'Mercredi',heure:'16h15/17h15',lieu:'MB',animateur:'F. CHARROIS',tarif:122.75,tarif_brut:144.42,remise:15,fnsmr:15,total_famille:160,mode_paiement:'cheque',code_dossier:'FRI-TEST',qs_sante:'Attestation OK',pass_aide:'',pass_sport_montant:70,ancv_montant:0,avoir_montant:0,fftt_price:0};
}

// ============================================================
// CORRIGER EN-TÊTES
// ============================================================
function corrigerEnTetes() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS) || ss.insertSheet(SHEET_INSCRIPTIONS);
  var NB_COLS = 41; // structure v8.8 : 41 colonnes

  // ── Étendre le Sheet si moins de 41 colonnes ──
  var currentCols = sheet.getMaxColumns();
  if (currentCols < NB_COLS) {
    sheet.insertColumnsAfter(currentCols, NB_COLS - currentCols);
    Logger.log('Colonnes ajoutées : ' + currentCols + ' → ' + NB_COLS);
  }

  var h = ['Numéro de licence','Civilité','Nom','Prénom','Date de naissance',
    'Appartement - Etage','Batiment - Résidence','N° et nom de voie',
    'Lieu-dit ou boîte postale','Code postal','Ville','Cedex','Code pays',
    'Tél. fixe','Tél. portable','Adresse e-mail','Commentaire','Tél. 2','Email 2',
    'N° Dossier','Date inscription','Statut paiement','Activité','Jour','Heure',
    'Lieu','Animateur',
    'Tarif brut EUR','Éligible remise','Tarif net EUR','FNSMR EUR',
    'Total famille EUR','Mode paiement','Avoir','QS Santé','Pass / Aide',
    'Sexe','ID Activité','Responsable','Statut inscription','Licence FFTT (€)'];

  // Insérer la ligne d'en-tête si besoin
  var firstCell = '';
  try { firstCell = String(sheet.getRange(1,1).getValue() || ''); } catch(e) {}
  if (firstCell && firstCell !== 'Numéro de licence' && firstCell !== '') {
    sheet.insertRowBefore(1);
  }

  // Écrire les en-têtes
  sheet.getRange(1, 1, 1, h.length).setValues([h])
    .setBackground('#2d6a4f').setFontColor('#ffffff').setFontWeight('bold');

  // Couleurs spécifiques par colonne
  sheet.getRange(1, 1, 1, 19).setBackground('#1b5e20');
  sheet.getRange(1, 20).setBackground('#1565c0').setFontColor('#ffffff');
  sheet.getRange(1, 29).setBackground('#e8f5e9').setFontColor('#1b5e20');  // Éligible
  sheet.getRange(1, 30).setBackground('#d8f3dc').setFontColor('#1b5e20');  // Tarif net
  sheet.getRange(1, 34).setBackground('#e65100').setFontColor('#ffffff');  // Avoir
  sheet.getRange(1, 40).setBackground('#4a148c').setFontColor('#ffffff');  // Statut inscription
  sheet.getRange(1, 41).setBackground('#bf360c').setFontColor('#ffffff');  // FFTT

  // Supprimer les colonnes excédentaires (> 41)
  var lastCol = sheet.getLastColumn();
  if (lastCol > NB_COLS) {
    sheet.deleteColumns(NB_COLS + 1, lastCol - NB_COLS);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, h.length);
  Logger.log('✅ En-têtes v8.8 corrigés — ' + NB_COLS + ' colonnes');
}

// ============================================================
// REMISE À ZÉRO — inchangé v8.2
// ============================================================
function remiseAZeroComplete(){
  Logger.log('=== REMISE À ZÉRO COMPLÈTE DÉMARRÉE ===');

  // ── 1. Sauvegarde préventive ─────────────────────────────────
  try { sauvegardeAvantReset(); } catch(e) { Logger.log('⚠️ Sauvegarde KO : ' + e.toString()); }

  var ss = SpreadsheetApp.openById(SHEET_ID);

  // ── 2. Vider et recréer tous les onglets fixes ───────────────
  // Ordre : Inscriptions → Récapitulatif → Places →
  //         Avoirs générés (lignes épuisées supprimées) → Avoirs utilisés (réinitialisé) →
  //         Aides ANCV/Pass'sport/Pass'jeunes/Atout →
  //         HelloAsso → Chèques 1/2/3 → Espèces → Licences FFTT
  try { viderSheet(ss); } catch(e) { Logger.log('viderSheet KO : ' + e); }

  // ── 3. Corriger les en-têtes (41 colonnes) ───────────────────
  try { corrigerEnTetes(); } catch(e) { Logger.log('corrigerEnTetes KO : ' + e); }

  // ── 4. Supprimer les onglets dynamiques {placesId} et LA-{id}─
  try { supprimerOngletsActivites(ss); } catch(e) { Logger.log('supprimerOngletsActivites KO : ' + e); }

  // ── 5. Réordonner tous les onglets dans l'ordre cible ────────
  // 1.Inscriptions  2.Récapitulatif  3.Places
  // 4.{ID activité} 5.LA-{ID activité}
  // 6.Avoirs  7.Aides ANCV  8.Aides Pass'sport  9.Aides Pass'jeunes  10.Aides Atout
  // 11.HelloAsso  12.Chèques 1  13.Chèques 2  14.Chèques 3  15.Espèces  16.Licences FFTT
  try { reordonnerOnglets(ss); } catch(e) { Logger.log('reordonnerOnglets KO : ' + e); }

  // ── 6. Vider les dossiers Drive ──────────────────────────────
  try { viderDossierDrive('4-Factures acquittées'); }   catch(e) { Logger.log('Drive Factures KO : ' + e); }
  try { viderDossierDrive('2-QS Santé Adhérents'); }        catch(e) { Logger.log('Drive QS KO : ' + e); }
  try { viderDossierDrive('3-Règlements intérieurs'); }        catch(e) { Logger.log('Drive RI KO : ' + e); }
  try { viderDossierDrive('1-Certificats médicaux'); }       catch(e) { Logger.log('Drive Certificats KO : ' + e); }
  try { viderDossierDrive('5-Bordereaux FFTT'); }             catch(e) { Logger.log('Drive Bordereaux KO : ' + e); }
  // ── 7. Effacer la col AP (compteur export Gestafill) ────────
  try {
    var shInscr2 = ss.getSheetByName(SHEET_INSCRIPTIONS);
    if (shInscr2 && shInscr2.getLastRow() > 1) {
      shInscr2.getRange(2, 42, shInscr2.getLastRow()-1, 1).clearContent();
      Logger.log('✅ Col AP (Gestafill) effacée');
    }
  } catch(eAP2) { Logger.log('Col AP reset KO : ' + eAP2); }

  Logger.log('=== REMISE À ZÉRO TERMINÉE ===');
}

function sauvegardeAvantReset(){
  try {
    var date = Utilities.formatDate(new Date(),'Europe/Paris','yyyy-MM-dd_HH-mm');
    var nomCopie = 'BACKUP_AVANT_RESET_FRI_' + date;
    // Copier le spreadsheet directement (pas d'export URL - fonctionne sur le meme compte)
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var copie = ss.copy(nomCopie);
    var dossiers = dossiersDriveParNom('FRI_Backup_Inscriptions');
    var dossier = dossiers.hasNext() ? dossiers.next() : creerDossierDrive('FRI_Backup_Inscriptions');
    // Deplacer la copie dans le dossier backup
    var fichierCopie = DriveApp.getFileById(copie.getId());
    dossier.addFile(fichierCopie);
    DriveApp.getRootFolder().removeFile(fichierCopie);
    Logger.log('✅ Sauvegarde créée : ' + copie.getUrl());
  } catch(e) { Logger.log('⚠️ Sauvegarde KO : ' + e.toString()); }
}
function viderSheet(ss){
  var protectedSheets=[
    SHEET_INSCRIPTIONS, SHEET_RECAPITULATIF, SHEET_PLACES,
    SHEET_AVOIRS, SHEET_AVOIRS_UTILISES,
    SHEET_AIDE_ANCV, SHEET_AIDE_PASS_S, SHEET_AIDE_PASS_J, SHEET_AIDE_ATOUT,
    SHEET_HELLOASSO, SHEET_CHEQUE_1, SHEET_CHEQUE_2, SHEET_CHEQUE_3,
    SHEET_ESPECES, 'Licences FFTT'
  ];

  protectedSheets.forEach(function(nom){
    var sheet = ss.getSheetByName(nom);

    // HelloAsso : supprimer/recréer
    if(nom===SHEET_HELLOASSO){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      getOrCreateHelloAssoSheet(ss);
      Logger.log('✅ Onglet ' + nom + ' recréé'); return;
    }
    // Avoirs générés : conserver inter-saison (supprimer seulement les épuisés)
    if(nom===SHEET_AVOIRS){
      var avoirSheet = getOrCreateAvoirSheet(ss);
      if (avoirSheet.getLastRow() > 1) {
        var avoirData = avoirSheet.getRange(2, 1, avoirSheet.getLastRow()-1, 9).getValues();
        for (var av = avoirData.length-1; av >= 0; av--) {
          var stAv  = String(avoirData[av][8]||'').toLowerCase();
          var solAv = parseFloat(avoirData[av][7])||0;
          if (stAv === 'épuisé' || solAv <= 0) avoirSheet.deleteRow(av + 2);
        }
        Logger.log('✅ Avoirs générés : lignes épuisées supprimées, soldes conservés');
      }
      return;
    }
    // Avoirs utilisés : supprimer/recréer
    if(nom===SHEET_AVOIRS_UTILISES){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      getOrCreateAvoirsUtilisesSheet(ss);
      Logger.log('✅ Onglet ' + nom + ' recréé'); return;
    }
    // Espèces : supprimer/recréer (nouvelle structure avec colonne Statut)
    if(nom===SHEET_ESPECES){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      getOrCreateEspecesSheet(ss);
      Logger.log('✅ Onglet ' + SHEET_ESPECES + ' recréé'); return;
    }
    // Licences FFTT : supprimer/recréer
    if(nom==='Licences FFTT'){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      try { getOrCreateFFTTSheet(ss); } catch(eFF) { Logger.log('FFTT recréation KO: '+eFF); }
      Logger.log('✅ Onglet ' + 'Licences FFTT' + ' recréé'); return;
    }
    // Aides : supprimer/recréer avec nouvelles couleurs pastels + colonne Statut
    if(nom===SHEET_AIDE_ANCV){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      getOrCreateAideSheet(ss,SHEET_AIDE_ANCV,'Aides Coupon Sport ANCV','#1565c0','#ffffff','TOTAL ANCV déduit');
      Logger.log('✅ Onglet ' + SHEET_AIDE_ANCV + ' recréé'); return;
    }
    if(nom===SHEET_AIDE_PASS_S){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      getOrCreateAideSheet(ss,SHEET_AIDE_PASS_S,"Aides Pass'sport État",'#1b5e20','#ffffff',"TOTAL Pass'sport déduit");
      Logger.log('✅ Onglet ' + SHEET_AIDE_PASS_S + ' recréé'); return;
    }
    if(nom===SHEET_AIDE_PASS_J){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      getOrCreateAideSheet(ss,SHEET_AIDE_PASS_J,"Aides Pass'jeunes 76 / Handipass'sport",'#6a1b9a','#ffffff',"TOTAL Pass'jeunes déduit");
      Logger.log('✅ Onglet ' + SHEET_AIDE_PASS_J + ' recréé'); return;
    }
    if(nom===SHEET_AIDE_ATOUT){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      getOrCreateAideSheet(ss,SHEET_AIDE_ATOUT,'Aides Atout Normandie','#e65100','#ffffff','TOTAL Atout Normandie déduit');
      Logger.log('✅ Onglet ' + SHEET_AIDE_ATOUT + ' recréé'); return;
    }
    // Chèques 1/2/3 : supprimer/recréer (colonne Statut + largeur col 7)
    if(nom===SHEET_CHEQUE_1||nom===SHEET_CHEQUE_2||nom===SHEET_CHEQUE_3){
      if(sheet){try{ss.deleteSheet(sheet);}catch(e){}}
      getOrCreateChequeSheet(ss, nom);
      Logger.log('✅ Onglet ' + nom + ' recréé'); return;
    }

    // Places : créer si absent, puis reconstruire depuis les capacités fixes
    if(nom===SHEET_PLACES) {
      if(!sheet){
        sheet=ss.insertSheet(SHEET_PLACES);
        var h=['ID Activite','Nom activite','Capacite max','Inscrits','Places restantes'];
        sheet.getRange(1,1,1,h.length).setValues([h])
          .setBackground('#c9a84c').setFontColor('#1a2e22').setFontWeight('bold');
        sheet.setFrozenRows(1);
      } else {
        var plLr = sheet.getLastRow();
        if(plLr > 1) sheet.getRange(2, 1, plLr - 1, 5).clearContent();
      }
      initAllActivities(sheet);
      Logger.log('✅ Places réinitialisées (capacités fixes saison 2026/2027)'); return;
    }

    // Inscriptions + Récapitulatif : traitement spécifique
    if(!sheet) return;

    var lastRow   = sheet.getLastRow();
    var frozen    = sheet.getFrozenRows();
    var firstData = Math.max(frozen + 1, 2);
    if(lastRow >= firstData + 1){
      sheet.deleteRows(firstData, lastRow - firstData);
      sheet.getRange(firstData, 1, 1, sheet.getLastColumn()).clearContent();
    } else if(lastRow === firstData){
      sheet.getRange(firstData, 1, 1, sheet.getLastColumn()).clearContent();
    }
    if(nom===SHEET_INSCRIPTIONS) { ecrireEnTeteInscriptions(sheet); Logger.log('✅ En-têtes Inscriptions recréés'); }
  });
}

function reordonnerOnglets(ss) {
  // ── Ordre cible ──────────────────────────────────────────────
  // 1. Inscriptions
  // 2. Récapitulatif
  // 3. Places
  // 4. {ID activité}   ← onglets dynamiques insérés ici
  // 5. LA-{ID activité}← onglets dynamiques insérés ici
  // 6. Avoirs générés
  // 6b. Avoirs utilisés
  // 7. Aides ANCV
  // 8. Aides Pass'sport
  // 9. Aides Pass'jeunes
  // 10. Aides Atout Normandie
  // 11. HelloAsso
  // 12. Chèques 1
  // 13. Chèques 2
  // 14. Chèques 3
  // 15. Espèces
  // 16. Licences FFTT

  var allSheets = ss.getSheets();

  // Identifier les onglets dynamiques {placesId} et LA-{placesId}
  var ongletsFixes = [
    SHEET_INSCRIPTIONS, SHEET_RECAPITULATIF, SHEET_PLACES,
    SHEET_AVOIRS, SHEET_AVOIRS_UTILISES,
    SHEET_AIDE_ANCV, SHEET_AIDE_PASS_S, SHEET_AIDE_PASS_J, SHEET_AIDE_ATOUT,
    SHEET_HELLOASSO, SHEET_CHEQUE_1, SHEET_CHEQUE_2, SHEET_CHEQUE_3,
    SHEET_ESPECES, 'Licences FFTT'
  ];
  var ongletsFixesSet = {};
  ongletsFixes.forEach(function(n){ ongletsFixesSet[n] = true; });

  var ongletsActivites = []; // {placesId}
  var ongletsLA        = []; // LA-{placesId}
  allSheets.forEach(function(sh) {
    var n = sh.getName();
    if (ongletsFixesSet[n]) return;
    if (n.indexOf('LA-') === 0) ongletsLA.push(n);
    else ongletsActivites.push(n);
  });
  ongletsActivites.sort();
  ongletsLA.sort();

  // Ordre complet
  var ordreFinal = [
    SHEET_INSCRIPTIONS,
    SHEET_RECAPITULATIF,
    SHEET_PLACES
  ].concat(ongletsActivites)
   .concat(ongletsLA)
   .concat([
    SHEET_AVOIRS,
    SHEET_AVOIRS_UTILISES,
    SHEET_AIDE_ANCV,
    SHEET_AIDE_PASS_S,
    SHEET_AIDE_PASS_J,
    SHEET_AIDE_ATOUT,
    SHEET_HELLOASSO,
    SHEET_CHEQUE_1,
    SHEET_CHEQUE_2,
    SHEET_CHEQUE_3,
    SHEET_ESPECES,
    'Licences FFTT'
  ]);

  var position = 1;
  ordreFinal.forEach(function(nom) {
    var sheet = ss.getSheetByName(nom);
    if (sheet) {
      try { ss.setActiveSheet(sheet); ss.moveActiveSheet(position); position++; }
      catch(e) { Logger.log('Réordonnancement ' + nom + ' KO: ' + e); }
    }
  });
  Logger.log('Onglets réordonnancés — ' + position + ' onglets positionnés');
}
function supprimerOngletsActivites(ss){
  var protectedSheets=[
    SHEET_INSCRIPTIONS, SHEET_RECAPITULATIF, SHEET_PLACES,
    SHEET_AVOIRS, SHEET_AVOIRS_UTILISES,
    SHEET_AIDE_ANCV, SHEET_AIDE_ATOUT, SHEET_AIDE_PASS_J, SHEET_AIDE_PASS_S,
    SHEET_ESPECES, SHEET_HELLOASSO, 
    SHEET_CHEQUE_1, SHEET_CHEQUE_2, SHEET_CHEQUE_3
  ];
  var aSupprimer = ss.getSheets().filter(function(sheet){
    return protectedSheets.indexOf(sheet.getName()) < 0;
  });
  var nbSupp = 0;
  aSupprimer.forEach(function(sheet){
    try { ss.deleteSheet(sheet); nbSupp++; } catch(e){}
  });
  Logger.log('Onglets activites supprimes : ' + nbSupp);
}
function viderDossierDrive(nomDossier){try{var dossiers=dossiersDriveParNom(nomDossier);if(!dossiers.hasNext())return;var dossier=dossiers.next();var fichiers=dossier.getFiles();var nbSupp=0;while(fichiers.hasNext()){fichiers.next().setTrashed(true);nbSupp++;}Logger.log('✅ Dossier "'+nomDossier+'" vidé — '+nbSupp+' fichier(s)');}catch(e){Logger.log('⚠️ Erreur vidage "'+nomDossier+'" : '+e.toString());}}
function simulerRemiseAZero(){Logger.log('=== SIMULATION REMISE À ZÉRO ===');var ss=SpreadsheetApp.openById(SHEET_ID);var protectedSheets=[SHEET_INSCRIPTIONS,SHEET_RECAPITULATIF,SHEET_PLACES,SHEET_CHEQUE_1,SHEET_CHEQUE_2,SHEET_CHEQUE_3,SHEET_AVOIRS,SHEET_AVOIRS_UTILISES,SHEET_AIDE_ANCV,SHEET_AIDE_ATOUT,SHEET_AIDE_PASS_J,SHEET_AIDE_PASS_S,SHEET_ESPECES,SHEET_HELLOASSO];protectedSheets.forEach(function(nom){var sheet=ss.getSheetByName(nom);if(sheet)Logger.log('  Onglet "'+nom+'" : '+Math.max(0,sheet.getLastRow()-1)+' ligne(s)');});var nbAct=0;ss.getSheets().forEach(function(sheet){if(protectedSheets.indexOf(sheet.getName())<0)nbAct++;});Logger.log('  Onglets activités : '+nbAct+' à supprimer');Logger.log('=== FIN SIMULATION ===');}

function initialiserOngletEspeces() {
  Logger.log('=== INITIALISATION ONGLET ESPÈCES ===');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var existing = ss.getSheetByName(SHEET_ESPECES);
  if (existing) {
    Logger.log('ℹ️ Onglet "'+SHEET_ESPECES+'" existe déjà — repositionnement uniquement.');
    try{var avoirs=ss.getSheetByName(SHEET_AVOIRS);if(avoirs){var pos=avoirs.getIndex();var posAct=existing.getIndex();if(posAct!==pos){ss.setActiveSheet(existing);ss.moveActiveSheet(pos);Logger.log('✅ Repositionné');}}}catch(e){}
  } else {
    getOrCreateEspecesSheet(ss);Logger.log('✅ Onglet Espèces créé');
    getOrCreateHelloAssoSheet(ss);Logger.log('✅ Onglet HelloAsso créé');
  }
  Logger.log('=== FIN INITIALISATION ===');
}

function testHelloAssoCheckout() {
  Logger.log('=== TEST HELLOASSO CHECKOUT ===');
  var token = helloassoGetToken();
  if (!token) { Logger.log('❌ Token KO'); return; }
  Logger.log('✅ Token OK : '+token.substring(0,20)+'...');
  var lien = helloassoCreerLienPaiement(16000,'Marie','DUPONT',EMAIL_ADMIN,'FRI-TEST','Test FRI 2026/2027');
  if (!lien) { Logger.log('❌ Lien KO'); return; }
  Logger.log('✅ Lien : '+lien);
  try{envoyerEmail(EMAIL_ADMIN,'[FRI] Test HelloAsso','Lien : '+lien);}catch(e){}
  Logger.log('=== FIN TEST ===');
}

// ============================================================
// GESTION FIN DE SAISON — inchangé v8.2
// ============================================================

// ══════════════════════════════════════════════════════════════
// RAPPEL AUTOMATIQUE — Adhérents non réglés après 7 jours
// ══════════════════════════════════════════════════════════════

// À appeler via un déclencheur quotidien (chaque jour à 9h)

// ══════════════════════════════════════════════════════════════
// ADMIN — Rappel manuel : facture PDF + email rappel
// ══════════════════════════════════════════════════════════════
function envoyerRappelManuelGAS(code) {
  if (!code) return {status:'error', message:'Code dossier manquant'};

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_INSCRIPTIONS);
  if (!sheet) return {status:'error', message:'Onglet Inscriptions introuvable'};

  var data = sheet.getRange(2, 1, Math.max(sheet.getLastRow()-1,1), 41).getValues();

  // Récupérer toutes les lignes du dossier
  var emailRows = [];
  data.forEach(function(row) {
    var rowCode  = String(row[19]||'').trim();
    var rowStatut = lireStatutInscription(row).toLowerCase();
    if (rowCode !== code) return;
    if (rowStatut.indexOf('supprim') >= 0) return;

    var respFull = String(row[38]||'').trim();
    var respParts = respFull.split(' ');
    emailRows.push({
      code_dossier:       rowCode,
      responsable_nom:    respParts.length > 1 ? respParts.slice(1).join(' ') : String(row[2]||''),
      responsable_prenom: respParts[0] || String(row[3]||''),
      email1:             String(row[15]||''),
      tel1:               String(row[14]||''),
      adresse:            String(row[7]||'')+', '+String(row[9]||'')+' '+String(row[10]||''),
      commune:            String(row[10]||'').toUpperCase().indexOf('ISNEAUVILLE')>=0?'isno':'',
      ville:              String(row[10]||''),
      membre_nom:         String(row[2]||''),
      membre_prenom:      String(row[3]||''),
      activite:           String(row[22]||''),
      activite_id:        lireActiviteId(row),
      jour:               String(row[23]||''),
      heure:              String(row[24]||''),
      lieu:               String(row[25]||''),
      tarif_brut:         Number(row[27]||0),
      tarif:              Number(row[29]||0),
      statut_inscription: lireStatutInscription(row),
      mode_paiement:      String(row[32]||''),
      total_famille:      String(row[31]||''),
      avoir_montant:      String(row[33]||''),
      qs_sante:           String(row[34]||''),
      pass_aide:          String(row[35]||''),
      pass_sport_montant: '0',
      ancv_montant:       '0',
      fnsmr:              15,
      remise:             Number(row[28]||0),
      note_tarif:         '',
      date:               formaterDateHeure(row[20])
    });
  });

  if (emailRows.length === 0) {
    return {status:'error', message:'Dossier ' + code + ' introuvable ou déjà réglé'};
  }

  var r0 = emailRows[0];
  var email = r0.email1;
  if (!email || email.indexOf('@') < 0) {
    return {status:'error', message:'Email invalide pour le dossier ' + code};
  }

  // Vérifier si déjà réglé
  var statutPaie = String(data.filter(function(row){return String(row[19]||'').trim()===code;})[0]&&data.filter(function(row){return String(row[19]||'').trim()===code;})[0][21]||'').toLowerCase();
  if (statutPaie.indexOf('pay') >= 0 || statutPaie.indexOf('valid') >= 0) {
    return {status:'error', message:'Le dossier ' + code + ' est déjà réglé (' + statutPaie + ')'};
  }

  // Générer la facture PDF
  var pdfBlob = null;
  try {
    var modeLabel = r0.mode_paiement || 'helloasso';
    pdfBlob = genererFacturePDF(emailRows, modeLabel, 'INSCRIPTION À RÉGLER');
  } catch(ePdf) {
    Logger.log('Facture PDF rappel KO: ' + ePdf);
  }

  // Réinitialiser le verrou de rappel pour permettre un nouvel envoi
  PropertiesService.getScriptProperties().deleteProperty('rappel_' + code);

  // Envoyer le mail rappel via envoyerMailRappel (même fonction que le rappel auto)
  var total = parseFloat(r0.total_famille) || 0;
  var dateInscription = r0.date || '';

  // Mail adhérent avec facture PDF
  try {
    var logoB = getLogoBlob();
    var sujet = '[FRI] Rappel - Votre inscription ' + code + ' est en attente de reglement';

    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
      + '<body style="font-family:Arial,sans-serif;background:#f0f4f2;margin:0;padding:20px">'
      + '<div style="max-width:600px;margin:0 auto;background:white;border-radius:12px;overflow:hidden">'
      + '<div style="background:#1a2e22;padding:20px 28px;text-align:center">'
      + '<img src="cid:logo_fri" alt="FRI" style="width:60px;height:auto;border-radius:8px;display:block;margin:0 auto 10px">'
      + '<h1 style="color:white;margin:0;font-size:18px">⏰ Rappel — Règlement de votre inscription</h1>'
      + '<p style="color:#a8d5c2;margin:4px 0 0;font-size:13px">Foyer Rural d\'Isneauville — Saison 2026/2027</p>'
      + '</div>'
      + '<div style="padding:24px 28px">'
      + '<p style="color:#333;font-size:15px">Bonjour <strong>' + r0.responsable_prenom + ' ' + r0.responsable_nom + '</strong>,</p>'
      + '<p style="color:#555;line-height:1.7">Votre inscription (dossier '
      + '<span style="font-family:monospace;color:#1565c0;font-weight:bold">' + code + '</span>)'
      + (dateInscription ? ' du <strong>' + dateInscription + '</strong>' : '')
      + ' est en attente de règlement.</p>'
      + (total > 0 ? '<div style="background:#fff8e1;border:1px solid #e8c84a;border-radius:10px;padding:16px;margin:16px 0;text-align:center">'
        + '<div style="font-size:13px;color:#856404;font-weight:700;margin-bottom:4px">💰 Montant à régler</div>'
        + '<div style="font-size:24px;font-weight:900;color:#1a2e22;">' + total.toFixed(2) + ' €</div>'
        + '<div style="font-size:12px;color:#888;margin-top:4px">(activités + adhésion FNSMR)</div>'
        + '</div>' : '')
      + '<p style="color:#555;line-height:1.7"><strong>Comment régler ?</strong></p>'
      + '<ul style="color:#555;line-height:2;padding-left:20px">'
      + '<li>💳 <strong>En ligne</strong> via HelloAsso (lien ci-dessous)</li>'
      + '<li>📅 <strong>Aux permanences</strong> : Mardi 16h30 – 18h30</li>'
      + '<li>✉️ <strong>Par courrier</strong> : Chèque à l\'ordre du Foyer Rural d\'Isneauville</li>'
      + '</ul>'
      + helloassoBoutonHtml(null, total > 0 ? total.toFixed(2) + ' €' : null)
      + (pdfBlob ? '<p style="color:#555;font-size:13px;margin-top:16px">📎 <em>Votre facture détaillée est jointe à cet email.</em></p>' : '')
      + '<p style="color:#888;font-size:12px;margin-top:16px">⚠️ <em>Sans règlement, votre inscription ne pourra pas être confirmée.</em></p>'
      + '<hr style="border:none;border-top:1px solid #eee;margin:20px 0">'
      + '<p style="color:#888;font-size:12px">Foyer Rural d\'Isneauville — '
      + '<a href="mailto:frisneauville@orange.fr" style="color:#2d6a4f">frisneauville@orange.fr</a></p>'
      + '</div></div></body></html>';

    var opts = {htmlBody: html, charset: 'UTF-8', name: NOM_ASSO, replyTo: EMAIL_ADMIN};
    if (logoB) opts.inlineImages = {logo_fri: logoB};
    if (pdfBlob) opts.attachments = [pdfBlob];

    var bodyText = 'Bonjour ' + r0.responsable_prenom + ','
      + '\n\nVotre inscription ' + code + ' est en attente de reglement.'
      + '\nMontant : ' + (total > 0 ? total.toFixed(2) + ' EUR' : '?')
      + '\n\nCordialement,\nFoyer Rural d\'Isneauville';
    envoyerEmail(email, sujet, bodyText, opts
    );
    Logger.log('✅ Rappel manuel envoyé à ' + email + ' pour ' + code);

    // Mail admin
    envoyerEmail(EMAIL_ADMIN,
      '[FRI] Rappel manuel envoye - ' + code + ' — ' + r0.responsable_prenom + ' ' + r0.responsable_nom,
      'Un rappel de règlement a été envoyé manuellement à ' + email + ' pour le dossier ' + code + '.',
      {name: NOM_ASSO}
    );

  } catch(eEmail) {
    Logger.log('Mail rappel manuel KO: ' + eEmail);
    return {status:'error', message:'Erreur envoi email : ' + eEmail.toString()};
  }

  return {status:'ok', email: email, code: code};
}

function verifierFinDeSaison(){var props=PropertiesService.getScriptProperties();var now=new Date();var today=Utilities.formatDate(now,'Europe/Paris','yyyy-MM-dd');var dateCible=props.getProperty('FDS_DATE_SUPPRESSION')||(now.getFullYear()+'-06-01');var statut=props.getProperty('FDS_STATUT')||'attente';if(statut==='annulee'){return;}var dateSuppression=new Date(dateCible+'T02:00:00');var dateAvertissement=new Date(dateSuppression);dateAvertissement.setDate(dateAvertissement.getDate()-7);var todayAvert=Utilities.formatDate(dateAvertissement,'Europe/Paris','yyyy-MM-dd');var avertOk=props.getProperty('FDS_AVERTISSEMENT_OK')==='true';if(today===todayAvert&&!avertOk){envoyerAvertissementSuppression(dateCible);props.setProperty('FDS_AVERTISSEMENT_OK','true');return;}if(today===dateCible&&(statut==='confirmee'||statut==='attente')){remiseAZeroComplete();props.deleteProperty('FDS_STATUT');props.deleteProperty('FDS_AVERTISSEMENT_OK');props.setProperty('FDS_DATE_SUPPRESSION',(now.getFullYear()+1)+'-06-01');}}
function envoyerAvertissementSuppression(dateCible){var scriptId=ScriptApp.getScriptId();var baseUrl=ScriptApp.getService().getUrl();var tokenConfirm=Utilities.base64Encode('confirmer:'+dateCible+':'+scriptId.substring(0,8));var tokenAnnuler=Utilities.base64Encode('annuler:'+dateCible+':'+scriptId.substring(0,8));var tokenReporter=Utilities.base64Encode('reporter30:'+dateCible+':'+scriptId.substring(0,8));PropertiesService.getScriptProperties().setProperty('FDS_TOKEN_CONFIRM',tokenConfirm);PropertiesService.getScriptProperties().setProperty('FDS_TOKEN_ANNULER',tokenAnnuler);PropertiesService.getScriptProperties().setProperty('FDS_TOKEN_REPORTER',tokenReporter);var urlConfirm=baseUrl+'?action=fds&token='+encodeURIComponent(tokenConfirm);var urlAnnuler=baseUrl+'?action=fds&token='+encodeURIComponent(tokenAnnuler);var urlReporter=baseUrl+'?action=fds&token='+encodeURIComponent(tokenReporter);envoyerEmail(EMAIL_ADMIN,'[FRI] Suppression donnees prevue le '+dateCible,'Confirmer : '+urlConfirm+'\nReporter 30j : '+urlReporter+'\nAnnuler : '+urlAnnuler,{name:NOM_ASSO});}
function traiterActionFinDeSaison(token){var props=PropertiesService.getScriptProperties();var tokenConfirm=props.getProperty('FDS_TOKEN_CONFIRM')||'',tokenAnnuler=props.getProperty('FDS_TOKEN_ANNULER')||'',tokenReporter=props.getProperty('FDS_TOKEN_REPORTER')||'';var message='';if(token===tokenConfirm){props.setProperty('FDS_STATUT','confirmee');message='✅ Suppression confirmée.';}else if(token===tokenAnnuler){props.setProperty('FDS_STATUT','annulee');message='🚫 Suppression annulée.';}else if(token===tokenReporter){var dateCible=props.getProperty('FDS_DATE_SUPPRESSION')||'';var nouvelleDate=new Date(dateCible+'T00:00:00');nouvelleDate.setDate(nouvelleDate.getDate()+30);var nouvelleDateStr=Utilities.formatDate(nouvelleDate,'Europe/Paris','yyyy-MM-dd');props.setProperty('FDS_DATE_SUPPRESSION',nouvelleDateStr);props.setProperty('FDS_STATUT','attente');props.deleteProperty('FDS_AVERTISSEMENT_OK');message='📅 Suppression reportée au '+nouvelleDateStr+'.';}else{message='❌ Lien invalide ou expiré.';}return HtmlService.createHtmlOutput('<html><head><meta charset="UTF-8"></head><body style="font-family:Arial;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f0f4f2"><div style="background:white;border-radius:14px;padding:40px;max-width:480px;text-align:center;box-shadow:0 10px 40px rgba(0,0,0,0.1)"><div style="font-size:48px;margin-bottom:16px">🏡</div><h2>Foyer Rural d\'Isneauville</h2><p>'+message+'</p></div></body></html>');}
function installerDeclencheurFinDeSaison(){ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='verifierFinDeSaison')ScriptApp.deleteTrigger(t);});ScriptApp.newTrigger('verifierFinDeSaison').timeBased().everyDays(1).atHour(2).create();var props=PropertiesService.getScriptProperties();if(!props.getProperty('FDS_DATE_SUPPRESSION'))props.setProperty('FDS_DATE_SUPPRESSION',new Date().getFullYear()+'-06-01');Logger.log('✅ Déclencheur fin de saison installé');}
function modifierDateSuppression(nouvelleDateStr){var props=PropertiesService.getScriptProperties();props.setProperty('FDS_DATE_SUPPRESSION',nouvelleDateStr);props.setProperty('FDS_STATUT','attente');props.deleteProperty('FDS_AVERTISSEMENT_OK');envoyerEmail(EMAIL_ADMIN,'[FRI] Date de suppression modifiee','Nouvelle date : '+nouvelleDateStr,{name:NOM_ASSO});}

// ============================================================
// AUDIT SÉCURITÉ DRIVE — inchangé v8.2
// ============================================================
function auditSecuriteDrive(){Logger.log('=== AUDIT SÉCURITÉ DRIVE ===');['4-Factures acquittées','2-QS Santé Adhérents','FRI_Backup_Inscriptions'].forEach(function(nom){try{var dossiers=dossiersDriveParNom(nom);if(!dossiers.hasNext())return;var dossier=dossiers.next();var access=dossier.getSharingAccess();if(access!==DriveApp.Access.PRIVATE){securiserDossier(dossier);}var fichiers=dossier.getFiles();var nbFich=0,nbProblemes=0;while(fichiers.hasNext()){var f=fichiers.next();nbFich++;if(f.getSharingAccess()!==DriveApp.Access.PRIVATE){securiserFichier(f);nbProblemes++;}}Logger.log('  '+nom+' : '+nbFich+' fichier(s), '+nbProblemes+' corrigé(s)');}catch(e){Logger.log('  ❌ Erreur : '+e.toString());}});try{envoyerEmail(EMAIL_ADMIN,'[FRI] Audit securite Drive - '+Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy'),'Audit effectué.');}catch(e){}}
function installerAuditMensuel(){ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='auditSecuriteDrive')ScriptApp.deleteTrigger(t);});ScriptApp.newTrigger('auditSecuriteDrive').timeBased().onMonthDay(1).atHour(3).create();Logger.log('✅ Audit mensuel Drive installé');}