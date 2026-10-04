/**
 * À exécuter UNE SEULE FOIS dans l'éditeur Apps Script
 * pour archiver tous les anciens déploiements web
 * et ne garder que le plus récent actif.
 *
 * ÉTAPES :
 * 1. Coller ce code dans un NOUVEAU fichier .gs (ex: Nettoyage.gs)
 * 2. Exécuter nettoyerDeploiements()
 * 3. Vérifier que 1 seul déploiement reste actif
 * 4. Supprimer ce fichier Nettoyage.gs
 */

function nettoyerDeploiements() {
  var scriptId = ScriptApp.getScriptId();
  var token    = ScriptApp.getOAuthToken();

  Logger.log('Script ID : ' + scriptId);

  // Récupérer tous les déploiements via l'API Apps Script
  var url      = 'https://script.googleapis.com/v1/projects/' + scriptId + '/deployments';
  var response = UrlFetchApp.fetch(url, {
    method:  'GET',
    headers: { 'Authorization': 'Bearer ' + token }
  });

  var data        = JSON.parse(response.getContentText());
  var deployments = data.deployments || [];
  Logger.log('Nombre total de déploiements : ' + deployments.length);

  // Identifier les déploiements Web App (type WEB_APP)
  var webApps = deployments.filter(function(d) {
    return d.entryPoints && d.entryPoints.some(function(e) {
      return e.entryPointType === 'WEB_APP';
    });
  });

  Logger.log('Déploiements Web App trouvés : ' + webApps.length);

  // Trier par date de mise à jour décroissante — garder le plus récent
  webApps.sort(function(a, b) {
    return (b.updateTime || '').localeCompare(a.updateTime || '');
  });

  var leGardien = webApps[0];
  Logger.log('Déploiement conservé : ' + leGardien.deploymentId + ' — ' + leGardien.updateTime);

  // Afficher l'URL du déploiement conservé
  leGardien.entryPoints.forEach(function(e) {
    if (e.entryPointType === 'WEB_APP') {
      Logger.log('URL À UTILISER : ' + e.webApp.url);
    }
  });

  // Supprimer tous les autres déploiements Web App
  var supprimes = 0;
  for (var i = 1; i < webApps.length; i++) {
    var dep = webApps[i];
    try {
      var delUrl = 'https://script.googleapis.com/v1/projects/' + scriptId + '/deployments/' + dep.deploymentId;
      UrlFetchApp.fetch(delUrl, {
        method:  'DELETE',
        headers: { 'Authorization': 'Bearer ' + token }
      });
      supprimes++;
      Logger.log('Supprimé : ' + dep.deploymentId);
    } catch(e) {
      Logger.log('Impossible de supprimer ' + dep.deploymentId + ' : ' + e.toString());
    }
  }

  Logger.log('=== TERMINÉ ===');
  Logger.log('Déploiements supprimés : ' + supprimes);
  Logger.log('Il reste 1 déploiement actif.');
  Logger.log('Copiez cette URL dans votre formulaire :');
  leGardien.entryPoints.forEach(function(e) {
    if (e.entryPointType === 'WEB_APP') {
      Logger.log('>>> ' + e.webApp.url + ' <<<');
    }
  });
}