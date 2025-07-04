'use strict';

const NomOngletPresences = 'Présences';
const PremiereLigneBenevole = 4;
const LigneSemaine = 1;
const PremiereLigneDistribution = 3;
const DerniereLigneDistribution = 24;
const LigneStock = 28;
const LigneTitreAvecCompteur = 6
const PremiereColonneDefaut = 4;
const ColoneTitre = 2
const NbDemiJournee = 4;
const ColNom = 1;
const ColPadawan = 2;

// opérationel

function onEdit(e) {
  RemettreDefaut(e.range);
  let d = new Date();
  doLog('fin onEdit pour ' + e.user + ' ' + d.toLocaleString('fr-FR'));
}

function afficherTousLesOnglets() {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  for (let sheet of spreadSheet.getSheets()) {
    sheet.showSheet();
  }
}

function preparaCampagne() {
  // faire 2 appels pour la campagne d'hiver, un appel pour chaque année
  // faire 2 appels pour la campagne d'été, un appel pour chaque demi campagne
  // ATTENTION avant à contrôler la case du numéro de ligne pour le stock en bas (case G28)
  // commencer par le 2e bloc
  duplicate(19, 36, 43);  // à adapter chaque campagne
  duplicate(19, 20, 26);  // à adapter chaque campagne
}

function duplicate(sSource, sMin, sMax) {
  let oDoc = SpreadsheetApp.getActive();
  let tabSource = oDoc.getSheetByName('S' + sSource);
  for (let i = sMax; i >= sMin; i--) {
    tabSource.activate();
    let newSheet = oDoc.getSheetByName('S' + i);
    if (newSheet) oDoc.deleteSheet(newSheet);
    newSheet = oDoc.duplicateActiveSheet();
    newSheet.setName('S' + i);
    newSheet.getRange(1, 2, 1, 1).setValue(i);
  }
}

class cSemainePlanning {
  /* déclaration des propriétés non supporté par googlesheet
  static oDoc;
  static oSheetPresences;
  static nLastRowPresences;
  static rEntetePresences;
  static nFirstSemaine;
  static instances;
  // */

  static init() {
    cSemainePlanning.oDoc = SpreadsheetApp.getActiveSpreadsheet();
    cSemainePlanning.oSheetPresences = cSemainePlanning.oDoc.getSheetByName('Présences');
    cSemainePlanning.nLastRowPresences = cSemainePlanning.oSheetPresences.getLastRow();
    cSemainePlanning.rEntetePresences = cSemainePlanning.oSheetPresences.getRange(PremiereLigneBenevole, 1, 1 + cSemainePlanning.nLastRowPresences - PremiereLigneBenevole, PremiereColonneDefaut + NbDemiJournee - 1);
    cSemainePlanning.nFirstSemaine = parseInt(cSemainePlanning.oSheetPresences.getRange(LigneSemaine, PremiereColonneDefaut + NbDemiJournee, 1, 1).getValue());
  }

  static getInstance(nSemaine) {
    if (!cSemainePlanning.oDoc) cSemainePlanning.init()
    nSemaine = parseInt(nSemaine);
    if (cSemainePlanning.instances) {
      if (cSemainePlanning.instances[nSemaine]) return cSemainePlanning.instances[nSemaine];
    } else {
        cSemainePlanning.instances = [];
    }
    let oInstance = new cSemainePlanning(nSemaine);
    cSemainePlanning.instances[nSemaine] = oInstance;
    return oInstance;
  }

  static loadAllInstances() {
    if (!cSemainePlanning.oDoc) cSemainePlanning.init()
    for (let oSheet of cSemainePlanning.oDoc.getSheets()) {
      let iSemaine = parseInt(oSheet.getName().substring(1));
      if (isNaN(iSemaine)) continue;
      cSemainePlanning.getInstance(iSemaine);
    }
  }

  // fin zone statique

  /* déclaration des propriétés non supporté par googlesheet
  nSemaine;
  rSemainePresences;
  //  */

  constructor(nSemaine) {
    this.nSemaine = nSemaine;
    //this.rSemainePresences = cSemainePlanning.oSheetPresences.getRange(PremiereLigneBenevole, PremiereColonneDefaut + (nSemaine - cSemainePlanning.nFirstSemaine) * NbDemiJournee, 1 + cSemainePlanning.nLastRowPresences - PremiereLigneBenevole, NbDemiJournee);
  }

  /* abandonné
  doPlanning(postesParSemaine, affichageBenevolesAbsents, code, compteur, formation, numeroPoste) {
    // liste des bénévoles
    // liste des bénévoles absents
    // affichage
    return `nSemaine=${this.nSemaine}, firstSemaine=${cSemainePlanning.nFirstSemaine}, col=${PremiereColonneDefaut + (this.nSemaine - cSemainePlanning.nFirstSemaine) * NbDemiJournee}`;
    return 'S' + this.nSemaine + ' (' + postesParSemaine + ') ' + (affichageBenevolesAbsents ? 'avec' : 'sans') + ' affichage absents';
  }
  // */
}

// test/debuging

function duplicateFormulaToAll() {
  // prendre la formule en C3 de la première semaine et la mettre partout
  // et aussi mettre la plage des présences en G2 de chaque onglet de semaine
  // et aussi copier la formule en B6 (titre avec compteur)

  cSemainePlanning.loadAllInstances();
  // trouver la première semaine
  let firstSemaine;
  for (let oSemaine of cSemainePlanning.instances) {
      if (!oSemaine) continue;
      doLog(`vu semaine ${oSemaine.nSemaine}`);
      if (firstSemaine === undefined || firstSemaine > oSemaine.nSemaine) firstSemaine = oSemaine.nSemaine;
  }
  doLog(`firstSemaine=${firstSemaine}`, true);
  let formula = cSemainePlanning.oDoc.getSheetByName('S' + firstSemaine).getRange(3, PremiereLigneDistribution, 1, 1).getFormulaR1C1();
  let formulaB6 = cSemainePlanning.oDoc.getSheetByName('S' + firstSemaine).getRange(LigneTitreAvecCompteur, ColoneTitre, 1, 1).getFormulaR1C1();
  doLog(formula.toString(), true);
  for (let oSemaine of cSemainePlanning.instances) {
      if (!oSemaine) continue;
      let oSheet = cSemainePlanning.oDoc.getSheetByName('S' + oSemaine.nSemaine);
      // mettre la formule dans la zone distribution
      oSheet.getRange(PremiereLigneDistribution, 3, 1 + DerniereLigneDistribution - PremiereLigneDistribution, 2).setFormulaR1C1(formula);
      // mettre la formule en B6
      oSheet.getRange(LigneTitreAvecCompteur, ColoneTitre, 1, 1).setFormulaR1C1(formulaB6);
      // mettre la formule dans la zone stock
      oSheet.getRange(LigneStock, 3, 1, 2).setFormulaR1C1(formula);
      // mettre le range special en G2
      let firstColPosteCetteSemaine = PremiereColonneDefaut + (1 + oSemaine.nSemaine - cSemainePlanning.nFirstSemaine) * NbDemiJournee;
      //oSheet.getRange(2, 7, 1, 1).setValue(null);
      oSheet.getRange(2, 7, 1, 1).setValue(`="'Présences'!$${col2Lettre(firstColPosteCetteSemaine)}:$${col2Lettre(firstColPosteCetteSemaine + NbDemiJournee - 1)}"`);
      doLog(`fait semaine ${oSemaine.nSemaine}`, true);
  }
}

function doTest() {
  // tests
  //let rDebug = sTest.getRange(1, 1, 1, 1);
  /*
  //SpreadsheetApp.getUi().alert(rDebug);
  let d = new Date();
  doLog(d.toLocaleString('fr-FR'));
  // */
  /*
  let sPresences = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NomOngletPresences);
  //rModified = sPresences.getRange(PremiereLigneBenevole, PremiereColonneDefaut + NbDemiJournee, 3, NbDemiJournee);
  //rModified = sPresences.getRange(PremiereLigneBenevole + 1, PremiereColonneDefaut + NbDemiJournee * 2, 1, 4);
  //rModified = sPresences.getRange(116, PremiereColonneDefaut + NbDemiJournee * 1, 10, 4);
  rModified = sPresences.getRange(PremiereLigneBenevole, PremiereColonneDefaut + NbDemiJournee * 1, 155 - 3, ((52+1-46) + 10)*NbDemiJournee);  
  //RemettreDefaut(rModified);
  RemettreFormule(rModified);
  // */
  /* *
  for (let i = 19; i <= 26; i++) patchV(i);
  for (let i = 36; i <= 43; i++) patchV(i);
  patchV(19);
  // */
}

function patchV(i) {
  //SpreadsheetApp.getUi().alert('i=' + i);
  let oDoc = SpreadsheetApp.getActive();
  let oSheet = oDoc.getSheetByName('S' + i);
  let r = oSheet.getRange(28, 7, 1, 1);
  let v = r.getValue();
  if (v != 27) {
    SpreadsheetApp.getUi().alert('v=' + v);
    return;
  }
  r.setValue(26);
}

function RemettreFormule(rModified) {
  let tabFormula = rModified.getFormulas();
  let tabValues = rModified.getValues();
  let sheetModified = rModified.getSheet()
  let tabValueDefault = sheetModified.getRange(PremiereLigneBenevole, PremiereColonneDefaut, 155-3, NbDemiJournee).getValues();
  nbCol = rModified.getNumColumns();
  let premiereColModif = rModified.getColumn();
  let premierRowModif = rModified.getRow();
  let nbRow = rModified.getNumRows();
  for (let iCol = 0; iCol < nbCol; iCol++) {
    let baseCol = PremiereColonneDefaut + (iCol % NbDemiJournee);
    for (let iRow = 0; iRow < nbRow; iRow++) {
      let thisFormula = tabFormula[iRow][iCol];
      let value = tabValues[iRow][iCol];
      let shouldBe = "=$" + col2Lettre(baseCol) + (iRow + premierRowModif);
      if (!thisFormula && (value == '')) {
        //doLog(col2Lettre(iCol + premiereColModif) + (iRow + premierRowModif) + " value blanche " + thisFormula + "->" + shouldBe, true);
        //if (tabValueDefault[iRow][iCol % NbDemiJournee])
          doLog(col2Lettre(iCol + premiereColModif) + (iRow + premierRowModif) + " value blanche=>" + tabValueDefault[iRow][iCol % NbDemiJournee] + ', formule ' + thisFormula + "->" + shouldBe, true);
        let r = sheetModified.getRange(iRow + premierRowModif, iCol + premiereColModif, 1, 1);
        //r.setFormula(shouldBe);
      }
      let badDollar = "=" + col2Lettre(baseCol) + (iRow + premierRowModif); 
      if (thisFormula == shouldBe) continue;
      if (badDollar == thisFormula) {
        doLog(col2Lettre(iCol + premiereColModif) + (iRow + premierRowModif) + " formule " + thisFormula + "==" + badDollar, true);
        let r = sheetModified.getRange(iRow + premierRowModif, iCol + premiereColModif, 1, 1);
        //r.setFormula(shouldBe);
      }
      if (value == tabValueDefault[iRow][iCol % NbDemiJournee])
      {
        doLog(col2Lettre(iCol + premiereColModif) + (iRow + premierRowModif) + " value " + value + '=>' + tabValueDefault[iRow][iCol % NbDemiJournee] + ', formule ' + thisFormula + "->" + shouldBe, true);
        let r = sheetModified.getRange(iRow + premierRowModif, iCol + premiereColModif, 1, 1);
        //r.setFormula(shouldBe);
      }
      //continue;
      doLog(col2Lettre(iCol + premiereColModif) + (iRow + premierRowModif) + " value " + value + '=>' + tabValueDefault[iRow][iCol % NbDemiJournee] + ', formule ' + thisFormula + "->" + shouldBe, true);
    }
  }
}

function RemettreDefaut(rModified) {
  //doLog(JSON.stringify(rModified));
  let premiereColModif = rModified.getColumn();
  let premierRowModif = rModified.getRow();
  let nbCol = rModified.getNumColumns();
  let nbRow = rModified.getNumRows();
  let sheetModified = rModified.getSheet()
  if (sheetModified.getName() != NomOngletPresences) {
    //doLog('Pas le bon onglet=' + sheetModified.getName());
    return;
  }
  doLog("c'est le bon onglet=" + sheetModified.getName() + ', row=' + premierRowModif + ', col=' + col2Lettre(premiereColModif) + ', nbRow=' + nbRow + ', nbCol=' + nbCol);
  let tabValues = rModified.getValues();
  let tabValidations = rModified.getDataValidations();
  let iRowMax = null;
  for (let iCol = 0; iCol < nbCol; iCol++) {
    let colModif = premiereColModif + iCol;
    if (colModif < PremiereColonneDefaut) {
      doLog('col modif=' + colModif + '<' + PremiereColonneDefaut);
      continue;
    }
    let rModeleValidator = null;
    let colHabituelle = PremiereColonneDefaut + ((colModif - PremiereColonneDefaut) % NbDemiJournee);
    for (let iRow = 0; iRow < nbRow; iRow++) {
      let rowModif = premierRowModif + iRow;
      //doLog('col modif=' + colModif + ', row modif=' + rowModif);
      if (rowModif < PremiereLigneBenevole) {
        doLog('row modif=' + rowModif + '<' + PremiereLigneBenevole);
        continue;
      }
      let value = tabValues[iRow][iCol];
      //doLog('Value=' + value);
      // vérification Validation
      let rThis = undefined;
      let thisValidation = tabValidations[iRow][iCol];
      let copyDone = false;
      if (!thisValidation) {
        // on a perdu la liste de postes pour cette cellule
        if (!rModeleValidator) {
          if (!iRowMax) iRowMax = sheetModified.getLastRow();
          // trouver une cellule dans la même colonne avec la règle de validation
          for (let iRow2 = PremiereLigneBenevole; iRow2 <= iRowMax; iRow2++) {
            let r = sheetModified.getRange(iRow2, colModif, 1, 1);
            let dataValidator = r.getDataValidation();
            if (dataValidator) {
              doLog('la cellule ' + col2Lettre(colModif) + iRow2 + ' est choisie comme modèle de liste')
              rModeleValidator = r;
              break;
            }
          }
        }
        if (rModeleValidator) {
          if (!rThis) rThis = sheetModified.getRange(rowModif, colModif, 1, 1);
          doLog('Copy (' + rModeleValidator.getRow() + ', ' + rModeleValidator.getColumn() + ') vers (' + rThis.getRow() + ', ' + rThis.getColumn() + ')');
          rModeleValidator.copyTo(rThis);
          copyDone = true;
        }
      }
      else
      {
        //doLog('Validation OK');
      }
      let bDefault = false;
      if (colModif >= PremiereColonneDefaut + NbDemiJournee) {
        bDefault = (value == '');
        if (!bDefault) {
          let rDefaut = sheetModified.getRange(rowModif, colHabituelle, 1, 1);
          bDefault = (value == rDefaut.getValue());
        }
      }
      if (bDefault) {
        doLog('reset à défaut col=' + colModif + ', row=' + rowModif);
        if (!rThis) rThis = sheetModified.getRange(rowModif, colModif, 1, 1);
        rThis.setFormula('=$' + col2Lettre(colHabituelle) + rowModif);
      } else if (copyDone) {
        doLog('reset val= [' + value + '], col=' + colModif + ', row=' + rowModif);
        if (!rThis) rThis = sheetModified.getRange(rowModif, colModif, 1, 1);
        rThis.setValue(value);
      } else {
        doLog('rien à faire, col=' + colModif + ', row=' + rowModif + ', value=' + value);
      }
    }
  }
}

function col2Lettre(iCol) {
  if (iCol <= 26) return String.fromCharCode(64 + iCol);
  return String.fromCharCode(64 + Math.floor((iCol - 1) / 26)) + String.fromCharCode(65 + ((iCol - 1) % 26));
}

var gNextRowDebug = 1;
var gSheetDebug;
function doLog(msg, bForce) {
  if (!bForce) return; // désactivation quand tout tourne bien, fin de debugging
  if (!gSheetDebug) {
      gSheetDebug = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Debug');
      if (!gSheetDebug) {
        SpreadsheetApp.getUi().alert("pas d'onglet Debug");
        // forcer une erreur pour terminer le script dans tous les cas
        foo.bar = 1;
      }
      gSheetDebug.getRange('A:A').setValue('');
  }
  let rDebug = gSheetDebug.getRange(gNextRowDebug++, 1, 1, 1);
  rDebug.setValue(msg);  
}
